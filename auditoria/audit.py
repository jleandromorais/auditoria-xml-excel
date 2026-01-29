import os
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from .excel_loader import carregar_excel
from .report import gerar_relatorio, gerar_relatorio_avisos
from .utils import safe_float
from .xml_parser import parse_xml_file

# <--- DB: Importa a classe de banco de dados
from .database import AuditDB 

# 1. CONFIGURAÇÃO (As Regras do Jogo)
# Esta classe serve apenas para guardar números.
# Define o quanto "deixamos passar" de diferença antes de considerar um erro.
@dataclass
class AuditConfig:
    tolerancia_cte: float = 50.0   # Aceita diferença de até 50.00 para CT-e (fretes)
    tolerancia_nfe: float = 5.0    # Aceita diferença de até 5.00 para Notas Fiscais (NF-e)
    tolerancia_volume: float = 1.0 # Aceita diferença de até 1 unidade no volume (m3, kg)

# 2. O COLECTOR (O "Estafeta")
# Esta função entra nas pastas das empresas e faz uma lista de todos os XMLs encontrados.
def coletar_xmls_por_empresas(pasta_pai: Path, empresas: List[Path]) -> List[Tuple[str, str]]:
    out: List[Tuple[str, str]] = [] # Lista vazia que vai guardar os resultados
    
    # Para cada empresa selecionada na interface...
    for empresa_dir in empresas:
        vistos = set() # Cria um conjunto para evitar ficheiros duplicados
        
        # Procura tanto por .xml minúsculo como .XML maiúsculo
        for extensao in ["*.xml", "*.XML"]:
            # .rglob() procura em TODAS as subpastas (recursivo)
            for arq in empresa_dir.rglob(extensao):
                p = str(arq)
                # Se já vimos este ficheiro, ignoramos para não duplicar
                if p.lower() in vistos: continue
                
                vistos.add(p.lower())
                # Adiciona à lista final: (Nome da Empresa, Caminho do Arquivo)
                out.append((empresa_dir.name, p))
    
    # Organiza a lista por nome da empresa e depois pelo nome do ficheiro
    out.sort(key=lambda t: (t[0].lower(), os.path.basename(t[1]).lower()))
    return out

# 3. O MAESTRO (A Função Principal)
# Esta é a função que controla tudo. Ela recebe os caminhos e decide o que fazer.
def auditar_pasta_pai(
    pasta_pai: Path,           # Onde estão as pastas das empresas
    empresas: List[Path],      # Quais empresas vamos auditar
    excel_path: str,           # Onde está a folha de Excel (Conta Gráfica)
    saida: Optional[str] = None,       # Onde vamos guardar o relatório final
    config: Optional[AuditConfig] = None, # As regras de tolerância (opcional)
    mes_filtro: Optional[str] = None,  # Se o utilizador quer filtrar só um mês (ex: "OUT")
) -> str:
    # Se ninguém passar configurações, usamos as padrão definidas lá em cima
    if config is None:
        config = AuditConfig()
    # <--- DB: Inicializa o banco de dados
    print("Inicializando banco de dados DuckDB...")
    db = AuditDB()
    db.inicializar()

    # 1. Carrega Excel Bruto
    # Chama a função especialista em ler o Excel (que lida com as abas e cabeçalhos).
    df_base = carregar_excel(excel_path)
    
    # Se o Excel estiver vazio ou ilegível, pára tudo e avisa o erro.
    if df_base.empty:
        raise RuntimeError("Não foi possível carregar os dados do Excel.")

    # 2. Aplica Filtro de Mês
    # Se o utilizador pediu um mês específico (ex: "OUT"), filtramos aqui.
    if mes_filtro and str(mes_filtro).strip():
        mes_busca = str(mes_filtro).upper().strip()
        # Mantém apenas as linhas onde a coluna "Mes" contém o texto procurado.
        df_base = df_base[df_base["Mes"].astype(str).str.upper().str.contains(mes_busca, na=False)].copy()
        
        # Se depois de filtrar não sobrar nada, avisa que aquele mês não existe no Excel.
        if df_base.empty:
            raise RuntimeError(f"Atenção: Não existem notas para o mês '{mes_filtro}' no Excel.")

    # Limpeza básica: Garante que o número da nota não tem espaços extras.
    df_base["NF_Clean"] = df_base["NF_Clean"].astype(str).str.strip()

    # <--- DB: Salva uma cópia de segurança no banco de dados para consultas futuras.
    db.salvar_excel(df_base)

    # ============================================================
    # 2.5 CAPTURA DE DUPLICATAS (O Detetive)
    # ============================================================
    # IMPORTANTE: Fazemos isto ANTES de agrupar.
    # O código procura se o mesmo número de nota aparece mais de uma vez.
    # Se aparecer, guarda numa lista separada para avisar no relatório de "AVISOS".
    # (Ex: Nota lançada duas vezes por engano no financeiro).
    df_duplicadas = df_base[df_base.duplicated(subset="NF_Clean", keep=False)].copy()

    # ============================================================
    # 3. AGRUPAMENTO (A "Magia" da Soma)
    # ============================================================
    cols_numericas = ["Vol_Excel", "Liq_Excel", "ICMS_Excel", "PIS_Excel", "COFINS_Excel"]
    
    # Converte tudo para número. Se houver texto ("R$ 10,00"), força a virar número para não dar erro na soma.
    for c in cols_numericas:
        if c in df_base.columns:
            df_base[c] = pd.to_numeric(df_base[c], errors='coerce').fillna(0.0)

    # Agrupa por Nota Fiscal (NF_Clean)
    # Se a nota 123 tem 3 linhas (3 produtos), aqui elas viram 1 linha só.
    # E o que fazemos com os valores? Somamos ("sum").
    agg_dict = {
        "Mes": "first",          # Pega o primeiro mês que encontrar (são iguais)
        "Vol_Excel": "sum",      # Soma o volume
        "Liq_Excel": "sum",      # Soma o valor líquido
        "ICMS_Excel": "sum",     # Soma o ICMS
        "PIS_Excel": "sum",      # Soma o PIS
        "COFINS_Excel": "sum",   # Soma o COFINS
    }
    
    # Se tiver nome da empresa, mantém o primeiro que achar (não dá para somar nomes).
    if "Empresa" in df_base.columns:
        agg_dict["Empresa"] = "first"

    # Cria o "df_agrupado": A tabela final, limpa e somada, pronta para a batalha contra o XML.
    df_agrupado = df_base.groupby("NF_Clean", as_index=False).agg(agg_dict)

   # ============================================================
    # 4. Leitura e Soma dos XMLs
    # ============================================================
    
    # Chama o "Estafeta" (função que explicámos antes) para nos dar a lista de todos os ficheiros.
    xmls_arquivos = coletar_xmls_por_empresas(pasta_pai, empresas)
    
    # Cria um Dicionário vazio.
    # Podes imaginar isto como um armário com gavetas. Cada gaveta terá o número da Nota Fiscal na etiqueta.
    xmls_agrupados: Dict[str, Dict] = {} 
    
    # <--- DB: Uma lista simples para guardar o histórico de tudo o que lemos.
    lista_dados_xml_brutos = []

    # O GRANDE CICLO (LOOP)
    # Para cada ficheiro encontrado na lista...
    for empresa_nome, xml_path in xmls_arquivos:
        try:
            # Tenta ler o XML usando a função 'parse_xml_file' (o "Tradutor").
            # Ela devolve um dicionário com: { "Nota": "123", "Valor": 100.0, ... }
            info = parse_xml_file(xml_path)
        except:
            # Se o ficheiro estiver estragado ou não for um XML válido, ignora e passa ao próximo.
            # O "continue" diz: "Esquece este, vai para o seguinte".
            continue

        # Se o tradutor não devolveu nada (None) ou se o XML não tem número de Nota, ignora.
        if not info or not info.get("Nota"): continue
        
        # <--- DB: Antes de somar, guardamos os dados originais (nome do ficheiro, caminho)
        # para salvar no Banco de Dados depois. É o nosso rasto de auditoria.
        info['Arquivo'] = os.path.basename(xml_path)
        info['CaminhoCompleto'] = str(xml_path)
        info['Empresa'] = empresa_nome
        lista_dados_xml_brutos.append(info) 

        # Normaliza o número da nota (remove espaços em branco)
        nota = str(info["Nota"]).strip()
        
        # A LÓGICA DE AGRUPAMENTO (O Armário)
        # Se esta nota ainda não tem uma gaveta no nosso armário 'xmls_agrupados'...
        if nota not in xmls_agrupados:
            # ... criamos uma gaveta nova com tudo a zero.
            xmls_agrupados[nota] = {
                "Empresa": empresa_nome, 
                "Tipo": info["Tipo"], # Se é NF-e ou CT-e
                "Arquivos": [os.path.basename(xml_path)], # Lista de ficheiros desta nota
                "Vol": 0.0, "Bruto": 0.0, "ICMS": 0.0, "PIS": 0.0, "COFINS": 0.0,
            }
        
        # Agora SOMAMOS os valores do ficheiro atual aos valores que já estão na gaveta.
        # Mesmo que seja o primeiro ficheiro (e a gaveta estivesse a zero), ele soma aqui.
        xmls_agrupados[nota]["Vol"] += info.get("Vol", 0.0)
        xmls_agrupados[nota]["Bruto"] += info.get("Bruto", 0.0)
        xmls_agrupados[nota]["ICMS"] += info.get("ICMS", 0.0)
        xmls_agrupados[nota]["PIS"] += info.get("PIS", 0.0)
        xmls_agrupados[nota]["COFINS"] += info.get("COFINS", 0.0)
        
        # Guarda o nome do ficheiro na lista (para sabermos que ficheiros compõem esta nota)
        nome_arq = os.path.basename(xml_path)
        if nome_arq not in xmls_agrupados[nota]["Arquivos"]:
            xmls_agrupados[nota]["Arquivos"].append(nome_arq)

    # <--- DB: Finalmente, despeja toda a lista de dados brutos no Banco de Dados de uma só vez.
    db.salvar_xmls(lista_dados_xml_brutos)
    # ============================================================
    # 5. Comparação Final (Excel Agrupado vs XML Agrupado)
    # ============================================================
    relatorio: List[Dict] = []       # A lista final que vai para o Excel de saída
    notas_sem_xml: List[Dict] = []   # Lista secundária para o relatório de avisos
    notas_xml_vistas = set()         # Um conjunto para marcar os XMLs que já usámos

    # PARTE 1: Olhamos para o Excel e procuramos o XML correspondente
    for _, row in df_agrupado.iterrows():
        nota_ex = str(row["NF_Clean"]).strip()
        # Se a nota estiver vazia ou inválida, ignora.
        if not nota_ex or nota_ex.upper() == "NAN": continue

        # Pega os valores do Excel com segurança (para não dar erro se for nulo)
        vol_ex = safe_float(row.get("Vol_Excel", 0))
        liq_ex = safe_float(row.get("Liq_Excel", 0))

        # Cria a ficha base da auditoria com os dados do Excel
        item: Dict = {
            "Nota": nota_ex,
            "Mes": row.get("Mes", "-"),
            "Vol Excel": vol_ex,
            "Liq Excel": liq_ex,
            "ICMS Excel": safe_float(row.get("ICMS_Excel", 0)),
            "PIS Excel": safe_float(row.get("PIS_Excel", 0)),
            "COFINS Excel": safe_float(row.get("COFINS_Excel", 0)),
            "Empresa": row.get("Empresa", "-"),
            "Status": "PENDENTE", # Começa como pendente até validarmos
             # ... outros campos vazios ...
        }

        # A PERGUNTA DE OURO: "Esta nota do Excel existe nos XMLs que li?"
        if nota_ex in xmls_agrupados:
            xml = xmls_agrupados[nota_ex]
            notas_xml_vistas.add(nota_ex) # Marca: "Já vi este XML!"

            # Preenche os dados vindos do XML na ficha
            item.update({
                "Tipo": xml["Tipo"],
                "Arquivo": ", ".join(xml["Arquivos"]), # Lista os nomes dos ficheiros
                "Vol XML": xml["Vol"],
                "Bruto XML": xml["Bruto"],
                "ICMS XML": xml["ICMS"],
            })
            
            # Lógica especial para Fretes (CT-e):
            # Às vezes o XML do frete não traz o imposto discriminado, mas o Excel traz.
            # Se for esse o caso, "copiamos" o imposto do Excel para não dar erro falso.
            p_xml, c_xml = xml["PIS"], xml["COFINS"]
            if xml["Tipo"] == "CT-e" and p_xml == 0 and item["PIS Excel"] != 0:
                p_xml, c_xml = item["PIS Excel"], item["COFINS Excel"]
                item["Obs"] = "CT-e: Usado impostos do Excel."
            
            # CÁLCULO DO LÍQUIDO DO XML
            # Líquido = Bruto - Impostos
            bruto = xml["Bruto"]
            liq_calc = bruto - sum(v for v in (xml["ICMS"], p_xml, c_xml) if 0 < v < bruto)
            item["Liq XML (Calc)"] = max(liq_calc, 0.0)
            
            # Calcula as diferenças (Diferença = XML - Excel)
            item["Diff Vol"] = "-" if vol_ex == 0 else (xml["Vol"] - vol_ex)
            item["Diff R$"] = item["Liq XML (Calc)"] - liq_ex

            # VERIFICA SE ESTÁ DENTRO DA TOLERÂNCIA (Aquela config do início)
            tol = config.tolerancia_cte if xml["Tipo"] == "CT-e" else config.tolerancia_nfe
            
            # Se o volume no Excel for 0, ignoramos erro de volume. Senão, testamos.
            v_ok = True if vol_ex == 0 else abs(float(item["Diff Vol"])) < config.tolerancia_volume
            # Testamos o valor financeiro
            f_ok = abs(item["Diff R$"]) < tol

            # Veredicto Final
            if v_ok and f_ok: item["Status"] = "OK ✅"
            else:
                errs = []
                if not v_ok: errs.append("VOL")   # Erro de Volume
                if not f_ok: errs.append("VALOR") # Erro de Valor
                item["Status"] = f"ERRO {'+'.join(errs)} ❌"
        else:
            # Se não encontrou no dicionário de XMLs
            item["Status"] = "SEM XML ❌"
            item["Diff R$"] = 0.0 - liq_ex
            notas_sem_xml.append(item.copy()) # Guarda para o relatório de avisos

        relatorio.append(item)

    # PARTE 2: XMLs SOBRANTES ("Sem Excel")
    # Percorremos todos os XMLs. Se algum NÃO foi marcado como "visto" no passo anterior,
    # significa que ele existe fisicamente mas não está na planilha do Excel.
    for nt, dados in xmls_agrupados.items():
        if nt not in notas_xml_vistas:
            relatorio.append({
                "Nota": nt, "Status": "SEM EXCEL ❌", 
                "Vol XML": dados["Vol"], 
                # ... preenche o resto ...
            })

    # <--- DB: Salva o resultado final no Banco de Dados
    if relatorio:
        db.salvar_relatorio_final(pd.DataFrame(relatorio))
    
    db.fechar() # Fecha a conexão com o banco

    # --- GERAÇÃO DOS ARQUIVOS FINAIS ---
    
    # Gera o Excel bonito principal
    caminho_resultado = gerar_relatorio(relatorio, saida=saida)
    
    # Gera o Excel de "AVISOS" (separado para não poluir o principal)
    # Só gera se houver duplicatas ou notas sem XML
    caminho_avisos = ""
    if not df_duplicadas.empty or notas_sem_xml:
        caminho_avisos = gerar_relatorio_avisos(df_duplicadas, notas_sem_xml, caminho_resultado)
        
        # Tenta abrir o arquivo de avisos automaticamente na tela do utilizador
        try:
            os.startfile(caminho_avisos)
        except:
            pass # Se não conseguir abrir (ex: não é Windows), não faz mal

    return f"{caminho_resultado}..."