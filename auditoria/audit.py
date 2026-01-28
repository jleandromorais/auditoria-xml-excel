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

@dataclass
class AuditConfig:
    tolerancia_cte: float = 50.0
    tolerancia_nfe: float = 5.0
    tolerancia_volume: float = 1.0

def coletar_xmls_por_empresas(pasta_pai: Path, empresas: List[Path]) -> List[Tuple[str, str]]:
    out: List[Tuple[str, str]] = []
    for empresa_dir in empresas:
        vistos = set()
        for extensao in ["*.xml", "*.XML"]:
            for arq in empresa_dir.rglob(extensao):
                p = str(arq)
                if p.lower() in vistos: continue
                vistos.add(p.lower())
                out.append((empresa_dir.name, p))
    out.sort(key=lambda t: (t[0].lower(), os.path.basename(t[1]).lower()))
    return out

def auditar_pasta_pai(
    pasta_pai: Path,
    empresas: List[Path],
    excel_path: str,
    saida: Optional[str] = None,
    config: Optional[AuditConfig] = None,
    mes_filtro: Optional[str] = None,
) -> str:
    if config is None:
        config = AuditConfig()

    # <--- DB: Inicializa o banco de dados
    print("Inicializando banco de dados DuckDB...")
    db = AuditDB()
    db.inicializar()

    # 1. Carrega Excel Bruto
    df_base = carregar_excel(excel_path)
    if df_base.empty:
        raise RuntimeError("Não foi possível carregar os dados do Excel.")

    # 2. Aplica Filtro de Mês
    if mes_filtro and str(mes_filtro).strip():
        mes_busca = str(mes_filtro).upper().strip()
        df_base = df_base[df_base["Mes"].astype(str).str.upper().str.contains(mes_busca, na=False)].copy()
        
        if df_base.empty:
            raise RuntimeError(f"Atenção: Não existem notas para o mês '{mes_filtro}' no Excel.")

    df_base["NF_Clean"] = df_base["NF_Clean"].astype(str).str.strip()

    # <--- DB: Salva os dados do Excel filtrados no banco
    db.salvar_excel(df_base)

    # ============================================================
    # 2.5 CAPTURA DE DUPLICATAS (PARA O RELATÓRIO DE AVISO)
    # ============================================================
    df_duplicadas = df_base[df_base.duplicated(subset="NF_Clean", keep=False)].copy()

    # ============================================================
    # 3. AGRUPAMENTO (CORREÇÃO PARA O RELATÓRIO PRINCIPAL)
    # ============================================================
    cols_numericas = ["Vol_Excel", "Liq_Excel", "ICMS_Excel", "PIS_Excel", "COFINS_Excel"]
    for c in cols_numericas:
        if c in df_base.columns:
            df_base[c] = pd.to_numeric(df_base[c], errors='coerce').fillna(0.0)

    # Agrupa por Nota Fiscal para a auditoria funcionar corretamente
    df_agrupado = df_base.groupby("NF_Clean", as_index=False).agg({
        "Mes": "first",
        "Vol_Excel": "sum",
        "Liq_Excel": "sum",
        "ICMS_Excel": "sum",
        "PIS_Excel": "sum",
        "COFINS_Excel": "sum"
    })

    # ============================================================
    # 4. Leitura e Soma dos XMLs
    # ============================================================
    xmls_arquivos = coletar_xmls_por_empresas(pasta_pai, empresas)
    xmls_agrupados: Dict[str, Dict] = {} 
    
    # <--- DB: Lista para armazenar todos os dados brutos dos XMLs para o banco
    lista_dados_xml_brutos = []

    for empresa_nome, xml_path in xmls_arquivos:
        try:
            info = parse_xml_file(xml_path)
        except:
            continue

        if not info or not info.get("Nota"): continue
        
        # <--- DB: Adiciona info de arquivo e empresa para salvar no banco
        info['Arquivo'] = os.path.basename(xml_path)
        info['CaminhoCompleto'] = str(xml_path)
        info['Empresa'] = empresa_nome
        lista_dados_xml_brutos.append(info) # Guarda na lista

        nota = str(info["Nota"]).strip()
        
        if nota not in xmls_agrupados:
            xmls_agrupados[nota] = {
                "Empresa": empresa_nome, "Tipo": info["Tipo"],
                "Arquivos": [os.path.basename(xml_path)],
                "Vol": 0.0, "Bruto": 0.0, "ICMS": 0.0, "PIS": 0.0, "COFINS": 0.0,
            }
        
        xmls_agrupados[nota]["Vol"] += info.get("Vol", 0.0)
        xmls_agrupados[nota]["Bruto"] += info.get("Bruto", 0.0)
        xmls_agrupados[nota]["ICMS"] += info.get("ICMS", 0.0)
        xmls_agrupados[nota]["PIS"] += info.get("PIS", 0.0)
        xmls_agrupados[nota]["COFINS"] += info.get("COFINS", 0.0)
        
        nome_arq = os.path.basename(xml_path)
        if nome_arq not in xmls_agrupados[nota]["Arquivos"]:
            xmls_agrupados[nota]["Arquivos"].append(nome_arq)

    # <--- DB: Salva todos os XMLs processados no banco de uma vez
    db.salvar_xmls(lista_dados_xml_brutos)

    # ============================================================
    # 5. Comparação Final (Excel Agrupado vs XML Agrupado)
    # ============================================================
    relatorio: List[Dict] = []
    notas_sem_xml: List[Dict] = [] 
    notas_xml_vistas = set()

    for _, row in df_agrupado.iterrows():
        nota_ex = str(row["NF_Clean"]).strip()
        if not nota_ex or nota_ex.upper() == "NAN": continue

        vol_ex = safe_float(row.get("Vol_Excel", 0))
        liq_ex = safe_float(row.get("Liq_Excel", 0))

        item: Dict = {
            "Nota": nota_ex, "Mes": row.get("Mes", "-"),
            "Vol Excel": vol_ex, "Liq Excel": liq_ex,
            "ICMS Excel": safe_float(row.get("ICMS_Excel", 0)),
            "PIS Excel": safe_float(row.get("PIS_Excel", 0)),
            "COFINS Excel": safe_float(row.get("COFINS_Excel", 0)),
            "Status": "PENDENTE", "Obs": "", "Empresa": "-", "Tipo": "-", "Arquivo": "-"
        }

        if nota_ex in xmls_agrupados:
            xml = xmls_agrupados[nota_ex]
            notas_xml_vistas.add(nota_ex)

            item.update({
                "Empresa": xml["Empresa"], "Tipo": xml["Tipo"],
                "Arquivo": ", ".join(xml["Arquivos"]),
                "Vol XML": xml["Vol"], "Bruto XML": xml["Bruto"], "ICMS XML": xml["ICMS"]
            })
            
            p_xml, c_xml = xml["PIS"], xml["COFINS"]
            if xml["Tipo"] == "CT-e" and p_xml == 0 and item["PIS Excel"] != 0:
                p_xml, c_xml = item["PIS Excel"], item["COFINS Excel"]
                item["Obs"] = "CT-e: Usado impostos do Excel."
            
            item["PIS"], item["COFINS"] = p_xml, c_xml
            
            bruto = xml["Bruto"]
            liq_calc = bruto - sum(v for v in (xml["ICMS"], p_xml, c_xml) if 0 < v < bruto)
            item["Liq XML (Calc)"] = max(liq_calc, 0.0)
            
            item["Diff Vol"] = "-" if vol_ex == 0 else (xml["Vol"] - vol_ex)
            item["Diff R$"] = item["Liq XML (Calc)"] - liq_ex

            tol = config.tolerancia_cte if xml["Tipo"] == "CT-e" else config.tolerancia_nfe
            v_ok = True if vol_ex == 0 else abs(float(item["Diff Vol"])) < config.tolerancia_volume
            f_ok = abs(item["Diff R$"]) < tol

            if v_ok and f_ok: item["Status"] = "OK ✅"
            else:
                errs = []
                if not v_ok: errs.append("VOL")
                if not f_ok: errs.append("VALOR")
                item["Status"] = f"ERRO {'+'.join(errs)} ❌"
        else:
            item["Status"] = "SEM XML ❌"
            item["Diff R$"] = 0.0 - liq_ex
            item["Diff Vol"] = "-"
            notas_sem_xml.append(item.copy())

        relatorio.append(item)

    # XMLs sobrantes
    for nt, dados in xmls_agrupados.items():
        if nt not in notas_xml_vistas:
            relatorio.append({
                "Nota": nt, "Status": "SEM EXCEL ❌", "Empresa": dados["Empresa"],
                "Tipo": dados["Tipo"], "Vol XML": dados["Vol"], "Bruto XML": dados["Bruto"],
                "Liq XML (Calc)": dados["Bruto"], "Arquivo": ", ".join(dados["Arquivos"]),
                "Mes": "-", "Liq Excel": 0, "Vol Excel": 0
            })

    # <--- DB: Salva o relatório final no DuckDB para BI
    if relatorio:
        db.salvar_relatorio_final(pd.DataFrame(relatorio))
    
    db.fechar()
    # <--- Fim DB

    # --- GERAÇÃO DOS ARQUIVOS ---
    
    # 1. Relatório Principal (Resultado da Auditoria)
    caminho_resultado = gerar_relatorio(relatorio, saida=saida)
    
    # 2. Relatório de Avisos (Duplicatas e Sem XML)
    caminho_avisos = ""
    if not df_duplicadas.empty or notas_sem_xml:
        caminho_avisos = gerar_relatorio_avisos(df_duplicadas, notas_sem_xml, caminho_resultado)
        
        try:
            os.startfile(caminho_avisos)
        except:
            pass

    return f"{caminho_resultado}\n\n(AVISOS também gerado em: {os.path.basename(caminho_avisos)})" if caminho_avisos else caminho_resultado