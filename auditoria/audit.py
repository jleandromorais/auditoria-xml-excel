import os
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from .excel_loader import carregar_excel
from .report import gerar_relatorio
from .utils import safe_float
from .xml_parser import parse_xml_file

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
                if p.lower() in vistos:
                    continue
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

    df_base = carregar_excel(excel_path)
    if df_base.empty:
        raise RuntimeError("Não foi possível carregar os dados do Excel.")

    # --- FILTRO DE MÊS ---
    # Se um mês for passado (ex: "DEZ"), mantém apenas as notas desse mês do Excel
    if mes_filtro:
        mes_txt = str(mes_filtro).upper().strip()
        df_base = df_base[df_base["Mes"].astype(str).str.upper().str.contains(mes_txt, na=False)].copy()
        
        if df_base.empty:
            raise RuntimeError(f"Nenhum dado encontrado no Excel para o mês: {mes_filtro}")

    df_base["NF_Clean"] = df_base["NF_Clean"].astype(str).str.strip()

    # 1) Agrupamento e Soma de XMLs
    xmls_arquivos = coletar_xmls_por_empresas(pasta_pai, empresas)
    xmls_agrupados: Dict[str, Dict] = {} 

    for empresa_nome, xml_path in xmls_arquivos:
        try:
            info = parse_xml_file(xml_path)
        except Exception:
            info = None

        if not info or not info.get("Nota"):
            continue

        nota = str(info["Nota"]).strip()
        
        if nota not in xmls_agrupados:
            xmls_agrupados[nota] = {
                "Empresa": empresa_nome,
                "Tipo": info["Tipo"],
                "Arquivos": [os.path.basename(xml_path)],
                "Vol": 0.0, "Bruto": 0.0, "ICMS": 0.0, "PIS": 0.0, "COFINS": 0.0,
            }
        
        # Soma valores (para casos de notas fracionadas em vários XMLs)
        xmls_agrupados[nota]["Vol"] += info.get("Vol", 0.0)
        xmls_agrupados[nota]["Bruto"] += info.get("Bruto", 0.0)
        xmls_agrupados[nota]["ICMS"] += info.get("ICMS", 0.0)
        xmls_agrupados[nota]["PIS"] += info.get("PIS", 0.0)
        xmls_agrupados[nota]["COFINS"] += info.get("COFINS", 0.0)
        
        nome_arq = os.path.basename(xml_path)
        if nome_arq not in xmls_agrupados[nota]["Arquivos"]:
            xmls_agrupados[nota]["Arquivos"].append(nome_arq)

    # 2) Comparação baseada no Excel (Referência Principal)
    relatorio: List[Dict] = []
    notas_xml_vistas = set()

    for _, row in df_base.iterrows():
        nota_ex = str(row["NF_Clean"]).strip()
        if not nota_ex or nota_ex == "NAN": continue

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
            
            # Fallback PIS/COFINS para CT-e
            pis_xml, cof_xml = xml["PIS"], xml["COFINS"]
            if xml["Tipo"] == "CT-e" and pis_xml == 0 and (item["PIS Excel"] != 0):
                pis_xml, cof_xml = item["PIS Excel"], item["COFINS Excel"]
                item["Obs"] = "CT-e sem impostos no XML; usado Excel como fallback."
            
            item["PIS"], item["COFINS"] = pis_xml, cof_xml
            
            # Cálculo Líquido somado
            bruto = xml["Bruto"]
            liq_calc = bruto - sum(v for v in (xml["ICMS"], pis_xml, cof_xml) if 0 < v < bruto)
            item["Liq XML (Calc)"] = max(liq_calc, 0.0)
            
            item["Diff Vol"] = "-" if vol_ex == 0 else (xml["Vol"] - vol_ex)
            item["Diff R$"] = item["Liq XML (Calc)"] - liq_ex

            # Validação
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

        relatorio.append(item)

    # 3) Notas que estão no XML mas NÃO no Excel
    for nt, dados in xmls_agrupados.items():
        if nt not in notas_xml_vistas:
            relatorio.append({
                "Nota": nt, "Status": "SEM EXCEL ❌", "Empresa": dados["Empresa"],
                "Tipo": dados["Tipo"], "Vol XML": dados["Vol"], "Bruto XML": dados["Bruto"],
                "Liq XML (Calc)": dados["Bruto"], "Arquivo": ", ".join(dados["Arquivos"])
            })

    return gerar_relatorio(relatorio, saida=saida)