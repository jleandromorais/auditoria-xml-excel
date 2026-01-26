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
    """
    Retorna lista de (empresa_nome, caminho_xml) buscando XMLs recursivamente.
    """
    out: List[Tuple[str, str]] = []

    for empresa_dir in empresas:
        # pega tanto *.xml quanto *.XML (sem repetir)
        vistos = set()

        for arq in empresa_dir.rglob("*.xml"):
            p = str(arq)
            if p.lower() in vistos:
                continue
            vistos.add(p.lower())
            out.append((empresa_dir.name, p))

        for arq in empresa_dir.rglob("*.XML"):
            p = str(arq)
            if p.lower() in vistos:
                continue
            vistos.add(p.lower())
            out.append((empresa_dir.name, p))

    # ordena pra ficar previsível (empresa, arquivo)
    out.sort(key=lambda t: (t[0].lower(), os.path.basename(t[1]).lower()))
    return out


def auditar_pasta_pai(
    pasta_pai: Path,
    empresas: List[Path],
    excel_path: str,
    saida: Optional[str] = None,
    config: Optional[AuditConfig] = None,
) -> str:
    if config is None:
        config = AuditConfig()

    df_base = carregar_excel(excel_path)
    if df_base.empty:
        raise RuntimeError("Não consegui carregar as abas do Excel alvo (verifique ANO_ALVO e MESES_ALVO).")

    if "NF_Clean" not in df_base.columns:
        raise RuntimeError("Coluna NF_Clean não existe no df_base. Verifique excel_loader.py.")

    # Normaliza as NFs do Excel para evitar mismatch por tipo/espaço
    df_base["NF_Clean"] = df_base["NF_Clean"].astype(str).str.strip()

    xmls = coletar_xmls_por_empresas(pasta_pai, empresas)

    relatorio: List[Dict] = []
    notas_xml_encontradas = set()

    for empresa_nome, xml_path in xmls:
        nome_arquivo = os.path.basename(xml_path)

        try:
            info = parse_xml_file(xml_path)
        except Exception:
            info = None

        nota_xml = str(info["Nota"]).strip() if info and info.get("Nota") else ""

        item: Dict = {
            "Arquivo": nome_arquivo,
            "Empresa": empresa_nome,
            "Tipo": info["Tipo"] if info else "-",
            "Mes": "-",
            "Nota": nota_xml,
            "Vol XML": info["Vol"] if info else 0.0,
            "Bruto XML": info["Bruto"] if info else 0.0,
            "ICMS XML": info["ICMS"] if info else 0.0,
            "PIS": info["PIS"] if info else 0.0,
            "COFINS": info["COFINS"] if info else 0.0,
            "Liq XML (Calc)": info["Liq_Calc"] if info else 0.0,
            "Vol Excel": 0.0,
            "Liq Excel": 0.0,
            "ICMS Excel": 0.0,
            "PIS Excel": 0.0,
            "COFINS Excel": 0.0,
            "Diff Vol": 0.0,
            "Diff R$": 0.0,
            "Status": "ERRO PARSE ❌" if not info else "PENDENTE",
            "Obs": "",
        }

        if info and info.get("Nota"):
            nota = str(info["Nota"]).strip()
            notas_xml_encontradas.add(nota)

            match = df_base[df_base["NF_Clean"] == nota]
            if not match.empty:
                row = match.iloc[0]
                item["Mes"] = row.get("Mes", "-")

                vol_excel = safe_float(row.get("Vol_Excel", 0))
                liq_excel = safe_float(row.get("Liq_Excel", 0))
                icms_excel = safe_float(row.get("ICMS_Excel", 0))
                pis_excel = safe_float(row.get("PIS_Excel", 0))
                cof_excel = safe_float(row.get("COFINS_Excel", 0))

                item["Vol Excel"] = vol_excel
                item["Liq Excel"] = liq_excel
                item["ICMS Excel"] = icms_excel
                item["PIS Excel"] = pis_excel
                item["COFINS Excel"] = cof_excel

                if vol_excel == 0:
                    item["Obs"] = (item["Obs"] + " | " if item["Obs"] else "") + "Volume não encontrado no Excel."

                # Ajuste do líquido (CT-e sem PIS/COFINS no XML):
                icms = info["ICMS"]
                pis = info["PIS"]
                cof = info["COFINS"]
                if info["Tipo"] == "CT-e" and pis == 0 and cof == 0 and (pis_excel != 0 or cof_excel != 0):
                    pis = pis_excel
                    cof = cof_excel
                    item["Obs"] = (item["Obs"] + " | " if item["Obs"] else "") + \
                                  "CT-e sem PIS/COFINS no XML; usei valores do Excel p/ calcular líquido."

                liq_calc = info["Bruto"] - sum(v for v in (icms, pis, cof) if 0 < v < info["Bruto"])
                liq_calc = max(liq_calc, 0.0)
                item["Liq XML (Calc)"] = liq_calc

                item["Diff Vol"] = "-" if vol_excel == 0 else (info["Vol"] - vol_excel)
                item["Diff R$"] = liq_calc - liq_excel

                tol_r = config.tolerancia_cte if info["Tipo"] == "CT-e" else config.tolerancia_nfe
                financeiro_ok = abs(item["Diff R$"]) < tol_r

                volume_ok = True if vol_excel == 0 else abs(float(item["Diff Vol"])) < config.tolerancia_volume

                if financeiro_ok and volume_ok:
                    item["Status"] = "OK ✅"
                else:
                    status = []
                    if not volume_ok and vol_excel != 0:
                        status.append("VOL")
                    if not financeiro_ok:
                        status.append("VALOR")
                    item["Status"] = f"ERRO {'+'.join(status)} ❌"

            else:
                # XML existe, mas não existe no Excel base
                item["Status"] = "SEM EXCEL ❌"
                item["Obs"] = "XML encontrado, mas a nota não existe no Excel base."

        elif info and not info.get("Nota"):
            item["Status"] = "SEM NOTA ⚠️"
            item["Obs"] = "XML parseado, mas não consegui extrair número da nota."

        relatorio.append(item)

    # ============================================================
    # NOTAS DO EXCEL QUE NÃO FORAM ENCONTRADAS EM NENHUM XML (SEM XML)
    # ============================================================
    notas_excel = set(df_base["NF_Clean"].dropna())
    notas_xml = set(str(n).strip() for n in notas_xml_encontradas)
    notas_faltando_xml = notas_excel - notas_xml

    for nf in sorted(notas_faltando_xml):
        match = df_base[df_base["NF_Clean"] == nf]
        if match.empty:
            continue
        row = match.iloc[0]

        vol_excel = safe_float(row.get("Vol_Excel", 0))
        liq_excel = safe_float(row.get("Liq_Excel", 0))
        icms_excel = safe_float(row.get("ICMS_Excel", 0))
        pis_excel = safe_float(row.get("PIS_Excel", 0))
        cof_excel = safe_float(row.get("COFINS_Excel", 0))

        relatorio.append({
            "Arquivo": "-",
            "Empresa": "-",
            "Tipo": "-",
            "Mes": row.get("Mes", "-"),
            "Nota": nf,
            "Vol XML": 0.0,
            "Bruto XML": 0.0,
            "ICMS XML": 0.0,
            "PIS": 0.0,
            "COFINS": 0.0,
            "Liq XML (Calc)": 0.0,
            "Vol Excel": vol_excel,
            "Liq Excel": liq_excel,
            "ICMS Excel": icms_excel,
            "PIS Excel": pis_excel,
            "COFINS Excel": cof_excel,
            "Diff Vol": "-",
            "Diff R$": 0.0 - liq_excel,
            "Status": "SEM XML ❌",
            "Obs": "Nota consta no Excel, mas não foi localizado XML.",
        })

    return gerar_relatorio(relatorio, saida=saida)
