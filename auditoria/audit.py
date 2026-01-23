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

    xmls = coletar_xmls_por_empresas(pasta_pai, empresas)

    relatorio: List[Dict] = []

    for empresa_nome, xml_path in xmls:
        nome_arquivo = os.path.basename(xml_path)

        try:
            info = parse_xml_file(xml_path)
        except Exception:
            info = None

        item: Dict = {
            "Arquivo": nome_arquivo,
            "Empresa": empresa_nome,
            "Tipo": info["Tipo"] if info else "-",
            "Mes": "-",
            "Nota": info["Nota"] if info else "",
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
            "Status": "ERRO PARSE ❌" if not info else "Ñ ENCONTRADO ⚠️",
            "Obs": "",
        }

        if info and info.get("Nota"):
            match = df_base[df_base["NF_Clean"] == info["Nota"]]
            if not match.empty:
                row = match.iloc[0]
                item["Mes"] = row["Mes"]

                vol_excel = safe_float(row["Vol_Excel"])
                liq_excel = safe_float(row["Liq_Excel"])
                icms_excel = safe_float(row["ICMS_Excel"])
                pis_excel = safe_float(row["PIS_Excel"])
                cof_excel = safe_float(row["COFINS_Excel"])

                item["Vol Excel"] = vol_excel if vol_excel != 0 else "NÃO NO EXCEL"
                item["Liq Excel"] = liq_excel
                item["ICMS Excel"] = icms_excel
                item["PIS Excel"] = pis_excel
                item["COFINS Excel"] = cof_excel

                # Ajuste do líquido (CT-e sem PIS/COFINS no XML):
                icms = info["ICMS"]
                pis = info["PIS"]
                cof = info["COFINS"]
                if info["Tipo"] == "CT-e" and pis == 0 and cof == 0 and (pis_excel != 0 or cof_excel != 0):
                    pis = pis_excel
                    cof = cof_excel
                    item["Obs"] = "CT-e sem PIS/COFINS no XML; usei valores do Excel p/ calcular líquido."

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

        relatorio.append(item)

    return gerar_relatorio(relatorio, saida=saida)
