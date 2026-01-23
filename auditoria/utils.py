import math
import re
from typing import Iterable, List

import pandas as pd


ANO_ALVO = "25"
MESES_ALVO = ["OUT", "NOV", "DEZ"]


def limpar_numero_nf_bruto(valor) -> str:
    if pd.isna(valor) or str(valor).strip() == "":
        return ""
    texto = str(valor).upper().strip()
    if "-" in texto:
        texto = texto.split("-")[0]
    if "/" in texto:
        texto = texto.split("/")[0]
    nums = re.findall(r"\d+", texto.replace(".", ""))
    if nums:
        return str(int(nums[0]))
    return ""


def to_float(texto) -> float:
    if texto is None or (isinstance(texto, float) and math.isnan(texto)):
        return 0.0
    if isinstance(texto, (int, float)):
        return float(texto)
    if not isinstance(texto, str):
        texto = str(texto)

    clean = texto.replace(" ", "")
    if clean == "":
        return 0.0

    # Corrige múltiplas vírgulas (milhar errado)
    if clean.count(",") > 1:
        partes = clean.split(",")
        inteiro = "".join(partes[:-1])
        decimal = partes[-1]
        clean = f"{inteiro}.{decimal}"

    limpo = re.sub(r"[^\d.,-]", "", clean)

    if "," in limpo and "." not in limpo:
        limpo = limpo.replace(",", ".")
    elif "," in limpo and "." in limpo:
        limpo = limpo.replace(".", "").replace(",", ".")

    try:
        return float(limpo)
    except Exception:
        return 0.0


def make_unique_columns(cols: Iterable[str]) -> List[str]:
    seen = {}
    out = []
    for c in cols:
        base = str(c).strip()
        if base in seen:
            seen[base] += 1
            out.append(f"{base}__{seen[base]}")
        else:
            seen[base] = 0
            out.append(base)
    return out


def safe_float(x) -> float:
    if x is None:
        return 0.0
    try:
        if isinstance(x, float) and math.isnan(x):
            return 0.0
    except Exception:
        pass
    try:
        return float(x)
    except Exception:
        return 0.0
