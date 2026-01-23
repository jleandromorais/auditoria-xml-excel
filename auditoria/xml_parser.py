import re
import xml.etree.ElementTree as ET
from typing import Dict, Optional

from .utils import to_float


def strip_ns(tag: str) -> str:
    return tag.split("}", 1)[1] if "}" in tag else tag


def iter_elems(root: ET.Element, name: str):
    for el in root.iter():
        if strip_ns(el.tag) == name:
            yield el


def get_first_text(root: ET.Element, path_names):
    def rec(node, idx):
        if idx == len(path_names):
            return node.text
        for ch in node:
            if strip_ns(ch.tag) == path_names[idx]:
                t = rec(ch, idx + 1)
                if t is not None:
                    return t
        return None

    return rec(root, 0)


def parse_nfe(root: ET.Element) -> Optional[Dict]:
    inf = next(iter(iter_elems(root, "infNFe")), None)
    if inf is None:
        return None

    nNF = get_first_text(inf, ["ide", "nNF"])
    nota = re.sub(r"\D", "", nNF or "")
    if nota:
        nota = str(int(nota))

    vNF = get_first_text(inf, ["total", "ICMSTot", "vNF"])
    vICMS = get_first_text(inf, ["total", "ICMSTot", "vICMS"])
    vPIS = get_first_text(inf, ["total", "ICMSTot", "vPIS"])
    vCOFINS = get_first_text(inf, ["total", "ICMSTot", "vCOFINS"])

    bruto = to_float(vNF)
    icms = to_float(vICMS)
    pis = to_float(vPIS)
    cof = to_float(vCOFINS)

    # Volume: soma itens com unidade M3/NM3
    vol = 0.0
    for det in iter_elems(inf, "det"):
        prod = next((ch for ch in det if strip_ns(ch.tag) == "prod"), None)
        if prod is None:
            continue
        uCom = get_first_text(prod, ["uCom"]) or ""
        qCom = get_first_text(prod, ["qCom"])
        u = uCom.upper().replace("³", "3")
        if "M3" in u:
            vol += to_float(qCom)

    # fallback transporte
    if vol == 0.0:
        qVol = get_first_text(inf, ["transp", "vol", "qVol"])
        if qVol:
            vol = to_float(qVol)

    liq = bruto
    for v in (icms, pis, cof):
        if bruto > 0 and 0 < v < bruto:
            liq -= v
    liq = max(liq, 0.0)

    return {
        "Tipo": "NF-e",
        "Nota": nota,
        "Vol": vol,
        "Bruto": bruto,
        "ICMS": icms,
        "PIS": pis,
        "COFINS": cof,
        "Liq_Calc": liq,
    }


def parse_cte(root: ET.Element) -> Optional[Dict]:
    inf = next(iter(iter_elems(root, "infCte")), None)
    if inf is None:
        return None

    nCT = get_first_text(inf, ["ide", "nCT"])
    nota = re.sub(r"\D", "", nCT or "")
    if nota:
        nota = str(int(nota))

    bruto = to_float(get_first_text(inf, ["vPrest", "vTPrest"]))

    icms = 0.0
    for el in iter_elems(inf, "vICMS"):
        icms = to_float(el.text)
        break

    pis = 0.0
    cof = 0.0
    for el in iter_elems(inf, "vPIS"):
        pis = to_float(el.text)
        break
    for el in iter_elems(inf, "vCOFINS"):
        cof = to_float(el.text)
        break

    vol = 0.0
    for infQ in iter_elems(inf, "infQ"):
        q = get_first_text(infQ, ["qCarga"])
        if q:
            vol = to_float(q)
            if vol > 0:
                break

    liq = bruto
    for v in (icms, pis, cof):
        if bruto > 0 and 0 < v < bruto:
            liq -= v
    liq = max(liq, 0.0)

    return {
        "Tipo": "CT-e",
        "Nota": nota,
        "Vol": vol,
        "Bruto": bruto,
        "ICMS": icms,
        "PIS": pis,
        "COFINS": cof,
        "Liq_Calc": liq,
    }


def parse_xml_file(path: str) -> Optional[Dict]:
    tree = ET.parse(path)
    root = tree.getroot()

    root_tag = strip_ns(root.tag).lower()
    tags = {strip_ns(el.tag) for el in root.iter()}

    # prioridade CT-e se for namespace/estrutura CT-e
    if "infCte" in tags and ("cte" in root_tag or "portalfiscal.inf.br/cte" in root.tag.lower()):
        return parse_cte(root)

    if "infNFe" in tags:
        return parse_nfe(root)

    if "infCte" in tags:
        return parse_cte(root)

    return None
