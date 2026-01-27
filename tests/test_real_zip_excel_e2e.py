from __future__ import annotations

import json
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional

import pytest

from auditoria.audit import auditar_pasta_pai, AuditConfig


CONFIG_FILE = Path(__file__).resolve().parents[1] / "test_config.json"


def _load_cfg() -> dict:
    if not CONFIG_FILE.exists():
        return {}
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


@pytest.fixture
def cfg() -> dict:
    c = _load_cfg()
    if "base_dir" not in c:
        pytest.fail('Crie test_config.json na raiz com: { "base_dir": "C:/Auditoria" }')
    return c


@pytest.fixture
def paths(cfg: dict, mes_alvo: str) -> Dict[str, Path]:
    base = Path(cfg["base_dir"]) / mes_alvo
    excel_path = base / "base.xlsx"
    zip_path = base / "xmls_empresas.zip"

    if not excel_path.exists():
        pytest.fail(f"Excel não encontrado: {excel_path}")
    if not zip_path.exists():
        pytest.fail(f"ZIP não encontrado: {zip_path}")

    return {"excel": excel_path, "zip": zip_path}


@pytest.fixture
def pasta_pai_extraida(tmp_path: Path, paths: Dict[str, Path]) -> Path:
    root = tmp_path / "pasta_pai"
    root.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(paths["zip"], "r") as z:
        z.extractall(root)

    return root


@pytest.fixture
def empresas(pasta_pai_extraida: Path) -> List[Path]:
    dirs = [p for p in pasta_pai_extraida.iterdir() if p.is_dir()]
    dirs.sort(key=lambda p: p.name.lower())
    return dirs


@pytest.fixture
def capturar_relatorio(monkeypatch: pytest.MonkeyPatch) -> Dict[str, Any]:
    captured: Dict[str, Any] = {"relatorio": None, "resumo": None}

    def fake_gerar_relatorio(relatorio: List[Dict], saida: Optional[str] = None, resumo: Optional[List[Dict]] = None) -> str:
        captured["relatorio"] = relatorio
        captured["resumo"] = resumo
        return "RELATORIO_OK"

    import auditoria.audit as audit_mod
    monkeypatch.setattr(audit_mod, "gerar_relatorio", fake_gerar_relatorio)
    return captured


def _find_items(relatorio: List[Dict[str, Any]], status: str) -> List[Dict[str, Any]]:
    return [r for r in relatorio if str(r.get("Status", "")).strip() == status]


def test_e2e_real_zip_excel_por_mes(
    mes_alvo: str,
    paths: Dict[str, Path],
    pasta_pai_extraida: Path,
    empresas: List[Path],
    capturar_relatorio: Dict[str, Any],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import auditoria.utils as utils_mod
    import auditoria.excel_loader as loader_mod

    utils_mod.MESES_ALVO = [mes_alvo]
    loader_mod.MESES_ALVO = [mes_alvo]
    utils_mod.ANO_ALVO = "25"
    loader_mod.ANO_ALVO = "25"

    out = auditar_pasta_pai(
        pasta_pai=pasta_pai_extraida,
        empresas=empresas,
        excel_path=str(paths["excel"]),
        saida=None,
        config=AuditConfig(tolerancia_cte=50.0, tolerancia_nfe=5.0, tolerancia_volume=1.0),
    )

    assert out == "RELATORIO_OK"

    relatorio = capturar_relatorio["relatorio"]
    assert isinstance(relatorio, list) and len(relatorio) > 0

    # SEM XML
    sem_xml = _find_items(relatorio, "SEM XML ❌")
    if sem_xml:
        notas = sorted({str(r.get("Nota", "")).strip() for r in sem_xml})
        pytest.fail("Encontrou notas no Excel sem XML:\n" + "\n".join(notas))

    # CT-e: valida Liq XML (Calc) quando usou PIS/COFINS do Excel
    ctes = [r for r in relatorio if str(r.get("Tipo", "")).strip() == "CT-e"]
    for r in ctes:
        obs = str(r.get("Obs", ""))
        if "CT-e sem PIS/COFINS no XML" in obs:
            bruto = float(r.get("Bruto XML", 0.0) or 0.0)
            icms = float(r.get("ICMS XML", 0.0) or 0.0)
            pis_ex = float(r.get("PIS Excel", 0.0) or 0.0)
            cof_ex = float(r.get("COFINS Excel", 0.0) or 0.0)

            esperado = bruto - sum(v for v in (icms, pis_ex, cof_ex) if 0 < v < bruto)
            esperado = max(esperado, 0.0)

            liq_calc = float(r.get("Liq XML (Calc)", 0.0) or 0.0)
            assert abs(liq_calc - esperado) < 1e-6
