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


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption("--mes", action="store", default="OUT/2025")


@pytest.fixture
def mes_alvo(request: pytest.FixtureRequest) -> str:
    mes = request.config.getoption("--mes")
    if not mes:
        pytest.fail("Passe --mes=OUT/2025 (ou similar).")
    return str(mes).strip()


@pytest.fixture
def cfg() -> dict:
    c = _load_cfg()
    if "excel_path" not in c or "zip_path" not in c:
        pytest.fail("Config não encontrada. Rode o GUI e selecione Excel e ZIP (test_config.json).")
    return c


@pytest.fixture
def pasta_pai_extraida(tmp_path: Path, cfg: dict) -> Path:
    zip_path = Path(cfg["zip_path"])
    if not zip_path.exists():
        pytest.fail(f"ZIP não existe: {zip_path}")

    root = tmp_path / "pasta_pai"
    root.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(root)

    return root


@pytest.fixture
def empresas(pasta_pai_extraida: Path) -> List[Path]:
    dirs = [p for p in pasta_pai_extraida.iterdir() if p.is_dir()]
    dirs.sort(key=lambda p: p.name.lower())
    return dirs


@pytest.fixture
def capturar_relatorio(monkeypatch: pytest.MonkeyPatch) -> Dict[str, Any]:
    """
    Captura relatorio e resumo sem necessariamente salvar arquivo final,
    mas mantém a lógica intacta.
    """
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


def test_e2e_real_zip_excel(
    mes_alvo: str,
    cfg: dict,
    pasta_pai_extraida: Path,
    empresas: List[Path],
    capturar_relatorio: Dict[str, Any],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """
    Regras checadas automaticamente:
      - Não pode existir SEM XML ❌ (Excel tem nota que não existe nos XML)
      - CT-e: valida Liq XML (Calc) quando XML não tem PIS/COFINS e Excel tem
      - Itens OK ✅ devem ter Diff dentro da tolerância (já é tua regra)
    """

    excel_path = Path(cfg["excel_path"])
    if not excel_path.exists():
        pytest.fail(f"Excel não existe: {excel_path}")

    # Filtro por mês: como teu carregar_excel já filtra por MESES_ALVO/ANO_ALVO,
    # aqui a gente patcha MESES_ALVO para o mês selecionado.
    # Se no teu projeto o mês for OUT/2025 e o loader usa OUT_25, adapte aqui.
    # Vou fazer patch do jeito mais seguro: tentar setar MESES_ALVO com a chave "OUT", "NOV", ...
    mes_key = mes_alvo.split("/")[0].strip().upper()

    import auditoria.utils as utils_mod
    import auditoria.excel_loader as loader_mod

    utils_mod.MESES_ALVO = [mes_key]
    loader_mod.MESES_ALVO = [mes_key]

    # Em geral seu ANO_ALVO é "25" (2025). Se variar, você pode puxar do mes_alvo.
    # Vou manter "25" para não quebrar tua lógica atual.
    utils_mod.ANO_ALVO = "25"
    loader_mod.ANO_ALVO = "25"

    out = auditar_pasta_pai(
        pasta_pai=pasta_pai_extraida,
        empresas=empresas,
        excel_path=str(excel_path),
        saida=None,
        config=AuditConfig(tolerancia_cte=50.0, tolerancia_nfe=5.0, tolerancia_volume=1.0),
    )

    assert out == "RELATORIO_OK"
    relatorio = capturar_relatorio["relatorio"]
    assert isinstance(relatorio, list) and len(relatorio) > 0

    # 1) SEM XML ❌ não pode existir (se existir, falha e lista as notas)
    sem_xml = _find_items(relatorio, "SEM XML ❌")
    if sem_xml:
        notas = sorted({str(r.get("Nota", "")).strip() for r in sem_xml})
        pytest.fail("Encontrou notas no Excel sem XML:\n" + "\n".join(notas))

    # 2) CT-e: se Obs indica “CT-e sem PIS/COFINS no XML; usei valores do Excel…”
    # então Liq XML (Calc) deve bater com: Bruto XML - (ICMS XML + PIS Excel + COFINS Excel)
    # OBS: teu código só faz isso quando PIS/COFINS no XML = 0 e no Excel != 0.
    ctes = [r for r in relatorio if str(r.get("Tipo", "")).strip() == "CT-e"]
    for r in ctes:
        obs = str(r.get("Obs", ""))
        if "CT-e sem PIS/COFINS no XML" in obs:
            bruto = float(r.get("Bruto XML", 0.0) or 0.0)
            icms = float(r.get("ICMS XML", 0.0) or 0.0)
            pis_ex = float(r.get("PIS Excel", 0.0) or 0.0)
            cof_ex = float(r.get("COFINS Excel", 0.0) or 0.0)
            esperado = bruto - sum(v for v in (icms, pis_ex, cof_ex) if 0 < v < bruto)
            if esperado < 0:
                esperado = 0.0
            liq_calc = float(r.get("Liq XML (Calc)", 0.0) or 0.0)

            assert abs(liq_calc - esperado) < 1e-6, (
                f"CT-e {r.get('Nota')} Liq XML (Calc) inválido.\n"
                f"Esperado={esperado} | Obtido={liq_calc}\n"
                f"Bruto={bruto} ICMS={icms} PIS_Excel={pis_ex} COFINS_Excel={cof_ex}"
            )

    # 3) Itens OK ✅: apenas valida que realmente estão OK pelo teu critério
    # (se tua lógica marcar OK, não deve haver “ERRO ... ❌” escondido)
    erros = [r for r in relatorio if str(r.get("Status", "")).startswith("ERRO")]
    if erros:
        exemplos = erros[:30]
        linhas = []
        for e in exemplos:
            linhas.append(f"{e.get('Nota')} | {e.get('Tipo')} | {e.get('Status')} | DiffR$={e.get('Diff R$')} | DiffVol={e.get('Diff Vol')}")
        pytest.fail("Encontrou itens com ERRO:\n" + "\n".join(linhas))
