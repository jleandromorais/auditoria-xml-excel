from __future__ import annotations

import sys
from pathlib import Path
import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        "--mes",
        action="store",
        default="OUT",
        help="Mês alvo (OUT, NOV, DEZ). Ex: python -m pytest -q --mes=OUT",
    )


@pytest.fixture
def mes_alvo(request: pytest.FixtureRequest) -> str:
    # IMPORTANTE: pega pelo dest "mes" (sem --)
    mes = request.config.getoption("mes")
    mes = str(mes).strip().upper()
    if mes not in {"OUT", "NOV", "DEZ"}:
        pytest.fail("Mês inválido. Use OUT, NOV ou DEZ.")
    return mes
