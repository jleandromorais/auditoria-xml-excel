import pandas as pd
import tempfile
from pathlib import Path

from auditoria.excel_loader import carregar_excel


def test_carregar_excel_minimo():
    # cria um excel com aba "OUT_25" e headers parecidos com seu padrão
    df = pd.DataFrame([
        ["NOTA", "S/TRIBUTOS", "VOL", "ICMS", "PIS", "COFINS"],
        ["123", "87,00", "0", "10,00", "1,00", "2,00"],
    ])

    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "base.xlsx"
        with pd.ExcelWriter(p, engine="openpyxl") as w:
            df.to_excel(w, sheet_name="OUT_25", index=False, header=False)

        base = carregar_excel(str(p))
        assert not base.empty
        assert base.iloc[0]["NF_Clean"] == "123"
