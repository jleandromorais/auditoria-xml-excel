import os
from datetime import datetime
from typing import List, Dict, Optional

import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment


def gerar_relatorio(lista: List[Dict], saida: Optional[str] = None) -> str:
    df = pd.DataFrame(lista)
    cols = [
        "Arquivo",
        "Empresa",
        "Tipo",
        "Mes",
        "Nota",
        "Vol XML",
        "Vol Excel",
        "Diff Vol",
        "Bruto XML",
        "ICMS XML",
        "PIS",
        "COFINS",
        "ICMS Excel",
        "PIS Excel",
        "COFINS Excel",
        "Liq XML (Calc)",
        "Liq Excel",
        "Diff R$",
        "Status",
        "Obs",
    ]
    for c in cols:
        if c not in df.columns:
            df[c] = "-"

    df = df[cols]

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    if saida is None:
        saida = os.path.join(os.environ.get("USERPROFILE", os.getcwd()), "Downloads", f"Auditoria_XML_{ts}.xlsx")

    with pd.ExcelWriter(saida, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado")
        ws = writer.sheets["Resultado"]

        header_fill = PatternFill("solid", fgColor="203764")
        header_font = Font(bold=True, color="FFFFFF")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        verde = PatternFill("solid", fgColor="C6EFCE")
        vermelho = PatternFill("solid", fgColor="FFC7CE")

        status_col = cols.index("Status") + 1
        for row in ws.iter_rows(min_row=2):
            status = str(row[status_col - 1].value)
            cor = verde if "OK" in status else vermelho
            for cell in row:
                cell.fill = cor
                if isinstance(cell.value, (int, float)):
                    # valores em dinheiro
                    if cell.col_idx >= cols.index("Bruto XML") + 1:
                        cell.number_format = "R$ #,##0.00"
                    # volumes
                    if cell.col_idx in [cols.index("Vol XML") + 1, cols.index("Vol Excel") + 1, cols.index("Diff Vol") + 1]:
                        cell.number_format = "#,##0.000"

        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 22

    try:
        os.startfile(saida)
    except Exception:
        pass

    return saida
