import os
from datetime import datetime
from typing import List, Dict, Optional

import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment


def gerar_relatorio(lista: List[Dict], saida: Optional[str] = None, resumo: Optional[List[Dict]] = None) -> str:

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
        saida = os.path.join(
            os.environ.get("USERPROFILE", os.getcwd()),
            "Downloads",
            f"Auditoria_XML_{ts}.xlsx",
        )

    with pd.ExcelWriter(saida, engine="openpyxl") as writer:
        # ============================
        # Aba principal
        # ============================
        df.to_excel(writer, index=False, sheet_name="Resultado")
        ws = writer.sheets["Resultado"]

         # ============================
        # Aba 0: Resumo (Excel x XML por mês)
        # ============================
        if resumo:
            df_resumo = pd.DataFrame(resumo)
            cols_r = ["Mes", "Notas no Excel", "Notas com XML", "Notas sem XML"]
            for c in cols_r:
                if c not in df_resumo.columns:
                    df_resumo[c] = 0
            df_resumo[cols_r].to_excel(writer, index=False, sheet_name="Resumo")
        # ============================
        # Aba 1: Excel_sem_XML
        # ============================
        df_excel_sem_xml = df[df["Status"].astype(str).str.contains(r"SEM XML", regex=True, na=False)].copy()
        if not df_excel_sem_xml.empty:
            cols1 = ["Nota", "Mes", "Liq Excel", "Vol Excel", "ICMS Excel", "PIS Excel", "COFINS Excel", "Status", "Obs"]
            for c in cols1:
                if c not in df_excel_sem_xml.columns:
                    df_excel_sem_xml[c] = "-"
            df_excel_sem_xml[cols1].to_excel(writer, index=False, sheet_name="Excel_sem_XML")

        # ============================
        # Aba 2: XML_sem_Excel
        # ============================
        df_xml_sem_excel = df[df["Status"].astype(str).str.contains(r"SEM EXCEL", regex=True, na=False)].copy()
        if not df_xml_sem_excel.empty:
            cols2 = ["Nota", "Empresa", "Tipo", "Bruto XML", "Liq XML (Calc)", "Vol XML", "ICMS XML", "PIS", "COFINS", "Status", "Obs", "Arquivo"]
            for c in cols2:
                if c not in df_xml_sem_excel.columns:
                    df_xml_sem_excel[c] = "-"
            df_xml_sem_excel[cols2].to_excel(writer, index=False, sheet_name="XML_sem_Excel")

        # ============================
        # Estilo da aba Resultado
        # ============================
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
                    # dinheiro a partir de Bruto XML (inclui ICMS/PIS/COFINS, Liq, Diff)
                    if cell.col_idx >= cols.index("Bruto XML") + 1:
                        cell.number_format = "R$ #,##0.00"

                    # volumes
                    if cell.col_idx in [
                        cols.index("Vol XML") + 1,
                        cols.index("Vol Excel") + 1,
                        cols.index("Diff Vol") + 1,
                    ]:
                        cell.number_format = "#,##0.000"

        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 22

    try:
        os.startfile(saida)
    except Exception:
        pass

    return saida
