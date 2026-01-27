import os
from datetime import datetime
from typing import List, Dict, Optional

import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment

def gerar_relatorio(lista: List[Dict], saida: Optional[str] = None, resumo: Optional[List[Dict]] = None) -> str:
    # Cria o DataFrame base com os resultados da auditoria
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

    # Garante que todas as colunas necessárias existem
    for c in cols:
        if c not in df.columns:
            df[c] = "-"

    # --- LÓGICA DE TOTAIS GERAIS ---
    # Criar uma cópia para calcular somas numéricas sem afetar os dados originais
    df_temp = df.copy()
    
    # Converter colunas numéricas (forçar erros para NaN e depois para 0)
    cols_numericas = [
        "Vol XML", "Vol Excel", "Bruto XML", "ICMS XML", "PIS", "COFINS",
        "ICMS Excel", "PIS Excel", "COFINS Excel", "Liq XML (Calc)", "Liq Excel", "Diff R$"
    ]
    
    for c in cols_numericas:
        df_temp[c] = pd.to_numeric(df_temp[c], errors='coerce').fillna(0)

    # Calcular somas
    totais = {col: df_temp[col].sum() for col in cols_numericas}
    totais["Arquivo"] = "TOTAIS GERAIS"
    totais["Status"] = "---"
    totais["Nota"] = "---"

    # Adicionar a linha de total ao DataFrame principal
    df = pd.concat([df[cols], pd.DataFrame([totais])], ignore_index=True)

    # Configuração do caminho de saída
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    if saida is None:
        saida = os.path.join(
            os.environ.get("USERPROFILE", os.getcwd()),
            "Downloads",
            f"Auditoria_XML_{ts}.xlsx",
        )

    with pd.ExcelWriter(saida, engine="openpyxl") as writer:
        # Aba Principal: Resultado
        df.to_excel(writer, index=False, sheet_name="Resultado")
        ws = writer.sheets["Resultado"]

        # Aba Resumo (Excel x XML por mês)
        if resumo:
            df_resumo = pd.DataFrame(resumo)
            cols_r = ["Mes", "Notas no Excel", "Notas com XML", "Notas sem XML"]
            for c in cols_r:
                if c not in df_resumo.columns:
                    df_resumo[c] = 0
            df_resumo[cols_r].to_excel(writer, index=False, sheet_name="Resumo")

        # Aba Excel_sem_XML
        df_excel_sem_xml = df[df["Status"].astype(str).str.contains(r"SEM XML", regex=True, na=False)].copy()
        if not df_excel_sem_xml.empty:
            cols1 = ["Nota", "Mes", "Liq Excel", "Vol Excel", "ICMS Excel", "PIS Excel", "COFINS Excel", "Status", "Obs"]
            df_excel_sem_xml[cols1].to_excel(writer, index=False, sheet_name="Excel_sem_XML")

        # Aba XML_sem_Excel
        df_xml_sem_excel = df[df["Status"].astype(str).str.contains(r"SEM EXCEL", regex=True, na=False)].copy()
        if not df_xml_sem_excel.empty:
            cols2 = ["Nota", "Empresa", "Tipo", "Bruto XML", "Liq XML (Calc)", "Vol XML", "ICMS XML", "PIS", "COFINS", "Status", "Obs", "Arquivo"]
            df_xml_sem_excel[cols2].to_excel(writer, index=False, sheet_name="XML_sem_Excel")

        # --- ESTILO E FORMATAÇÃO ---
        # Cabeçalho
        header_fill = PatternFill("solid", fgColor="203764")
        header_font = Font(bold=True, color="FFFFFF")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        verde = PatternFill("solid", fgColor="C6EFCE")
        vermelho = PatternFill("solid", fgColor="FFC7CE")
        cinza_total = PatternFill("solid", fgColor="D3D3D3")
        
        status_col_idx = cols.index("Status")
        max_row = ws.max_row

        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            status = str(row[status_col_idx].value)
            
            # Formatação especial para a linha de Total (última linha)
            if row_idx == max_row:
                for cell in row:
                    cell.fill = cinza_total
                    cell.font = Font(bold=True)
            else:
                # Cores baseadas no Status
                cor = verde if "OK" in status else vermelho
                for cell in row:
                    cell.fill = cor

            # Formatação numérica
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    # Moeda
                    if cell.col_idx >= cols.index("Bruto XML") + 1:
                        cell.number_format = "R$ #,##0.00"
                    # Volumes
                    if cell.col_idx in [cols.index("Vol XML")+1, cols.index("Vol Excel")+1, cols.index("Diff Vol")+1]:
                        cell.number_format = "#,##0.000"

        # Ajuste de largura das colunas
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 22

    try:
        os.startfile(saida)
    except Exception:
        pass

    return saida