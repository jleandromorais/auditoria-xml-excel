import os
from datetime import datetime
from typing import List, Dict, Optional
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment

def gerar_relatorio(lista: List[Dict], saida: Optional[str] = None) -> str:
    """Gera o relatório principal da auditoria (Comparativo)."""
    df = pd.DataFrame(lista)
    cols = [
        "Arquivo", "Empresa", "Tipo", "Mes", "Nota", "Vol XML", "Vol Excel", "Diff Vol",
        "Bruto XML", "ICMS XML", "PIS", "COFINS", "ICMS Excel", "PIS Excel", "COFINS Excel",
        "Liq XML (Calc)", "Liq Excel", "Diff R$", "Status", "Obs"
    ]

    for c in cols:
        if c not in df.columns: df[c] = "-"

    # --- TOTAIS GERAIS ---
    df_temp = df.copy()
    cols_num = ["Vol XML", "Vol Excel", "Bruto XML", "ICMS XML", "PIS", "COFINS", 
                "ICMS Excel", "PIS Excel", "COFINS Excel", "Liq XML (Calc)", "Liq Excel", "Diff R$"]
    
    for c in cols_num:
        df_temp[c] = pd.to_numeric(df_temp[c], errors='coerce').fillna(0)

    totais = {c: df_temp[c].sum() for c in cols_num}
    totais.update({"Arquivo": "TOTAIS GERAIS", "Status": "---", "Nota": "---", "Mes": "---"})
    
    df = pd.concat([df[cols], pd.DataFrame([totais])], ignore_index=True)

    if not saida:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        saida = os.path.join(os.path.expanduser("~"), "Downloads", f"Auditoria_Resultado_{ts}.xlsx")

    with pd.ExcelWriter(saida, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado")
        ws = writer.sheets["Resultado"]
        _estilizar_planilha(ws, cols)

    return saida

def gerar_relatorio_avisos(df_duplicadas: pd.DataFrame, lista_sem_xml: List[Dict], pasta_destino: str) -> str:
    """
    Gera um relatório SECUNDÁRIO apenas com os problemas para a chefia.
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho = os.path.join(os.path.dirname(pasta_destino), f"Auditoria_AVISOS_GERENCIA_{ts}.xlsx")

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        
        # --- ABA 1: DUPLICIDADES (PROVA DO ERRO DO EXCEL) ---
        if not df_duplicadas.empty:
            # Seleciona colunas relevantes para provar a duplicidade
            cols_dup = ["Mes", "NF_Clean", "Vol_Excel", "Liq_Excel", "ICMS_Excel", "PIS_Excel", "COFINS_Excel"]
            # Garante que existem
            cols_existentes = [c for c in cols_dup if c in df_duplicadas.columns]
            
            df_dup_export = df_duplicadas[cols_existentes].sort_values(by="NF_Clean")
            df_dup_export.to_excel(writer, index=False, sheet_name="Excel_Duplicados")
            
            # Estilo simples
            ws = writer.sheets["Excel_Duplicados"]
            ws.column_dimensions["B"].width = 20 # NF
            # Cabeçalho Vermelho para chamar atenção
            red_fill = PatternFill("solid", fgColor="FFC7CE")
            for cell in ws[1]:
                cell.fill = red_fill
                cell.font = Font(bold=True)

        # --- ABA 2: FALTAM XMLS ---
        if lista_sem_xml:
            df_sem = pd.DataFrame(lista_sem_xml)
            cols_sem = ["Nota", "Mes", "Liq Excel", "Vol Excel", "Obs"]
            # Filtra colunas
            df_sem = df_sem[[c for c in cols_sem if c in df_sem.columns]]
            
            df_sem.to_excel(writer, index=False, sheet_name="Faltam_XMLs")
            ws2 = writer.sheets["Faltam_XMLs"]
            for cell in ws2[1]:
                cell.font = Font(bold=True)
            ws2.column_dimensions["A"].width = 20

    return caminho

def _estilizar_planilha(ws, cols):
    """Função auxiliar de estilo para o relatório principal."""
    header_f = PatternFill("solid", fgColor="203764")
    verde_f = PatternFill("solid", fgColor="C6EFCE")
    vermelho_f = PatternFill("solid", fgColor="FFC7CE")
    total_f = PatternFill("solid", fgColor="D3D3D3")

    for cell in ws[1]:
        cell.fill, cell.font = header_f, Font(bold=True, color="FFFFFF")

    max_r = ws.max_row
    st_col_idx = cols.index("Status") + 1

    for r_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        if r_idx == max_r:
            for cell in row:
                cell.fill, cell.font = total_f, Font(bold=True)
        else:
            status_val = str(row[st_col_idx-1].value)
            cor = verde_f if "OK" in status_val else vermelho_f
            for cell in row: cell.fill = cor

        for cell in row:
            if isinstance(cell.value, (int, float)):
                if cell.col_idx >= cols.index("Bruto XML") + 1:
                    cell.number_format = "R$ #,##0.00"
                else:
                    cell.number_format = "#,##0.000"

    for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 20