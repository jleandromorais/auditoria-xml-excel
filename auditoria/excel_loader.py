import pandas as pd

from .utils import ANO_ALVO, MESES_ALVO, limpar_numero_nf_bruto, make_unique_columns, to_float


def carregar_excel(caminho: str) -> pd.DataFrame:
    dados = []
    xls = pd.read_excel(caminho, sheet_name=None, header=None)

    for aba, df in xls.items():
        aba_upper = str(aba).upper()
        if ANO_ALVO not in aba_upper:
            continue
        if not any(mes in aba_upper for mes in MESES_ALVO):
            continue

        idx = -1
        for i, row in df.head(120).iterrows():
            linha = [str(x).upper() for x in row.values]
            if any("NOTA" in x for x in linha) and (
                any("S/TRIBUTOS" in x for x in linha)
                or any("C/TRIBUTOS" in x for x in linha)
                or any("TOTAL" in x for x in linha)
            ):
                idx = i
                break

        if idx == -1:
            continue

        cols = make_unique_columns([str(c).upper().strip() for c in df.iloc[idx]])
        df2 = df[idx + 1 :].copy()
        df2.columns = cols

        c_nf = next((c for c in df2.columns if "NOTA" in c or c == "NF"), None)
        c_liq = next((c for c in df2.columns if "S/TRIBUTOS" in c), None)
        c_vol = next(
            (
                c
                for c in df2.columns
                if "VOL" in c or "M³" in c or "M3" in c or "QTDE" in c or "QTD" in c or "QUANT" in c
            ),
            None,
        )

        c_icms = [c for c in df2.columns if c.startswith("ICMS")]
        c_pis = [c for c in df2.columns if c.startswith("PIS")]
        c_cof = [c for c in df2.columns if c.startswith("COFINS")]

        c_icms = c_icms[-1] if c_icms else None
        c_pis = c_pis[-1] if c_pis else None
        c_cof = c_cof[-1] if c_cof else None

        if not (c_nf and c_liq):
            continue

        
        temp = df2.copy()

        # ============================================================
        # Correção para planilhas com células mescladas/linhas repetidas:
        # Em muitos relatórios, o número da Nota Fiscal aparece só na 1ª linha
        # e as linhas seguintes ficam "em branco" (visual), mas ainda fazem parte da mesma NF.
        # Aqui, nós propagamos (ffill) a NF para as linhas vazias APENAS quando a linha tem
        # valores relevantes (ex.: líquidos/volume/impostos), evitando preencher totais/linhas de separação.
        # ============================================================
        def _has_content(v) -> bool:
            if pd.isna(v):
                return False
            s = str(v).strip()
            return s != "" and s.upper() != "NAN"

        nf_raw = temp[c_nf]
        nf_norm = nf_raw.where(nf_raw.apply(_has_content), pd.NA)

        cols_relevantes = [c for c in [c_liq, c_vol, c_icms, c_pis, c_cof] if c]
        if cols_relevantes:
            row_has_values = temp[cols_relevantes].applymap(_has_content).any(axis=1)
        else:
            row_has_values = pd.Series([True] * len(temp), index=temp.index)

        nf_ffill = nf_norm.ffill()
        nf_final = nf_norm.copy()
        mask_preencher = nf_norm.isna() & row_has_values
        nf_final.loc[mask_preencher] = nf_ffill.loc[mask_preencher]
        temp["NF_Clean"] = nf_final.apply(limpar_numero_nf_bruto)
        temp["Vol_Excel"] = temp[c_vol].apply(to_float) if c_vol else 0.0
        temp["Liq_Excel"] = temp[c_liq].apply(to_float)

        temp["ICMS_Excel"] = temp[c_icms].apply(to_float) if c_icms else 0.0
        temp["PIS_Excel"] = temp[c_pis].apply(to_float) if c_pis else 0.0
        temp["COFINS_Excel"] = temp[c_cof].apply(to_float) if c_cof else 0.0

        temp["Mes"] = aba
        temp = temp[temp["NF_Clean"] != ""]

        if not temp.empty:
            dados.append(
                temp[
                    ["NF_Clean", "Vol_Excel", "Liq_Excel", "ICMS_Excel", "PIS_Excel", "COFINS_Excel", "Mes"]
                ]
            )

    return pd.concat(dados, ignore_index=True) if dados else pd.DataFrame()
