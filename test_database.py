import pandas as pd

from database import AuditDB


def main():
    db = AuditDB("auditoria_test.db")
    try:
        db.inicializar()

        # 1) Teste de XMLs (lista de dicts)
        dados_xml = [
            {
                "Chave": "NFe123",
                "Nota": "0001",
                "Data": "2026-01-28",
                "Emitente": "Empresa A",
                "CNPJ": "00.000.000/0001-00",
                "Valor": 123.45,
                "Vol": 1.0,
                "ICMS": 10.0,
                "PIS": 2.0,
                "COFINS": 3.0,
                "Arquivo": "arquivo1.xml",
            },
            {
                "Chave": "NFe456",
                "Nota": "0002",
                "Data": "2026-01-28",
                "Emitente": "Empresa B",
                "CNPJ": "11.111.111/0001-11",
                "Valor": 999.99,
                "Vol": 2.0,
                "ICMS": 0.0,
                "PIS": 0.0,
                "COFINS": 0.0,
                "Arquivo": "arquivo2.xml",
            },
        ]
        db.salvar_xmls(dados_xml)

        df_xmls = db.con.execute(
            "SELECT chave, nota, data_emissao, emitente, valor_total, arquivo_origem, importado_em FROM raw_xmls"
        ).df()
        print("\n[TESTE] raw_xmls:")
        print(df_xmls)

        # 2) Teste de Excel (DataFrame)
        df_excel = pd.DataFrame(
            [
                {"Nota": "0001", "Valor Total": 123.45, "Alguma Coluna": "x"},
                {"Nota": "0002", "Valor Total": 999.99, "Alguma Coluna": "y"},
            ]
        )
        db.salvar_excel(df_excel)

        df_excel_out = db.con.execute("SELECT * FROM raw_excel").df()
        print("\n[TESTE] raw_excel:")
        print(df_excel_out)

        print("\n[OK] Testes concluídos. Banco gerado: auditoria_test.db")
    finally:
        db.fechar()


if __name__ == "__main__":
    main()

