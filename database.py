# auditoria/database.py
import duckdb
import pandas as pd
from typing import List, Dict

class AuditDB:
    def __init__(self, db_path='auditoria.db'):
        self.con = duckdb.connect(db_path)

    def inicializar(self):
        """Cria as tabelas necessárias"""
        # Tabela de XMLs (Dados Brutos)
        self.con.execute("""
            CREATE TABLE IF NOT EXISTS raw_xmls (
                chave VARCHAR,
                nota VARCHAR,
                data_emissao VARCHAR,
                emitente VARCHAR,
                cnpj_emitente VARCHAR,
                valor_total DOUBLE,
                vol DOUBLE,
                icms DOUBLE,
                pis DOUBLE,
                cofins DOUBLE,
                arquivo_origem VARCHAR,
                importado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        """)
        
        # Limpa tabelas temporárias para nova carga
        self.con.execute("DROP TABLE IF EXISTS raw_excel")
        self.con.execute("DROP TABLE IF EXISTS relatorio_final")

    def salvar_xmls(self, dados_xml: List[Dict]):
        """Salva a lista de dicionários dos XMLs"""
        if not dados_xml:
            return
        
        # Normaliza dados para o DF
        df = pd.DataFrame(dados_xml)
        
        # Seleciona colunas úteis e renomeia se necessário para bater com a tabela
        # Ajuste conforme as chaves reais que vêm do seu xml_parser.py
        colunas_map = {
            'Chave': 'chave', 'Nota': 'nota', 'Data': 'data_emissao',
            'Emitente': 'emitente', 'CNPJ': 'cnpj_emitente', 
            'Valor': 'valor_total', 'Vol': 'vol', 
            'ICMS': 'icms', 'PIS': 'pis', 'COFINS': 'cofins',
            'Arquivo': 'arquivo_origem'
        }
        
        # Garante que as colunas existam
        for k in colunas_map.keys():
            if k not in df.columns:
                df[k] = None

        df_final = df[list(colunas_map.keys())].rename(columns=colunas_map)
        
        # Insere só as 11 colunas; `importado_em` fica com o DEFAULT da tabela
        self.con.execute("""
            INSERT INTO raw_xmls (
                chave, nota, data_emissao, emitente, cnpj_emitente,
                valor_total, vol, icms, pis, cofins, arquivo_origem
            )
            SELECT * FROM df_final
        """)
        print(f"[DB] {len(df_final)} registros de XML salvos.")

    def salvar_excel(self, df_excel: pd.DataFrame):
        """Salva o DataFrame do Excel"""
        # Limpeza básica nos nomes das colunas para o SQL não reclamar
        df_excel.columns = [c.replace(" ", "_").replace(".", "") for c in df_excel.columns]
        self.con.execute("CREATE TABLE raw_excel AS SELECT * FROM df_excel")
        print(f"[DB] Tabela do Excel salva ({len(df_excel)} linhas).")

    def salvar_relatorio_final(self, df_relatorio: pd.DataFrame):
        """Salva o resultado final da auditoria (o que vai para o Excel)"""
        if df_relatorio.empty:
            return
        
        # Limpa nomes de colunas
        df_relatorio.columns = [c.replace(" ", "_").replace("(", "").replace(")", "").replace("$", "") for c in df_relatorio.columns]
        
        self.con.execute("CREATE TABLE relatorio_final AS SELECT * FROM df_relatorio")
        print("[DB] Relatório Final salvo no banco de dados para BI.")

    def fechar(self):
        self.con.close()