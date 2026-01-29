import sys
import os
import zipfile
import tempfile
from pathlib import Path

# Adiciona a pasta atual ao caminho do Python
sys.path.append(str(Path(__file__).parent))

try:
    from auditoria.audit import auditar_pasta_pai, AuditConfig
except ImportError:
    print("❌ Erro crítico: Não foi possível importar o sistema 'auditoria'.")
    input("Pressione ENTER para sair...")
    sys.exit(1)

def main():
    print("="*60)
    print("🚀 AUDITORIA RÁPIDA POR MÊS (Automática)")
    print("="*60)
    
    # 1. Pergunta qual mês o usuário quer processar
    print("\nQual mês deseja auditar? (Você deve ter a pasta: auditoria/MES)")
    print("Exemplos: OUT, NOV, DEZ")
    mes_input = input(">> Digite o mês: ").strip().upper()
    
    if not mes_input:
        print("❌ Nenhum mês digitado. Saindo.")
        return

    # 2. Localiza a pasta: auditoria/{MES}
    base_dir = Path(__file__).parent
    pasta_mes = base_dir / "auditoria" / mes_input
    
    if not pasta_mes.exists():
        print(f"\n❌ A pasta não existe: {pasta_mes}")
        print(f"   Crie a pasta 'auditoria/{mes_input}' e coloque o ZIP e o Excel lá.")
        input("\nPressione ENTER para sair...")
        return

    # 3. Caça os arquivos automaticamente
    zips = list(pasta_mes.glob("*.zip"))
    excels = list(pasta_mes.glob("*.xlsx"))

    if not zips:
        print(f"\n❌ Nenhum arquivo .zip encontrado em: {pasta_mes.name}")
        return
    if not excels:
        print(f"\n❌ Nenhum arquivo .xlsx encontrado em: {pasta_mes.name}")
        return

    # Pega os primeiros encontrados
    arquivo_zip = zips[0]
    arquivo_excel = excels[0]

    print(f"\n📂 Pasta: {pasta_mes.name}")
    print(f"   📦 ZIP:   {arquivo_zip.name}")
    print(f"   📊 Excel: {arquivo_excel.name}")

    confirm = input("\nConfirma a auditoria destes arquivos? (S/N): ").upper()
    if confirm != 'S':
        print("Cancelado.")
        return

    # 4. Executa a Auditoria
    # Define o nome do relatório final na raiz
    arquivo_saida = base_dir / f"Relatorio_Final_{mes_input}.xlsx"

    with tempfile.TemporaryDirectory() as tmp_dir:
        pasta_temp = Path(tmp_dir)
        print("\n⏳ Extraindo e processando... Aguarde.")
        
        try:
            with zipfile.ZipFile(arquivo_zip, "r") as z:
                z.extractall(pasta_temp)
            
            # Pega as pastas extraídas
            empresas = [p for p in pasta_temp.iterdir() if p.is_dir()]
            if not empresas: empresas = [pasta_temp]

            caminho_final = auditar_pasta_pai(
                pasta_pai=pasta_temp,
                empresas=empresas,
                excel_path=str(arquivo_excel),
                saida=str(arquivo_saida),
                mes_filtro=mes_input,  # Filtra o Excel pelo mês digitado
                config=AuditConfig()
            )

            print("\n" + "="*60)
            print(f"✅ SUCESSO! Relatório gerado.")
            print(f"📄 Arquivo: {caminho_final}")
            print("="*60)

            # Abre automaticamente no Windows
            if os.name == 'nt':
                os.startfile(caminho_final)
                
        except Exception as e:
            print(f"\n❌ Erro fatal: {e}")
            import traceback
            traceback.print_exc()
            input("\nPressione ENTER para ver o erro...")

if __name__ == "__main__":
    main()