
# 📊 Auditoria XML NF-e / CT-e x Excel

Aplicação desktop em **Python** para **auditoria automática de XMLs fiscais (NF-e e CT-e)**, comparando os dados dos arquivos XML com uma planilha Excel e gerando um **relatório detalhado em XLSX**.

O sistema percorre uma **pasta PAI**, identifica **empresas em subpastas**, coleta todos os XMLs de forma **recursiva** e realiza a validação financeira e de volume.

---

## 🚀 Funcionalidades

- 📁 Seleção de **pasta PAI** com múltiplas empresas  
- 🏢 Seleção das empresas a serem auditadas  
- 🔍 Leitura **recursiva** de XMLs (`.xml` / `.XML`)  
- 🧾 Suporte a **NF-e** e **CT-e**  
- 📊 Comparação com dados de **Excel** (S/Tributos, ICMS, PIS, COFINS, Volume)  
- ✅ Identificação automática de divergências  
- 🟢 Status claro: `OK`, `ERRO VALOR`, `ERRO VOLUME`, `ERRO PARSE`  
- 📄 Geração de **relatório XLSX formatado**  
- 🖥️ Interface gráfica com **Tkinter**  

---

## 🖼️ Interface

- Botão **AUDITAR XMLs** em destaque (usável em telas pequenas)
- Status em tempo real (quantidade de XMLs encontrados)
- Fluxo guiado: Pasta → Empresas → Excel → Auditoria

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.13+**
- **Tkinter** – interface gráfica
- **Pandas** – manipulação de dados
- **OpenPyXL** – geração do Excel
- **Pytest** – testes automatizados
- **PyInstaller** – empacotamento em `.exe`

---

## 📂 Estrutura do Projeto

```text
auditoria-xml-excel/
│ app.py
│ requirements.txt
│ README.md
│ .gitignore
│
├── auditoria/
│   ├── __init__.py
│   ├── gui.py
│   ├── audit.py
│   ├── excel_loader.py
│   ├── xml_parser.py
│   ├── report.py
│   └── utils.py
│
└── tests/
    ├── test_utils.py
    ├── test_xml_parser.py
    └── test_excel_loader.py
▶️ Como Executar o Projeto
1️⃣ Clonar o repositório
git clone https://github.com/jleandromorais/auditoria-xml-excel.git
cd auditoria-xml-excel
2️⃣ Criar ambiente virtual
python -m venv .venv
Ativar no Windows:

.venv\Scripts\activate
3️⃣ Instalar dependências
pip install -r requirements.txt
4️⃣ Executar a aplicação
python app.py
🧪 Rodar Testes
pytest
ou

python -m pytest
📦 Gerar Executável (.exe)
pip install pyinstaller
pyinstaller --onefile --windowed app.py --name AuditoriaXML
O executável será gerado em:

dist/AuditoriaXML.exe
📌 Regras de Auditoria
📄 NF-e: tolerância de R$ 5,00

🚚 CT-e: tolerância de R$ 50,00

📦 Volume: tolerância de 1 unidade

CT-e sem PIS/COFINS no XML usa valores do Excel como fallback

🎯 Objetivo do Projeto
Este projeto foi desenvolvido com foco em:

praticar arquitetura modular em Python

manipulação de dados fiscais reais

criação de aplicação desktop

boas práticas para nível júnior

👤 Autor
José Leandro de Morais Alves Luz
GitHub: @jleandromorais

📄 Licença
Projeto open-source para fins educacionais e profissionais.


