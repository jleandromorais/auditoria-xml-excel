from pathlib import Path

# Carrega variáveis do .env (opcional)
try:
    from dotenv import load_dotenv

    # tenta primeiro o .env ao lado do app.py (auditoria-xml-excel/.env)
    load_dotenv(Path(__file__).resolve().with_name(".env"))
    # e depois tenta o .env no nível acima (raiz do workspace, se você criou lá)
    load_dotenv(Path(__file__).resolve().parent.parent / ".env")
except Exception:
    pass

from auditoria.gui import App

if __name__ == "__main__":
    App().mainloop()
