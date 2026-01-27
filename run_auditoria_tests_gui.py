from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

CONFIG_FILE = Path(__file__).parent / "test_config.json"


def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_config(cfg: dict) -> None:
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")


def pick_file(title: str, filetypes: list[tuple[str, str]]) -> str:
    p = filedialog.askopenfilename(title=title, filetypes=filetypes)
    return p or ""


def run_pytest(mes: str) -> tuple[int, str]:
    """
    Executa pytest e retorna (exit_code, output).
    """
    cmd = [
        sys.executable,
        "-m",
        "pytest",
        "-q",
        f"--mes={mes}",
        "tests/test_real_zip_excel_e2e.py",
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True, cwd=str(Path(__file__).parent))
    out = (proc.stdout or "") + ("\n" + proc.stderr if proc.stderr else "")
    return proc.returncode, out


def main() -> None:
    cfg = load_config()

    root = tk.Tk()
    root.title("Auditoria - Teste Automático (1 clique)")
    root.geometry("720x320")

    mes_var = tk.StringVar(value=cfg.get("mes", "OUT/2025"))
    excel_var = tk.StringVar(value=cfg.get("excel_path", ""))
    zip_var = tk.StringVar(value=cfg.get("zip_path", ""))

    def choose_excel():
        p = pick_file("Escolha o Excel base", [("Excel", "*.xlsx *.xls")])
        if p:
            excel_var.set(p)
            cfg["excel_path"] = p
            save_config(cfg)

    def choose_zip():
        p = pick_file("Escolha o ZIP com XMLs", [("ZIP", "*.zip")])
        if p:
            zip_var.set(p)
            cfg["zip_path"] = p
            save_config(cfg)

    def save_mes(*_):
        cfg["mes"] = mes_var.get().strip()
        save_config(cfg)

    def run_tests():
        mes = mes_var.get().strip()
        excel_path = excel_var.get().strip()
        zip_path = zip_var.get().strip()

        if not mes:
            messagebox.showerror("Erro", "Informe o mês (ex.: OUT/2025).")
            return
        if not excel_path or not Path(excel_path).exists():
            messagebox.showerror("Erro", "Selecione um Excel válido.")
            return
        if not zip_path or not Path(zip_path).exists():
            messagebox.showerror("Erro", "Selecione um ZIP válido.")
            return

        # salva config
        cfg["mes"] = mes
        cfg["excel_path"] = excel_path
        cfg["zip_path"] = zip_path
        save_config(cfg)

        # roda pytest
        code, out = run_pytest(mes)

        if code == 0:
            messagebox.showinfo("PASSOU ✅", "Teste passou! Excel e XML batem para o mês selecionado.")
        else:
            # Mostra output do pytest com as NFs problemáticas
            messagebox.showerror("FALHOU ❌", out[-3500:] if len(out) > 3500 else out)

    # UI
    frm = tk.Frame(root, padx=12, pady=12)
    frm.pack(fill="both", expand=True)

    tk.Label(frm, text="Mês alvo (igual ao Excel, ex.: OUT/2025)").grid(row=0, column=0, sticky="w")
    ent_mes = tk.Entry(frm, textvariable=mes_var, width=30)
    ent_mes.grid(row=0, column=1, sticky="w", padx=6)
    ent_mes.bind("<FocusOut>", save_mes)

    tk.Label(frm, text="Excel base").grid(row=1, column=0, sticky="w", pady=(12, 0))
    tk.Entry(frm, textvariable=excel_var, width=70).grid(row=1, column=1, sticky="w", padx=6, pady=(12, 0))
    tk.Button(frm, text="Escolher Excel", command=choose_excel).grid(row=1, column=2, padx=6, pady=(12, 0))

    tk.Label(frm, text="ZIP com XMLs").grid(row=2, column=0, sticky="w", pady=(12, 0))
    tk.Entry(frm, textvariable=zip_var, width=70).grid(row=2, column=1, sticky="w", padx=6, pady=(12, 0))
    tk.Button(frm, text="Escolher ZIP", command=choose_zip).grid(row=2, column=2, padx=6, pady=(12, 0))

    tk.Button(frm, text="▶ Rodar Teste Automático", command=run_tests, height=2).grid(
        row=3, column=1, sticky="w", pady=(20, 0)
    )

    tk.Label(
        frm,
        text="Dica: ele salva suas escolhas em test_config.json. Depois é só 1 clique.",
        fg="#444",
    ).grid(row=4, column=1, sticky="w", pady=(14, 0))

    root.mainloop()


if __name__ == "__main__":
    main()
