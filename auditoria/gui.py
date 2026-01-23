import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from datetime import datetime

from .audit import auditar_pasta_pai, coletar_xmls_por_empresas


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Auditoria XML (pasta PAI -> empresas -> XMLs -> comparação Excel)")
        self.geometry("900x600")
        self.minsize(760, 520)

        self.pasta_pai: Path | None = None
        self.excel_path: str | None = None
        self.destino_relatorio: Path | None = None

        # nome_empresa -> (Path, BooleanVar)
        self.empresas_vars: dict[str, tuple[Path, tk.BooleanVar]] = {}

        # ======== TOPO ========
        topo = tk.Frame(self)
        topo.pack(fill="x", padx=12, pady=10)

        tk.Label(
            topo,
            text="1) Escolha a pasta PAI (onde estão as empresas):",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w")

        linha_pai = tk.Frame(topo)
        linha_pai.pack(fill="x", pady=6)

        tk.Button(linha_pai, text="Escolher pasta PAI", command=self.escolher_pasta_pai).pack(side="left")
        self.lbl_pai = tk.Label(linha_pai, text="(não selecionada)", fg="gray")
        self.lbl_pai.pack(side="left", padx=10)

        # ======== MEIO (EMPRESAS) ========
        mid = tk.Frame(self)
        mid.pack(fill="both", expand=True, padx=12, pady=8)

        tk.Label(
            mid,
            text="2) Selecione as empresas:",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w")

        actions = tk.Frame(mid)
        actions.pack(fill="x", pady=6)

        tk.Button(actions, text="Marcar todas", command=self.marcar_todas).pack(side="left")
        tk.Button(actions, text="Desmarcar todas", command=self.desmarcar_todas).pack(side="left", padx=8)

        # Área com scroll
        self.canvas = tk.Canvas(mid, borderwidth=0)
        self.scroll = tk.Scrollbar(mid, orient="vertical", command=self.canvas.yview)
        self.frame_checks = tk.Frame(self.canvas)

        self.frame_checks.bind(
            "<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.frame_checks, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scroll.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scroll.pack(side="right", fill="y")

        # ======== ABAIXO (EXCEL + DESTINO) ========
        bottom = tk.Frame(self)
        bottom.pack(fill="x", padx=12, pady=10)

        tk.Label(
            bottom,
            text="3) Escolha o Excel:",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w")

        linha_excel = tk.Frame(bottom)
        linha_excel.pack(fill="x", pady=6)

        tk.Button(linha_excel, text="Escolher Excel", command=self.escolher_excel).pack(side="left")
        self.lbl_excel = tk.Label(linha_excel, text="(não selecionado)", fg="gray")
        self.lbl_excel.pack(side="left", padx=10)

        tk.Label(
            bottom,
            text="4) (Opcional) Pasta para salvar o relatório:",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w")

        linha_dest = tk.Frame(bottom)
        linha_dest.pack(fill="x", pady=6)

        tk.Button(linha_dest, text="Escolher pasta", command=self.escolher_destino).pack(side="left")
        self.lbl_dest = tk.Label(linha_dest, text="(Downloads por padrão)", fg="gray")
        self.lbl_dest.pack(side="left", padx=10)

        # ======== RODAPÉ (BOTÃO GRANDE + STATUS) ========
        footer = tk.Frame(self, bg="#f0f0f0")
        footer.pack(fill="x", padx=8, pady=8)

        self.status = tk.Label(
            footer,
            text="Pronto. Nenhum XML carregado ainda.",
            anchor="w",
            bg="#f0f0f0",
            font=("Segoe UI", 10),
        )
        self.status.pack(side="left", fill="x", expand=True, padx=6)

        self.btn_auditar = tk.Button(
            footer,
            text="🔍 AUDITAR XMLs",
            command=self.rodar_auditoria,
            bg="#1f6feb",
            fg="white",
            font=("Segoe UI", 12, "bold"),
            padx=20,
            pady=10,
        )
        self.btn_auditar.pack(side="right", padx=6)

    # ===== Ações =====
    def escolher_pasta_pai(self):
        pasta = filedialog.askdirectory(title="Escolha a pasta PAI (onde estão as empresas)")
        if not pasta:
            return
        self.pasta_pai = Path(pasta)
        self.lbl_pai.config(text=str(self.pasta_pai), fg="black")
        self.carregar_empresas()
        self.atualizar_preview_xmls()

    def carregar_empresas(self):
        # limpa checks antigos
        for widget in self.frame_checks.winfo_children():
            widget.destroy()
        self.empresas_vars.clear()

        if not self.pasta_pai:
            return

        subpastas = [p for p in self.pasta_pai.iterdir() if p.is_dir()]
        subpastas.sort(key=lambda p: p.name.lower())

        if not subpastas:
            tk.Label(self.frame_checks, text="Nenhuma subpasta (empresa) encontrada na pasta PAI.").pack(anchor="w")
            self.status.config(text="Nenhuma empresa encontrada.")
            return

        for p in subpastas:
            var = tk.BooleanVar(value=True)  # marca todas por padrão
            chk = tk.Checkbutton(self.frame_checks, text=p.name, variable=var, command=self.atualizar_preview_xmls)
            chk.pack(anchor="w")
            self.empresas_vars[p.name] = (p, var)

        self.status.config(text=f"{len(subpastas)} empresa(s) carregada(s).")

    def marcar_todas(self):
        for _, var in self.empresas_vars.values():
            var.set(True)
        self.atualizar_preview_xmls()

    def desmarcar_todas(self):
        for _, var in self.empresas_vars.values():
            var.set(False)
        self.atualizar_preview_xmls()

    def escolher_excel(self):
        excel = filedialog.askopenfilename(title="Escolha o Excel", filetypes=[("Excel", "*.xlsx")])
        if not excel:
            return
        self.excel_path = excel
        self.lbl_excel.config(text=str(self.excel_path), fg="black")

    def escolher_destino(self):
        pasta = filedialog.askdirectory(title="Escolha a pasta para salvar o relatório")
        if not pasta:
            return
        self.destino_relatorio = Path(pasta)
        self.lbl_dest.config(text=str(self.destino_relatorio), fg="black")

    def _empresas_selecionadas(self):
        return [p for (p, var) in (v for v in self.empresas_vars.values()) if var.get()]

    def atualizar_preview_xmls(self):
        """
        Mostra no status quantos XMLs seriam encontrados com a seleção atual.
        (Ajuda a dar confiança e evita 'tá travado?')
        """
        if not self.pasta_pai or not self.empresas_vars:
            return

        empresas = self._empresas_selecionadas()
        if not empresas:
            self.status.config(text="Nenhuma empresa selecionada.")
            return

        try:
            xmls = coletar_xmls_por_empresas(self.pasta_pai, empresas)
            self.status.config(text=f"Prévia: {len(xmls)} XML(s) encontrados nas empresas selecionadas.")
        except Exception:
            self.status.config(text="Prévia: não consegui contar os XMLs (mas a auditoria pode rodar).")

    def rodar_auditoria(self):
        if not self.pasta_pai:
            messagebox.showwarning("Atenção", "Escolha a pasta PAI.")
            return
        if not self.excel_path:
            messagebox.showwarning("Atenção", "Escolha o Excel.")
            return

        empresas = self._empresas_selecionadas()
        if not empresas:
            messagebox.showwarning("Atenção", "Selecione pelo menos uma empresa.")
            return

        saida = None
        if self.destino_relatorio:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            saida = str(self.destino_relatorio / f"Auditoria_XML_{ts}.xlsx")

        try:
            # 1) Conta XMLs antes (mostra no status)
            xmls = coletar_xmls_por_empresas(self.pasta_pai, empresas)
            total_xmls = len(xmls)

            self.status.config(text=f"{total_xmls} XML(s) encontrados. Iniciando auditoria...")
            self.btn_auditar.config(state="disabled")
            self.update_idletasks()

            # 2) Roda auditoria
            out = auditar_pasta_pai(self.pasta_pai, empresas, self.excel_path, saida=saida)

            self.status.config(text=f"Concluído! Relatório gerado: {out}")
            messagebox.showinfo("Finalizado", f"Concluído!\n\nRelatório:\n{out}")

        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.status.config(text="Erro ao auditar.")
        finally:
            self.btn_auditar.config(state="normal")
