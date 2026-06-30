"""
gui_fiche_ppi.py
Interface graphique pour la génération de fiches PPI.
Usage: python gui_fiche_ppi.py
"""

import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from .fiche_ppi import (
    build_fiche,
    derive_output,
    find_pairs,
    generate_one,
    stringify,
)
from .format_excel import format_ppi_bold


# ── Couleurs ───────────────────────────────────────────────────────────────────

BG        = "#1e1e2e"
BG_PANEL  = "#2a2a3d"
ACCENT    = "#7c6af7"
ACCENT_H  = "#9d8fff"
TEXT      = "#e0e0f0"
TEXT_DIM  = "#888aaa"
SUCCESS   = "#4caf87"
ERROR     = "#e05c6a"
BORDER    = "#3a3a55"
FONT_MONO = ("Consolas", 10)
FONT_UI   = ("Segoe UI", 10)
FONT_H    = ("Segoe UI", 12, "bold")


# ── Widget helpers ─────────────────────────────────────────────────────────────

def styled_btn(parent, text, command, primary=False, **kw):
    return tk.Button(
        parent, text=text, command=command,
        bg=ACCENT if primary else BG,
        fg="white" if primary else ACCENT,
        activebackground=ACCENT_H, activeforeground="white",
        relief="flat", font=("Segoe UI", 10, "bold") if primary else FONT_UI,
        cursor="hand2", bd=0,
        highlightthickness=1, highlightbackground=BORDER,
        padx=14, pady=6,
        **kw
    )


def file_row(parent, label: str, row: int,
             filetypes=None, save=False) -> tk.StringVar:
    """Label + Entry + bouton Parcourir. Retourne la StringVar."""
    if filetypes is None:
        filetypes = [("Excel", "*.xlsx *.xls"), ("Tous", "*.*")]

    var = tk.StringVar()
    tk.Label(parent, text=label, bg=BG_PANEL, fg=TEXT_DIM,
             font=FONT_UI, anchor="w").grid(
        row=row, column=0, sticky="w", padx=(16, 8), pady=(10, 2))

    tk.Entry(
        parent, textvariable=var, bg=BG, fg=TEXT,
        insertbackground=TEXT, relief="flat", font=FONT_MONO, bd=0,
        highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
        width=50
    ).grid(row=row+1, column=0, sticky="ew", padx=(16, 4), pady=(0, 6), ipady=5)

    def browse():
        if save:
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=filetypes)
        else:
            path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            var.set(path)

    styled_btn(parent, "Parcourir", browse).grid(
        row=row+1, column=1, padx=(0, 16), pady=(0, 6), sticky="w")

    return var


# ── App ────────────────────────────────────────────────────────────────────────

class FichePPIApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Générateur de Fiches PPI — LIDILEM · ANR PREFAB")
        self.configure(bg=BG)
        self.resizable(False, False)
        self._build_ui()

    # ── UI ─────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # En-tête
        hdr = tk.Frame(self, bg=ACCENT, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Générateur de Fiches PPI",
                 bg=ACCENT, fg="white",
                 font=("Segoe UI", 14, "bold")).pack()
        tk.Label(hdr, text="LIDILEM · ANR PREFAB",
                 bg=ACCENT, fg="#d0caff",
                 font=("Segoe UI", 9)).pack()

        # Notebook
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("TNotebook",        background=BG,       borderwidth=0)
        style.configure("TNotebook.Tab",    background=BG_PANEL, foreground=TEXT_DIM,
                        font=FONT_UI, padding=[16, 6])
        style.map("TNotebook.Tab",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", "white")])
        style.configure("TProgressbar",
                        troughcolor=BG_PANEL, background=ACCENT, thickness=4)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=0, pady=0)

        self._build_simple_tab(nb)
        self._build_batch_tab(nb)

        # Log partagé
        log_frame = tk.Frame(self, bg=BG)
        log_frame.pack(fill="both", expand=True, padx=20, pady=(6, 16))

        self.log = tk.Text(
            log_frame, height=8, bg=BG_PANEL, fg=TEXT,
            font=FONT_MONO, relief="flat", bd=0,
            state="disabled", wrap="word",
            highlightthickness=1, highlightbackground=BORDER,
        )
        self.log.pack(fill="both", expand=True)
        self.log.tag_config("ok",   foreground=SUCCESS)
        self.log.tag_config("err",  foreground=ERROR)
        self.log.tag_config("info", foreground=TEXT_DIM)

        # Barre de progression partagée
        self.progress = ttk.Progressbar(self, mode="indeterminate", length=500)
        self.progress.pack(padx=20, pady=(0, 10))

    # ── Onglet Simple ──────────────────────────────────────────────────────────

    def _build_simple_tab(self, nb):
        tab = tk.Frame(nb, bg=BG_PANEL)
        nb.add(tab, text="  Simple  ")
        tab.columnconfigure(0, weight=1)

        tk.Label(tab, text="Fichiers d'entrée", bg=BG_PANEL, fg=TEXT,
                 font=FONT_H, anchor="w").grid(
            row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(14, 0))

        self.var_oral  = file_row(tab, "Fichier oral (.xlsx)",  row=1)
        self.var_ecrit = file_row(tab, "Fichier écrit (.xlsx)", row=3)

        tk.Label(tab, text="Fichier de sortie", bg=BG_PANEL, fg=TEXT,
                 font=FONT_H, anchor="w").grid(
            row=5, column=0, columnspan=2, sticky="w", padx=16, pady=(12, 0))

        self.var_output = tk.StringVar()
        tk.Entry(
            tab, textvariable=self.var_output, bg=BG, fg=TEXT,
            insertbackground=TEXT, relief="flat", font=FONT_MONO, bd=0,
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
            width=50
        ).grid(row=7, column=0, sticky="ew", padx=(16, 4), pady=(0, 6), ipady=5)

        tk.Label(tab, text="Éditable — pré-rempli depuis le fichier oral",
                 bg=BG_PANEL, fg=TEXT_DIM, font=FONT_UI, anchor="w").grid(
            row=6, column=0, sticky="w", padx=(16, 8), pady=(8, 2))

        def browse_out():
            p = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
            if p:
                self.var_output.set(p)

        styled_btn(tab, "Choisir", browse_out).grid(
            row=7, column=1, padx=(0, 16), pady=(0, 6), sticky="w")

        def _update_out(*_):
            p = self.var_oral.get().strip() or self.var_ecrit.get().strip()
            if p:
                self.var_output.set(derive_output(p))

        self.var_oral.trace_add("write",  _update_out)
        self.var_ecrit.trace_add("write", _update_out)

        styled_btn(tab, "⚡  Générer la fiche", self._run_simple,
                   primary=True).grid(
            row=8, column=0, columnspan=2, pady=16)

    # ── Onglet Batch ───────────────────────────────────────────────────────────

    def _build_batch_tab(self, nb):
        tab = tk.Frame(nb, bg=BG_PANEL)
        nb.add(tab, text="  Batch  ")
        tab.columnconfigure(0, weight=1)

        tk.Label(tab, text="Dossier source", bg=BG_PANEL, fg=TEXT,
                 font=FONT_H, anchor="w").grid(
            row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(14, 0))

        tk.Label(tab,
                 text="Doit contenir des paires *_Or.xlsx / *_Ph.xlsx",
                 bg=BG_PANEL, fg=TEXT_DIM, font=FONT_UI, anchor="w").grid(
            row=1, column=0, columnspan=2, sticky="w", padx=16, pady=(4, 2))

        self.var_folder = tk.StringVar()
        tk.Entry(
            tab, textvariable=self.var_folder, bg=BG, fg=TEXT,
            insertbackground=TEXT, relief="flat", font=FONT_MONO, bd=0,
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
            width=50
        ).grid(row=2, column=0, sticky="ew", padx=(16, 4), pady=(0, 6), ipady=5)

        def browse_folder():
            p = filedialog.askdirectory()
            if p:
                self.var_folder.set(p)
                self._refresh_pairs(p)

        styled_btn(tab, "Parcourir", browse_folder).grid(
            row=2, column=1, padx=(0, 16), pady=(0, 6), sticky="w")

        # Liste des paires détectées
        tk.Label(tab, text="Paires détectées", bg=BG_PANEL, fg=TEXT,
                 font=FONT_H, anchor="w").grid(
            row=3, column=0, columnspan=2, sticky="w", padx=16, pady=(12, 2))

        list_frame = tk.Frame(tab, bg=BG_PANEL)
        list_frame.grid(row=4, column=0, columnspan=2, sticky="ew",
                        padx=16, pady=(0, 8))
        list_frame.columnconfigure(0, weight=1)

        self.pairs_list = tk.Listbox(
            list_frame, bg=BG, fg=TEXT, font=FONT_MONO,
            selectbackground=ACCENT, selectforeground="white",
            relief="flat", bd=0,
            highlightthickness=1, highlightbackground=BORDER,
            height=6, activestyle="none"
        )
        self.pairs_list.pack(fill="both", expand=True)

        self._pairs: list[tuple[str, str]] = []

        styled_btn(tab, "⚡  Générer toutes les fiches", self._run_batch,
                   primary=True).grid(
            row=5, column=0, columnspan=2, pady=16)

    def _refresh_pairs(self, folder: str):
        self.pairs_list.delete(0, "end")
        self._pairs = find_pairs(folder)
        if not self._pairs:
            self.pairs_list.insert("end", "  Aucune paire trouvée.")
            return
        for oral, ecrit in self._pairs:
            label = re.sub(r'_Or.*$', '', os.path.basename(oral), flags=re.IGNORECASE)
            self.pairs_list.insert("end", f"  {label}")
        self._log(f"{len(self._pairs)} paire(s) détectée(s) dans {folder}")

    # ── Logging ────────────────────────────────────────────────────────────────

    def _log(self, msg: str, tag: str = "info"):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _lock(self):
        self.progress.start(12)

    def _unlock(self):
        self.progress.stop()

    # ── Run simple ─────────────────────────────────────────────────────────────

    def _run_simple(self):
        oral  = self.var_oral.get().strip()
        ecrit = self.var_ecrit.get().strip()
        if not oral:
            messagebox.showwarning("Fichier manquant", "Sélectionne le fichier oral.")
            return
        if not ecrit:
            messagebox.showwarning("Fichier manquant", "Sélectionne le fichier écrit.")
            return
        output = self.var_output.get().strip() or derive_output(oral)
        self._lock()
        self._log(f"Traitement de {os.path.basename(oral)} …")
        threading.Thread(
            target=self._worker_simple,
            args=(oral, ecrit, output),
            daemon=True
        ).start()

    def _worker_simple(self, oral, ecrit, output):
        try:
            out, warnings = generate_one(oral, ecrit)
            for w in warnings:
                self.after(0, self._log, w, "err")
            self.after(0, self._unlock)
            self.after(0, self._log, f"✓ {out}", "ok")
            self.after(0, messagebox.showinfo, "Succès", f"Fiche exportée :\n{out}")
        except Exception as e:
            self.after(0, self._unlock)
            self.after(0, self._log, f"✗ {e}", "err")
            self.after(0, messagebox.showerror, "Erreur", str(e))

    # ── Run batch ──────────────────────────────────────────────────────────────

    def _run_batch(self):
        if not self._pairs:
            messagebox.showwarning("Aucune paire",
                                   "Sélectionne un dossier contenant des paires _Or/_Ph.")
            return
        self._lock()
        self._log(f"Batch : {len(self._pairs)} paire(s) à traiter …")
        threading.Thread(target=self._worker_batch, daemon=True).start()

    def _worker_batch(self):
        ok = 0
        errors = []
        for oral, ecrit in self._pairs:
            name = os.path.basename(oral)
            try:
                out, warnings = generate_one(oral, ecrit)
                for w in warnings:
                    self.after(0, self._log, w, "err")
                ok += 1
                self.after(0, self._log, f"  ✓ {os.path.basename(out)}", "ok")
            except Exception as e:
                errors.append((name, str(e)))
                self.after(0, self._log, f"  ✗ {name} : {e}", "err")

        self.after(0, self._unlock)
        summary = f"{ok}/{len(self._pairs)} fiche(s) générée(s)."
        if errors:
            summary += f"\n{len(errors)} erreur(s)."
        self.after(0, self._log, summary, "ok" if not errors else "err")
        self.after(0, messagebox.showinfo, "Batch terminé", summary)


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    """Launch the PPI Fiche Generator GUI."""
    app = FichePPIApp()
    app.mainloop()


if __name__ == "__main__":
    main()
