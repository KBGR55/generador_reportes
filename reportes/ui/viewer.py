"""Ventana «Reportes generados»."""
import os
import sys
import datetime
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox

from reportes.config import OUTPUT_DIR
from reportes.ui.theme import BG, CARD, BORDER, MUTED, ACCENT_BAR


def _open_path(path):
    """Abre un archivo o carpeta con el visor del sistema."""
    try:
        if sys.platform == "linux":
            subprocess.Popen(["xdg-open", path])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif sys.platform == "win32":
            os.startfile(path)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir:\n{e}")


def _classify(fname):
    if fname.startswith("Acta_"):
        return "Acta de Examen"
    if fname.startswith("Analitico"):
        return "Analítico"
    return "Otro"


def open_reports_viewer(parent):
    """Abre la ventana modal con la lista de reportes generados."""
    win = tk.Toplevel(parent)
    win.title("Reportes generados")
    win.geometry("760x480")
    win.minsize(700, 420)
    win.configure(bg=BG)
    win.transient(parent)

    tk.Frame(win, bg=ACCENT_BAR, height=4).pack(fill="x", side="top")

    head = ttk.Frame(win)
    head.pack(fill="x", padx=24, pady=(16, 6))
    ttk.Label(head, text="Reportes generados",
              style="Title.TLabel").pack(anchor="w")
    ttk.Label(head, text=OUTPUT_DIR,
              style="Subtitle.TLabel").pack(anchor="w", pady=(2, 0))

    tk.Frame(win, height=1, bg=BORDER).pack(fill="x", padx=24, pady=(10, 12))

    card = tk.Frame(win, bg=CARD,
                    highlightbackground=BORDER, highlightthickness=1, bd=0)
    card.pack(padx=24, pady=(0, 12), fill="both", expand=True)

    list_wrap = tk.Frame(card, bg=CARD)
    list_wrap.pack(padx=12, pady=12, fill="both", expand=True)

    cols = ("nombre", "tipo", "fecha", "tamano")
    tree = ttk.Treeview(list_wrap, columns=cols, show="headings", height=10)
    tree.heading("nombre", text="Nombre")
    tree.heading("tipo", text="Tipo")
    tree.heading("fecha", text="Modificado")
    tree.heading("tamano", text="Tamaño")
    tree.column("nombre", width=360, anchor="w")
    tree.column("tipo", width=110, anchor="w")
    tree.column("fecha", width=140, anchor="w")
    tree.column("tamano", width=80, anchor="e")

    sb = ttk.Scrollbar(list_wrap, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    empty_var = tk.StringVar(value="")
    empty_lbl = ttk.Label(card, textvariable=empty_var,
                          background=CARD, foreground=MUTED,
                          font=("Helvetica", 10, "italic"))

    def populate():
        for item in tree.get_children():
            tree.delete(item)
        empty_lbl.pack_forget()

        if not os.path.isdir(OUTPUT_DIR):
            empty_var.set("Aún no se han generado reportes.")
            empty_lbl.pack(pady=(0, 14))
            return

        archivos = [f for f in os.listdir(OUTPUT_DIR)
                    if os.path.isfile(os.path.join(OUTPUT_DIR, f))]
        archivos.sort(
            key=lambda f: os.path.getmtime(os.path.join(OUTPUT_DIR, f)),
            reverse=True,
        )

        if not archivos:
            empty_var.set("Aún no se han generado reportes.")
            empty_lbl.pack(pady=(0, 14))
            return

        for fname in archivos:
            fpath = os.path.join(OUTPUT_DIR, fname)
            mtime = datetime.datetime.fromtimestamp(os.path.getmtime(fpath))
            size_kb = os.path.getsize(fpath) / 1024
            tree.insert("", "end", values=(
                fname, _classify(fname),
                mtime.strftime("%d/%m/%Y %H:%M"),
                f"{size_kb:.0f} KB",
            ))

    def open_selected(_event=None):
        sel = tree.selection()
        if not sel:
            return
        fname = tree.item(sel[0], "values")[0]
        _open_path(os.path.join(OUTPUT_DIR, fname))

    def open_folder():
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        _open_path(OUTPUT_DIR)

    def delete_selected():
        sel = tree.selection()
        if not sel:
            return
        fname = tree.item(sel[0], "values")[0]
        if not messagebox.askyesno(
                "Eliminar reporte",
                f"¿Eliminar definitivamente «{fname}»?", parent=win):
            return
        try:
            os.remove(os.path.join(OUTPUT_DIR, fname))
            populate()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo eliminar:\n{e}")

    tree.bind("<Double-1>", open_selected)
    tree.bind("<Return>", open_selected)

    btns = tk.Frame(win, bg=BG)
    btns.pack(fill="x", padx=24, pady=(0, 18))
    ttk.Button(btns, text="Abrir", style="Primary.TButton",
               cursor="hand2", command=open_selected).pack(side="left")
    ttk.Button(btns, text="Abrir carpeta", style="Secondary.TButton",
               cursor="hand2", command=open_folder).pack(side="left", padx=(10, 0))
    ttk.Button(btns, text="Refrescar", style="Secondary.TButton",
               cursor="hand2", command=populate).pack(side="left", padx=(10, 0))
    ttk.Button(btns, text="Eliminar", style="Secondary.TButton",
               cursor="hand2", command=delete_selected).pack(side="left", padx=(10, 0))
    ttk.Button(btns, text="Cerrar", style="Secondary.TButton",
               cursor="hand2", command=win.destroy).pack(side="right")

    populate()
