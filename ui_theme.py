"""Paleta de colores y configuración de estilos ttk."""
from tkinter import ttk


# Paleta (inspirada en Tailwind)
BG = "#f1f5f9"          # slate-100
CARD = "#ffffff"
PRIMARY = "#4f46e5"     # indigo-600
PRIMARY_DARK = "#4338ca"
PRIMARY_SOFT = "#eef2ff"
TEXT = "#0f172a"        # slate-900
MUTED = "#64748b"       # slate-500
BORDER = "#e2e8f0"      # slate-200
SUCCESS = "#059669"     # emerald-600
ACCENT_BAR = "#6366f1"  # indigo-500
DISABLED = "#94a3b8"
HEADING_BG = "#f8fafc"


def apply_theme():
    """Configura todos los estilos ttk usados por la aplicación.

    Idempotente — se puede llamar más de una vez sin efectos secundarios.
    """
    style = ttk.Style()
    style.theme_use("clam")

    # Base
    style.configure(".", background=BG, foreground=TEXT,
                    font=("Helvetica", 10))
    style.configure("TFrame", background=BG)
    style.configure("TLabel", background=BG, foreground=TEXT)

    # Encabezados
    style.configure("Title.TLabel", font=("Helvetica", 19, "bold"),
                    background=BG, foreground=TEXT)
    style.configure("Subtitle.TLabel", font=("Helvetica", 10),
                    background=BG, foreground=MUTED)
    style.configure("Info.TLabel", font=("Helvetica", 9),
                    background=BG, foreground=MUTED)

    # Sobre tarjeta blanca
    style.configure("Section.TLabel", font=("Helvetica", 12, "bold"),
                    background=CARD, foreground=TEXT)
    style.configure("Form.TLabel", font=("Helvetica", 10),
                    background=CARD, foreground=TEXT)
    style.configure("Card.TLabel", background=CARD, foreground=TEXT)

    # Inputs
    style.configure("TEntry", fieldbackground=CARD, bordercolor=BORDER,
                    lightcolor=BORDER, darkcolor=BORDER, padding=6)
    style.configure("TCombobox", fieldbackground=CARD, bordercolor=BORDER,
                    lightcolor=BORDER, darkcolor=BORDER, padding=4,
                    arrowsize=14)
    style.map("TCombobox",
              fieldbackground=[("readonly", CARD), ("disabled", "#f8fafc")],
              bordercolor=[("focus", PRIMARY)])
    style.map("TEntry", bordercolor=[("focus", PRIMARY)])

    # Radiobutton sobre tarjeta
    style.configure("Type.TRadiobutton", background=CARD, foreground=TEXT,
                    font=("Helvetica", 11), padding=4)
    style.map("Type.TRadiobutton",
              background=[("active", CARD)],
              foreground=[("selected", PRIMARY)])

    # Botones
    style.configure("Primary.TButton",
                    background=PRIMARY, foreground="white",
                    font=("Helvetica", 12, "bold"),
                    padding=(28, 12), borderwidth=0, relief="flat",
                    focuscolor=PRIMARY)
    style.map("Primary.TButton",
              background=[("active", PRIMARY_DARK), ("pressed", PRIMARY_DARK),
                          ("disabled", DISABLED)],
              foreground=[("disabled", "white")])

    style.configure("Secondary.TButton",
                    background=CARD, foreground=PRIMARY,
                    font=("Helvetica", 10, "bold"),
                    padding=(14, 8), borderwidth=1, relief="flat",
                    bordercolor=BORDER, focuscolor=CARD)
    style.map("Secondary.TButton",
              background=[("active", PRIMARY_SOFT)],
              bordercolor=[("active", PRIMARY)])

    # Treeview
    style.configure("Treeview", background=CARD, fieldbackground=CARD,
                    foreground=TEXT, rowheight=26, borderwidth=0,
                    font=("Helvetica", 10))
    style.configure("Treeview.Heading", background=HEADING_BG,
                    foreground=TEXT, font=("Helvetica", 10, "bold"),
                    relief="flat", padding=6)
    style.map("Treeview", background=[("selected", PRIMARY)],
              foreground=[("selected", "white")])

    # Estado
    style.configure("Status.TLabel", background=BG, foreground=SUCCESS,
                    font=("Helvetica", 10, "bold"))
