#!/usr/bin/env python3
"""Generador de Reportes — Estudiantado Santa Mariana de Jesús.

Punto de entrada. La lógica está dividida en módulos:

    config.py         → rutas
    data_loader.py    → lectura del Excel
    docx_helpers.py   → utilidades comunes de docx
    acta_examen.py    → generación del Acta de Examen
    analitico.py      → generación del Analítico
    ui_theme.py       → paleta y estilos ttk
    ui_app.py         → ventana principal
    ui_viewer.py      → ventana «Reportes generados»
"""
import tkinter as tk

from ui_app import ReportGeneratorApp


def main():
    root = tk.Tk()
    ReportGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
