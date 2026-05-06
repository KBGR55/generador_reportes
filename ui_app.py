"""Ventana principal del Generador de Reportes."""
import os
import sys
import datetime
import subprocess
import tkinter as tk
import traceback
from tkinter import ttk, messagebox

from data_loader import load_excel_data
from acta_examen import generate_acta_examen
from analitico import generate_analitico
from ui_theme import apply_theme, BG, CARD, BORDER, ACCENT_BAR
from ui_viewer import open_reports_viewer


class ReportGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title(
            "Generador de Reportes - Estudiantado Santa Mariana de Jesús")
        self.root.geometry("760x640")
        self.root.minsize(720, 600)
        self.root.configure(bg=BG)

        try:
            self.notas, self.datos, self.profesores, self.plan = load_excel_data()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer esquema.xlsx:\n{e}")
            sys.exit(1)

        # Listas únicas para los combos
        self.materias_unicas = sorted({n["materia"] for n in self.notas
                                       if n["materia"]})
        self.alumnas_unicas = sorted({n["alumna"] for n in self.notas
                                      if n["alumna"]})
        self.profesores_unicos = sorted({n["profesor"] for n in self.notas
                                         if n["profesor"]})

        for d in self.datos:
            if d["nombre_religioso"] and d["nombre_religioso"] not in self.alumnas_unicas:
                self.alumnas_unicas.append(d["nombre_religioso"])

        self.materias_unicas = sorted(set(
            list(self.materias_unicas) +
            [m["materia"] for m in self.plan if m["materia"]]
        ))

        apply_theme()
        self.build_ui()

    # ------------------------------------------------------------------ UI
    def build_ui(self):
        # Barra de acento
        tk.Frame(self.root, bg=ACCENT_BAR, height=4).pack(fill="x", side="top")

        # Header con título y botón "Ver reportes"
        header = ttk.Frame(self.root)
        header.pack(fill="x", padx=28, pady=(18, 4))

        header_left = ttk.Frame(header)
        header_left.pack(side="left", anchor="w")
        ttk.Label(header_left, text="Generador de Reportes",
                  style="Title.TLabel").pack(anchor="w")
        ttk.Label(header_left, text="Estudiantado Santa Mariana de Jesús",
                  style="Subtitle.TLabel").pack(anchor="w", pady=(2, 0))

        ttk.Button(header, text="Ver reportes generados",
                   style="Secondary.TButton", cursor="hand2",
                   command=lambda: open_reports_viewer(self.root)
                   ).pack(side="right", anchor="ne")

        tk.Frame(self.root, height=1, bg=BORDER).pack(
            fill="x", padx=28, pady=(12, 16))

        # Tarjeta selector de tipo
        type_card = self._make_card()
        type_card.pack(padx=28, pady=(0, 14), fill="x")

        type_inner = tk.Frame(type_card, bg=CARD)
        type_inner.pack(padx=20, pady=16, fill="x")

        ttk.Label(type_inner, text="Tipo de reporte",
                  style="Section.TLabel").pack(anchor="w", pady=(0, 10))

        self.report_type = tk.StringVar(value="acta")
        radio_row = tk.Frame(type_inner, bg=CARD)
        radio_row.pack(anchor="w")
        ttk.Radiobutton(radio_row, text="Acta de Examen",
                        variable=self.report_type, value="acta",
                        style="Type.TRadiobutton",
                        command=self.toggle_panels
                        ).pack(side="left", padx=(0, 36))
        ttk.Radiobutton(radio_row, text="Analítico de Estudios",
                        variable=self.report_type, value="analitico",
                        style="Type.TRadiobutton",
                        command=self.toggle_panels).pack(side="left")

        self._build_acta_card()
        self._build_analitico_card()

        self.status_var = tk.StringVar(value="")
        self.status_label = ttk.Label(self.root, textvariable=self.status_var,
                                      style="Status.TLabel")

        self.generate_btn = ttk.Button(self.root, text="Generar Reporte",
                                       style="Primary.TButton",
                                       command=self.generate_report,
                                       cursor="hand2")

        # Footer
        self.info_var = tk.StringVar(
            value=f"Datos cargados: {len(self.notas)} notas  ·  "
                  f"{len(self.datos)} alumnas  ·  "
                  f"{len(self.plan)} materias en plan"
        )
        info_bar = ttk.Frame(self.root)
        info_bar.pack(side="bottom", fill="x", padx=28, pady=(0, 14))
        tk.Frame(info_bar, height=1, bg=BORDER).pack(fill="x", pady=(0, 8))
        ttk.Label(info_bar, textvariable=self.info_var,
                  style="Info.TLabel").pack(anchor="w")

        self.toggle_panels()

    def _make_card(self):
        return tk.Frame(self.root, bg=CARD,
                        highlightbackground=BORDER, highlightthickness=1, bd=0)

    def _build_acta_card(self):
        self.acta_frame = self._make_card()
        inner = tk.Frame(self.acta_frame, bg=CARD)
        inner.pack(padx=20, pady=16, fill="x", expand=True)

        ttk.Label(inner, text="Datos del Acta de Examen",
                  style="Section.TLabel").pack(anchor="w", pady=(0, 14))

        form = tk.Frame(inner, bg=CARD)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        def add_field(row, label, widget):
            ttk.Label(form, text=label, style="Form.TLabel").grid(
                row=row, column=0, sticky="w", pady=7, padx=(0, 14))
            return widget

        self.acta_materia = ttk.Combobox(form, values=self.materias_unicas)
        add_field(0, "Materia", self.acta_materia)
        self.acta_materia.grid(row=0, column=1, sticky="ew", pady=7)

        self.acta_profesor = ttk.Combobox(
            form, values=self.profesores_unicos + self.profesores)
        add_field(1, "Profesor", self.acta_profesor)
        self.acta_profesor.grid(row=1, column=1, sticky="ew", pady=7)

        self.acta_anio = ttk.Combobox(
            form, values=["2023", "2024", "2025", "2026"], width=14)
        add_field(2, "Año", self.acta_anio)
        self.acta_anio.grid(row=2, column=1, sticky="w", pady=7)
        self.acta_anio.set("2025")

        self.acta_semestre = ttk.Combobox(form, values=["I", "II"], width=14)
        add_field(3, "Semestre", self.acta_semestre)
        self.acta_semestre.grid(row=3, column=1, sticky="w", pady=7)
        self.acta_semestre.set("I")

        self.acta_fecha = ttk.Entry(form, width=22)
        add_field(4, "Fecha (DD/MM/AAAA)", self.acta_fecha)
        self.acta_fecha.grid(row=4, column=1, sticky="w", pady=7)
        self.acta_fecha.insert(0, datetime.date.today().strftime("%d/%m/%Y"))

    def _build_analitico_card(self):
        self.analitico_frame = self._make_card()
        inner = tk.Frame(self.analitico_frame, bg=CARD)
        inner.pack(padx=20, pady=16, fill="x", expand=True)

        ttk.Label(inner, text="Datos del Analítico de Estudios",
                  style="Section.TLabel").pack(anchor="w", pady=(0, 14))

        form = tk.Frame(inner, bg=CARD)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Alumna", style="Form.TLabel").grid(
            row=0, column=0, sticky="w", pady=7, padx=(0, 14))
        self.analitico_alumna = ttk.Combobox(form, values=self.alumnas_unicas)
        self.analitico_alumna.grid(row=0, column=1, sticky="ew", pady=7)
        if self.alumnas_unicas:
            self.analitico_alumna.set(self.alumnas_unicas[0])

    def toggle_panels(self):
        self.acta_frame.pack_forget()
        self.analitico_frame.pack_forget()
        self.status_label.pack_forget()
        self.generate_btn.pack_forget()

        if self.report_type.get() == "acta":
            self.acta_frame.pack(padx=28, pady=(0, 14), fill="x")
        else:
            self.analitico_frame.pack(padx=28, pady=(0, 14), fill="x")

        self.status_label.pack(padx=28, pady=(4, 0))
        self.generate_btn.pack(pady=(14, 8))

    # -------------------------------------------------------------- Acciones
    def generate_report(self):
        try:
            # Recargamos por si el Excel cambió mientras la app estaba abierta
            self.notas, self.datos, self.profesores, self.plan = load_excel_data()

            if self.report_type.get() == "acta":
                filepath = self._do_generate_acta()
            else:
                filepath = self._do_generate_analitico()

            self.status_var.set(f"Reporte generado: {os.path.basename(filepath)}")
            messagebox.showinfo(
                "Éxito", f"Reporte generado exitosamente!\n\n{filepath}")

            self._open_with_system(filepath)

        except Exception as e:
            messagebox.showerror("Error", f"Error al generar reporte:\n{e}")
            traceback.print_exc()

    def _do_generate_acta(self):
        materia = self.acta_materia.get().strip()
        profesor = self.acta_profesor.get().strip()
        anio = self.acta_anio.get().strip()
        semestre = self.acta_semestre.get().strip()
        fecha_str = self.acta_fecha.get().strip()

        if not materia:
            raise ValueError("Seleccione una materia")
        if not profesor:
            raise ValueError("Seleccione un profesor")

        try:
            parts = fecha_str.split("/")
            fecha = datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
        except Exception:
            fecha = fecha_str

        materia_norm = materia.strip().lower()
        notas_filtradas = [n for n in self.notas
                           if n["materia"].strip().lower() == materia_norm]

        return generate_acta_examen(materia, profesor, anio, semestre, fecha,
                                    notas_filtradas, self.datos)

    def _do_generate_analitico(self):
        alumna = self.analitico_alumna.get().strip()
        if not alumna:
            raise ValueError("Seleccione una alumna")
        return generate_analitico(alumna, self.notas, self.datos, self.plan)

    @staticmethod
    def _open_with_system(filepath):
        if sys.platform == "linux":
            subprocess.Popen(["xdg-open", filepath])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", filepath])
        elif sys.platform == "win32":
            os.startfile(filepath)
