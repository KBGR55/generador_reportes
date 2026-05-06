#!/usr/bin/env python3
"""
Generador de Reportes - Estudiantado Santa Mariana de Jesús
Genera Actas de Examen y Analíticos de Estudios desde esquema.xlsx
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import datetime
import subprocess
import sys

EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "esquema.xlsx")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_generados")


def load_excel_data():
    """Load all data from the Excel file."""
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

    # Notas
    ws = wb["notas"]
    notas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None and row[1] is None:
            continue
        notas.append({
            "profesor": row[0] or "",
            "alumna": row[1] or "",
            "materia": row[2] or "",
            "anio": row[3],
            "semestre": row[4] or "",
            "fecha": row[5],
            "nota": row[6],
        })

    # Datos de alumnas
    ws = wb["datos"]
    datos = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[2] is None and row[1] is None:
            continue
        datos.append({
            "nacionalidad": row[0] or "",
            "nombre_religioso": row[1] or "",
            "alumna": row[2] or "",
            "num_documento": row[3] or "",
            "tipo_documento": row[4] or "",
            "fecha_nacimiento": row[5],
        })

    # Profesores
    ws = wb["profesores"]
    profesores = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1]:
            profesores.append(row[1])

    # Plan de estudios
    ws = wb["Plan de estudios"]
    plan = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1] is None:
            continue
        codigo = (row[0] or "").strip()
        plan.append({
            "codigo": codigo,
            "materia": row[1],
            "orden": row[2],
            "anio_desc": row[3] or "",
            "ects": row[4],
        })

    wb.close()
    return notas, datos, profesores, plan


def set_cell_border(cell, **kwargs):
    """Set cell border. Usage: set_cell_border(cell, top={"sz": 4, "val": "single"})"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            element = OxmlElement(f'w:{edge}')
            for key in ["sz", "val", "color", "space"]:
                if key in edge_data:
                    element.set(qn(f'w:{key}'), str(edge_data[key]))
            tcBorders.append(element)
    tcPr.append(tcBorders)


def set_table_borders(table):
    """Apply borders to entire table."""
    border = {"sz": 4, "val": "single", "color": "000000", "space": "0"}
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, top=border, bottom=border, start=border, end=border)


def set_cell_shading(cell, color):
    """Set cell background color."""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    shading.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(shading)


def format_date(d):
    """Format a date or datetime to DD/MM/YYYY."""
    if isinstance(d, datetime.datetime):
        return d.strftime("%d/%m/%Y")
    elif isinstance(d, datetime.date):
        return d.strftime("%d/%m/%Y")
    elif isinstance(d, str):
        return d
    return ""


def anio_academico_from_nota(anio, semestre, plan):
    """Determine academic year label (Primero, Segundo, Tercero) from year number."""
    mapping = {2023: "Primero", 2024: "Segundo", 2025: "Tercero"}
    return mapping.get(anio, str(anio) if anio else "")


def generate_acta_examen(materia, profesor, anio, semestre, fecha, notas_filtradas, datos_alumnas):
    """Generate Acta de Examen document."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)

    # Map year to academic year name
    anio_map = {"2023": "Primero", "2024": "Segundo", "2025": "Tercero"}
    anio_academico = anio_map.get(str(anio), str(anio))

    # === Header section (NO borders) ===
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Estudiantado Santa Mariana de Jesús")
    run.bold = True
    run.font.size = Pt(14)

    doc.add_paragraph("")

    # Info fields as simple paragraphs (no table borders)
    info_fields = [
        ("Año:", str(anio)),
        ("Materia:", materia),
        ("Profesor:", profesor),
        ("Fecha:", format_date(fecha)),
        ("Año académico:", anio_academico),
        ("Semestre:", str(semestre)),
    ]
    for label, value in info_fields:
        p = doc.add_paragraph()
        run = p.add_run(label + "  ")
        run.bold = True
        run.font.size = Pt(11)
        run = p.add_run(value)
        run.font.size = Pt(11)

    doc.add_paragraph("")

    # === Student table (WITH borders) ===
    table = doc.add_table(rows=0, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row for students
    row = table.add_row()
    headers = ["N°", "Apellido, nombre", "Documento", "Nota"]
    for i, h in enumerate(headers):
        p = row.cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.bold = True
        run.font.size = Pt(9)
        pass  # No background color

    # Build lookup for alumna data
    datos_map = {}
    for d in datos_alumnas:
        datos_map[d["nombre_religioso"]] = d
        datos_map[d["alumna"]] = d

    # Student rows (at least 12 rows like template)
    num_rows = max(12, len(notas_filtradas))
    for i in range(num_rows):
        row = table.add_row()
        row.cells[0].text = str(i + 1)
        row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if i < len(notas_filtradas):
            nota = notas_filtradas[i]
            alumna_name = nota["alumna"]
            # Get nombre religioso (nombre de monja)
            dato = datos_map.get(alumna_name, {})
            nombre_mostrar = dato.get("nombre_religioso", "") or alumna_name
            doc_num = dato.get("num_documento", "")
            row.cells[1].text = nombre_mostrar
            row.cells[2].text = str(doc_num)
            row.cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row.cells[3].text = str(nota["nota"]) if nota["nota"] is not None else ""
            row.cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Set font size for student table
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    if run.font.size is None:
                        run.font.size = Pt(10)

    # Column widths
    for row in table.rows:
        row.cells[0].width = Cm(1.5)
        row.cells[1].width = Cm(7)
        row.cells[2].width = Cm(4)
        row.cells[3].width = Cm(2)

    set_table_borders(table)

    # Signature line
    doc.add_paragraph("")
    doc.add_paragraph("")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(".............................................")
    run.font.size = Pt(10)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Firma del titular")
    run.font.size = Pt(10)

    # Observations
    doc.add_paragraph("")
    p = doc.add_paragraph()
    run = p.add_run("Observaciones:")
    run.bold = True
    run.font.size = Pt(10)
    doc.add_paragraph("_" * 80)
    doc.add_paragraph("_" * 80)

    # Save
    safe_materia = materia.replace("/", "-").replace(":", "-")[:40]
    filename = f"Acta_Examen_{safe_materia}_{anio}_{semestre}.docx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    doc.save(filepath)
    return filepath


def generate_analitico(alumna_nombre, notas, datos_alumnas, plan):
    """Generate Analitico de Estudios document."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)

    # Get student data
    dato_alumna = None
    for d in datos_alumnas:
        if d["nombre_religioso"] == alumna_nombre or d["alumna"] == alumna_nombre:
            dato_alumna = d
            break

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Certificado de Estudios")
    run.bold = True
    run.font.size = Pt(14)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('Instituto "Servidoras del Señor y de la Virgen de Matará"')
    run.font.size = Pt(11)

    # Student name - mostrar nombre religioso
    nombre_mostrar = (dato_alumna.get("nombre_religioso") or dato_alumna.get("alumna", alumna_nombre)) if dato_alumna else alumna_nombre
    p = doc.add_paragraph()
    run = p.add_run(f"Alumna: ")
    run.bold = True
    run.font.size = Pt(11)
    run = p.add_run(nombre_mostrar)
    run.font.size = Pt(11)
    run.underline = True

    # Build notas lookup by materia name (normalize for matching)
    notas_alumna = {}
    for n in notas:
        if n["alumna"] == alumna_nombre:
            notas_alumna[n["materia"].strip().lower()] = n

    # Define semester groups matching the template structure
    semester_groups = [
        (1, "PRIMER AÑO - I Semestre", "1 ISR, I Sem"),
        (2, "PRIMER AÑO - II Semestre", "1 ISR, II Sem"),
        (3, "SEGUNDO AÑO - I Semestre", "2 ISR, I Sem"),
        (4, "SEGUNDO AÑO - II Semestre", "2 ISR, II Sem"),
        (5, "TERCER AÑO - I Semestre", "3 ISR, I Sem"),
        (6, "TERCER AÑO - II Semestre", "3 ISR, II Sem"),
    ]

    institucion = "Estudiantado Santa Mariana de Jesús, Ecuador"

    for orden, titulo, anio_desc in semester_groups:
        # Filter plan entries for this semester
        materias_sem = [m for m in plan if m["orden"] == orden]
        # Add Latin for relevant semesters
        latin_map = {1: ("L 101", "Latín I"), 2: ("L 102", "Latín II"),
                     3: ("L 103", "Latín III"), 4: ("L 104", "Latín IV"),
                     5: ("L 105", "Latín V"), 6: ("L 106", "Latín VI")}
        if orden in latin_map:
            cod, mat = latin_map[orden]
            # Check it's not already there
            if not any(m["codigo"] == cod for m in materias_sem):
                materias_sem.append({"codigo": cod, "materia": mat, "orden": orden,
                                     "anio_desc": "Lingue", "ects": 6})

        if not materias_sem:
            continue

        doc.add_paragraph("")

        # Create table
        num_rows = len(materias_sem) + 2  # title + header + data
        table = doc.add_table(rows=num_rows, cols=5)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Title row
        cell = table.rows[0].cells[0]
        cell.merge(table.rows[0].cells[4])
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(titulo)
        run.bold = True
        run.font.size = Pt(9)

        # Header row
        headers = ["Código", "Materia", "Nota", "Fecha", "Institución"]
        for i, h in enumerate(headers):
            cell = table.rows[1].cells[i]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(h)
            run.bold = True
            run.font.size = Pt(8)
            pass  # No background color

        # Data rows
        for idx, mat in enumerate(materias_sem):
            row = table.rows[idx + 2]
            row.cells[0].text = mat["codigo"]
            row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row.cells[1].text = mat["materia"]

            # Look up grade
            mat_key = mat["materia"].strip().lower()
            nota_data = notas_alumna.get(mat_key)
            if nota_data:
                row.cells[2].text = str(nota_data["nota"]) if nota_data["nota"] is not None else ""
                row.cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                row.cells[3].text = format_date(nota_data["fecha"])
                row.cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row.cells[4].text = institucion

            # Font size
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(8)

        # Set column widths
        widths = [Cm(2), Cm(6), Cm(1.5), Cm(2.5), Cm(6)]
        for row in table.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = w

        set_table_borders(table)

    # Extra courses section
    doc.add_paragraph("")
    p = doc.add_paragraph()
    run = p.add_run("CURSOS EXTRA:")
    run.bold = True
    run.font.size = Pt(10)

    # Extra materias (orden 20)
    extras = [m for m in plan if m["orden"] == 20 and m["materia"]]
    if extras:
        table = doc.add_table(rows=len(extras) + 1, cols=5)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        headers = ["Código", "Materia", "Nota", "Fecha", "Institución"]
        for i, h in enumerate(headers):
            cell = table.rows[0].cells[i]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(h)
            run.bold = True
            run.font.size = Pt(8)
            pass  # No background color

        for idx, mat in enumerate(extras):
            row = table.rows[idx + 1]
            row.cells[0].text = mat["codigo"]
            row.cells[1].text = mat["materia"]
            mat_key = mat["materia"].strip().lower()
            nota_data = notas_alumna.get(mat_key)
            if nota_data:
                row.cells[2].text = str(nota_data["nota"]) if nota_data["nota"] is not None else ""
                row.cells[3].text = format_date(nota_data["fecha"])
            row.cells[4].text = institucion
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(8)

        set_table_borders(table)

    # Signatures
    doc.add_paragraph("")
    doc.add_paragraph("")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("________________\t\t\t\t\t___________________")
    run.font.size = Pt(10)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Superiora Local\t\t\t\t\tEncargada de Estudios")
    run.font.size = Pt(10)

    # Save
    safe_name = alumna_nombre.replace(" ", "_")[:30]
    filename = f"Analitico_{safe_name}.docx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    doc.save(filepath)
    return filepath


class ReportGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Reportes - Estudiantado Santa Mariana de Jesús")
        self.root.geometry("700x580")
        self.root.resizable(False, False)

        # Try to set a nice background
        self.root.configure(bg="#f0f4f8")

        # Load data
        try:
            self.notas, self.datos, self.profesores, self.plan = load_excel_data()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer esquema.xlsx:\n{e}")
            sys.exit(1)

        # Get unique values for dropdowns
        self.materias_unicas = sorted(set(n["materia"] for n in self.notas if n["materia"]))
        self.alumnas_unicas = sorted(set(n["alumna"] for n in self.notas if n["alumna"]))
        self.profesores_unicos = sorted(set(n["profesor"] for n in self.notas if n["profesor"]))
        # Also add from datos sheet
        for d in self.datos:
            if d["nombre_religioso"] and d["nombre_religioso"] not in self.alumnas_unicas:
                self.alumnas_unicas.append(d["nombre_religioso"])
        # Add all materias from plan too
        all_materias = sorted(set(
            list(self.materias_unicas) +
            [m["materia"] for m in self.plan if m["materia"]]
        ))
        self.materias_unicas = all_materias

        self.build_ui()

    def build_ui(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Title.TLabel", font=("Helvetica", 16, "bold"), background="#f0f4f8")
        style.configure("Subtitle.TLabel", font=("Helvetica", 10), background="#f0f4f8", foreground="#555")
        style.configure("Section.TLabelframe.Label", font=("Helvetica", 11, "bold"))
        style.configure("Big.TButton", font=("Helvetica", 12, "bold"), padding=10)
        style.configure("TLabelframe", background="#f0f4f8")
        style.configure("TLabel", background="#f0f4f8")
        style.configure("TFrame", background="#f0f4f8")

        # Title
        title_frame = ttk.Frame(self.root)
        title_frame.pack(pady=(15, 5), fill="x")
        ttk.Label(title_frame, text="Generador de Reportes", style="Title.TLabel").pack()
        ttk.Label(title_frame, text="Estudiantado Santa Mariana de Jesús", style="Subtitle.TLabel").pack()

        # Report type selection
        type_frame = ttk.LabelFrame(self.root, text=" Tipo de Reporte ", style="Section.TLabelframe", padding=10)
        type_frame.pack(padx=20, pady=10, fill="x")

        self.report_type = tk.StringVar(value="acta")
        ttk.Radiobutton(type_frame, text="Acta de Examen", variable=self.report_type,
                        value="acta", command=self.toggle_panels).pack(side="left", padx=20)
        ttk.Radiobutton(type_frame, text="Analítico de Estudios", variable=self.report_type,
                        value="analitico", command=self.toggle_panels).pack(side="left", padx=20)

        # === ACTA panel ===
        self.acta_frame = ttk.LabelFrame(self.root, text=" Datos del Acta de Examen ",
                                         style="Section.TLabelframe", padding=15)

        row = 0
        ttk.Label(self.acta_frame, text="Materia:").grid(row=row, column=0, sticky="w", pady=5)
        self.acta_materia = ttk.Combobox(self.acta_frame, values=self.materias_unicas, width=50)
        self.acta_materia.grid(row=row, column=1, pady=5, padx=5)

        row += 1
        ttk.Label(self.acta_frame, text="Profesor:").grid(row=row, column=0, sticky="w", pady=5)
        self.acta_profesor = ttk.Combobox(self.acta_frame, values=self.profesores_unicos + self.profesores, width=50)
        self.acta_profesor.grid(row=row, column=1, pady=5, padx=5)

        row += 1
        ttk.Label(self.acta_frame, text="Año:").grid(row=row, column=0, sticky="w", pady=5)
        self.acta_anio = ttk.Combobox(self.acta_frame, values=["2023", "2024", "2025", "2026"], width=15)
        self.acta_anio.grid(row=row, column=1, pady=5, padx=5, sticky="w")
        self.acta_anio.set("2025")

        row += 1
        ttk.Label(self.acta_frame, text="Semestre:").grid(row=row, column=0, sticky="w", pady=5)
        self.acta_semestre = ttk.Combobox(self.acta_frame, values=["I", "II"], width=10)
        self.acta_semestre.grid(row=row, column=1, pady=5, padx=5, sticky="w")
        self.acta_semestre.set("I")

        row += 1
        ttk.Label(self.acta_frame, text="Fecha (DD/MM/AAAA):").grid(row=row, column=0, sticky="w", pady=5)
        self.acta_fecha = ttk.Entry(self.acta_frame, width=20)
        self.acta_fecha.grid(row=row, column=1, pady=5, padx=5, sticky="w")
        today = datetime.date.today().strftime("%d/%m/%Y")
        self.acta_fecha.insert(0, today)

        # === ANALITICO panel ===
        self.analitico_frame = ttk.LabelFrame(self.root, text=" Datos del Analítico de Estudios",
                                               style="Section.TLabelframe", padding=15)

        ttk.Label(self.analitico_frame, text="Alumna:").grid(row=0, column=0, sticky="w", pady=5)
        self.analitico_alumna = ttk.Combobox(self.analitico_frame, values=self.alumnas_unicas, width=50)
        self.analitico_alumna.grid(row=0, column=1, pady=5, padx=5)
        if self.alumnas_unicas:
            self.analitico_alumna.set(self.alumnas_unicas[0])

        # Status label
        self.status_var = tk.StringVar(value="")
        self.status_label = ttk.Label(self.root, textvariable=self.status_var,
                                      foreground="#2563eb", font=("Helvetica", 10))

        # Generate button
        self.generate_btn = ttk.Button(self.root, text="Generar Reporte", style="Big.TButton",
                                       command=self.generate_report)

        # Info label
        self.info_var = tk.StringVar(
            value=f"Datos cargados: {len(self.notas)} notas, {len(self.datos)} alumnas, "
                  f"{len(self.plan)} materias en plan"
        )
        ttk.Label(self.root, textvariable=self.info_var, style="Subtitle.TLabel").pack(side="bottom", pady=5)

        # Initial layout
        self.toggle_panels()

    def toggle_panels(self):
        self.acta_frame.pack_forget()
        self.analitico_frame.pack_forget()
        self.status_label.pack_forget()
        self.generate_btn.pack_forget()

        if self.report_type.get() == "acta":
            self.acta_frame.pack(padx=20, pady=10, fill="x")
        else:
            self.analitico_frame.pack(padx=20, pady=10, fill="x")

        self.status_label.pack(padx=20, pady=5)
        self.generate_btn.pack(pady=15)

    def generate_report(self):
        try:
            # Reload data each time in case Excel was updated
            self.notas, self.datos, self.profesores, self.plan = load_excel_data()

            if self.report_type.get() == "acta":
                filepath = self.generate_acta()
            else:
                filepath = self.generate_analitico()

            self.status_var.set(f"Reporte generado: {os.path.basename(filepath)}")
            messagebox.showinfo("Éxito",
                                f"Reporte generado exitosamente!\n\n{filepath}")

            # Open the file
            if sys.platform == "linux":
                subprocess.Popen(["xdg-open", filepath])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", filepath])
            elif sys.platform == "win32":
                os.startfile(filepath)

        except Exception as e:
            messagebox.showerror("Error", f"Error al generar reporte:\n{e}")
            import traceback
            traceback.print_exc()

    def generate_acta(self):
        materia = self.acta_materia.get().strip()
        profesor = self.acta_profesor.get().strip()
        anio = self.acta_anio.get().strip()
        semestre = self.acta_semestre.get().strip()
        fecha_str = self.acta_fecha.get().strip()

        if not materia:
            raise ValueError("Seleccione una materia")
        if not profesor:
            raise ValueError("Seleccione un profesor")

        # Parse fecha
        try:
            parts = fecha_str.split("/")
            fecha = datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
        except Exception:
            fecha = fecha_str

        # Filter notas for this materia
        notas_filtradas = []
        for n in self.notas:
            materia_match = n["materia"].strip().lower() == materia.strip().lower()
            if materia_match:
                notas_filtradas.append(n)

        return generate_acta_examen(materia, profesor, anio, semestre, fecha, notas_filtradas, self.datos)

    def generate_analitico(self):
        alumna = self.analitico_alumna.get().strip()
        if not alumna:
            raise ValueError("Seleccione una alumna")

        return generate_analitico(alumna, self.notas, self.datos, self.plan)


if __name__ == "__main__":
    root = tk.Tk()
    app = ReportGeneratorApp(root)
    root.mainloop()