"""Ventana «Cargar datos al Excel».

Tres pestañas: Nueva nota, Nueva alumna, Nuevo profesor.
"""
import datetime
import tkinter as tk
from tkinter import ttk, messagebox

from reportes.core import (
    ExcelLockedError,
    add_alumna,
    add_nota,
    add_profesor,
    load_excel_data,
)
from reportes.ui.theme import BG, CARD, BORDER, ACCENT_BAR


def _parse_fecha(s):
    """DD/MM/YYYY → datetime.date. Lanza ValueError si no es válida."""
    parts = s.split("/")
    if len(parts) != 3:
        raise ValueError("Formato de fecha esperado: DD/MM/AAAA")
    return datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))


def _parse_nota(s):
    """Convierte a float aceptando coma o punto. Valida rango 0-10."""
    s = s.replace(",", ".").strip()
    n = float(s)
    if not 0 <= n <= 10:
        raise ValueError("La nota debe estar entre 0 y 10")
    return n


def open_data_entry(parent, on_saved=None):
    """Abre la ventana de carga de datos.

    `on_saved` se ejecuta tras cada guardado exitoso (para refrescar combos).
    """
    win = tk.Toplevel(parent)
    win.title("Cargar datos al Excel")
    win.geometry("640x540")
    win.minsize(580, 480)
    win.configure(bg=BG)
    win.transient(parent)

    tk.Frame(win, bg=ACCENT_BAR, height=4).pack(fill="x", side="top")

    head = ttk.Frame(win)
    head.pack(fill="x", padx=24, pady=(16, 6))
    ttk.Label(head, text="Cargar datos al Excel",
              style="Title.TLabel").pack(anchor="w")
    ttk.Label(head, text="Los cambios se guardan en esquema.xlsx (con backup)",
              style="Subtitle.TLabel").pack(anchor="w", pady=(2, 0))

    tk.Frame(win, height=1, bg=BORDER).pack(fill="x", padx=24, pady=(10, 12))

    notebook = ttk.Notebook(win)
    notebook.pack(padx=24, pady=(0, 16), fill="both", expand=True)

    # Datos para los combos (se recargan a demanda)
    state = {"data": load_excel_data()}

    def reload_state():
        state["data"] = load_excel_data()

    def alumnas():
        notas, datos, _, _ = state["data"]
        nombres = sorted({n["alumna"] for n in notas if n["alumna"]})
        for d in datos:
            if d["nombre_religioso"] and d["nombre_religioso"] not in nombres:
                nombres.append(d["nombre_religioso"])
        return sorted(nombres)

    def profesores():
        _, _, profs, _ = state["data"]
        return sorted(set(profs))

    def materias():
        _, _, _, plan = state["data"]
        return sorted({m["materia"] for m in plan if m["materia"]})

    def after_save():
        reload_state()
        if on_saved:
            on_saved()

    _build_nota_tab(notebook, alumnas, profesores, materias,
                    parent=win, after_save=after_save)
    _build_alumna_tab(notebook, parent=win, after_save=after_save)
    _build_profesor_tab(notebook, parent=win, after_save=after_save)


# --------------------------------------------------------------------- helpers
def _make_card_frame(notebook):
    frame = tk.Frame(notebook, bg=CARD)
    return frame


def _grid_field(parent, row, label, widget, sticky="ew"):
    ttk.Label(parent, text=label, style="Form.TLabel").grid(
        row=row, column=0, sticky="w", pady=7, padx=(0, 14))
    widget.grid(row=row, column=1, sticky=sticky, pady=7)
    return widget


def _handle_save(parent, action):
    """Ejecuta `action`, muestra error/éxito y devuelve True si grabó.

    `action` debe devolver el ID asignado (o None) para mostrarlo al usuario.
    """
    try:
        new_id = action()
    except ExcelLockedError as e:
        messagebox.showerror("Excel abierto", str(e), parent=parent)
        return False
    except ValueError as e:
        messagebox.showerror("Datos inválidos", str(e), parent=parent)
        return False
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar:\n{e}",
                             parent=parent)
        return False

    if new_id:
        msg = f"Registro guardado en esquema.xlsx\n\nID asignado: {new_id}"
    else:
        msg = "Registro guardado en esquema.xlsx"
    messagebox.showinfo("Guardado", msg, parent=parent)
    return True


# ------------------------------------------------------------------ Nueva nota
def _build_nota_tab(notebook, alumnas, profesores, materias,
                    parent, after_save):
    frame = _make_card_frame(notebook)
    notebook.add(frame, text="Nueva nota")

    inner = tk.Frame(frame, bg=CARD)
    inner.pack(padx=24, pady=20, fill="both", expand=True)
    inner.columnconfigure(1, weight=1)

    cb_alumna = ttk.Combobox(inner, values=alumnas())
    cb_profesor = ttk.Combobox(inner, values=profesores())
    cb_materia = ttk.Combobox(inner, values=materias())
    cb_anio = ttk.Combobox(inner, values=["2023", "2024", "2025", "2026", "2027"], width=14)
    cb_semestre = ttk.Combobox(inner, values=["I", "II"], width=14)
    en_fecha = ttk.Entry(inner, width=22)
    en_nota = ttk.Entry(inner, width=14)

    _grid_field(inner, 0, "Alumna", cb_alumna)
    _grid_field(inner, 1, "Profesor", cb_profesor)
    _grid_field(inner, 2, "Materia", cb_materia)
    _grid_field(inner, 3, "Año", cb_anio, sticky="w")
    _grid_field(inner, 4, "Semestre", cb_semestre, sticky="w")
    _grid_field(inner, 5, "Fecha (DD/MM/AAAA)", en_fecha, sticky="w")
    _grid_field(inner, 6, "Nota (0-10)", en_nota, sticky="w")

    cb_anio.set(str(datetime.date.today().year))
    cb_semestre.set("I")
    en_fecha.insert(0, datetime.date.today().strftime("%d/%m/%Y"))

    def guardar():
        def action():
            alumna = cb_alumna.get().strip()
            profesor = cb_profesor.get().strip()
            materia = cb_materia.get().strip()
            anio_s = cb_anio.get().strip()
            semestre = cb_semestre.get().strip()
            if not alumna:
                raise ValueError("Seleccione la alumna")
            if not profesor:
                raise ValueError("Seleccione el profesor")
            if not materia:
                raise ValueError("Seleccione la materia")
            if not anio_s:
                raise ValueError("Seleccione el año")
            if not semestre:
                raise ValueError("Seleccione el semestre")
            try:
                anio = int(anio_s)
            except ValueError:
                raise ValueError("Año inválido")
            fecha = _parse_fecha(en_fecha.get().strip())
            nota = _parse_nota(en_nota.get().strip())
            return add_nota(profesor, alumna, materia, anio, semestre, fecha, nota)

        if _handle_save(parent, action):
            en_nota.delete(0, "end")
            after_save()
            cb_alumna["values"] = alumnas()
            cb_profesor["values"] = profesores()
            cb_materia["values"] = materias()

    btns = tk.Frame(frame, bg=CARD)
    btns.pack(fill="x", padx=24, pady=(0, 20))
    ttk.Button(btns, text="Guardar nota", style="Primary.TButton",
               cursor="hand2", command=guardar).pack(side="left")


# ---------------------------------------------------------------- Nueva alumna
def _build_alumna_tab(notebook, parent, after_save):
    frame = _make_card_frame(notebook)
    notebook.add(frame, text="Nueva alumna")

    inner = tk.Frame(frame, bg=CARD)
    inner.pack(padx=24, pady=20, fill="both", expand=True)
    inner.columnconfigure(1, weight=1)

    en_nac = ttk.Entry(inner)
    en_relig = ttk.Entry(inner)
    en_civil = ttk.Entry(inner)
    en_doc = ttk.Entry(inner)
    cb_tipo = ttk.Combobox(inner, values=["Cédula", "Pasaporte", "DNI"], width=20)
    en_nac_fecha = ttk.Entry(inner, width=22)

    _grid_field(inner, 0, "Nacionalidad", en_nac)
    _grid_field(inner, 1, "Nombre religioso", en_relig)
    _grid_field(inner, 2, "Nombre civil", en_civil)
    _grid_field(inner, 3, "N° documento", en_doc, sticky="w")
    _grid_field(inner, 4, "Tipo documento", cb_tipo, sticky="w")
    _grid_field(inner, 5, "Fecha nacimiento (DD/MM/AAAA)", en_nac_fecha, sticky="w")

    cb_tipo.set("Cédula")

    def guardar():
        def action():
            relig = en_relig.get().strip()
            civil = en_civil.get().strip()
            if not relig and not civil:
                raise ValueError(
                    "Ingrese al menos un nombre (religioso o civil)")
            fecha_str = en_nac_fecha.get().strip()
            fecha_nac = _parse_fecha(fecha_str) if fecha_str else None
            return add_alumna(en_nac.get().strip(), relig, civil,
                              en_doc.get().strip(), cb_tipo.get().strip(),
                              fecha_nac)

        if _handle_save(parent, action):
            for w in (en_nac, en_relig, en_civil, en_doc, en_nac_fecha):
                w.delete(0, "end")
            after_save()

    btns = tk.Frame(frame, bg=CARD)
    btns.pack(fill="x", padx=24, pady=(0, 20))
    ttk.Button(btns, text="Guardar alumna", style="Primary.TButton",
               cursor="hand2", command=guardar).pack(side="left")


# -------------------------------------------------------------- Nuevo profesor
def _build_profesor_tab(notebook, parent, after_save):
    frame = _make_card_frame(notebook)
    notebook.add(frame, text="Nuevo profesor")

    inner = tk.Frame(frame, bg=CARD)
    inner.pack(padx=24, pady=20, fill="both", expand=True)
    inner.columnconfigure(1, weight=1)

    en_nombre = ttk.Entry(inner)
    _grid_field(inner, 0, "Nombre completo", en_nombre)

    def guardar():
        def action():
            nombre = en_nombre.get().strip()
            if not nombre:
                raise ValueError("Ingrese el nombre del profesor")
            return add_profesor(nombre)

        if _handle_save(parent, action):
            en_nombre.delete(0, "end")
            after_save()

    btns = tk.Frame(frame, bg=CARD)
    btns.pack(fill="x", padx=24, pady=(0, 20))
    ttk.Button(btns, text="Guardar profesor", style="Primary.TButton",
               cursor="hand2", command=guardar).pack(side="left")
