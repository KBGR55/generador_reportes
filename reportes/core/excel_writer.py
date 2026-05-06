"""Escritura segura sobre esquema.xlsx.

- Detecta si el archivo está bloqueado por Excel/LibreOffice.
- Hace un backup con timestamp antes de cada escritura.
- Anexa filas a las hojas correspondientes.
"""
import datetime
import os
import shutil

import openpyxl

from reportes.config import EXCEL_PATH, PROJECT_ROOT


BACKUPS_DIR = os.path.join(PROJECT_ROOT, "backups")


class ExcelLockedError(Exception):
    """El archivo esquema.xlsx está abierto en otra aplicación."""


def _lock_file_path():
    folder = os.path.dirname(EXCEL_PATH)
    name = os.path.basename(EXCEL_PATH)
    return os.path.join(folder, f".~lock.{name}#")


def excel_is_locked():
    """True si LibreOffice/Excel tienen el archivo abierto."""
    return os.path.exists(_lock_file_path())


def backup_excel():
    """Copia esquema.xlsx a backups/ con timestamp y devuelve la ruta."""
    os.makedirs(BACKUPS_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(BACKUPS_DIR, f"esquema_{ts}.xlsx")
    shutil.copy2(EXCEL_PATH, dest)
    return dest


def _append_row(sheet_name, row_values):
    if excel_is_locked():
        raise ExcelLockedError(
            "El archivo esquema.xlsx está abierto en Excel/LibreOffice. "
            "Cerralo antes de guardar."
        )
    backup_excel()
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[sheet_name]
    ws.append(row_values)
    wb.save(EXCEL_PATH)
    wb.close()


def add_nota(profesor, alumna, materia, anio, semestre, fecha, nota):
    """Anexa una fila a la hoja `notas`."""
    _append_row("notas",
                [profesor, alumna, materia, anio, semestre, fecha, nota])


def add_alumna(nacionalidad, nombre_religioso, alumna,
               num_documento, tipo_documento, fecha_nacimiento):
    """Anexa una fila a la hoja `datos`."""
    _append_row("datos",
                [nacionalidad, nombre_religioso, alumna,
                 num_documento, tipo_documento, fecha_nacimiento])


def add_profesor(nombre):
    """Anexa una fila a la hoja `profesores` (columna B)."""
    _append_row("profesores", [None, nombre])
