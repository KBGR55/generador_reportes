"""Rutas de archivos del proyecto."""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "esquema.xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, "reportes_generados")
