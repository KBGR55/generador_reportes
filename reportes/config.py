"""Rutas de archivos del proyecto.

`PROJECT_ROOT` se resuelve desde la ubicación de este módulo, dos niveles
arriba (`reportes/config.py` → `reportes/` → raíz del proyecto).
"""
import os

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH = os.path.join(PROJECT_ROOT, "esquema.xlsx")
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "reportes_generados")
