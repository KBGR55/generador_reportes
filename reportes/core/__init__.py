"""Lógica de negocio (sin dependencias de UI).

Expone la API pública que consumen las capas superiores.
"""
from reportes.core.acta_examen import generate_acta_examen
from reportes.core.analitico import generate_analitico
from reportes.core.excel_loader import load_excel_data

__all__ = [
    "load_excel_data",
    "generate_acta_examen",
    "generate_analitico",
]
