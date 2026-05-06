"""Utilidades comunes para construir documentos docx."""
import datetime

from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_cell_border(cell, **kwargs):
    """Aplica bordes a una celda de tabla.

    Ejemplo: set_cell_border(cell, top={"sz": 4, "val": "single"})
    """
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_borders = OxmlElement('w:tcBorders')
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            element = OxmlElement(f'w:{edge}')
            for key in ("sz", "val", "color", "space"):
                if key in edge_data:
                    element.set(qn(f'w:{key}'), str(edge_data[key]))
            tc_borders.append(element)
    tc_pr.append(tc_borders)


def set_table_borders(table):
    """Bordes simples en toda la tabla."""
    border = {"sz": 4, "val": "single", "color": "000000", "space": "0"}
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, top=border, bottom=border,
                            start=border, end=border)


def set_cell_shading(cell, color):
    """Color de fondo de una celda."""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    shading.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(shading)


def format_date(d):
    """Formatea fecha/datetime a DD/MM/YYYY. Strings se devuelven sin tocar."""
    if isinstance(d, (datetime.datetime, datetime.date)):
        return d.strftime("%d/%m/%Y")
    if isinstance(d, str):
        return d
    return ""
