# Generador de Reportes — Estudiantado Santa Mariana de Jesús

Aplicación de escritorio en Python/Tkinter que genera dos tipos de documentos
Word (`.docx`) a partir de los datos de una planilla Excel:

- **Acta de Examen** — listado de alumnas con sus calificaciones para una
  materia, profesor, año y semestre dados.
- **Analítico de Estudios** — certificado completo de una alumna con todas
  las materias del plan agrupadas por año y semestre.

Incluye un visor integrado para ver, abrir y eliminar los reportes ya
generados sin salir de la aplicación.

## Requisitos

- Python 3.10 o superior
- `tkinter` (en Linux: `sudo apt install python3-tk`)
- Dependencias Python:

```bash
pip install python-docx openpyxl
```

## Ejecución

```bash
python3 generador_reportes.py
```

La ventana principal se ve así:

- Header con barra de acento y botón **Ver reportes generados**
- Selector de tipo de reporte (Acta / Analítico)
- Formulario con los campos del reporte
- Botón primario **Generar Reporte**

## Estructura del Excel (`esquema.xlsx`)

El programa lee un único archivo `esquema.xlsx` ubicado en la misma carpeta
que el script. Debe tener **cuatro hojas**:

### Hoja `notas`
| Profesor | Alumna | Materia | Año | semestre | fecha | nota |
|---|---|---|---|---|---|---|
| P. Leonardo López | María Reina de los Santos | Exégesis NT: San Juan I | 2025 | I | 11/06/2025 | 10 |

### Hoja `datos`
| nacionalidad | nombre_religioso | alumna | num_documento | tipo_documento | fecha_nacimiento |
|---|---|---|---|---|---|

### Hoja `profesores`
| (col A) | profesor |
|---|---|
| | P. Leonardo López |

### Hoja `Plan de estudios`
| codigo | materia | orden | anio_desc | ects |
|---|---|---|---|---|
| SR 119 | Metafísica I | 3 | 2 ISR, I Sem | 6 |

`orden` es un entero del 1 al 6 que indica la posición en el plan
(1 = Primer año I sem, 2 = Primer año II sem, … 6 = Tercer año II sem).
`orden = 20` se reserva para los **cursos extra**.

> **Importante:** los nombres de materia en `Plan de estudios` deben coincidir
> exactamente (ignorando mayúsculas/minúsculas y espacios al inicio/fin) con
> los nombres usados en la hoja `notas`. Si no coinciden, la nota saldrá
> vacía en el Analítico.

## Salida

Los `.docx` generados se guardan en `reportes_generados/`. El nombre tiene
el formato:

- `Acta_Examen_<materia>_<año>_<semestre>.docx`
- `Analitico_<nombre_alumna>.docx`

Al finalizar, el archivo se abre automáticamente con el visor del sistema.

## Generar ejecutable (opcional)

El proyecto incluye `GeneradorReportes.spec` para empaquetar con PyInstaller:

```bash
pip install pyinstaller
pyinstaller GeneradorReportes.spec
```

El binario queda en `dist/`.

## Privacidad

El repositorio **no incluye** `esquema.xlsx` ni la carpeta
`reportes_generados/` porque contienen datos personales de las alumnas
(nombres religiosos, documentos, fechas de nacimiento, calificaciones).
Ver `.gitignore` para el detalle. Cada quien debe proveer su propio
`esquema.xlsx` con la estructura descrita arriba.

## Estructura del proyecto

```
reportes_monjas/
├── generador_reportes.py          # aplicación principal (GUI Tkinter)
├── esquema.xlsx                   # datos fuente (NO se commitea)
├── ACTA DE EXAMEN-1.docx          # plantilla de referencia visual
├── Analítico de Estudios.docx     # plantilla de referencia visual
├── GeneradorReportes.spec         # config PyInstaller
├── reportes_generados/            # salida (NO se commitea)
├── .gitignore
└── README.md
```

## Licencia

Uso interno del Estudiantado Santa Mariana de Jesús.
