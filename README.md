# Generador de Reportes — Estudiantado Santa Mariana de Jesús

Aplicación de escritorio en Python/Tkinter que genera dos tipos de documentos
Word (`.docx`) a partir de los datos de una planilla Excel:

- **Acta de Examen** — listado de alumnas con sus calificaciones para una
  materia, profesor, año y semestre dados.
- **Analítico de Estudios** — certificado completo de una alumna con todas
  las materias del plan agrupadas por año y semestre.

Incluye un visor integrado para ver, abrir y eliminar los reportes ya
generados sin salir de la aplicación, y un formulario para cargar nuevas
notas, alumnas o profesores directamente al Excel (con backup automático).

## Requisitos

- Python 3.10 o superior
- `tkinter` (en Linux: `sudo apt install python3-tk`)

## Instalación

Clonar el repo y, desde la raíz del proyecto, instalar las dependencias.
Hay dos opciones:

**Opción 1 — instalar el paquete (recomendado):**

```bash
pip install -e .
```

Esto instala `python-docx` y `openpyxl`, y deja disponible el comando
`generador-reportes` en la terminal.

**Opción 2 — solo dependencias:**

```bash
pip install python-docx openpyxl
```

## Ejecución

Cualquiera de estas tres formas funciona:

```bash
python3 main.py            # script de entrada
python3 -m reportes        # ejecutar el paquete
generador-reportes         # si instalaste con `pip install -e .`
```

## Estructura del Excel (`esquema.xlsx`)

El programa lee un único archivo `esquema.xlsx` ubicado en la raíz del
proyecto. Debe tener **cuatro hojas**:

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

## Cargar datos al Excel

El botón **Cargar datos** del header abre una ventana con tres pestañas:

- **Nueva nota** — anexa una fila a la hoja `notas` (alumna + profesor +
  materia + año + semestre + fecha + nota).
- **Nueva alumna** — anexa una fila a la hoja `datos`.
- **Nuevo profesor** — anexa una fila a la hoja `profesores`.

Antes de cada escritura el programa:

1. Verifica que `esquema.xlsx` no esté abierto en Excel/LibreOffice (si lo
   está, se aborta y se avisa).
2. Copia `esquema.xlsx` a `backups/esquema_<timestamp>.xlsx` para no perder
   datos en caso de error.

Tras cada guardado, los combos de la app principal se refrescan
automáticamente con los nuevos valores.

### Identificadores únicos

Cada registro nuevo recibe un ID auto-generado en una columna `id` que se
agrega al final de cada hoja:

- Notas → `NOT-000001`, `NOT-000002`, …
- Alumnas → `ALU-000001`, `ALU-000002`, …
- Profesores → `PRO-000001`, `PRO-000002`, …

El ID se muestra en el mensaje de éxito al guardar y queda en la hoja para
poder referenciar el registro sin depender del nombre (que es propenso a
errores de tipeo). **Los registros que ya estaban en la planilla no son
modificados** — solo los nuevos llevan ID; si se quiere migrar los antiguos
se puede hacer manualmente desde Excel.

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
generador_reportes/
├── pyproject.toml                 # metadata + dependencias
├── main.py                        # punto de entrada
├── README.md
├── .gitignore
├── GeneradorReportes.spec         # config PyInstaller
├── esquema.xlsx                   # datos fuente (NO se commitea)
├── reportes_generados/            # salida (NO se commitea)
├── backups/                       # backups automáticos (NO se commitea)
└── reportes/                      # paquete principal
    ├── __init__.py
    ├── __main__.py                # `python -m reportes`
    ├── config.py                  # rutas (EXCEL_PATH, OUTPUT_DIR)
    ├── core/                      # lógica de negocio (sin tkinter)
    │   ├── __init__.py
    │   ├── excel_loader.py        # lectura de esquema.xlsx
    │   ├── excel_writer.py        # escritura + lock + backup
    │   ├── docx_helpers.py        # utilidades de docx
    │   ├── acta_examen.py         # generador del Acta de Examen
    │   └── analitico.py           # generador del Analítico
    └── ui/                        # capa de presentación (tkinter)
        ├── __init__.py
        ├── theme.py               # paleta + apply_theme()
        ├── app.py                 # ventana principal
        ├── viewer.py              # ventana «Reportes generados»
        └── data_entry.py          # ventana «Cargar datos al Excel»
```

### Capas

```
main.py
  └── reportes.ui.app
        ├── reportes.core ── reportes.config
        │       └── (excel_loader, acta_examen, analitico, docx_helpers)
        ├── reportes.ui.theme
        └── reportes.ui.viewer
```

`reportes.core` no depende de tkinter y puede usarse desde un script CLI o
desde tests sin levantar la UI.

## Licencia

Uso interno del Estudiantado Santa Mariana de Jesús.
