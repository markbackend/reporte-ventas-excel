# Reporte de ventas automático en Excel

Script en Python que toma una **planilla de ventas** y genera en segundos un **reporte en Excel** con resumen,
tablas, gráficos y control de errores. Lo que a mano lleva horas de tablas dinámicas, acá es un comando.

Hecho con **pandas** (cálculos) y **openpyxl** (formato y gráficos nativos de Excel, editables).

## Qué genera

| Hoja | Contenido |
|---|---|
| **Resumen** | Ventas totales, operaciones, unidades, ticket promedio, período, mejor mes, categoría, vendedor y producto |
| **Por mes** | Ventas por mes con variación respecto al mes anterior y gráfico de línea |
| **Por categoría** | Ventas y % del total, gráfico de torta |
| **Por vendedor** | Ranking de vendedores con gráfico de barras |
| **Top productos** | Los 10 productos que más facturaron |
| **Datos** | Todas las ventas válidas, con filtros |
| **Errores** | Filas que no se pudieron usar, con el **número de fila** y el **motivo** para corregirlas |

Mirá [`ejemplo/reporte_ejemplo.xlsx`](ejemplo/reporte_ejemplo.xlsx), generado a partir de
[`ejemplo/ventas.xlsx`](ejemplo/ventas.xlsx) (400 ventas inventadas, con 4 errores a propósito).

## Planilla de entrada

Columnas necesarias (en cualquier orden; acepta mayúsculas, tildes y espacios):

| fecha | producto | categoria | vendedor | cantidad | precio_unitario |
|---|---|---|---|---|---|
| 15/01/2026 | Mouse inalámbrico | Electrónica | Ana | 2 | 12000 |

También entiende `Precio` y `Unidades`. Las columnas extra se ignoran.

**Validaciones:** fecha inválida, cantidad cero o negativa, precio negativo o no numérico y campos vacíos.
Esas filas no rompen el reporte: van a la hoja *Errores*.

## Cómo usarlo

```bash
python -m venv .venv
.venv\Scripts\activate        # en Linux/Mac: source .venv/bin/activate
pip install -r requirements.txt

python -m reporte ventas.xlsx                          # crea ventas_reporte.xlsx
python -m reporte ventas.xlsx -o reporte_junio.xlsx    # elegir el nombre
python -m reporte ventas.xlsx --desde 2026-03-01 --hasta 2026-03-31   # solo marzo
```

Salida en consola:

```
Reporte generado: ventas_reporte.xlsx
  Ventas incluidas: 396
  Filas con errores (no incluidas): 4 -> ver hoja 'Errores'
```

Para crear de nuevo la planilla de ejemplo: `python ejemplo/generar_ejemplo.py`.

## Pruebas

```bash
pytest -q
```

8 pruebas: nombres de columnas, validación con fila y motivo, cálculos, filtro de fechas, estructura del
Excel generado (hojas, valores y gráficos) y la línea de comandos.

## Estructura

```
reporte/
  datos.py      # lectura y validación
  analisis.py   # cálculos
  excel.py      # formato y gráficos
  __main__.py   # línea de comandos
ejemplo/
tests/
```

## Adaptaciones posibles

Leer CSV o Google Sheets, enviar el reporte por email todos los lunes, sumar otras métricas (margen,
clientes, zonas) o generar también un PDF.
