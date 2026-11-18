"""Escribe el reporte en Excel con formato y gráficos nativos (editables en Excel)."""

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import analisis

AZUL = "1F4E78"
CELESTE = "DDEBF7"
MONEDA = '"$"#,##0.00'
PORCENTAJE = "0.0%"
FINO = Side(style="thin", color="BFBFBF")

TITULOS = {
    "mes": "Mes",
    "categoria": "Categoría",
    "vendedor": "Vendedor",
    "producto": "Producto",
    "ventas": "Ventas",
    "unidades": "Unidades",
    "operaciones": "Operaciones",
    "participacion": "% del total",
    "variacion": "Variación vs. mes anterior",
}
FORMATOS = {"ventas": MONEDA, "participacion": PORCENTAJE, "variacion": PORCENTAJE}


def _tabla(hoja, df: pd.DataFrame, fila: int = 1, titulos: dict | None = None, formatos: dict | None = None):
    titulos, formatos = titulos or TITULOS, formatos or FORMATOS
    for j, col in enumerate(df.columns, start=1):
        celda = hoja.cell(row=fila, column=j, value=titulos.get(col, col))
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=AZUL)
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, registro in enumerate(df.itertuples(index=False), start=fila + 1):
        for j, (col, valor) in enumerate(zip(df.columns, registro), start=1):
            if pd.isna(valor):
                valor = None
            elif isinstance(valor, pd.Timestamp):
                valor = valor.to_pydatetime()
            celda = hoja.cell(row=i, column=j, value=valor)
            celda.border = Border(top=FINO, bottom=FINO, left=FINO, right=FINO)
            if col in formatos:
                celda.number_format = formatos[col]
            if (i - fila) % 2 == 0:
                celda.fill = PatternFill("solid", fgColor=CELESTE)
    for j, col in enumerate(df.columns, start=1):
        largo = max([len(str(titulos.get(col, col)))] + [len(str(v)) for v in df[col].head(200)])
        hoja.column_dimensions[get_column_letter(j)].width = min(max(12, largo + 3), 45)
    hoja.freeze_panes = hoja.cell(row=fila + 1, column=1)


def _grafico(hoja, tipo, titulo: str, filas: int, ancla: str, columna_valor: int = 2):
    if filas == 0:
        return
    grafico = tipo()
    grafico.title = titulo
    grafico.height, grafico.width = 8, 16
    datos = Reference(hoja, min_col=columna_valor, min_row=1, max_row=filas + 1)
    categorias = Reference(hoja, min_col=1, min_row=2, max_row=filas + 1)
    grafico.add_data(datos, titles_from_data=True)
    grafico.set_categories(categorias)
    if tipo is not PieChart:
        grafico.legend = None
        grafico.y_axis.numFmt = '"$"#,##0'
    hoja.add_chart(grafico, ancla)


def _hoja_resumen(libro: Workbook, datos: dict, errores: int):
    hoja = libro.active
    hoja.title = "Resumen"
    hoja["A1"] = "Reporte de ventas"
    hoja["A1"].font = Font(bold=True, size=16, color=AZUL)
    fila = 3
    for clave, valor in datos.items():
        hoja.cell(row=fila, column=1, value=clave).font = Font(bold=True)
        celda = hoja.cell(row=fila, column=2, value=valor)
        if clave in ("Ventas totales", "Ticket promedio"):
            celda.number_format = MONEDA
        fila += 1
    if errores:
        aviso = hoja.cell(row=fila + 1, column=1, value=f"⚠ {errores} filas con errores no se incluyeron (ver hoja Errores).")
        aviso.font = Font(bold=True, color="C00000")
    hoja.column_dimensions["A"].width = 24
    hoja.column_dimensions["B"].width = 28


def escribir_reporte(destino: str | Path, ventas: pd.DataFrame, errores: pd.DataFrame) -> Path:
    destino = Path(destino)
    libro = Workbook()
    _hoja_resumen(libro, analisis.resumen(ventas), len(errores))

    meses = analisis.por_mes(ventas)
    hoja = libro.create_sheet("Por mes")
    _tabla(hoja, meses)
    _grafico(hoja, LineChart, "Ventas por mes", len(meses), "H2")

    categorias = analisis.por_categoria(ventas)
    hoja = libro.create_sheet("Por categoría")
    _tabla(hoja, categorias)
    _grafico(hoja, PieChart, "Participación por categoría", len(categorias), "G2")

    vendedores = analisis.por_vendedor(ventas)
    hoja = libro.create_sheet("Por vendedor")
    _tabla(hoja, vendedores)
    _grafico(hoja, BarChart, "Ventas por vendedor", len(vendedores), "G2")

    productos = analisis.top_productos(ventas)
    hoja = libro.create_sheet("Top productos")
    _tabla(hoja, productos)
    _grafico(hoja, BarChart, "Top 10 productos", len(productos), "G2")

    hoja = libro.create_sheet("Datos")
    columnas = ["fecha", "producto", "categoria", "vendedor", "cantidad", "precio_unitario", "total"]
    _tabla(
        hoja,
        ventas[columnas],
        titulos={**TITULOS, "fecha": "Fecha", "cantidad": "Cantidad", "precio_unitario": "Precio unitario", "total": "Total"},
        formatos={"fecha": "DD/MM/YYYY", "precio_unitario": MONEDA, "total": MONEDA},
    )
    hoja.auto_filter.ref = hoja.dimensions

    if len(errores):
        _tabla(libro.create_sheet("Errores"), errores.astype(object), titulos={"fila_excel": "Fila", "motivo": "Motivo"})

    destino.parent.mkdir(parents=True, exist_ok=True)
    libro.save(destino)
    return destino
