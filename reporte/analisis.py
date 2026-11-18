"""Cálculos del reporte."""

import pandas as pd


def _agrupar(df: pd.DataFrame, columna: str) -> pd.DataFrame:
    tabla = (
        df.groupby(columna, as_index=False)
        .agg(ventas=("total", "sum"), unidades=("cantidad", "sum"), operaciones=("total", "size"))
        .sort_values("ventas", ascending=False, ignore_index=True, kind="stable")
    )
    total = tabla["ventas"].sum()
    tabla["participacion"] = tabla["ventas"] / total if total else 0.0
    return tabla


def por_mes(df: pd.DataFrame) -> pd.DataFrame:
    tabla = _agrupar(df, "mes").sort_values("mes", ignore_index=True)
    tabla["variacion"] = tabla["ventas"].pct_change()
    return tabla


def por_categoria(df: pd.DataFrame) -> pd.DataFrame:
    return _agrupar(df, "categoria")


def por_vendedor(df: pd.DataFrame) -> pd.DataFrame:
    return _agrupar(df, "vendedor")


def top_productos(df: pd.DataFrame, n: int = 10) -> pd.DataFrame:
    return _agrupar(df, "producto").head(n)


def resumen(df: pd.DataFrame) -> dict:
    if df.empty:
        return {"Ventas totales": 0.0, "Operaciones": 0, "Unidades vendidas": 0, "Ticket promedio": 0.0}
    return {
        "Ventas totales": float(df["total"].sum()),
        "Operaciones": int(len(df)),
        "Unidades vendidas": int(df["cantidad"].sum()),
        "Ticket promedio": float(df["total"].mean()),
        "Período": f"{df['fecha'].min():%d/%m/%Y} al {df['fecha'].max():%d/%m/%Y}",
        "Mejor mes": por_mes(df).sort_values("ventas", ascending=False).iloc[0]["mes"],
        "Mejor categoría": por_categoria(df).iloc[0]["categoria"],
        "Mejor vendedor": por_vendedor(df).iloc[0]["vendedor"],
        "Producto más vendido": top_productos(df, 1).iloc[0]["producto"],
    }
