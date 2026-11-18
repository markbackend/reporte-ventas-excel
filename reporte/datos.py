"""Lectura y validación de la planilla de ventas."""

import unicodedata
from pathlib import Path

import pandas as pd

COLUMNAS = ["fecha", "producto", "categoria", "vendedor", "cantidad", "precio_unitario"]
SINONIMOS = {"precio": "precio_unitario", "unidades": "cantidad", "categoría": "categoria"}


class PlanillaInvalida(ValueError):
    pass


def normalizar_columna(nombre) -> str:
    """'Precio Unitario ' -> 'precio_unitario', 'Categoría' -> 'categoria'."""
    texto = str(nombre).strip().lower()
    texto = SINONIMOS.get(texto, texto)
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    texto = "_".join(texto.split())
    return SINONIMOS.get(texto, texto)


def leer_ventas(origen: str | Path | pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Devuelve (ventas válidas, filas con errores).

    Cada error indica la fila de Excel (contando el encabezado) y el motivo, para que se pueda corregir.
    """
    crudo = origen.copy() if isinstance(origen, pd.DataFrame) else pd.read_excel(origen)
    crudo.columns = [normalizar_columna(c) for c in crudo.columns]
    faltan = [c for c in COLUMNAS if c not in crudo.columns]
    if faltan:
        raise PlanillaInvalida("Faltan columnas: " + ", ".join(faltan))

    df = crudo[COLUMNAS].copy()
    df["fila_excel"] = df.index + 2
    for col in ("producto", "categoria", "vendedor"):
        df[col] = df[col].astype("string").str.strip()
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce", dayfirst=True)
    df["cantidad"] = pd.to_numeric(df["cantidad"], errors="coerce")
    df["precio_unitario"] = pd.to_numeric(df["precio_unitario"], errors="coerce")

    motivos = pd.Series("", index=df.index, dtype="object")

    def marcar(mascara, texto):
        mascara = mascara.fillna(True).astype(bool)
        motivos[mascara] = motivos[mascara] + texto + "; "

    for col in ("producto", "categoria", "vendedor"):
        marcar(df[col].isna() | (df[col] == ""), f"{col} vacío")
    marcar(df["fecha"].isna(), "fecha inválida")
    marcar(df["cantidad"].isna() | (df["cantidad"] <= 0), "cantidad inválida")
    marcar(df["precio_unitario"].isna() | (df["precio_unitario"] < 0), "precio inválido")

    con_error = motivos != ""
    errores = crudo.loc[con_error].copy()
    errores.insert(0, "motivo", motivos[con_error].str.rstrip("; "))
    errores.insert(0, "fila_excel", df.loc[con_error, "fila_excel"])

    validas = df.loc[~con_error].copy()
    validas["total"] = validas["cantidad"] * validas["precio_unitario"]
    validas["mes"] = validas["fecha"].dt.to_period("M").astype(str)
    return validas.reset_index(drop=True), errores.reset_index(drop=True)


def filtrar_fechas(df: pd.DataFrame, desde: str | None = None, hasta: str | None = None) -> pd.DataFrame:
    if desde:
        df = df[df["fecha"] >= pd.Timestamp(desde)]
    if hasta:
        df = df[df["fecha"] <= pd.Timestamp(hasta)]
    return df.reset_index(drop=True)
