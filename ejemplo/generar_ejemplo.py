"""Crea ejemplo/ventas.xlsx con datos inventados (siempre los mismos) y algunas filas con errores a propósito."""

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

PRODUCTOS = {
    "Electrónica": [("Auriculares Bluetooth", 25000), ("Mouse inalámbrico", 12000), ("Teclado mecánico", 48000)],
    "Hogar": [("Lámpara LED", 15000), ("Set de sartenes", 52000), ("Cafetera", 61000)],
    "Deportes": [("Pelota de fútbol", 18000), ("Mancuernas 5 kg", 30000), ("Botella térmica", 14000)],
    "Librería": [("Cuaderno A4", 4500), ("Lapicera gel x10", 6000), ("Mochila", 35000)],
}
VENDEDORES = ["Ana", "Bruno", "Carla", "Diego"]


def generar(filas: int = 400, semilla: int = 7) -> pd.DataFrame:
    azar = random.Random(semilla)
    inicio = date(2026, 1, 1)
    registros = []
    for _ in range(filas):
        categoria = azar.choice(list(PRODUCTOS))
        producto, precio = azar.choice(PRODUCTOS[categoria])
        registros.append(
            {
                "Fecha": inicio + timedelta(days=azar.randint(0, 180)),
                "Producto": producto,
                "Categoría": categoria,
                "Vendedor": azar.choice(VENDEDORES),
                "Cantidad": azar.choice([1, 1, 1, 2, 2, 3, 5]),
                "Precio unitario": round(precio * azar.uniform(0.9, 1.1), 2),
            }
        )
    df = pd.DataFrame(registros).sort_values("Fecha", ignore_index=True)
    df["Fecha"] = pd.to_datetime(df["Fecha"])
    # Errores típicos de una planilla real, para mostrar la validación:
    df = df.astype({"Fecha": object, "Cantidad": object, "Precio unitario": object})
    df.loc[10, "Cantidad"] = -2
    df.loc[55, "Fecha"] = "no sé"
    df.loc[120, "Vendedor"] = None
    df.loc[200, "Precio unitario"] = "gratis"
    return df


if __name__ == "__main__":
    destino = Path(__file__).with_name("ventas.xlsx")
    generar().to_excel(destino, index=False)
    print(f"Ejemplo creado: {destino}")
