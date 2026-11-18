import pandas as pd
import pytest
from openpyxl import load_workbook

from reporte import analisis
from reporte.__main__ import main
from reporte.datos import PlanillaInvalida, filtrar_fechas, leer_ventas, normalizar_columna
from reporte.excel import escribir_reporte


def planilla(**cambios):
    datos = {
        "Fecha": ["01/01/2026", "15/01/2026", "03/02/2026", "10/02/2026"],
        "Producto": ["Mouse", "Teclado", "Mouse", "Lámpara"],
        "Categoría": ["Electrónica", "Electrónica", "Electrónica", "Hogar"],
        "Vendedor": ["Ana", "Bruno", "Ana", "Bruno"],
        "Cantidad": [2, 1, 3, 1],
        "Precio unitario": [100, 400, 100, 250],
    }
    datos.update(cambios)
    return pd.DataFrame(datos)


def test_normalizar_columna():
    assert normalizar_columna(" Precio Unitario ") == "precio_unitario"
    assert normalizar_columna("Categoría") == "categoria"
    assert normalizar_columna("Precio") == "precio_unitario"


def test_lectura_valida():
    ventas, errores = leer_ventas(planilla())
    assert len(ventas) == 4 and errores.empty
    assert ventas["total"].tolist() == [200, 400, 300, 250]
    assert ventas["mes"].tolist() == ["2026-01", "2026-01", "2026-02", "2026-02"]


def test_errores_con_fila_y_motivo():
    ventas, errores = leer_ventas(planilla(Cantidad=[2, -1, 3, 1], Vendedor=["Ana", "Bruno", None, "Bruno"]))
    assert len(ventas) == 2
    assert errores["fila_excel"].tolist() == [3, 4]
    assert errores["motivo"].tolist() == ["cantidad inválida", "vendedor vacío"]


def test_falta_columna():
    with pytest.raises(PlanillaInvalida, match="vendedor"):
        leer_ventas(planilla().drop(columns="Vendedor"))


def test_analisis():
    ventas, _ = leer_ventas(planilla())
    meses = analisis.por_mes(ventas)
    assert meses["ventas"].tolist() == [600, 550]
    assert meses["variacion"].iloc[1] == pytest.approx(550 / 600 - 1)
    categorias = analisis.por_categoria(ventas)
    assert categorias.iloc[0]["categoria"] == "Electrónica"
    assert categorias["participacion"].sum() == pytest.approx(1)
    datos = analisis.resumen(ventas)
    assert datos["Ventas totales"] == 1150
    assert datos["Mejor vendedor"] == "Bruno"  # 650 contra 500
    assert datos["Producto más vendido"] == "Mouse"


def test_filtrar_fechas():
    ventas, _ = leer_ventas(planilla())
    assert len(filtrar_fechas(ventas, desde="2026-02-01")) == 2
    assert len(filtrar_fechas(ventas, hasta="2026-01-31")) == 2


def test_reporte_excel(tmp_path):
    ventas, errores = leer_ventas(planilla(Cantidad=[2, "x", 3, 1]))
    destino = escribir_reporte(tmp_path / "reporte.xlsx", ventas, errores)
    libro = load_workbook(destino)
    assert libro.sheetnames == ["Resumen", "Por mes", "Por categoría", "Por vendedor", "Top productos", "Datos", "Errores"]
    assert libro["Resumen"]["B3"].value == 750  # sin la fila del Teclado (cantidad "x")
    assert len(libro["Por mes"]._charts) == 1
    assert libro["Errores"]["B2"].value == "cantidad inválida"


def test_linea_de_comandos(tmp_path, capsys):
    entrada = tmp_path / "ventas.xlsx"
    planilla().to_excel(entrada, index=False)
    assert main([str(entrada)]) == 0
    assert (tmp_path / "ventas_reporte.xlsx").exists()
    assert "Ventas incluidas: 4" in capsys.readouterr().out
    assert main([str(tmp_path / "no_existe.xlsx")]) == 1
