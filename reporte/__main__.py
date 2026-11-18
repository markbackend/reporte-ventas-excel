"""Uso: python -m reporte ventas.xlsx [-o reporte.xlsx] [--desde 2026-01-01] [--hasta 2026-06-30]"""

import argparse
import sys
from pathlib import Path

from .datos import PlanillaInvalida, filtrar_fechas, leer_ventas
from .excel import escribir_reporte


def main(argumentos: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="reporte", description="Genera un reporte de ventas en Excel.")
    parser.add_argument("entrada", type=Path, help="Planilla de ventas (.xlsx)")
    parser.add_argument("-o", "--salida", type=Path, help="Archivo de salida (por defecto: <entrada>_reporte.xlsx)")
    parser.add_argument("--desde", help="Fecha inicial, AAAA-MM-DD")
    parser.add_argument("--hasta", help="Fecha final, AAAA-MM-DD")
    args = parser.parse_args(argumentos)

    if not args.entrada.exists():
        print(f"No existe el archivo {args.entrada}", file=sys.stderr)
        return 1
    try:
        ventas, errores = leer_ventas(args.entrada)
    except PlanillaInvalida as error:
        print(f"Planilla inválida: {error}", file=sys.stderr)
        return 1

    ventas = filtrar_fechas(ventas, args.desde, args.hasta)
    salida = args.salida or args.entrada.with_name(args.entrada.stem + "_reporte.xlsx")
    escribir_reporte(salida, ventas, errores)

    print(f"Reporte generado: {salida}")
    print(f"  Ventas incluidas: {len(ventas)}")
    if len(errores):
        print(f"  Filas con errores (no incluidas): {len(errores)} -> ver hoja 'Errores'")
    return 0


if __name__ == "__main__":
    sys.exit(main())
