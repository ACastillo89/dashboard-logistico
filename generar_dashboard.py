#!/usr/bin/env python3
"""
Generador del Dashboard Comparativo de Pedidos por Mes.

Lee el archivo Excel de pedidos (hoja "DATA"), calcula la cantidad de
pedidos únicos (por N° ORDEN) por día, y genera un archivo HTML
autocontenido (template.html -> dashboard.html) con la tabla comparativa
mes vs. día y el selector de meses.

USO:
    python generar_dashboard.py RUTA_AL_EXCEL.xlsx [--out dashboard.html]

EJEMPLO:
    python generar_dashboard.py PRUEBA_RESUMEN_RUTAS.xlsx

Para actualizar el dashboard en GitHub:
    1. Reemplaza/actualiza el archivo Excel con los datos más recientes.
    2. Corre este script.
    3. Haz commit del nuevo dashboard.html (y del Excel, si también se versiona).

Requiere: pandas, openpyxl (pip install pandas openpyxl)
"""

import argparse
import calendar
import json
import sys
from pathlib import Path

import pandas as pd

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

SHEET_NAME = "DATA"
COL_FECHA = "FECHA DE SOLICITUD"
COL_ORDEN = "N° ORDEN"

SCRIPT_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = SCRIPT_DIR / "template.html"
DATA_MARKER = "__DATA_JSON__"


def construir_datos(ruta_excel: Path) -> dict:
    """Lee el Excel y arma el diccionario {mes: {...}} usado por el dashboard."""
    df = pd.read_excel(ruta_excel, sheet_name=SHEET_NAME)

    # Filas válidas: con fecha y con N° de pedido asignado
    df = df[df[COL_FECHA].notna()].copy()
    df = df[df[COL_ORDEN].notna()].copy()

    df["N_ORDEN"] = df[COL_ORDEN].astype(str).str.strip()
    df[COL_FECHA] = pd.to_datetime(df[COL_FECHA])
    df["mes_key"] = df[COL_FECHA].dt.strftime("%Y-%m")
    df["dia"] = df[COL_FECHA].dt.day

    resultado = {}
    for mes_key, grupo in df.groupby("mes_key"):
        anio, mes = (int(x) for x in mes_key.split("-"))
        nombre = f"{MESES_ES[mes]} {anio}"
        dias_en_mes = calendar.monthrange(anio, mes)[1]

        # cantidad de pedidos ÚNICOS (por N° ORDEN) por día
        conteo_por_dia = (
            grupo.drop_duplicates("N_ORDEN")
            .groupby("dia")["N_ORDEN"]
            .count()
            .to_dict()
        )
        conteos = [int(conteo_por_dia.get(d, 0)) for d in range(1, dias_en_mes + 1)]

        resultado[mes_key] = {
            "nombre": nombre,
            "dias_en_mes": dias_en_mes,
            "conteos": conteos,
            "total": sum(conteos),
        }

    # Ordenar por clave de mes (YYYY-MM) para que el dict quede prolijo
    return {k: resultado[k] for k in sorted(resultado)}


def generar_html(datos: dict, ruta_salida: Path) -> None:
    """Inserta el JSON de datos en la plantilla y escribe el HTML final."""
    if not TEMPLATE_PATH.exists():
        sys.exit(
            f"No se encontró la plantilla en {TEMPLATE_PATH}. "
            "Asegúrate de que template.html esté junto a este script."
        )

    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    if DATA_MARKER not in template:
        sys.exit(f"La plantilla no contiene el marcador {DATA_MARKER}.")

    data_json = json.dumps(datos, ensure_ascii=False)
    html_final = template.replace(DATA_MARKER, data_json)
    ruta_salida.write_text(html_final, encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("excel", type=Path, help="Ruta al archivo Excel de pedidos")
    parser.add_argument(
        "--out", type=Path, default=Path("dashboard.html"),
        help="Ruta del HTML de salida (default: dashboard.html)",
    )
    args = parser.parse_args()

    if not args.excel.exists():
        sys.exit(f"No se encontró el archivo: {args.excel}")

    datos = construir_datos(args.excel)

    if not datos:
        sys.exit("No se encontraron pedidos válidos (con fecha y N° de orden) en el Excel.")

    generar_html(datos, args.out)

    print(f"Dashboard generado en: {args.out.resolve()}")
    for mes_key, info in datos.items():
        print(f"  - {info['nombre']}: {info['total']} pedidos ({info['dias_en_mes']} días)")


if __name__ == "__main__":
    main()
