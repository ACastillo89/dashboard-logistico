# -*- coding: utf-8 -*-
"""
Dashboard de Gestión Logística y Análisis de Pedidos
=====================================================
Herramienta interactiva en Streamlit que permite cargar un archivo Excel
en cualquier momento y recalcula automáticamente todos los KPIs, gráficos,
rankings y reportes, sin necesidad de modificar el código.

Autor: Generado con Claude (Anthropic)
"""

import io
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONFIGURACIÓN GENERAL DE LA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Dashboard Logístico | Análisis de Pedidos",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded",
)

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre",
    12: "Diciembre",
}
DIAS_ES = {
    0: "Lunes", 1: "Martes", 2: "Miércoles", 3: "Jueves",
    4: "Viernes", 5: "Sábado", 6: "Domingo",
}

# Paleta de colores profesional (usada en gráficos y encabezados de tablas
# en vez de los colores por defecto de Plotly/Streamlit).
PALETA = ["#2563EB", "#0EA5E9", "#10B981", "#8B5CF6", "#F59E0B", "#EC4899", "#14B8A6", "#F43F5E"]
COLOR_HEADER_TABLA = "#1E3A5F"  # azul marino, fondo de encabezado de tabla
COLOR_PRINCIPAL = "#2563EB"     # azul, color base de gráficos de una sola serie

# ============================================================
# UTILIDADES DE NORMALIZACIÓN Y DETECCIÓN DE COLUMNAS
# ============================================================

def normalizar_texto(texto: str) -> str:
    """Convierte a mayúsculas, quita tildes/símbolos y espacios extra."""
    if texto is None:
        return ""
    texto = str(texto)
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("utf-8")
    texto = texto.upper().strip()
    texto = "".join(ch for ch in texto if ch.isalnum())
    return texto


# Candidatos posibles de nombres de columna para cada campo lógico simple
# (campos que casi siempre vienen como una columna directa y sin ambigüedad).
CANDIDATOS_COLUMNAS = {
    "orden": ["NORDEN", "NUMERODEORDEN", "NROORDEN", "NUMORDEN", "ORDEN", "IDPEDIDO", "NUMEROPEDIDO", "PEDIDO"],
    "ciudad": ["CIUDAD", "CIUDADDESTINO", "MUNICIPIO", "LOCALIDAD"],
}

# Nombres que indican explícitamente una columna de RUTA/ZONA ya existente.
CANDIDATOS_RUTA_EXPLICITA = ["RUTA", "RUTAENTREGA", "RUTADEENTREGA", "ZONA", "RUTALOGISTICA", "RUTAREPARTO"]

# Nombres de columnas de texto (dirección / punto de entrega) desde las que se
# puede DERIVAR la ruta cuando el archivo no trae una columna de ruta explícita
# (como ocurre en el formato real: "CLIENTE / DIRECCIÓN / ZONA" dentro de "PTO ENTREGA").
CANDIDATOS_FUENTE_RUTA = ["PTOENTREGA", "PUNTOENTREGA", "DIRECCION", "DIRECCIONENTREGA", "LUGARENTREGA"]

# Nombres de columnas de fecha, en orden de prioridad para el valor por defecto.
# FECHA DE SOLICITUD va primero porque es la fecha estándar de análisis para
# este negocio (fecha en que se solicitó el pedido).
CANDIDATOS_FECHA_PRIORIDAD = [
    "FECHADESOLICITUD", "FECHA", "FECHADEENTREGA", "FECHAENTREGA",
    "FECHADEPEDIDO", "FECHAPEDIDO", "FECHAFACTURA", "FECHADESPACHO",
]

# Columna opcional de tipo de movimiento (si existe, se ofrece como filtro extra).
CANDIDATOS_TIPO_MOVIMIENTO = ["TIPO", "TIPODEMOVIMIENTO", "TIPOMOVIMIENTO", "TIPOPEDIDO"]

CAMPOS_REQUERIDOS = {
    "orden": "N° ORDEN (identificador único de pedido)",
    "ciudad": "Ciudad",
}


def _buscar_columna(columnas_normalizadas: dict, candidatos: list) -> str:
    """Busca la primera columna cuyo nombre normalizado coincida (exacto o
    parcial) con alguno de los candidatos dados."""
    for col_original, col_norm in columnas_normalizadas.items():
        if col_norm in candidatos:
            return col_original
    for col_original, col_norm in columnas_normalizadas.items():
        if any(cand in col_norm or col_norm in cand for cand in candidatos):
            return col_original
    return None


def detectar_columnas(df: pd.DataFrame) -> dict:
    """Detecta automáticamente las columnas simples (orden, ciudad).

    Si el archivo no trae una columna de "Ciudad" con ese nombre explícito
    (como ocurre en el formato real de este negocio, donde la ubicación se
    guarda en la columna "RUTA": GUAYAQUIL, QUITO, CUENCA, etc.), se usa esa
    misma columna de ruta como Ciudad por defecto.
    """
    columnas_normalizadas = {col: normalizar_texto(col) for col in df.columns}
    mapeo = {}
    for campo, candidatos in CANDIDATOS_COLUMNAS.items():
        mapeo[campo] = _buscar_columna(columnas_normalizadas, candidatos)

    if mapeo.get("ciudad") is None:
        mapeo["ciudad"] = detectar_ruta_explicita(df)

    return mapeo


def detectar_columnas_fecha(df: pd.DataFrame) -> list:
    """Devuelve la lista de columnas que parecen ser de tipo fecha, ya sea
    por su nombre o por su dtype, ordenadas por prioridad de uso."""
    columnas_normalizadas = {col: normalizar_texto(col) for col in df.columns}
    candidatas = []
    for col, col_norm in columnas_normalizadas.items():
        es_por_nombre = any(cand in col_norm or col_norm in cand for cand in CANDIDATOS_FECHA_PRIORIDAD)
        es_por_tipo = pd.api.types.is_datetime64_any_dtype(df[col])
        # Se ignoran columnas casi vacías (ej. columnas fantasma tipo "Unnamed: N"
        # con un par de celdas sueltas) para no ensuciar el selector.
        tiene_datos_suficientes = len(df) == 0 or df[col].notna().mean() >= 0.3
        if (es_por_nombre or es_por_tipo) and tiene_datos_suficientes:
            candidatas.append(col)

    def prioridad(col):
        norm = columnas_normalizadas[col]
        if norm in CANDIDATOS_FECHA_PRIORIDAD:
            return CANDIDATOS_FECHA_PRIORIDAD.index(norm)
        return len(CANDIDATOS_FECHA_PRIORIDAD) + 1

    return sorted(candidatas, key=prioridad)


def detectar_ruta_explicita(df: pd.DataFrame) -> str:
    columnas_normalizadas = {col: normalizar_texto(col) for col in df.columns}
    return _buscar_columna(columnas_normalizadas, CANDIDATOS_RUTA_EXPLICITA)


def detectar_fuente_derivacion_ruta(df: pd.DataFrame) -> str:
    columnas_normalizadas = {col: normalizar_texto(col) for col in df.columns}
    return _buscar_columna(columnas_normalizadas, CANDIDATOS_FUENTE_RUTA)


def detectar_columna_tipo_movimiento(df: pd.DataFrame) -> str:
    columnas_normalizadas = {col: normalizar_texto(col) for col in df.columns}
    return _buscar_columna(columnas_normalizadas, CANDIDATOS_TIPO_MOVIMIENTO)


def derivar_ruta_desde_texto(serie: pd.Series, modo: str = "segmento_final") -> pd.Series:
    """Deriva la ruta/zona de entrega a partir de un campo de texto libre.

    En el formato real observado, la columna 'PTO ENTREGA' tiene la forma
    'CLIENTE / DIRECCIÓN / ZONA-RUTA' (ej: 'FANTAPE / KM 7 / GYE NORTE 1').
    El último segmento después de la última '/' es, en la práctica, la
    ruta o zona de reparto.
    """
    texto = serie.astype(str).str.strip()
    if modo == "segmento_final":
        resultado = texto.str.split("/").str[-1].str.strip()
    else:
        resultado = texto
    resultado = resultado.replace({"nan": None, "None": None, "": None})
    return resultado


def limpiar_valor_orden(valor) -> str:
    """Normaliza el identificador de pedido a texto, evitando artefactos
    como '100287670.0' cuando Excel entrega el número como float."""
    if pd.isna(valor):
        return None
    if isinstance(valor, float):
        if valor.is_integer():
            return str(int(valor))
        return str(valor)
    return str(valor).strip()


def centrar(obj):
    """Centra el texto de todas las celdas (y encabezados) de una tabla,
    ya sea un DataFrame o un Styler ya con formato/resaltado aplicado.
    Además le da al encabezado un fondo de color profesional."""
    styler = obj.style if isinstance(obj, pd.DataFrame) else obj
    return styler.set_properties(**{"text-align": "center"}).set_table_styles(
        [{
            "selector": "th",
            "props": [
                ("text-align", "center"),
                ("background-color", COLOR_HEADER_TABLA),
                ("color", "#ffffff"),
                ("font-weight", "700"),
            ],
        }]
    )


def formatear_numeros(df: pd.DataFrame, na_rep: str = "—") -> pd.DataFrame:
    """Convierte cada celda numérica a texto (sin decimales), reemplazando
    vacíos por `na_rep`. st.dataframe NO respeta Styler.format()/na_rep, así
    que hay que dejar el texto ya armado para que se vea bien."""
    return df.apply(lambda col: col.map(lambda v: na_rep if pd.isna(v) else f"{v:.0f}"))


# ============================================================
# CARGA Y PREPROCESAMIENTO DE DATOS
# ============================================================

@st.cache_data(show_spinner=False)
def leer_excel(archivo_bytes: bytes, nombre_hoja=None) -> pd.DataFrame:
    """Lee el Excel subido. Cacheado por contenido del archivo: si subes
    un archivo distinto, el cache se invalida automáticamente."""
    excel_file = pd.ExcelFile(io.BytesIO(archivo_bytes))
    hoja = nombre_hoja if nombre_hoja else excel_file.sheet_names[0]
    df = pd.read_excel(excel_file, sheet_name=hoja)
    df.columns = [str(c).strip() for c in df.columns]
    # Muchos exportadores de Excel dejan miles de filas totalmente vacías
    # después de los datos reales; se descartan aquí por rendimiento y
    # para que la detección de tipos de columna no se vea afectada.
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def preparar_datos(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    """Aplica el mapeo de columnas, limpia datos y genera columnas derivadas
    (año, mes, semana, día) necesarias para todos los análisis.

    `config` espera las claves:
        - orden: nombre de la columna con el N° de orden
        - fecha: nombre de la columna de fecha a usar para el análisis
        - ciudad: nombre de la columna de ciudad
        - ruta_modo: "columna" (usar una columna de ruta ya existente),
                     "ciudad" (usar la misma columna de Ciudad como Ruta), o
                     "derivada" (extraer la ruta desde un campo de texto)
        - ruta_columna: nombre de columna a usar cuando ruta_modo == "columna"
        - ruta_fuente: nombre de columna fuente cuando ruta_modo == "derivada"
        - ruta_derivacion_modo: "segmento_final" o "valor_completo"
        - tipo_movimiento: nombre de columna opcional (ej. "TIPO") o None
    """
    df = df.copy()

    # ---- Extraer las columnas fuente ANTES de tocar nombres ----
    # (se guardan como Series independientes para evitar el problema de
    # terminar con dos columnas con el mismo nombre, por ejemplo si el
    # archivo ya tiene una columna "FECHA" y además se elige otra columna
    # como fecha de análisis y también se renombra a "FECHA").
    orden_serie = df[config["orden"]]
    fecha_serie = df[config["fecha"]]
    ciudad_serie = df[config["ciudad"]]

    if config["ruta_modo"] == "columna":
        ruta_fuente_serie = df[config["ruta_columna"]]
    elif config["ruta_modo"] == "derivada":
        ruta_fuente_serie = df[config["ruta_fuente"]]
    else:
        ruta_fuente_serie = None  # se usará CIUDAD ya procesada

    tipo_serie = df[config["tipo_movimiento"]] if config.get("tipo_movimiento") else None

    # ---- Eliminar cualquier columna original que choque con los nombres
    # de destino, para evitar columnas duplicadas tras la asignación ----
    columnas_destino = ["N_ORDEN", "FECHA", "CIUDAD", "RUTA", "TIPO_MOVIMIENTO"]
    df = df.drop(columns=[c for c in columnas_destino if c in df.columns], errors="ignore")

    # ---- Orden, Fecha, Ciudad ----
    df["N_ORDEN"] = orden_serie.apply(limpiar_valor_orden)
    df["FECHA"] = fecha_serie
    df["CIUDAD"] = ciudad_serie.astype(str).str.strip().str.upper().replace({"NAN": None})

    # ---- Ruta: columna directa, igual a Ciudad, o derivada desde texto ----
    if config["ruta_modo"] == "columna":
        df["RUTA"] = ruta_fuente_serie.astype(str).str.strip().str.upper().replace({"NAN": None})
    elif config["ruta_modo"] == "ciudad":
        df["RUTA"] = df["CIUDAD"]
    else:
        ruta_derivada = derivar_ruta_desde_texto(
            ruta_fuente_serie, modo=config.get("ruta_derivacion_modo", "segmento_final")
        )
        df["RUTA"] = ruta_derivada.astype(str).str.strip().str.upper().replace({"NAN": None, "NONE": None})

    # ---- Tipo de movimiento (opcional) ----
    if tipo_serie is not None:
        df["TIPO_MOVIMIENTO"] = tipo_serie.astype(str).str.strip().str.upper()

    # ---- Fecha: forzar a datetime, descartar filas sin fecha válida ----
    df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce", dayfirst=True)
    filas_antes = len(df)
    df = df.dropna(subset=["FECHA"])
    filas_sin_fecha = filas_antes - len(df)

    # ---- Descartar filas sin N° de Orden: no se pueden contar como pedido ----
    filas_antes = len(df)
    df = df.dropna(subset=["N_ORDEN"])
    filas_sin_orden = filas_antes - len(df)

    # ---- Quitar duplicados exactos de línea (mismo pedido/fecha/ciudad/ruta repetido íntegro) ----
    df = df.drop_duplicates()

    # ---- Un mismo N° de Orden a veces trae líneas con fechas distintas
    # (ej. una parte despachada el 31/08 y otra el 01/09). Para que el
    # pedido cuente como UNO SOLO en todos lados (total, por mes, por
    # ciudad, etc.) se usa la fecha MÁS ANTIGUA de ese N° de Orden como su
    # fecha "oficial" para todos los análisis. ----
    fecha_canonica = df.groupby("N_ORDEN")["FECHA"].transform("min")
    df["FECHA"] = fecha_canonica

    # ---- Columnas derivadas de tiempo ----
    df["ANIO"] = df["FECHA"].dt.year
    df["MES_NUM"] = df["FECHA"].dt.month
    df["MES_NOMBRE"] = df["MES_NUM"].map(MESES_ES)
    df["DIA"] = df["FECHA"].dt.day
    df["DIA_SEMANA_NUM"] = df["FECHA"].dt.dayofweek
    df["DIA_SEMANA"] = df["DIA_SEMANA_NUM"].map(DIAS_ES)
    df["SEMANA_ISO"] = df["FECHA"].dt.isocalendar().week.astype(int)
    df["SEMANA_ANIO_ID"] = df["FECHA"].dt.strftime("%G-W%V")
    # Semana dentro del mes (1 a 5), útil para "semana más fuerte del mes"
    df["SEMANA_MES"] = ((df["DIA"] - 1) // 7) + 1
    df["ANIO_MES"] = df["FECHA"].dt.strftime("%Y-%m")
    df["ANIO_MES_LABEL"] = df["MES_NOMBRE"] + " " + df["ANIO"].astype(str)
    df["FECHA_DIA"] = df["FECHA"].dt.date

    df.attrs["filas_sin_fecha"] = filas_sin_fecha
    df.attrs["filas_sin_orden"] = filas_sin_orden
    return df


# ============================================================
# FUNCIONES DE AGREGACIÓN (KPIs, RANKINGS, COMPARATIVOS)
# ============================================================

def pedidos_unicos(df: pd.DataFrame) -> int:
    return df["N_ORDEN"].nunique()


def pedidos_por_dia(df: pd.DataFrame) -> pd.DataFrame:
    out = df.groupby("FECHA_DIA")["N_ORDEN"].nunique().reset_index()
    out.columns = ["FECHA", "PEDIDOS"]
    return out.sort_values("FECHA")


def pedidos_por_semana(df: pd.DataFrame) -> pd.DataFrame:
    out = df.groupby("SEMANA_ANIO_ID")["N_ORDEN"].nunique().reset_index()
    out.columns = ["SEMANA", "PEDIDOS"]
    return out.sort_values("SEMANA")


def pedidos_por_mes(df: pd.DataFrame) -> pd.DataFrame:
    out = df.groupby(["ANIO", "MES_NUM", "MES_NOMBRE", "ANIO_MES"])["N_ORDEN"].nunique().reset_index()
    out.columns = ["ANIO", "MES_NUM", "MES_NOMBRE", "ANIO_MES", "PEDIDOS"]
    return out.sort_values(["ANIO", "MES_NUM"])


def mes_mayor_volumen_por_anio(df: pd.DataFrame) -> pd.DataFrame:
    mensual = pedidos_por_mes(df)
    idx = mensual.groupby("ANIO")["PEDIDOS"].idxmax()
    return mensual.loc[idx].sort_values("ANIO")


def semana_mas_fuerte_por_mes(df: pd.DataFrame) -> pd.DataFrame:
    agrupado = df.groupby(["ANIO", "MES_NUM", "MES_NOMBRE", "SEMANA_MES"])["N_ORDEN"].nunique().reset_index()
    agrupado.columns = ["ANIO", "MES_NUM", "MES_NOMBRE", "SEMANA_MES", "PEDIDOS"]
    idx = agrupado.groupby(["ANIO", "MES_NUM"])["PEDIDOS"].idxmax()
    return agrupado.loc[idx].sort_values(["ANIO", "MES_NUM"])


def ranking(df: pd.DataFrame, campo: str, top_n: int = None) -> pd.DataFrame:
    out = df.groupby(campo)["N_ORDEN"].nunique().reset_index()
    out.columns = [campo, "PEDIDOS"]
    out = out.sort_values("PEDIDOS", ascending=False).reset_index(drop=True)
    out.index = out.index + 1
    if top_n:
        out = out.head(top_n)
    return out


def comparativo_pivot(df: pd.DataFrame, mes_inicio: int = 1, mes_fin: int = 12) -> pd.DataFrame:
    """Tabla Año x Mes con total de pedidos, para comparativos mensuales/anuales.

    `mes_inicio`/`mes_fin` acotan qué filas (meses) se muestran, ej. 8 y 9
    para ver solo Agosto y Septiembre. Si mes_fin < mes_inicio, el rango se
    entiende como que da la vuelta al año (ej. Nov a Feb)."""
    mensual = pedidos_por_mes(df)
    pivot = mensual.pivot_table(index="MES_NOMBRE", columns="ANIO", values="PEDIDOS", aggfunc="sum")
    if mes_fin >= mes_inicio:
        rango_meses = list(range(mes_inicio, mes_fin + 1))
    else:
        rango_meses = list(range(mes_inicio, 13)) + list(range(1, mes_fin + 1))
    orden_meses = [MESES_ES[i] for i in rango_meses]
    pivot = pivot.reindex(orden_meses)
    return pivot


# ============================================================
# EXPORTACIÓN A EXCEL
# ============================================================

def generar_reporte_excel(df_filtrado: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        resumen = pd.DataFrame({
            "Indicador": [
                "Total de pedidos", "Ciudades distintas",
                "Fecha mínima", "Fecha máxima",
            ],
            "Valor": [
                pedidos_unicos(df_filtrado),
                df_filtrado["CIUDAD"].nunique(),
                df_filtrado["FECHA"].min().strftime("%Y-%m-%d") if len(df_filtrado) else "-",
                df_filtrado["FECHA"].max().strftime("%Y-%m-%d") if len(df_filtrado) else "-",
            ],
        })
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        pedidos_por_dia(df_filtrado).to_excel(writer, sheet_name="Pedidos por Dia", index=False)
        pedidos_por_semana(df_filtrado).to_excel(writer, sheet_name="Pedidos por Semana", index=False)
        pedidos_por_mes(df_filtrado).rename(columns={"ANIO": "AÑO"}).to_excel(
            writer, sheet_name="Pedidos por Mes", index=False
        )
        mes_mayor_volumen_por_anio(df_filtrado).rename(columns={"ANIO": "AÑO"}).to_excel(
            writer, sheet_name="Mes Top por Año", index=False
        )
        semana_mas_fuerte_por_mes(df_filtrado).rename(columns={"ANIO": "AÑO"}).to_excel(
            writer, sheet_name="Semana Top por Mes", index=False
        )
        ranking(df_filtrado, "CIUDAD").to_excel(writer, sheet_name="Ranking Ciudades")
        comparativo_pivot(df_filtrado).rename_axis(columns="AÑO").to_excel(
            writer, sheet_name="Comparativo Año-Mes"
        )
        df_filtrado.drop(columns=["FECHA_DIA"], errors="ignore").rename(
            columns={"ANIO": "AÑO"}
        ).to_excel(writer, sheet_name="Detalle Filtrado", index=False)
    return buffer.getvalue()


# ============================================================
# BARRA LATERAL: CARGA DE ARCHIVO Y MAPEO DE COLUMNAS
# ============================================================

st.sidebar.title("🚚 Panel de Control")
st.sidebar.markdown("### 1. Cargar datos")
archivo = st.sidebar.file_uploader(
    "Sube tu archivo Excel actualizado (.xlsx)",
    type=["xlsx", "xls"],
    help="Cada vez que subas un archivo nuevo, todo el dashboard se recalcula automáticamente.",
)

if archivo is not None:
    contenido = archivo.getvalue()
    excel_file_preview = pd.ExcelFile(io.BytesIO(contenido))
    hojas = excel_file_preview.sheet_names
    hoja_seleccionada = hojas[0]
    if len(hojas) > 1:
        hoja_seleccionada = st.sidebar.selectbox("Hoja del Excel a usar", hojas)

    df_crudo = leer_excel(contenido, hoja_seleccionada)

    # Guardar en session_state para persistir entre interacciones de filtros
    st.session_state["archivo_hash"] = hash(contenido) 
    st.session_state["df_crudo"] = df_crudo
    st.session_state["archivo_nombre"] = archivo.name

if "df_crudo" not in st.session_state:
    st.title("🚚 Dashboard de Gestión Logística y Análisis de Pedidos")
    st.info(
        "👈 Sube un archivo Excel desde la barra lateral para comenzar.\n\n"
        "La columna principal para contar pedidos debe ser **N° ORDEN**. "
        "También se necesitan columnas de **Fecha**, **Ciudad** y **Ruta** "
        "(la app intenta detectarlas automáticamente)."
    )
    st.stop()

df_crudo = st.session_state["df_crudo"]
st.sidebar.success(f"Archivo cargado: {st.session_state['archivo_nombre']} ({len(df_crudo)} filas con datos)")

# ---- Detección automática: orden y ciudad (columnas directas) ----
mapeo_auto = detectar_columnas(df_crudo)

# ---- Detección automática: fecha (puede haber varias columnas de fecha) ----
columnas_fecha_detectadas = detectar_columnas_fecha(df_crudo)

# ---- Detección automática: ruta (columna explícita o derivable) ----
col_ruta_explicita = detectar_ruta_explicita(df_crudo)
col_fuente_ruta = detectar_fuente_derivacion_ruta(df_crudo)

# ---- Detección automática: tipo de movimiento (opcional) ----
col_tipo_movimiento_auto = detectar_columna_tipo_movimiento(df_crudo)

# ---- Mapeo de columnas: 100% automático, sin preguntar nada en pantalla ----
mapeo_final = {
    "orden": mapeo_auto.get("orden"),
    "ciudad": mapeo_auto.get("ciudad"),
    "fecha": columnas_fecha_detectadas[0] if columnas_fecha_detectadas else None,
    "tipo_movimiento": col_tipo_movimiento_auto,
}

if col_ruta_explicita and col_ruta_explicita == mapeo_final.get("ciudad"):
    # La columna de Ciudad ya ES la columna de ruta (caso: "RUTA" trae
    # GUAYAQUIL, QUITO, etc.) → se reutiliza, sin pedirla aparte.
    mapeo_final["ruta_modo"] = "ciudad"
else:
    # Solo se considera la columna de Ciudad para todo el análisis de
    # ubicación; no se deriva una Ruta aparte desde PTO ENTREGA u otra
    # columna, aunque el archivo la traiga.
    mapeo_final["ruta_modo"] = "ciudad"

# Validación de columnas obligatorias
faltan = [mapeo_final.get("orden"), mapeo_final.get("ciudad"), mapeo_final.get("fecha")]
ruta_incompleta = (
    (mapeo_final.get("ruta_modo") == "columna" and not mapeo_final.get("ruta_columna"))
    or (mapeo_final.get("ruta_modo") == "derivada" and not mapeo_final.get("ruta_fuente"))
)
if any(f is None for f in faltan) or ruta_incompleta:
    st.error(
        "⚠️ No se pudieron identificar automáticamente todas las columnas necesarias "
        "(N° ORDEN, Fecha, Ciudad, Ruta) en este archivo. Verifica que el Excel tenga "
        "la misma estructura de columnas de siempre."
    )
    st.stop()

# ---- Preprocesamiento ----
df = preparar_datos(df_crudo, mapeo_final)

avisos = []
if df.attrs.get("filas_sin_fecha", 0) > 0:
    avisos.append(
        f"{df.attrs['filas_sin_fecha']} fila(s) descartada(s) por fecha inválida o vacía "
        f"en la columna **{mapeo_final['fecha']}** (probablemente filas vacías/fantasma del Excel)."
    )
if df.attrs.get("filas_sin_orden", 0) > 0:
    avisos.append(f"{df.attrs['filas_sin_orden']} filas descartadas por no tener N° de Orden (no cuentan como pedido).")
for aviso in avisos:
    st.sidebar.warning(aviso)

if df.empty:
    st.error(
        "No quedaron registros válidos después de procesar el archivo. "
        "Revisa el mapeo de columnas de Fecha y N° ORDEN."
    )
    st.stop()

# ============================================================
# FILTROS GLOBALES
# ============================================================
st.sidebar.markdown("### 2. Filtros")

hoy = datetime.now()
anio_vigente = hoy.year
mes_vigente = hoy.month

# ---- Año: por defecto el año vigente; se recalcula solo en cada carga ----
anios_disponibles = sorted(df["ANIO"].unique())
if anio_vigente in anios_disponibles:
    anio_default = [anio_vigente]
elif anios_disponibles:
    anio_default = [max(anios_disponibles)]  # año más reciente disponible en los datos
else:
    anio_default = []

anios_sel = st.sidebar.multiselect(
    "Año",
    anios_disponibles,
    default=anio_default,
    help="Por defecto muestra el año vigente. Agrega otros años si necesitas comparar.",
)

df_tmp = df[df["ANIO"].isin(anios_sel)] if anios_sel else df

# ---- Mes: UN SOLO selector que filtra el dashboard Y alimenta los
# comparativos (Semana y Día). Usa mes+año como identificador (ej. "Agosto
# 2026") para que comparar entre años no sea ambiguo. ----
claves_mes_disponibles = sorted(df_tmp["ANIO_MES"].unique())
etiquetas_mes = dict(zip(df_tmp["ANIO_MES"], df_tmp["ANIO_MES_LABEL"]))
opciones_meses = [etiquetas_mes[k] for k in claves_mes_disponibles]

mes_vigente_key = f"{anio_vigente}-{mes_vigente:02d}"
if mes_vigente_key in claves_mes_disponibles:
    meses_default_labels = [etiquetas_mes[mes_vigente_key]]
elif claves_mes_disponibles:
    meses_default_labels = [etiquetas_mes[claves_mes_disponibles[-1]]]  # mes más reciente disponible
else:
    meses_default_labels = []

meses_sel_labels = st.sidebar.multiselect(
    "Mes",
    opciones_meses,
    default=meses_default_labels,
    help=(
        "Por defecto muestra el mes vigente. Agrega otros meses para comparar "
        "(ej. Agosto 2026 vs Septiembre 2026). Este mismo selector alimenta las "
        "tablas de Pedidos por Semana y el Comparativo por Día."
    ),
)
claves_sel = [k for k in claves_mes_disponibles if etiquetas_mes[k] in meses_sel_labels]

df_tmp = df_tmp[df_tmp["ANIO_MES"].isin(claves_sel)] if claves_sel else df_tmp

semanas_disponibles = sorted(df_tmp["SEMANA_ANIO_ID"].unique())
semanas_sel = st.sidebar.multiselect("Semana (ISO)", semanas_disponibles, default=[])

# Nota: no hay selector de Ciudad, Ruta ni Tipo de Movimiento — el análisis
# siempre considera TODAS las ciudades, TODAS las rutas y TODOS los tipos
# de movimiento, sin necesidad de filtrarlos.

# Aplicar filtros finales
df_filtrado = df.copy()
if anios_sel:
    df_filtrado = df_filtrado[df_filtrado["ANIO"].isin(anios_sel)]
if claves_sel:
    df_filtrado = df_filtrado[df_filtrado["ANIO_MES"].isin(claves_sel)]
if semanas_sel:
    df_filtrado = df_filtrado[df_filtrado["SEMANA_ANIO_ID"].isin(semanas_sel)]

if df_filtrado.empty:
    st.warning("No hay datos para la combinación de filtros seleccionada.")
    st.stop()

# ============================================================
# COMPARATIVO POR MES: SEMANA Y DÍA
# ============================================================
# El día siempre se agrupa por FECHA DE SOLICITUD, sin importar qué columna
# de fecha se eligió para el resto del dashboard. Usa el mismo selector de
# "Mes" de arriba (claves_sel) — no hay un selector aparte.
col_fecha_solicitud = _buscar_columna(
    {c: normalizar_texto(c) for c in df_crudo.columns}, ["FECHADESOLICITUD"]
)

df_fs = None
if col_fecha_solicitud is not None:
    df_fs = pd.DataFrame({
        "N_ORDEN": df_crudo[mapeo_final["orden"]].apply(limpiar_valor_orden),
        "FECHA_SOLICITUD": pd.to_datetime(
            df_crudo[col_fecha_solicitud], errors="coerce", dayfirst=True
        ),
    })
    df_fs = df_fs.dropna(subset=["N_ORDEN", "FECHA_SOLICITUD"])
    # Mismo criterio que en el resto del dashboard: si un N° de Orden tiene
    # líneas en más de una fecha, se usa la más antigua como su fecha única.
    df_fs["FECHA_SOLICITUD"] = df_fs.groupby("N_ORDEN")["FECHA_SOLICITUD"].transform("min")
    ordenes_filtrados = set(df_filtrado["N_ORDEN"].unique())
    df_fs = df_fs[df_fs["N_ORDEN"].isin(ordenes_filtrados)]
    df_fs = df_fs.drop_duplicates(subset=["N_ORDEN", "FECHA_SOLICITUD"])
    if not df_fs.empty:
        df_fs["ANIO"] = df_fs["FECHA_SOLICITUD"].dt.year
        df_fs["MES_NUM"] = df_fs["FECHA_SOLICITUD"].dt.month
        df_fs["DIA"] = df_fs["FECHA_SOLICITUD"].dt.day
        df_fs["MES_KEY"] = df_fs["FECHA_SOLICITUD"].dt.strftime("%Y-%m")
        df_fs["MES_LABEL"] = df_fs["MES_NUM"].map(MESES_ES) + " " + df_fs["ANIO"].astype(str)
    else:
        df_fs = None

# ---- Tabla: Pedidos por Semana del Mes (comparativo) ----
tabla_semanal = None
if claves_sel:
    df_sem_base = df_filtrado[df_filtrado["ANIO_MES"].isin(claves_sel)]
    if not df_sem_base.empty:
        agg_sem = (
            df_sem_base.groupby(["ANIO_MES", "SEMANA_MES"])["N_ORDEN"]
            .nunique()
            .reset_index(name="PEDIDOS")
        )
        tabla_semanal = agg_sem.pivot(index="SEMANA_MES", columns="ANIO_MES", values="PEDIDOS")
        max_semana = int(tabla_semanal.index.max())
        tabla_semanal = tabla_semanal.reindex(range(1, max_semana + 1))
        cols_orden_sem = [k for k in claves_sel if k in tabla_semanal.columns]
        tabla_semanal = tabla_semanal[cols_orden_sem]
        tabla_semanal.columns = [etiquetas_mes[k] for k in cols_orden_sem]
        tabla_semanal.index = [f"Semana {i}" for i in tabla_semanal.index]
        tabla_semanal.index.name = "Semana del mes"

# ---- Tabla y datos largos: Pedidos por Día (comparativo, Fecha de Solicitud) ----
tabla_comparativo = None
totales_comparativo = None
conteo_comp_largo = None
if df_fs is not None and claves_sel:
    df_fs_sel = df_fs[df_fs["MES_KEY"].isin(claves_sel)]
    if not df_fs_sel.empty:
        conteo_comp = (
            df_fs_sel.groupby(["MES_KEY", "DIA"])["N_ORDEN"]
            .nunique()
            .reset_index()
            .rename(columns={"N_ORDEN": "PEDIDOS"})
        )
        tabla_comparativo = conteo_comp.pivot(index="DIA", columns="MES_KEY", values="PEDIDOS")
        max_dia_comp = int(tabla_comparativo.index.max())
        tabla_comparativo = tabla_comparativo.reindex(range(1, max_dia_comp + 1))
        cols_orden_dia = [k for k in claves_sel if k in tabla_comparativo.columns]
        tabla_comparativo = tabla_comparativo[cols_orden_dia]
        tabla_comparativo.columns = [etiquetas_mes[k] for k in cols_orden_dia]
        tabla_comparativo.index.name = "Día"
        totales_comparativo = tabla_comparativo.sum(numeric_only=True)

        # Datos largos para el gráfico, construidos a partir de la MISMA tabla
        # ya completa (todos los días del mes, con 0 donde no hubo pedidos),
        # así el eje X del gráfico muestra siempre el mes entero.
        conteo_comp_largo = (
            tabla_comparativo.fillna(0)
            .astype(int)
            .reset_index()
            .melt(id_vars="Día", var_name="MES_LABEL", value_name="PEDIDOS")
        )

# ============================================================
# ENCABEZADO Y KPIs EJECUTIVOS
# ============================================================
st.title("🚚 Dashboard de Gestión Logística y Análisis de Pedidos")
st.caption(f"Última actualización de datos: **{st.session_state['archivo_nombre']}**  |  "
           f"Rango: {df_filtrado['FECHA'].min().strftime('%d/%m/%Y')} – {df_filtrado['FECHA'].max().strftime('%d/%m/%Y')}  |  "
           f"Fecha usada en todos los cálculos: **{mapeo_final['fecha']}**")

total_pedidos = pedidos_unicos(df_filtrado)
dias_activos = df_filtrado["FECHA_DIA"].nunique()
promedio_diario = total_pedidos / dias_activos if dias_activos else 0
ciudad_top_df = ranking(df_filtrado, "CIUDAD", top_n=1)
mes_top_df = mes_mayor_volumen_por_anio(df_filtrado)

# ---- Semana Pico GLOBAL (entre todos los meses/años filtrados) ----
semanas_kpi = (
    df_filtrado.groupby(["ANIO", "MES_NUM", "MES_NOMBRE", "SEMANA_MES"])["N_ORDEN"]
    .nunique()
    .reset_index(name="PEDIDOS")
)
semanas_kpi["MES"] = semanas_kpi["MES_NOMBRE"] + " " + semanas_kpi["ANIO"].astype(str)

# ---- Top 3 Clientes (con nombre canónico por N° de Orden, para no
# duplicar por variaciones de texto del mismo pedido) ----
top3_clientes_kpi = None
if "CLIENTE" in df_filtrado.columns:
    df_cli_kpi = df_filtrado.copy()
    df_cli_kpi["CLIENTE"] = df_cli_kpi["CLIENTE"].astype(str).str.strip()
    df_cli_kpi = df_cli_kpi[
        df_cli_kpi["CLIENTE"].notna() & (df_cli_kpi["CLIENTE"] != "") & (df_cli_kpi["CLIENTE"] != "nan")
    ]
    if not df_cli_kpi.empty:
        cliente_canonico_kpi = df_cli_kpi.groupby("N_ORDEN")["CLIENTE"].transform(
            lambda s: s.mode().iat[0] if not s.mode().empty else s.iloc[0]
        )
        df_cli_kpi["CLIENTE"] = cliente_canonico_kpi
        top3_clientes_kpi = (
            df_cli_kpi.groupby("CLIENTE")["N_ORDEN"].nunique()
            .sort_values(ascending=False).head(3)
        )

col1, col2, col3, col4, col5, col6 = st.columns(6)
col1.metric("Total de Pedidos", f"{total_pedidos:,}")
col2.metric(
    "Promedio Diario",
    f"{promedio_diario:,.1f}",
    f"sobre {dias_activos} día(s) activo(s)",
    delta_color="off",
    help="Total de Pedidos ÷ cantidad de días con al menos un pedido dentro de lo filtrado (no se cuentan los días sin actividad).",
)
col3.metric("Ciudad Líder", ciudad_top_df.iloc[0]["CIUDAD"] if len(ciudad_top_df) else "-",
            f"{int(ciudad_top_df.iloc[0]['PEDIDOS']) if len(ciudad_top_df) else 0} pedidos")
if len(mes_top_df):
    fila = mes_top_df.iloc[-1]
    col4.metric(f"Mes Pico {int(fila['ANIO'])}", fila["MES_NOMBRE"], f"{int(fila['PEDIDOS'])} pedidos")
else:
    col4.metric("Mes Pico", "-")

if not semanas_kpi.empty:
    f = semanas_kpi.loc[semanas_kpi["PEDIDOS"].idxmax()]
    col5.metric(f"Semana Pico — {f['MES']}", f"Semana {int(f['SEMANA_MES'])}", f"{int(f['PEDIDOS'])} pedidos")
else:
    col5.metric("Semana Pico", "-")

with col6:
    st.markdown("**Top 3 Clientes**")
    if top3_clientes_kpi is not None and not top3_clientes_kpi.empty:
        for i, (nombre, pedidos) in enumerate(top3_clientes_kpi.items(), start=1):
            nombre_corto = nombre if len(nombre) <= 26 else nombre[:24] + "…"
            st.caption(f"{i}. {nombre_corto} — **{int(pedidos)}**")
    else:
        st.caption("-")

st.divider()

# ============================================================
# PESTAÑAS DE ANÁLISIS
# ============================================================
tab1, tab2, tab6, tab7, tab5 = st.tabs([
    "📈 Pedidos por Tiempo",
    "🏙️ Ciudades",
    "🗓️ Comparativo Mes x Día",
    "👥 Clientes/Ciudad",
    "📥 Exportar",
])

# ---------------- TAB 1: PEDIDOS POR TIEMPO ----------------
with tab1:
    st.subheader("Pedidos por Día")
    df_dia = pedidos_por_dia(df_filtrado)
    if not df_dia.empty:
        # Completar TODOS los días del rango (incluye días sin pedidos, en 0)
        # para que el eje X muestre la secuencia completa, sin saltos.
        rango_dias = pd.date_range(df_dia["FECHA"].min(), df_dia["FECHA"].max(), freq="D")
        df_dia = (
            df_dia.set_index("FECHA")
            .reindex(rango_dias)
            .rename_axis("FECHA")
            .fillna(0)
            .reset_index()
        )
        df_dia["PEDIDOS"] = df_dia["PEDIDOS"].astype(int)

    fig_dia = px.line(
        df_dia, x="FECHA", y="PEDIDOS", markers=True, text="PEDIDOS",
        color_discrete_sequence=[COLOR_PRINCIPAL],
    )
    fig_dia.update_traces(
        textposition="top center",
        textfont=dict(color="#1a1f2b", size=9),
        line=dict(width=3),
        marker=dict(size=7, line=dict(width=1, color="#ffffff")),
        fill="tozeroy",
        fillcolor="rgba(37, 99, 235, 0.08)",
        cliponaxis=False,
    )
    fig_dia.update_xaxes(
        dtick=86400000,  # un día, en milisegundos: fuerza una marca por cada día, sin saltos
        tickformat="%d %b",
        tickangle=-90,
        title="Fecha",
    )
    fig_dia.update_layout(hovermode="x unified", yaxis_title="Pedidos", height=420)
    st.plotly_chart(fig_dia, use_container_width=True)

    # ---- 1. Pedidos por Mes ----
    st.subheader("Pedidos por Mes")
    df_mes = pedidos_por_mes(df_filtrado)
    tabla_mes = df_mes[["MES_NOMBRE", "ANIO", "PEDIDOS"]].rename(
        columns={"MES_NOMBRE": "Mes", "ANIO": "Año"}
    )
    st.dataframe(
        centrar(tabla_mes),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Pedidos": st.column_config.ProgressColumn(
                "Pedidos",
                min_value=0,
                max_value=int(tabla_mes["PEDIDOS"].max()) if len(tabla_mes) else 1,
                format="%d",
            )
        },
    )

    # ---- 2. Mes y Semana Pico ----
    st.subheader("Mes y Semana Pico")

    todas_semanas = (
        df_filtrado.groupby(["ANIO", "MES_NUM", "MES_NOMBRE", "SEMANA_MES"])["N_ORDEN"]
        .nunique()
        .reset_index()
    )
    todas_semanas.columns = ["ANIO", "MES_NUM", "MES_NOMBRE", "SEMANA_MES", "PEDIDOS"]
    todas_semanas["MES"] = todas_semanas["MES_NOMBRE"] + " " + todas_semanas["ANIO"].astype(str)
    todas_semanas = todas_semanas.sort_values(["ANIO", "MES_NUM", "SEMANA_MES"]).reset_index(drop=True)

    # Mes pico y Semana pico GLOBALES entre todos los meses seleccionados
    # (no solo el mes/año vigente).
    df_mes_general = pedidos_por_mes(df_filtrado)  # una fila por (Año, Mes) ya filtrado

    colMesPico, colSemanaPico = st.columns(2)
    with colMesPico:
        if not df_mes_general.empty:
            f = df_mes_general.loc[df_mes_general["PEDIDOS"].idxmax()]
            colMesPico.metric(
                f"Mes Pico {int(f['ANIO'])}",
                f["MES_NOMBRE"],
                f"{int(f['PEDIDOS'])} pedidos",
            )
        else:
            colMesPico.metric("Mes Pico", "-")

    with colSemanaPico:
        if not todas_semanas.empty:
            f = todas_semanas.loc[todas_semanas["PEDIDOS"].idxmax()]
            colSemanaPico.metric(
                f"Semana Pico — {f['MES']}",
                f"Semana {int(f['SEMANA_MES'])}",
                f"{int(f['PEDIDOS'])} pedidos",
            )
        else:
            colSemanaPico.metric("Semana Pico", "-")

    # ---- 3. Pedidos por Semana ----
    st.subheader("Pedidos por Semana")
    st.caption(
        "Semana del mes vs. mes, según el selector **Mes** de la barra lateral. "
        "Solo la semana más fuerte de cada mes se resalta en amarillo."
    )
    if tabla_semanal is None or tabla_semanal.empty:
        st.info("Selecciona al menos un mes en la barra lateral (sección 3) para ver esta sección.")
    else:
        def _resaltar_max_por_mes(col):
            col_numerico = tabla_semanal[col.name]
            maximo = col_numerico.max(skipna=True)
            return [
                "background-color: #FFEB3B; color: #000000; font-weight: 600;"
                if (pd.notna(v) and v == maximo) else ""
                for v in col_numerico
            ]

        st.dataframe(
            centrar(formatear_numeros(tabla_semanal).style.apply(_resaltar_max_por_mes, axis=0)),
            use_container_width=True,
        )

        tabla_semanal_largo = (
            tabla_semanal.rename_axis("Semana").reset_index()
            .melt(id_vars="Semana", var_name="Mes", value_name="Pedidos")
            .dropna(subset=["Pedidos"])
        )
        tabla_semanal_largo["Pedidos"] = tabla_semanal_largo["Pedidos"].astype(int)

        # Colores fijos para las barras que NO son el pico: azul para el
        # primer mes, celeste para el segundo (se repite el patrón si hay
        # más meses seleccionados). El pico siempre es amarillo.
        colores_base = ["#3B82F6", "#7DD3FC", "#2563EB", "#93C5FD"]
        fig_semanal = go.Figure()
        for i, mes in enumerate(tabla_semanal_largo["Mes"].unique()):
            sub = tabla_semanal_largo[tabla_semanal_largo["Mes"] == mes]
            maximo = sub["Pedidos"].max()
            color_base = colores_base[i % len(colores_base)]
            colores_barras = ["#FFEB3B" if v == maximo else color_base for v in sub["Pedidos"]]
            fig_semanal.add_trace(go.Bar(
                x=sub["Semana"], y=sub["Pedidos"], name=mes,
                marker_color=colores_barras,
                text=sub["Pedidos"],
                textposition="inside",
                textfont=dict(color="white", size=16),
            ))
        fig_semanal.update_traces(cliponaxis=False)
        fig_semanal.update_layout(
            barmode="group", height=380, xaxis_title="", yaxis_title="Pedidos",
        )
        st.plotly_chart(fig_semanal, use_container_width=True)

# ---------------- TAB 2: CIUDADES ----------------
with tab2:
    st.caption("Se considera la columna **Ciudad**. Se muestran todas las ciudades recorridas, sin selector de cantidad.")

    rank_ciudad = ranking(df_filtrado, "CIUDAD")  # sin top_n: todas las ciudades

    st.subheader("Ranking de Ciudades")
    st.markdown("**Ranking completo, de la 1ª a la última ciudad**")
    top3_ciudades = set(rank_ciudad.head(3)["CIUDAD"]) if not rank_ciudad.empty else set()
    tabla_ranking_ciudad = rank_ciudad.rename_axis("N°").reset_index().rename(columns={"CIUDAD": "Ciudad", "PEDIDOS": "Pedidos"})

    fila_resumen_ciudad = pd.DataFrame([{
        "N°": "",
        "Ciudad": "N° de Ciudades visitadas",
        "Pedidos": len(rank_ciudad),
    }])
    tabla_ranking_ciudad = pd.concat([fila_resumen_ciudad, tabla_ranking_ciudad], ignore_index=True)

    def _resaltar_top3_ciudad(row):
        if row["Ciudad"] == "N° de Ciudades visitadas":
            return ["background-color: #2b3245; font-weight: 700;"] * len(row)
        if row["Ciudad"] in top3_ciudades:
            return ["background-color: #FFEB3B; color: #000000; font-weight: 600;"] * len(row)
        return [""] * len(row)

    st.dataframe(
        centrar(tabla_ranking_ciudad.style.apply(_resaltar_top3_ciudad, axis=1)),
        use_container_width=True,
        hide_index=True,
    )

    st.divider()
    st.subheader("Ciudades Visitadas por Mes")
    st.markdown("**Ciudades × Mes**")
    st.caption(
        "Cada fila es una ciudad, cada columna un mes. Las 3 ciudades con más pedidos en "
        "total quedan resaltadas en amarillo. La primera fila muestra cuántas ciudades "
        "distintas se visitaron cada mes."
    )

    conteo_ciudad_mes = df_filtrado.groupby(["CIUDAD", "ANIO_MES"])["N_ORDEN"].nunique().reset_index(name="PEDIDOS")
    tabla_ciudad_mes = conteo_ciudad_mes.pivot(index="CIUDAD", columns="ANIO_MES", values="PEDIDOS")
    cols_orden_cm = sorted(tabla_ciudad_mes.columns)
    tabla_ciudad_mes = tabla_ciudad_mes[cols_orden_cm]
    tabla_ciudad_mes.columns = [etiquetas_mes.get(k, k) for k in cols_orden_cm]
    # OJO: la columna "Total" NO se suma de las columnas de mes, porque un
    # mismo N° de orden puede tener líneas en dos meses distintos (ej. pedido
    # despachado a caballo entre fin de agosto y principio de septiembre),
    # y sumar los conteos mensuales lo contaría dos veces. Se recalcula con
    # nunique() sobre TODO el período, igual que el ranking de arriba.
    total_real_por_ciudad = df_filtrado.groupby("CIUDAD")["N_ORDEN"].nunique()
    tabla_ciudad_mes["Total"] = total_real_por_ciudad
    tabla_ciudad_mes = tabla_ciudad_mes.sort_values("Total", ascending=False)
    tabla_ciudad_mes.index.name = "Ciudad"

    # Fila resumen al final: cuántas ciudades distintas tuvieron pedidos cada mes
    fila_resumen = (tabla_ciudad_mes.drop(columns=["Total"]) > 0).sum()
    fila_resumen["Total"] = len(tabla_ciudad_mes)
    fila_resumen.name = "N° de ciudades visitadas"
    tabla_ciudad_mes_final = pd.concat([fila_resumen.to_frame().T, tabla_ciudad_mes])

    top3_nombres = set(tabla_ciudad_mes.head(3).index)

    def _resaltar_ciudad_mes(row):
        if row.name == "N° de ciudades visitadas":
            return ["background-color: #2b3245; font-weight: 700;"] * len(row)
        if row.name in top3_nombres:
            return ["background-color: #FFEB3B; color: #000000; font-weight: 600;"] * len(row)
        return [""] * len(row)

    st.dataframe(
        centrar(formatear_numeros(tabla_ciudad_mes_final).style.apply(_resaltar_ciudad_mes, axis=1)),
        use_container_width=True,
    )

    st.divider()
    st.subheader("Gráfico — Ranking de Ciudades")
    fig_ciudad = px.bar(
        rank_ciudad.sort_values("PEDIDOS"), x="PEDIDOS", y="CIUDAD",
        orientation="h", text="PEDIDOS",
        color="PEDIDOS", color_continuous_scale=["#93C5FD", "#2563EB", "#1E3A5F"],
    )
    fig_ciudad.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0)
    fig_ciudad.update_layout(
        yaxis_title="", xaxis_title="Pedidos",
        margin=dict(r=40),
        height=max(500, 28 * len(rank_ciudad)),
        bargap=0.1,
        coloraxis_showscale=False,
    )
    fig_ciudad.update_xaxes(range=[0, rank_ciudad["PEDIDOS"].max() * 1.15])
    st.plotly_chart(fig_ciudad, use_container_width=True)

# ---------------- TAB 7: CLIENTES/CIUDAD ----------------
with tab7:
    st.caption(
        "Análisis de clientes: a cuántos se atendió cada mes, con qué frecuencia repiten "
        "pedidos (por semana y por mes), qué SKU solicitan, y en qué ciudad están."
    )

    if "CLIENTE" not in df_filtrado.columns:
        st.warning("No se encontró una columna **CLIENTE** en el archivo cargado. Esta pestaña necesita esa columna.")
    else:
        df_cli = df_filtrado.copy()
        df_cli["CLIENTE"] = df_cli["CLIENTE"].astype(str).str.strip()
        df_cli = df_cli[df_cli["CLIENTE"].notna() & (df_cli["CLIENTE"] != "") & (df_cli["CLIENTE"] != "nan")]

        # Igual que con la fecha: un mismo N° de Orden a veces trae el
        # nombre del cliente escrito de forma levemente distinta entre sus
        # líneas (ej. con o sin punto final, con la dirección pegada). Se
        # usa el nombre MÁS FRECUENTE de ese N° de Orden como su cliente
        # "oficial", para que el pedido no se cuente dos veces.
        cliente_canonico = df_cli.groupby("N_ORDEN")["CLIENTE"].transform(
            lambda s: s.mode().iat[0] if not s.mode().empty else s.iloc[0]
        )
        df_cli["CLIENTE"] = cliente_canonico

        # ---- 1. Clientes atendidos por mes ----
        st.subheader("Clientes Atendidos por Mes")
        tabla_clientes_mes = (
            df_cli.groupby(["ANIO_MES", "ANIO_MES_LABEL"])["CLIENTE"]
            .nunique()
            .reset_index(name="Clientes Distintos")
            .sort_values("ANIO_MES")
            .rename(columns={"ANIO_MES_LABEL": "Mes"})[["Mes", "Clientes Distintos"]]
        )
        st.dataframe(centrar(tabla_clientes_mes), use_container_width=True, hide_index=True)

        st.divider()

        if df_cli.empty:
            st.info("No hay datos de clientes para los filtros actuales.")
        else:
            # Ciudad "oficial" de cada cliente: la más frecuente en sus pedidos
            ciudad_por_cliente = (
                df_cli.groupby("CLIENTE")["CIUDAD"]
                .agg(lambda s: s.mode().iat[0] if not s.mode().empty else "-")
            )

            st.subheader("Recurrencia de Clientes")
            st.caption(
                "Cuántas veces solicitó pedidos cada cliente, comparando TODAS las semanas de "
                "TODOS los meses marcados en la barra lateral (sin necesidad de elegir un mes). "
                "Los 5 clientes más recurrentes quedan resaltados en amarillo."
            )

            # ---- 2a. Recurrencia por semana, con tarjetas de mes CLICKEABLES
            # para filtrar la tabla a un solo mes ----
            st.markdown("**Por semana**")
            st.caption(
                "Formato de columna: MM.AA · S# (ej. '08.26 · S1' = Agosto 2026, Semana 1). "
                "Haz clic en una tarjeta para ver solo ese mes abajo; haz clic de nuevo para "
                "volver a ver todos los meses juntos."
            )

            if "cliente_mes_seleccionado" not in st.session_state:
                st.session_state["cliente_mes_seleccionado"] = None

            total_por_mes_cli = df_cli.groupby("ANIO_MES")["N_ORDEN"].nunique().sort_index()
            cols_tarjetas = st.columns(len(total_por_mes_cli))
            for i, (col_widget, (clave, total)) in enumerate(zip(cols_tarjetas, total_por_mes_cli.items())):
                etiqueta = etiquetas_mes.get(clave, clave)
                seleccionada = st.session_state["cliente_mes_seleccionado"] == clave
                with col_widget:
                    texto_boton = f"{'✓ ' if seleccionada else ''}Total {etiqueta}\n{int(total):,}"
                    if st.button(
                        texto_boton,
                        key=f"tarjeta_mes_cli_{clave}",
                        use_container_width=True,
                        type="primary" if seleccionada else "secondary",
                    ):
                        st.session_state["cliente_mes_seleccionado"] = None if seleccionada else clave

            mes_filtro_cli = st.session_state["cliente_mes_seleccionado"]
            df_cli_semanal = df_cli if mes_filtro_cli is None else df_cli[df_cli["ANIO_MES"] == mes_filtro_cli]

            rec_semana = (
                df_cli_semanal.groupby(["CLIENTE", "ANIO_MES", "SEMANA_MES"])["N_ORDEN"]
                .nunique()
                .reset_index(name="PEDIDOS")
            )
            rec_semana["MES_COMPACTO"] = (
                rec_semana["ANIO_MES"].str.slice(5, 7) + "." + rec_semana["ANIO_MES"].str.slice(2, 4)
            )
            rec_semana["COL"] = rec_semana["MES_COMPACTO"] + " · S" + rec_semana["SEMANA_MES"].astype(str)
            orden_cols_sem = (
                rec_semana[["ANIO_MES", "SEMANA_MES", "COL"]]
                .drop_duplicates()
                .sort_values(["ANIO_MES", "SEMANA_MES"])["COL"]
                .tolist()
            )

            tabla_rec_semana = rec_semana.pivot(index="CLIENTE", columns="COL", values="PEDIDOS")
            tabla_rec_semana = tabla_rec_semana[orden_cols_sem]
            tabla_rec_semana.insert(0, "Ciudad", tabla_rec_semana.index.map(ciudad_por_cliente))
            cols_sem_num = orden_cols_sem
            total_real_cliente_sem = df_cli_semanal.groupby("CLIENTE")["N_ORDEN"].nunique()
            tabla_rec_semana["Total"] = total_real_cliente_sem
            tabla_rec_semana = tabla_rec_semana.sort_values("Total", ascending=False)
            tabla_rec_semana.index.name = "Cliente"

            top5_semana = set(tabla_rec_semana.head(5).index)

            # Fila resumen (primera fila): total de pedidos de TODOS los
            # clientes en cada semana de cada mes.
            fila_total_semana = rec_semana.groupby("COL")["PEDIDOS"].sum()
            fila_resumen_sem = pd.Series(index=tabla_rec_semana.columns, dtype=object)
            fila_resumen_sem["Ciudad"] = ""
            for col in cols_sem_num:
                fila_resumen_sem[col] = fila_total_semana.get(col, 0)
            fila_resumen_sem["Total"] = df_cli_semanal["N_ORDEN"].nunique()
            fila_resumen_sem.name = "N° de pedidos"
            tabla_rec_semana = pd.concat([fila_resumen_sem.to_frame().T, tabla_rec_semana])

            def _resaltar_top5(row, top5set):
                if row.name == "N° de pedidos":
                    return ["background-color: #2b3245; font-weight: 700;"] * len(row)
                if row.name in top5set:
                    return ["background-color: #FFEB3B; color: #000000; font-weight: 600;"] * len(row)
                return [""] * len(row)

            tabla_rec_semana_fmt = tabla_rec_semana.copy()
            tabla_rec_semana_fmt[cols_sem_num + ["Total"]] = formatear_numeros(tabla_rec_semana[cols_sem_num + ["Total"]])

            st.dataframe(
                centrar(tabla_rec_semana_fmt.style.apply(lambda r: _resaltar_top5(r, top5_semana), axis=1)),
                use_container_width=True,

            )

            # ---- 2b. Recurrencia por mes (todos los meses seleccionados) ----
            st.markdown("**Por mes (todos los meses marcados en la barra lateral)**")
            rec_mes = (
                df_cli.groupby(["CLIENTE", "ANIO_MES"])["N_ORDEN"]
                .nunique()
                .reset_index(name="PEDIDOS")
            )
            tabla_rec_mes = rec_mes.pivot(index="CLIENTE", columns="ANIO_MES", values="PEDIDOS")
            cols_orden_rm = sorted(tabla_rec_mes.columns)
            tabla_rec_mes = tabla_rec_mes[cols_orden_rm]
            tabla_rec_mes.columns = [etiquetas_mes.get(k, k) for k in cols_orden_rm]
            tabla_rec_mes.insert(0, "Ciudad", tabla_rec_mes.index.map(ciudad_por_cliente))
            cols_mes_num = [c for c in tabla_rec_mes.columns if c != "Ciudad"]
            total_real_cliente = df_cli.groupby("CLIENTE")["N_ORDEN"].nunique()
            tabla_rec_mes["Total"] = total_real_cliente
            tabla_rec_mes = tabla_rec_mes.sort_values("Total", ascending=False)
            tabla_rec_mes.index.name = "Cliente"

            top5_mes = set(tabla_rec_mes.head(5).index)

            # Fila resumen (primera fila): total de pedidos de TODOS los
            # clientes en cada mes.
            fila_total_mes = df_cli.groupby("ANIO_MES")["N_ORDEN"].nunique()
            fila_resumen_mes = pd.Series(index=tabla_rec_mes.columns, dtype=object)
            fila_resumen_mes["Ciudad"] = ""
            for col, clave in zip(cols_mes_num, cols_orden_rm):
                fila_resumen_mes[col] = fila_total_mes.get(clave, 0)
            fila_resumen_mes["Total"] = df_cli["N_ORDEN"].nunique()
            fila_resumen_mes.name = "N° de pedidos"
            tabla_rec_mes = pd.concat([fila_resumen_mes.to_frame().T, tabla_rec_mes])

            tabla_rec_mes_fmt = tabla_rec_mes.copy()
            tabla_rec_mes_fmt[cols_mes_num + ["Total"]] = formatear_numeros(tabla_rec_mes[cols_mes_num + ["Total"]])

            st.dataframe(
                centrar(tabla_rec_mes_fmt.style.apply(lambda r: _resaltar_top5(r, top5_mes), axis=1)),
                use_container_width=True,
            )

# ---------------- TAB 5: EXPORTAR ----------------
with tab5:
    st.subheader("Exportar Reporte Completo a Excel")
    st.write(
        "Genera un archivo Excel con hojas separadas para Resumen, Pedidos por Día/Semana/Mes, "
        "Rankings, Comparativos y el Detalle de los datos filtrados actualmente."
    )
    reporte_bytes = generar_reporte_excel(df_filtrado)
    st.download_button(
        label="⬇️ Descargar Reporte Excel",
        data=reporte_bytes,
        file_name=f"reporte_logistico_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.divider()
    st.subheader("Vista previa de datos filtrados")
    st.dataframe(centrar(df_filtrado.drop(columns=["FECHA_DIA"], errors="ignore")), use_container_width=True)

# ---------------- TAB 6: COMPARATIVO MES X DÍA ----------------
with tab6:
    st.subheader("Pedidos por Día — Comparativo entre Meses")
    st.caption(
        "Agrupado por **Fecha de Solicitud**. Marca o desmarca los meses en el selector "
        "**Mes** de la barra lateral."
    )

    if col_fecha_solicitud is None:
        st.warning(
            "No se encontró una columna **FECHA DE SOLICITUD** en el archivo cargado. "
            "Esta vista necesita esa columna para poder generarse."
        )
    elif tabla_comparativo is None or tabla_comparativo.empty:
        st.info("Selecciona al menos un mes en la barra lateral (sección 3) para ver la comparación.")
    else:
        # La tabla se separa en dos bloques de días (1–15 y 16–fin de mes),
        # cada uno como una tabla delgada y compacta aparte. El gráfico de
        # abajo sigue mostrando el mes completo, sin cortar.
        ultimo_dia = int(tabla_comparativo.index.max())
        tabla_bloque1 = tabla_comparativo.loc[tabla_comparativo.index <= 15]
        tabla_bloque2 = tabla_comparativo.loc[tabla_comparativo.index > 15]

        colBloque1, colBloque2 = st.columns(2)
        with colBloque1:
            st.caption("**Días 1–15**")
            st.dataframe(
                centrar(formatear_numeros(tabla_bloque1)),
                use_container_width=True,
            )
        with colBloque2:
            st.caption(f"**Días 16–{ultimo_dia}**")
            st.dataframe(
                centrar(formatear_numeros(tabla_bloque2)),
                use_container_width=True,
            )

        fig_dia_comp = px.bar(
            conteo_comp_largo, x="Día", y="PEDIDOS", color="MES_LABEL",
            barmode="group", text="PEDIDOS", labels={"MES_LABEL": "Mes"},
            color_discrete_sequence=PALETA,
        )
        fig_dia_comp.update_traces(textposition="outside", cliponaxis=False, textfont_size=10, marker_line_width=0)
        fig_dia_comp.update_xaxes(tickmode="linear", dtick=1, title="Día")
        fig_dia_comp.update_layout(height=420, yaxis_title="Pedidos")
        st.plotly_chart(fig_dia_comp, use_container_width=True)

        cols_totales = st.columns(len(totales_comparativo))
        for c, (mes_label, total) in zip(cols_totales, totales_comparativo.items()):
            c.metric(mes_label, f"{int(total):,}")

st.sidebar.divider()
st.sidebar.caption(
    "Para actualizar el dashboard, simplemente sube un nuevo archivo Excel arriba. "
    "Todos los indicadores se recalculan automáticamente."
)
