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
CANDIDATOS_FECHA_PRIORIDAD = [
    "FECHA", "FECHADEENTREGA", "FECHAENTREGA", "FECHADEPEDIDO", "FECHAPEDIDO",
    "FECHADESOLICITUD", "FECHAFACTURA", "FECHADESPACHO",
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
    """Detecta automáticamente las columnas simples (orden, ciudad)."""
    columnas_normalizadas = {col: normalizar_texto(col) for col in df.columns}
    mapeo = {}
    for campo, candidatos in CANDIDATOS_COLUMNAS.items():
        mapeo[campo] = _buscar_columna(columnas_normalizadas, candidatos)
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


def comparativo_pivot(df: pd.DataFrame) -> pd.DataFrame:
    """Tabla Año x Mes con total de pedidos, para comparativos mensuales/anuales."""
    mensual = pedidos_por_mes(df)
    pivot = mensual.pivot_table(index="MES_NOMBRE", columns="ANIO", values="PEDIDOS", aggfunc="sum")
    orden_meses = [MESES_ES[i] for i in range(1, 13)]
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
                "Total de pedidos", "Ciudades distintas", "Rutas distintas",
                "Fecha mínima", "Fecha máxima",
            ],
            "Valor": [
                pedidos_unicos(df_filtrado),
                df_filtrado["CIUDAD"].nunique(),
                df_filtrado["RUTA"].nunique(),
                df_filtrado["FECHA"].min().strftime("%Y-%m-%d") if len(df_filtrado) else "-",
                df_filtrado["FECHA"].max().strftime("%Y-%m-%d") if len(df_filtrado) else "-",
            ],
        })
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        pedidos_por_dia(df_filtrado).to_excel(writer, sheet_name="Pedidos por Dia", index=False)
        pedidos_por_semana(df_filtrado).to_excel(writer, sheet_name="Pedidos por Semana", index=False)
        pedidos_por_mes(df_filtrado).to_excel(writer, sheet_name="Pedidos por Mes", index=False)
        mes_mayor_volumen_por_anio(df_filtrado).to_excel(writer, sheet_name="Mes Top por Anio", index=False)
        semana_mas_fuerte_por_mes(df_filtrado).to_excel(writer, sheet_name="Semana Top por Mes", index=False)
        ranking(df_filtrado, "CIUDAD").to_excel(writer, sheet_name="Ranking Ciudades")
        ranking(df_filtrado, "RUTA").to_excel(writer, sheet_name="Ranking Rutas")
        comparativo_pivot(df_filtrado).to_excel(writer, sheet_name="Comparativo Anio-Mes")
        df_filtrado.drop(columns=["FECHA_DIA"], errors="ignore").to_excel(
            writer, sheet_name="Detalle Filtrado", index=False
        )
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

columnas_faltantes = [campo for campo, col in mapeo_auto.items() if col is None]
faltan_fecha = len(columnas_fecha_detectadas) == 0
faltan_ruta = col_ruta_explicita is None and col_fuente_ruta is None

with st.sidebar.expander(
    "⚙️ Mapeo de columnas",
    expanded=bool(columnas_faltantes or faltan_fecha or faltan_ruta),
):
    st.caption("La app detecta las columnas automáticamente. Ajusta aquí si algo no coincide.")

    # -- Orden y Ciudad --
    mapeo_final = {}
    opciones = ["-- Selecciona --"] + list(df_crudo.columns)
    for campo, descripcion in CAMPOS_REQUERIDOS.items():
        valor_default = mapeo_auto.get(campo)
        idx_default = opciones.index(valor_default) if valor_default in opciones else 0
        seleccion = st.selectbox(descripcion, opciones, index=idx_default, key=f"map_{campo}")
        mapeo_final[campo] = None if seleccion == "-- Selecciona --" else seleccion

    # -- Fecha: selector dedicado entre las columnas tipo fecha detectadas --
    st.markdown("**Fecha a usar para el análisis**")
    opciones_fecha = columnas_fecha_detectadas if columnas_fecha_detectadas else list(df_crudo.columns)
    if not columnas_fecha_detectadas:
        st.caption("No se detectaron columnas de fecha automáticamente; selecciona una manualmente.")
    col_fecha_sel = st.selectbox(
        "Columna de fecha (ej. FECHA = fecha del pedido, FECHA DE ENTREGA = fecha de entrega)",
        opciones_fecha, index=0, key="map_fecha",
    )
    mapeo_final["fecha"] = col_fecha_sel

    # -- Ruta: columna explícita, igual a Ciudad, derivada, o manual --
    st.markdown("**Ruta / Zona de entrega**")
    opciones_modo_ruta = [
        "Usar columna existente",
        "Usar la misma columna que Ciudad",
        "Derivar desde un campo de texto",
    ]
    if col_ruta_explicita:
        modo_ruta_default = "Usar columna existente"
        st.caption(f"Se detectó una columna de ruta explícita: **{col_ruta_explicita}**.")
    else:
        modo_ruta_default = "Usar la misma columna que Ciudad"
        st.caption(
            "No hay columna de RUTA explícita en este archivo, así que por defecto se usa "
            "**CIUDAD** como Ruta. Si prefieres derivarla desde otro campo (ej. punto de entrega), "
            "cambia el método abajo."
        )

    modo_ruta = st.radio(
        "Método para obtener la Ruta",
        opciones_modo_ruta,
        index=opciones_modo_ruta.index(modo_ruta_default),
        key="modo_ruta",
    )

    ruta_config = {}
    if modo_ruta == "Usar columna existente":
        idx_default = opciones.index(col_ruta_explicita) if col_ruta_explicita in opciones else 0
        col_ruta_manual = st.selectbox("Columna de Ruta", opciones, index=idx_default, key="map_ruta_columna")
        ruta_config["ruta_modo"] = "columna"
        ruta_config["ruta_columna"] = None if col_ruta_manual == "-- Selecciona --" else col_ruta_manual
    elif modo_ruta == "Usar la misma columna que Ciudad":
        ruta_config["ruta_modo"] = "ciudad"
    else:
        idx_default = opciones.index(col_fuente_ruta) if col_fuente_ruta in opciones else 0
        col_fuente_manual = st.selectbox(
            "Columna de texto desde la que derivar la ruta", opciones, index=idx_default, key="map_ruta_fuente"
        )
        modo_derivacion = st.selectbox(
            "Cómo extraer la ruta de ese texto",
            ["Último segmento tras '/' (recomendado)", "Usar el valor completo"],
            key="map_ruta_derivacion_modo",
        )
        ruta_config["ruta_modo"] = "derivada"
        ruta_config["ruta_fuente"] = None if col_fuente_manual == "-- Selecciona --" else col_fuente_manual
        ruta_config["ruta_derivacion_modo"] = (
            "segmento_final" if "Último segmento" in modo_derivacion else "valor_completo"
        )
    mapeo_final.update(ruta_config)

    # -- Tipo de movimiento (opcional) --
    st.markdown("**Tipo de movimiento (opcional)**")
    opciones_tipo = ["(No usar)"] + list(df_crudo.columns)
    idx_tipo_default = opciones_tipo.index(col_tipo_movimiento_auto) if col_tipo_movimiento_auto in opciones_tipo else 0
    col_tipo_sel = st.selectbox(
        "Columna que distingue tipos de pedido (ej. DESPACHO DE PEDIDO, TRANSFERENCIA, CROSSDOCKING)",
        opciones_tipo, index=idx_tipo_default, key="map_tipo",
    )
    mapeo_final["tipo_movimiento"] = None if col_tipo_sel == "(No usar)" else col_tipo_sel

# Validación de columnas obligatorias
faltan = [mapeo_final.get("orden"), mapeo_final.get("ciudad"), mapeo_final.get("fecha")]
ruta_incompleta = (
    (mapeo_final.get("ruta_modo") == "columna" and not mapeo_final.get("ruta_columna"))
    or (mapeo_final.get("ruta_modo") == "derivada" and not mapeo_final.get("ruta_fuente"))
)
if any(f is None for f in faltan) or ruta_incompleta:
    st.error(
        "⚠️ No se pudieron identificar todas las columnas necesarias "
        "(N° ORDEN, Fecha, Ciudad, Ruta). Ajusta el mapeo de columnas en la barra lateral."
    )
    st.stop()

# ---- Preprocesamiento ----
df = preparar_datos(df_crudo, mapeo_final)

avisos = []
if df.attrs.get("filas_sin_fecha", 0) > 0:
    avisos.append(f"{df.attrs['filas_sin_fecha']} filas descartadas por fecha inválida o vacía.")
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

anios_disponibles = sorted(df["ANIO"].unique())
anios_sel = st.sidebar.multiselect("Año", anios_disponibles, default=anios_disponibles)

df_tmp = df[df["ANIO"].isin(anios_sel)] if anios_sel else df

meses_disponibles = sorted(df_tmp["MES_NUM"].unique())
meses_labels = [MESES_ES[m] for m in meses_disponibles]
meses_sel_labels = st.sidebar.multiselect("Mes", meses_labels, default=meses_labels)
meses_sel = [k for k, v in MESES_ES.items() if v in meses_sel_labels]

df_tmp = df_tmp[df_tmp["MES_NUM"].isin(meses_sel)] if meses_sel else df_tmp

semanas_disponibles = sorted(df_tmp["SEMANA_ANIO_ID"].unique())
semanas_sel = st.sidebar.multiselect("Semana (ISO)", semanas_disponibles, default=[])

ciudades_disponibles = sorted(df_tmp["CIUDAD"].unique())
ciudades_sel = st.sidebar.multiselect("Ciudad", ciudades_disponibles, default=[])

rutas_disponibles = sorted([r for r in df_tmp["RUTA"].unique() if r is not None])
rutas_sel = st.sidebar.multiselect("Ruta", rutas_disponibles, default=[])

tipo_sel = []
if "TIPO_MOVIMIENTO" in df.columns:
    tipos_disponibles = sorted(df_tmp["TIPO_MOVIMIENTO"].dropna().unique())
    tipo_sel = st.sidebar.multiselect("Tipo de movimiento", tipos_disponibles, default=[])

# Aplicar filtros finales
df_filtrado = df.copy()
if anios_sel:
    df_filtrado = df_filtrado[df_filtrado["ANIO"].isin(anios_sel)]
if meses_sel:
    df_filtrado = df_filtrado[df_filtrado["MES_NUM"].isin(meses_sel)]
if semanas_sel:
    df_filtrado = df_filtrado[df_filtrado["SEMANA_ANIO_ID"].isin(semanas_sel)]
if ciudades_sel:
    df_filtrado = df_filtrado[df_filtrado["CIUDAD"].isin(ciudades_sel)]
if rutas_sel:
    df_filtrado = df_filtrado[df_filtrado["RUTA"].isin(rutas_sel)]
if tipo_sel:
    df_filtrado = df_filtrado[df_filtrado["TIPO_MOVIMIENTO"].isin(tipo_sel)]

if df_filtrado.empty:
    st.warning("No hay datos para la combinación de filtros seleccionada.")
    st.stop()

# ============================================================
# ENCABEZADO Y KPIs EJECUTIVOS
# ============================================================
st.title("🚚 Dashboard de Gestión Logística y Análisis de Pedidos")
st.caption(f"Última actualización de datos: **{st.session_state['archivo_nombre']}**  |  "
           f"Rango: {df_filtrado['FECHA'].min().strftime('%d/%m/%Y')} – {df_filtrado['FECHA'].max().strftime('%d/%m/%Y')}")

total_pedidos = pedidos_unicos(df_filtrado)
dias_activos = df_filtrado["FECHA_DIA"].nunique()
promedio_diario = total_pedidos / dias_activos if dias_activos else 0
ciudad_top_df = ranking(df_filtrado, "CIUDAD", top_n=1)
ruta_top_df = ranking(df_filtrado, "RUTA", top_n=1)
mes_top_df = mes_mayor_volumen_por_anio(df_filtrado)

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Total de Pedidos", f"{total_pedidos:,}")
col2.metric("Promedio Diario", f"{promedio_diario:,.1f}")
col3.metric("Ciudad Líder", ciudad_top_df.iloc[0]["CIUDAD"] if len(ciudad_top_df) else "-",
            f"{int(ciudad_top_df.iloc[0]['PEDIDOS']) if len(ciudad_top_df) else 0} pedidos")
col4.metric("Ruta Líder", ruta_top_df.iloc[0]["RUTA"] if len(ruta_top_df) else "-",
            f"{int(ruta_top_df.iloc[0]['PEDIDOS']) if len(ruta_top_df) else 0} pedidos")
if len(mes_top_df):
    fila = mes_top_df.iloc[-1]
    col5.metric(f"Mes Pico {int(fila['ANIO'])}", fila["MES_NOMBRE"], f"{int(fila['PEDIDOS'])} pedidos")
else:
    col5.metric("Mes Pico", "-")

st.divider()

# ============================================================
# PESTAÑAS DE ANÁLISIS
# ============================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📈 Pedidos por Tiempo",
    "🏙️ Ciudades y Rutas",
    "📊 Comparativos",
    "🏆 Mes y Semana Pico",
    "📥 Exportar",
])

# ---------------- TAB 1: PEDIDOS POR TIEMPO ----------------
with tab1:
    st.subheader("Pedidos por Día")
    df_dia = pedidos_por_dia(df_filtrado)
    fig_dia = px.line(df_dia, x="FECHA", y="PEDIDOS", markers=True)
    fig_dia.update_layout(hovermode="x unified", xaxis_title="Fecha", yaxis_title="Pedidos")
    st.plotly_chart(fig_dia, use_container_width=True)

    colA, colB = st.columns(2)
    with colA:
        st.subheader("Pedidos por Semana")
        df_sem = pedidos_por_semana(df_filtrado)
        fig_sem = px.bar(df_sem, x="SEMANA", y="PEDIDOS")
        fig_sem.update_layout(xaxis_title="Semana (Año-Semana ISO)", yaxis_title="Pedidos")
        st.plotly_chart(fig_sem, use_container_width=True)
    with colB:
        st.subheader("Pedidos por Mes")
        df_mes = pedidos_por_mes(df_filtrado)
        fig_mes = px.bar(df_mes, x="ANIO_MES", y="PEDIDOS", color="ANIO",
                          labels={"ANIO_MES": "Mes"})
        st.plotly_chart(fig_mes, use_container_width=True)

# ---------------- TAB 2: CIUDADES Y RUTAS ----------------
with tab2:
    top_n = st.slider("Cantidad de elementos en el ranking (Top N)", 3, 30, 10)
    ruta_igual_a_ciudad = df_filtrado["RUTA"].equals(df_filtrado["CIUDAD"])

    colA, colB = st.columns(2)
    with colA:
        st.subheader("Ranking de Ciudades")
        rank_ciudad = ranking(df_filtrado, "CIUDAD", top_n=top_n)
        fig_ciudad = px.bar(
            rank_ciudad.sort_values("PEDIDOS"), x="PEDIDOS", y="CIUDAD",
            orientation="h", text="PEDIDOS",
        )
        fig_ciudad.update_layout(yaxis_title="", xaxis_title="Pedidos")
        st.plotly_chart(fig_ciudad, use_container_width=True)
        st.dataframe(rank_ciudad, use_container_width=True)

    with colB:
        st.subheader("Ranking de Rutas")
        if ruta_igual_a_ciudad:
            st.caption("ℹ️ La Ruta está configurada igual a la Ciudad, por lo que este ranking coincide con el de Ciudades.")
        rank_ruta = ranking(df_filtrado, "RUTA", top_n=top_n)
        fig_ruta = px.bar(
            rank_ruta.sort_values("PEDIDOS"), x="PEDIDOS", y="RUTA",
            orientation="h", text="PEDIDOS", color_discrete_sequence=["#EF553B"],
        )
        fig_ruta.update_layout(yaxis_title="", xaxis_title="Pedidos")
        st.plotly_chart(fig_ruta, use_container_width=True)
        st.dataframe(rank_ruta, use_container_width=True)

    if not ruta_igual_a_ciudad:
        st.subheader("Distribución Ciudad x Ruta")
        tabla_cruzada = df_filtrado.groupby(["CIUDAD", "RUTA"])["N_ORDEN"].nunique().reset_index()
        tabla_cruzada.columns = ["CIUDAD", "RUTA", "PEDIDOS"]
        fig_heat = px.density_heatmap(
            tabla_cruzada, x="RUTA", y="CIUDAD", z="PEDIDOS",
            color_continuous_scale="Blues", text_auto=True,
        )
        st.plotly_chart(fig_heat, use_container_width=True)

# ---------------- TAB 3: COMPARATIVOS ----------------
with tab3:
    st.subheader("Comparativo Mensual entre Años")
    pivot = comparativo_pivot(df_filtrado)
    st.dataframe(pivot.style.format("{:.0f}", na_rep="-"), use_container_width=True)

    pivot_reset = pivot.reset_index().melt(id_vars="MES_NOMBRE", var_name="ANIO", value_name="PEDIDOS")
    pivot_reset = pivot_reset.dropna(subset=["PEDIDOS"])
    fig_comp = px.bar(
        pivot_reset, x="MES_NOMBRE", y="PEDIDOS", color=pivot_reset["ANIO"].astype(str),
        barmode="group", labels={"color": "Año", "MES_NOMBRE": "Mes"},
    )
    st.plotly_chart(fig_comp, use_container_width=True)

    st.subheader("Comparativo Anual (Total de Pedidos por Año)")
    anual = df_filtrado.groupby("ANIO")["N_ORDEN"].nunique().reset_index()
    anual.columns = ["ANIO", "PEDIDOS"]
    anual["ANIO"] = anual["ANIO"].astype(str)
    fig_anual = px.bar(anual, x="ANIO", y="PEDIDOS", text="PEDIDOS")
    st.plotly_chart(fig_anual, use_container_width=True)

    if len(anual) > 1:
        anual_sorted = anual.sort_values("ANIO")
        variacion = anual_sorted["PEDIDOS"].pct_change().iloc[-1] * 100
        st.metric(
            f"Variación {anual_sorted.iloc[-1]['ANIO']} vs {anual_sorted.iloc[-2]['ANIO']}",
            f"{variacion:+.1f}%",
        )

# ---------------- TAB 4: MES Y SEMANA PICO ----------------
with tab4:
    st.subheader("Mes con Mayor Volumen por Año")
    df_mes_top = mes_mayor_volumen_por_anio(df_filtrado)
    st.dataframe(df_mes_top[["ANIO", "MES_NOMBRE", "PEDIDOS"]], use_container_width=True)
    fig_top_mes = px.bar(df_mes_top, x="ANIO", y="PEDIDOS", color="MES_NOMBRE", text="MES_NOMBRE")
    st.plotly_chart(fig_top_mes, use_container_width=True)

    st.subheader("Semana Más Fuerte de Cada Mes")
    df_sem_top = semana_mas_fuerte_por_mes(df_filtrado)
    df_sem_top_view = df_sem_top.copy()
    df_sem_top_view["MES"] = df_sem_top_view["MES_NOMBRE"] + " " + df_sem_top_view["ANIO"].astype(str)
    st.dataframe(
        df_sem_top_view[["MES", "SEMANA_MES", "PEDIDOS"]].rename(
            columns={"SEMANA_MES": "N° Semana del Mes"}
        ),
        use_container_width=True,
    )
    fig_sem_top = px.bar(df_sem_top_view, x="MES", y="PEDIDOS", color="SEMANA_MES",
                          labels={"SEMANA_MES": "Semana del mes"})
    st.plotly_chart(fig_sem_top, use_container_width=True)

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
    st.dataframe(df_filtrado.drop(columns=["FECHA_DIA"], errors="ignore"), use_container_width=True)

st.sidebar.divider()
st.sidebar.caption(
    "Para actualizar el dashboard, simplemente sube un nuevo archivo Excel arriba. "
    "Todos los indicadores se recalculan automáticamente."
)
