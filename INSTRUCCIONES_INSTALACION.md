# Instrucciones de Instalación y Uso
## Dashboard de Gestión Logística y Análisis de Pedidos

---

## 1. Requisitos previos

- Python 3.9 o superior instalado ([python.org/downloads](https://www.python.org/downloads/))
- Los archivos: `app.py` y `requirements.txt` en la misma carpeta.

Verifica tu versión de Python:

```bash
python --version
```

---

## 2. Instalación (solo la primera vez)

### Paso 1 — Crear un entorno virtual (recomendado)

```bash
python -m venv venv
```

Actívalo:

- **Windows:**
  ```bash
  venv\Scripts\activate
  ```
- **macOS / Linux:**
  ```bash
  source venv/bin/activate
  ```

### Paso 2 — Instalar las dependencias

```bash
pip install -r requirements.txt
```

Esto instalará: Streamlit, Pandas, NumPy, Plotly, Openpyxl y XlsxWriter.

---

## 3. Ejecutar la aplicación

Desde la carpeta donde están `app.py` y `requirements.txt`, ejecuta:

```bash
streamlit run app.py
```

Se abrirá automáticamente en tu navegador en una dirección como:
`http://localhost:8501`

Si no se abre solo, copia esa URL en tu navegador.

---

## 4. Uso de la herramienta

### 4.1 Cargar el archivo Excel

1. En la barra lateral izquierda, haz clic en **"Sube tu archivo Excel actualizado (.xlsx)"**.
2. Selecciona tu archivo. Si tiene varias hojas, la app te dejará elegir cuál usar.
3. **Todo el dashboard se calcula automáticamente** en cuanto el archivo se carga.

### 4.2 Columnas necesarias en tu Excel

La herramienta fue ajustada para el formato real de tu archivo (`PRUEBA_RESUMEN_RUTAS`, hoja `DATA`), que tiene esta particularidad: **no existe una columna "RUTA" explícita**. En su lugar, la ruta/zona de reparto va incluida como el último segmento de texto dentro de la columna **"PTO ENTREGA"** (ejemplo: `MEGAMETALES S.A. / AV. LEON FEBRES CORDERO / GYE NORTE 3` → la ruta es `GYE NORTE 3`).

| Campo lógico | Columna detectada en tu archivo | Cómo se obtiene |
|---|---|---|
| N° de orden (conteo de pedidos) | `N° ORDEN` | Directo |
| Fecha | `FECHA` (también detecta `FECHA DE ENTREGA`, `FECHA DE SOLICITUD`, `FECHA FACTURA` como alternativas seleccionables) | Directo, seleccionable |
| Ciudad | `CIUDAD` | Directo |
| Ruta | *(no existe columna RUTA)* | **Usa la misma columna que Ciudad** por defecto (configurable) |
| Tipo de movimiento (filtro opcional) | `TIPO` (`DESPACHO DE PEDIDO`, `CROSSDOCKING`, `TRANSFERENCIA`, `RETIRO Y ENTREGA`) | Directo, opcional |

Como tu archivo no trae una columna de ruta/zona propia, por defecto la app usa **CIUDAD** también como valor de **Ruta** (así que "Ranking de Rutas" mostrará lo mismo que "Ranking de Ciudades" hasta que exista una columna de ruta más granular). Si más adelante agregas una columna real de ruta/zona de reparto, la app la detectará y la usará automáticamente en su lugar.

Si en el futuro subes un archivo con estructura distinta (por ejemplo, con una columna `RUTA` explícita, o con otros nombres de columna), la app:

1. Intenta detectar automáticamente cada columna.
2. Si encuentra una columna `RUTA`/`ZONA` explícita, la usa directamente en lugar de Ciudad.
3. Si no detecta ninguna columna esperada, abre el panel **"⚙️ Mapeo de columnas"** en la barra lateral para que la selecciones manualmente con un par de clics (incluye un selector para elegir qué columna de fecha usar, y un interruptor para elegir entre "usar columna de ruta existente", "usar la misma columna que Ciudad", o "derivar la ruta desde un texto" como un campo de dirección/punto de entrega).

### 4.3 Filas que se excluyen automáticamente

- Filas sin fecha válida.
- Filas sin `N° ORDEN` (en tu archivo esto ocurre en algunas filas de tipo `TRANSFERENCIA`, que son movimientos internos de bodega, no pedidos de cliente). Estas filas no se cuentan como "pedido", pero puedes verlas si desactivas el filtro de Tipo de movimiento en el detalle exportado.

La barra lateral muestra un aviso con la cantidad de filas excluidas por cada motivo, para que siempre sepas qué se está contando.

### 4.4 Actualizar los datos periódicamente

Cada vez que tengas un archivo Excel nuevo:

1. Vuelve a hacer clic en el cargador de archivos.
2. Sube el nuevo archivo.
3. La aplicación descarta automáticamente los datos anteriores y **recalcula
   todos los KPIs, gráficos, rankings y comparativos** con la nueva información.

**No necesitas modificar ni una línea de código.**

### 4.5 Filtros

En la barra lateral puedes filtrar por:
- Año
- Mes
- Semana (formato ISO, ej: `2026-W10`)
- Ciudad
- Ruta
- Tipo de movimiento (solo aparece si la columna `TIPO` está presente, ej. para excluir `TRANSFERENCIA` o `CROSSDOCKING` y quedarte solo con `DESPACHO DE PEDIDO`)

Todos los filtros se combinan entre sí y afectan a **todas** las pestañas del dashboard.

### 4.6 Pestañas disponibles

- **📈 Pedidos por Tiempo** — pedidos por día, semana y mes.
- **🏙️ Ciudades y Rutas** — rankings configurables (Top N) y mapa de calor cruzado.
- **📊 Comparativos** — tabla y gráfico comparativo mes a mes entre distintos años, y comparativo anual total.
- **🏆 Mes y Semana Pico** — mes con mayor volumen por año y semana más fuerte de cada mes.
- **📥 Exportar** — descarga un reporte Excel completo (multi-hoja) con los datos ya filtrados.

---

## 5. Detener la aplicación

En la terminal donde se está ejecutando, presiona `Ctrl + C`.

---

## 6. Notas técnicas importantes

- **Conteo de pedidos:** se calcula como el número de valores **únicos** de `N° ORDEN`,
  no el número de filas. Esto evita contar dos veces un mismo pedido si tu Excel tiene
  varias filas por pedido (por ejemplo, una fila por producto).
- **Fechas inválidas:** las filas sin una fecha reconocible se descartan automáticamente,
  y verás un aviso en la barra lateral indicando cuántas filas fueron excluidas.
- **Rendimiento:** la app está optimizada con caché de Streamlit para archivos grandes
  (varias decenas de miles de filas funcionan sin problema).
- **Despliegue en la nube (opcional):** si más adelante quieres que otras personas del
  equipo accedan a este dashboard desde un enlace web (sin instalar nada), puedes
  desplegarlo gratis en [Streamlit Community Cloud](https://streamlit.io/cloud) subiendo
  estos mismos archivos a un repositorio de GitHub.

---

## 7. Solución de problemas comunes

| Problema | Solución |
|---|---|
| `streamlit: command not found` | Verifica que el entorno virtual esté activado y que `pip install -r requirements.txt` se ejecutó sin errores. |
| La app dice que faltan columnas | Abre "⚙️ Mapeo de columnas" en la barra lateral y selecciona manualmente las columnas correctas. |
| Las fechas no se reconocen | Asegúrate de que la columna de fecha en el Excel tenga formato de fecha real (no texto libre). |
| El navegador no abre solo | Copia la URL que aparece en la terminal (usualmente `http://localhost:8501`) y pégala en tu navegador. |
