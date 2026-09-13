# Dashboard Comparativo de Pedidos por Mes

## Archivos

- `template.html` — plantilla del dashboard (HTML + CSS + JS). No editar los datos aquí directamente; contiene el marcador `__DATA_JSON__` que el script reemplaza.
- `generar_dashboard.py` — lee el Excel de pedidos y genera el HTML final con los datos embebidos.
- `dashboard.html` — archivo generado (es el que se sube/sirve). Se regenera cada vez que corres el script.

## Requisitos

```bash
pip install pandas openpyxl
```

## Uso

```bash
python generar_dashboard.py RUTA_AL_EXCEL.xlsx --out dashboard.html
```

Ejemplo:

```bash
python generar_dashboard.py PRUEBA_RESUMEN_RUTAS.xlsx
```

Esto imprime un resumen en consola y escribe `dashboard.html`, listo para abrir en el navegador o publicar (GitHub Pages, etc.).

## Flujo recomendado en GitHub

1. Reemplaza el Excel de pedidos con la versión más reciente (mismo nombre de hoja: `DATA`).
2. Corre `python generar_dashboard.py tu_excel.xlsx`.
3. Haz commit de `dashboard.html` (y del Excel, si también lo versionas).
4. Si usas GitHub Pages, apunta a `dashboard.html` como página de inicio.

## Qué espera el Excel

Hoja llamada `DATA`, con al menos estas columnas:

- `FECHA DE SOLICITUD` — fecha usada para agrupar los pedidos por día.
- `N° ORDEN` — número de pedido. Solo se cuentan filas con este campo lleno; cada N° ORDEN se cuenta una sola vez por día (aunque tenga varias líneas de producto).

## Modificar el diseño

Cualquier cambio visual (colores, columnas, textos) se hace en `template.html`. El script solo inyecta datos — no necesitas tocar `generar_dashboard.py` para cambios de estilo.
