import os
import json
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from copy import copy
import urllib.request
import urllib.error
import time
import streamlit as st
import io
import zipfile

# ─────────────────────────────────────────────────────────────────
#   CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────────

CAMPOS_PLANTILLA = ["fecha", "vendedor", "cc", "referencia", "Valor Total", "cantidad", "cliente", "nit"]
RUTA_JSON = "config_distribuidores.json"

st.set_page_config(page_title="Liquidador de Distribuidores", page_icon="🐐", layout="centered")

st.title("Sistema de Liquidación de Distribuidores")
st.caption("Creado por 🐐 Nicolas Tovar y 🐐 Cristian Morales")
st.markdown("---")

# ─────────────────────────────────────────────────────────────────
#   SIDEBAR: API KEY
# ─────────────────────────────────────────────────────────────────

st.sidebar.header("🔑 Configuración de Acceso")
api_key_input = st.sidebar.text_input(
    "Anthropic API Key", type="password",
    help="Ingresa tu API Key para habilitar el mapeo automático con IA."
)
api_key = api_key_input.strip() if api_key_input else os.environ.get("ANTHROPIC_API_KEY", "")

if not api_key:
    st.sidebar.warning("⚠️ Sin API key. Solo se usarán mapeos del JSON existente.")
else:
    st.sidebar.success("✅ API Key cargada.")

# ─────────────────────────────────────────────────────────────────
#   DETECCIÓN DE FILA DE ENCABEZADOS (FILAS 1, 2 O 3)
# ─────────────────────────────────────────────────────────────────

def detectar_fila_encabezados(file_bytes: bytes, max_filas: int = 3) -> tuple[pd.DataFrame, int]:
    """
    Intenta leer el Excel probando la fila de encabezados en las posiciones
    0, 1 y 2 (equivalentes a filas 1, 2 y 3 en Excel).

    Una fila es válida como encabezado si:
      - Tiene al menos 3 columnas con texto (no nulas, no numéricas puras)
      - Al menos el 50% de sus valores son strings no vacíos

    Retorna (DataFrame con encabezados correctos, número de fila 1-based encontrada).
    Si no encuentra nada válido, retorna la lectura con header=0 como fallback.
    """
    for fila_idx in range(max_filas):
        try:
            df = pd.read_excel(
                io.BytesIO(file_bytes),
                header=fila_idx,
                dtype=str,
                nrows=5  # solo leemos unas pocas filas para validar rápido
            )
            df.columns = df.columns.astype(str).str.strip()

            # Filtrar columnas "Unnamed: X" que genera pandas cuando la celda está vacía
            cols_validas = [
                c for c in df.columns
                if c and not c.startswith("Unnamed:") and c.lower() != "nan"
            ]

            # Criterio: al menos 3 columnas con nombre real
            if len(cols_validas) >= 3:
                # Releer completo con la fila correcta
                df_completo = pd.read_excel(
                    io.BytesIO(file_bytes),
                    header=fila_idx,
                    dtype=str
                )
                df_completo.columns = df_completo.columns.astype(str).str.strip()
                # Eliminar columnas sin nombre
                df_completo = df_completo.loc[
                    :, ~df_completo.columns.str.startswith("Unnamed:")
                ]
                df_completo = df_completo.loc[
                    :, df_completo.columns.str.lower() != "nan"
                ]
                return df_completo, fila_idx + 1  # fila 1-based

        except Exception:
            continue

    # Fallback: leer con header=0
    df_fallback = pd.read_excel(io.BytesIO(file_bytes), header=0, dtype=str)
    df_fallback.columns = df_fallback.columns.astype(str).str.strip()
    return df_fallback, 1


# ─────────────────────────────────────────────────────────────────
#   MAPEO CON IA
# ─────────────────────────────────────────────────────────────────

def mapear_con_ia(encabezados: list, nombre_archivo: str, api_key: str) -> tuple[dict, str] | tuple[None, None]:
    prompt = f"""Eres un asistente experto en normalización de datos de ventas.

Tengo un archivo Excel de un distribuidor llamado "{nombre_archivo}" con estos encabezados exactos:
{json.dumps(encabezados, ensure_ascii=False)}

Necesito dos cosas:

1. Mapear los encabezados a estos campos estándar:
- fecha        → columna con fechas de venta/documento
- vendedor     → columna con nombre o código del vendedor/asesor
- cc           → columna con cédula o identificación del vendedor (puede no existir)
- referencia   → columna con el nombre DESCRIPTIVO del producto (texto, no código numérico). Si hay nombre y código de producto, elige el nombre.
- Valor Total  → columna con el valor/precio total de la venta (no el costo, no descuentos)
- cantidad     → columna con cantidad de unidades o cajas vendidas
- cliente      → columna con nombre descriptivo del cliente (texto, no código). Si hay nombre y código, elige el nombre.
- nit          → columna con NIT o identificación numérica del cliente (puede no existir)

2. Identificar un nombre corto y estable del distribuidor a partir del nombre del archivo, ignorando palabras como meses (enero, febrero, marzo, abril, mayo, junio, julio, agosto, septiembre, octubre, noviembre, diciembre), años (2024, 2025, 2026), palabras genéricas (ventas, liquidacion, formato, planilla, reporte, cierre, corregido) y números de versión. El nombre corto debe ser la palabra o palabras que identifican ÚNICAMENTE a ese distribuidor.

Ejemplos de nombre corto:
- "Ventas GMJ abril 2026" → "GMJ"
- "LIQUIDACION TITANES MABE CIERRE ABRIL 2026-KIRAMAR" → "KIRAMAR"
- "VENTA MULTIELECTO ABRIL 2026" → "MULTIELECTO"

Reglas del mapeo:
1. Cada valor del mapeo debe ser UNA LISTA con el encabezado exacto del Excel.
2. Si un campo no tiene columna clara, usa lista vacía: [].
3. Si la columna de vendedor parece agrupada (pocos valores únicos repetidos), agrega "rellenar_vendedor": true.
4. Para "referencia" y "cliente": si hay nombre + código, SIEMPRE prioriza el nombre en texto.

Responde ÚNICAMENTE con un objeto JSON válido con esta estructura exacta, sin explicaciones ni bloques de código:
{{"_clave": "NOMBRE_CORTO_AQUI", "fecha": [...], "vendedor": [...], "cc": [...], "referencia": [...], "Valor Total": [...], "cantidad": [...], "cliente": [...], "nit": [...], "rellenar_vendedor": false}}"""

    body = json.dumps({
        "model": "claude-haiku-4-5-20251001",
        "max_tokens": 500,
        "messages": [{"role": "user", "content": prompt}]
    }).encode("utf-8")

    def construir_req():
        return urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=body,
            headers={
                "Content-Type": "application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01"
            },
            method="POST"
        )

    req = construir_req()

    for intento in range(3):
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                respuesta = json.loads(resp.read().decode("utf-8"))
                texto = respuesta["content"][0]["text"].strip()
                texto = texto.replace("```json", "").replace("```", "").strip()
                mapeo = json.loads(texto)

                clave_corta = mapeo.pop("_clave", None)
                if not clave_corta:
                    clave_corta = "_".join(nombre_archivo.split()[:2])
                clave_corta = clave_corta.strip().upper()

                return mapeo, clave_corta

        except urllib.error.HTTPError as e:
            if e.code == 429:
                espera = 20 * (intento + 1)
                time.sleep(espera)
                req = construir_req()
                continue
            return None, None
        except Exception:
            return None, None

    return None, None


# ─────────────────────────────────────────────────────────────────
#   JSON DE CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────────

def cargar_json() -> dict:
    if os.path.exists(RUTA_JSON):
        with open(RUTA_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def guardar_json(config: dict):
    with open(RUTA_JSON, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# ─────────────────────────────────────────────────────────────────
#   RESOLUCIÓN DE MAPEO
# ─────────────────────────────────────────────────────────────────

def resolver_mapeo(nombre_archivo: str, encabezados: list, config: dict, api_key: str, log):
    texto_busqueda = nombre_archivo.upper()

    # 1. Buscar en JSON existente
    for llave, datos in config.items():
        if llave.upper() in ("GENERAL",):
            continue
        if llave.upper() in texto_busqueda:
            log.info(f"🎯 Mapeo encontrado en configuración para: `{llave}`")
            return datos, datos.get("rellenar_vendedor", False)

    # 2. Bloque GENERAL como fallback
    if "GENERAL" in config:
        general = config["GENERAL"]
        coincidencias = sum(
            1 for sinonimos in general.values()
            if isinstance(sinonimos, list) and any(s in encabezados for s in sinonimos)
        )
        if coincidencias >= 2:
            log.info(f"ℹ️ Usando bloque GENERAL (coincidencias: {coincidencias})")
            return general, general.get("rellenar_vendedor", False)

    # 3. Llamada a la IA
    if api_key:
        log.info("🤖 Consultando a Claude para generar el mapeo...")
        nuevo_mapeo, clave_corta = mapear_con_ia(encabezados, nombre_archivo, api_key)
        if nuevo_mapeo and clave_corta:
            config[clave_corta] = nuevo_mapeo
            guardar_json(config)
            log.info(f"✅ IA generó mapeo y guardó clave: `{clave_corta}`")
            return nuevo_mapeo, nuevo_mapeo.get("rellenar_vendedor", False)

    return None, False


# ─────────────────────────────────────────────────────────────────
#   PROCESAMIENTO DE UN ARCHIVO
# ─────────────────────────────────────────────────────────────────

def procesar_archivo(uploaded_file, template_bytes: bytes, config: dict, api_key: str, log) -> bytes | None:
    nombre_archivo = uploaded_file.name
    file_bytes = uploaded_file.read()

    # ── 1. Detectar fila de encabezados ──────────────────────────
    df_origen, fila_detectada = detectar_fila_encabezados(file_bytes)
    log.info(f"📋 Encabezados detectados en fila {fila_detectada}: `{list(df_origen.columns)}`")

    encabezados = list(df_origen.columns)

    # ── 2. Obtener mapeo ─────────────────────────────────────────
    mapeo_sinonimos, debe_rellenar = resolver_mapeo(
        nombre_archivo, encabezados, config, api_key, log
    )

    if mapeo_sinonimos is None:
        log.error("❌ No se pudo obtener un mapeo válido.")
        return None

    # ── 3. Construir mapeo columna_origen → campo_plantilla ──────
    #
    #   CORRECCIÓN DEL BUG PRINCIPAL:
    #   Comparamos los sinónimos devueltos por la IA contra los encabezados
    #   REALES del DataFrame (ya detectados en la fila correcta).
    #   Antes, si la fila era 2 o 3, df_origen.columns tenía índices numéricos
    #   y nunca coincidía.
    #
    columnas_a_extraer = {}   # { col_en_excel: campo_plantilla }
    campos_sin_columna = []

    for campo_plantilla, sinonimos in mapeo_sinonimos.items():
        if campo_plantilla == "rellenar_vendedor":
            continue
        if not isinstance(sinonimos, list) or len(sinonimos) == 0:
            campos_sin_columna.append(campo_plantilla)
            continue

        encontrado = False
        for sinonimo in sinonimos:
            # Comparación exacta primero
            if sinonimo in df_origen.columns:
                columnas_a_extraer[sinonimo] = campo_plantilla
                encontrado = True
                break
            # Comparación case-insensitive como fallback
            coincidencia = next(
                (c for c in df_origen.columns if c.strip().lower() == sinonimo.strip().lower()),
                None
            )
            if coincidencia:
                columnas_a_extraer[coincidencia] = campo_plantilla
                encontrado = True
                break

        if not encontrado:
            campos_sin_columna.append(campo_plantilla)

    if campos_sin_columna:
        log.warning(f"⚠️ Campos sin columna encontrada: {campos_sin_columna}")

    if not columnas_a_extraer:
        log.error("❌ Ninguna columna del mapeo coincidió con los encabezados del archivo.")
        return None

    log.info(f"🔗 Columnas mapeadas: {columnas_a_extraer}")

    # ── 4. Extraer y renombrar columnas ──────────────────────────
    df_filtrado = df_origen[list(columnas_a_extraer.keys())].rename(columns=columnas_a_extraer)

    # Eliminar filas donde TODAS las columnas están vacías (basura post-encabezado)
    df_filtrado = df_filtrado.dropna(how="all").reset_index(drop=True)

    # ── 5. Cargar plantilla y obtener columnas finales ────────────
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = wb.active
    columnas_finales = [cell.value for cell in ws[1] if cell.value is not None]

    # Agregar columnas faltantes como vacías
    for col in columnas_finales:
        if col not in df_filtrado.columns:
            df_filtrado[col] = None

    df_final = df_filtrado[columnas_finales].copy()

    # ── 6. Transformaciones estándar ─────────────────────────────
    NULOS = ['nan', 'NAN', 'None', 'none', 'NONE', '', 'NaT']

    if 'vendedor' in df_final.columns:
        df_final['vendedor'] = df_final['vendedor'].replace(NULOS, None)
        if debe_rellenar:
            log.info("⬇️ Aplicando fill-down en columna 'vendedor'.")
            df_final['vendedor'] = df_final['vendedor'].ffill()

    if 'cantidad' in df_final.columns:
        df_final['cantidad'] = (
            pd.to_numeric(df_final['cantidad'], errors='coerce')
            .fillna(0).round(0).astype(int)
        )

    if 'Valor Total' in df_final.columns:
        df_final['Valor Total'] = (
            pd.to_numeric(df_final['Valor Total'], errors='coerce')
            .fillna(0).round(0).astype(int)
        )

    if 'fecha' in df_final.columns:
        df_final['fecha'] = df_final['fecha'].replace(NULOS, None)

    if 'cc' in df_final.columns:
        df_final['cc'] = df_final['cc'].fillna('Sin cc').replace(NULOS, 'Sin cc')

    if 'nit' in df_final.columns:
        df_final['nit'] = df_final['nit'].fillna('Sin nit').replace(NULOS, 'Sin nit')

    # ── 7. Escribir datos en plantilla conservando estilos ────────
    estilos_columnas = {}
    if ws.max_row >= 2:
        for col_idx in range(1, ws.max_column + 1):
            celda = ws.cell(row=2, column=col_idx)
            estilos_columnas[col_idx] = {
                'font':          copy(celda.font),
                'border':        copy(celda.border),
                'fill':          copy(celda.fill),
                'number_format': celda.number_format,
                'alignment':     copy(celda.alignment)
            }
        ws.delete_rows(2, amount=ws.max_row)

    for r_idx, row in enumerate(dataframe_to_rows(df_final, index=False, header=False), start=2):
        for c_idx, value in enumerate(row, start=1):
            nueva_celda = ws.cell(row=r_idx, column=c_idx, value=value)
            if c_idx in estilos_columnas:
                for prop in ('font', 'border', 'fill', 'number_format', 'alignment'):
                    val = estilos_columnas[c_idx].get(prop)
                    if val:
                        setattr(nueva_celda, prop, val)

    # Limpiar última fila (totales/basura al final)
    ultima_fila = ws.max_row
    if ultima_fila >= 2:
        indices_columnas = {nombre: idx + 1 for idx, nombre in enumerate(columnas_finales)}
        for campo in ['fecha', 'cc', 'nit', 'vendedor', 'cliente', 'referencia']:
            if campo in indices_columnas:
                ws.cell(row=ultima_fila, column=indices_columnas[campo], value=None)

    # ── 8. Guardar en buffer y retornar ──────────────────────────
    out_buffer = io.BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer.getvalue()


# ═══════════════════════════════════════════════════════════════════
#   INTERFAZ PRINCIPAL
# ═══════════════════════════════════════════════════════════════════

st.markdown("### 📁 Carga de archivos")

template_file = st.file_uploader(
    "1. Archivo PLANTILLA (estructura de destino)",
    type=["xlsx"],
    help="El archivo con los encabezados estándar: fecha, vendedor, cc, referencia, Valor Total, cantidad, cliente, nit"
)

uploaded_files = st.file_uploader(
    "2. Archivos de DISTRIBUIDORES",
    type=["xlsx"],
    accept_multiple_files=True,
    help="Puedes subir varios archivos a la vez. El sistema detecta automáticamente en qué fila están los encabezados."
)

# Mostrar aviso de detección automática
if uploaded_files:
    st.info(f"📂 {len(uploaded_files)} archivo(s) cargado(s). El sistema buscará encabezados en las primeras 3 filas de cada uno.")

st.markdown("---")

# ── BOTÓN DE PROCESAMIENTO ───────────────────────────────────────
listo = bool(uploaded_files and template_file)

if st.button("🚀 Iniciar Procesamiento", disabled=not listo, use_container_width=True):

    config = cargar_json()
    template_bytes = template_file.read()

    exitosos: dict[str, bytes] = {}
    fallidos: list[str] = []

    progress_bar = st.progress(0, text="Iniciando...")
    total = len(uploaded_files)

    for idx, archivo in enumerate(uploaded_files):
        nombre = archivo.name
        progress_bar.progress((idx) / total, text=f"Procesando {idx+1}/{total}: {nombre}")

        # Expander con log en tiempo real por archivo
        with st.expander(f"📄 {nombre}", expanded=True):
            log = st.status(f"Procesando {nombre}...", expanded=True)

            resultado = procesar_archivo(archivo, template_bytes, config, api_key, log)

            if resultado:
                exitosos[f"Limpio_{nombre}"] = resultado
                log.update(label=f"✅ {nombre} — procesado correctamente", state="complete", expanded=False)
            else:
                fallidos.append(nombre)
                log.update(label=f"❌ {nombre} — falló (ver detalle arriba)", state="error", expanded=True)

    progress_bar.progress(1.0, text="¡Procesamiento completado!")

    # ── RESUMEN ──────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📊 Resumen")
    col1, col2 = st.columns(2)
    col1.metric("✅ Exitosos", len(exitosos))
    col2.metric("❌ Fallidos", len(fallidos))

    if fallidos:
        with st.expander("Ver archivos fallidos"):
            for f in fallidos:
                st.error(f"• {f}")

    # ── DESCARGA ─────────────────────────────────────────────────
    if exitosos:
        if len(exitosos) == 1:
            # Un solo archivo: descarga directa
            nombre_unico, bytes_unico = next(iter(exitosos.items()))
            st.download_button(
                label=f"📥 Descargar {nombre_unico}",
                data=bytes_unico,
                file_name=nombre_unico,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            # Varios archivos: ZIP
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname, fbytes in exitosos.items():
                    zf.writestr(fname, fbytes)
            zip_buffer.seek(0)

            st.download_button(
                label="📥 Descargar todos los archivos procesados (.ZIP)",
                data=zip_buffer,
                file_name="Distribuciones_Procesadas.zip",
                mime="application/zip",
                use_container_width=True
            )