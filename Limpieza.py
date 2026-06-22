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
#   CONFIGURACIÓN E INTERFAZ WEB
# ─────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Sistema de Liquidación", page_icon="🐐", layout="centered")

st.title("╔════════════════════════════════════╗")
st.subheader("  SISTEMA DE LIQUIDACIÓN DE DISTRIBUIDORES")
st.caption("Creado por Nicolas Tovar y Cristian Morales")
st.title("╚════════════════════════════════════╝")

CAMPOS_PLANTILLA = ["fecha", "vendedor", "cc", "referencia", "Valor Total", "cantidad", "cliente", "nit"]
RUTA_JSON = "config_distribuidores.json"

# --- SIDEBAR: Configuración de API Key ---
st.sidebar.header("🔑 Configuración de Acceso")
api_key_input = st.sidebar.text_input("Anthropic API Key", type="password", help="Ingresa tu API Key para habilitar el mapeo con IA.")

# fallback si está en variables de entorno del servidor
api_key = api_key_input.strip() if api_key_input else os.environ.get("ANTHROPIC_API_KEY", "")

if not api_key:
    st.sidebar.warning("⚠️ Sin API key. Solo se usarán mapeos del JSON existente.")
else:
    st.sidebar.success("⚡ API Key cargada exitosamente.")

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
#   MANEJO DE CONFIGURACIÓN JSON
# ─────────────────────────────────────────────────────────────────

def cargar_json() -> dict:
    if os.path.exists(RUTA_JSON):
        with open(RUTA_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def guardar_json(config: dict):
    with open(RUTA_JSON, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

def resolver_mapeo(nombre_archivo: str, encabezados: list, config: dict, api_key: str, status_placeholder):
    texto_busqueda = nombre_archivo.upper()

    # 1. Buscar en JSON existente
    for llave, datos in config.items():
        if llave.upper() in ("GENERAL",):
            continue
        if llave.upper() in texto_busqueda:
            status_placeholder.write(f"🎯 Mapeo encontrado en configuración para: `{llave}`")
            return datos, datos.get("rellenar_vendedor", False)

    # 2. Bloque GENERAL como fallback
    if "GENERAL" in config:
        general = config["GENERAL"]
        coincidencias = sum(
            1 for sinonimos in general.values()
            if isinstance(sinonimos, list) and any(s in encabezados for s in sinonimos)
        )
        if coincidencias >= 2:
            status_placeholder.write(f"ℹ️ Usando bloque GENERAL (coincidencias: {coincidencias})")
            return general, general.get("rellenar_vendedor", False)

    # 3. Llamada a la IA
    if api_key:
        status_placeholder.write(f"🤖 No se encontró mapeo local. Consultando a Claude...")
        nuevo_mapeo, clave_corta = mapear_con_ia(encabezados, nombre_archivo, api_key)
        if nuevo_mapeo and clave_corta:
            config[clave_corta] = nuevo_mapeo
            guardar_json(config)
            status_placeholder.write(f"🤖 IA generó mapeo exitoso y guardó clave: `{clave_corta}`")
            return nuevo_mapeo, nuevo_mapeo.get("rellenar_vendedor", False)

    return None, False

# ─────────────────────────────────────────────────────────────────
#   CARGA DE ARCHIVOS EN WEB
# ─────────────────────────────────────────────────────────────────

st.markdown("### 📁 Carga de insumos")
template_file = st.file_uploader("1. Selecciona el archivo PLANTILLA (.xlsx)", type=["xlsx"])
uploaded_files = st.file_uploader("2. Selecciona los archivos de los DISTRIBUIDORES (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if st.button("🚀 Comenzar Procesamiento", use_container_width=True) and uploaded_files and template_file:
    
    config = cargar_json()
    exitosos = {}
    fallidos = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Leer plantilla en memoria para duplicar sus estilos eficientemente
    template_bytes = template_file.read()
    
    for idx, uploaded_file in enumerate(uploaded_files):
        nombre_archivo = uploaded_file.name
        status_text.markdown(f"**Procesando:** `{nombre_archivo}`...")
        
        try:
            # Leer origen
            df_origen = pd.read_excel(uploaded_file, dtype=str)
            df_origen.columns = df_origen.columns.str.strip()
            encabezados = list(df_origen.columns)
            
            sub_status = st.empty()
            mapeo_sinonimos, debe_rellenar = resolver_mapeo(nombre_archivo, encabezados, config, api_key, sub_status)
            
            if mapeo_sinonimos is None:
                fallidos.append(nombre_archivo)
                continue
                
            columnas_a_extraer = {}
            for col_plantilla, sinonimos in mapeo_sinonimos.items():
                if col_plantilla == "rellenar_vendedor" or not isinstance(sinonimos, list):
                    continue
                for sinonimo in sinonimos:
                    if sinonimo in df_origen.columns:
                        columnas_a_extraer[sinonimo] = col_plantilla
                        break
                        
            if not columnas_a_extraer:
                fallidos.append(nombre_archivo)
                continue
                
            df_filtrado = df_origen[list(columnas_a_extraer.keys())].rename(columns=columnas_a_extraer)
            
            # Cargar estructura de plantilla limpia desde los bytes cargados
            wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
            ws = wb.active
            columnas_finales = [cell.value for cell in ws[1]]
            
            for col in columnas_finales:
                if col not in df_filtrado.columns:
                    df_filtrado[col] = None
                    
            df_final = df_filtrado[columnas_finales].copy()
            
            # Transformaciones estándar
            if 'vendedor' in df_final.columns:
                df_final['vendedor'] = df_final['vendedor'].replace(['nan', 'NAN', 'None', ''], None)
                if debe_rellenar:
                    df_final['vendedor'] = df_final['vendedor'].ffill()
                    
            if 'cantidad' in df_final.columns:
                df_final['cantidad'] = pd.to_numeric(df_final['cantidad'], errors='coerce').fillna(0).round(0).astype(int)
            if 'Valor Total' in df_final.columns:
                df_final['Valor Total'] = pd.to_numeric(df_final['Valor Total'], errors='coerce').fillna(0).round(0).astype(int)
            if 'fecha' in df_final.columns:
                df_final['fecha'] = df_final['fecha'].replace(['nan', 'NAN', 'None', ''], None)
            if 'cc' in df_final.columns:
                df_final['cc'] = df_final['cc'].fillna('Sin cc').replace(['nan', 'NAN', 'None', ''], 'Sin cc')
            if 'nit' in df_final.columns:
                df_final['nit'] = df_final['nit'].fillna('Sin nit').replace(['nan', 'NAN', 'None', ''], 'Sin nit')
                
            estilos_columnas = {}
            if ws.max_row >= 2:
                for col_idx in range(1, ws.max_column + 1):
                    celda = ws.cell(row=2, column=col_idx)
                    estilos_columnas[col_idx] = {
                        'font': copy(celda.font), 'border': copy(celda.border),
                        'fill': copy(celda.fill), 'number_format': celda.number_format,
                        'alignment': copy(celda.alignment)
                    }
                    
            if ws.max_row >= 2:
                ws.delete_rows(2, amount=ws.max_row)
                
            for r_idx, row in enumerate(dataframe_to_rows(df_final, index=False, header=False), start=2):
                for c_idx, value in enumerate(row, start=1):
                    nueva_celda = ws.cell(row=r_idx, column=c_idx, value=value)
                    if c_idx in estilos_columnas:
                        for prop in ('font', 'border', 'fill', 'number_format', 'alignment'):
                            if estilos_columnas[c_idx][prop]:
                                setattr(nueva_celda, prop, estilos_columnas[c_idx][prop])
                                
            ultima_fila = ws.max_row
            if ultima_fila >= 2:
                indices_columnas = {nombre: idx + 1 for idx, nombre in enumerate(columnas_finales)}
                for campo in ['fecha', 'cc', 'nit', 'vendedor', 'cliente', 'referencia']:
                    if campo in indices_columnas:
                        ws.cell(row=ultima_fila, column=indices_columnas[campo], value=None)
            
            # Guardar en buffer de memoria
            out_buffer = io.BytesIO()
            wb.save(out_buffer)
            out_buffer.seek(0)
            exitosos[f"Limpio_{nombre_archivo}"] = out_buffer.getvalue()
            
        except Exception as e:
            fallidos.append(nombre_archivo)
            
        progress_bar.progress((idx + 1) / len(uploaded_files))
        
    status_text.success("🏁 ¡Procesamiento completado!")
    
    # --- RESULTADOS ---
    st.markdown("### 📊 Resumen")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Exitosos", len(exitosos))
    with col2:
        st.metric("Fallidos", len(fallidos))
        
    if fallidos:
        with st.expander("Ver archivos fallidos"):
            for f in fallidos:
                st.error(f"• {f}")
                
    if exitosos:
        # Empaquetar todo en un ZIP listo para descarga web
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for file_name, file_bytes in exitosos.items():
                zip_file.writestr(file_name, file_bytes)
        zip_buffer.seek(0)
        
        st.markdown("---")
        st.download_button(
            label="📥 Descargar Todos los Archivos Procesados (.ZIP)",
            data=zip_buffer,
            file_name="Distribuciones_Procesadas.zip",
            mime="application/zip",
            use_container_width=True
        )