# ======================================================
# PUMI 2026 - DASHBOARD NACIONAL EJECUTIVO
# Lectura automática de Excel regionales del repositorio
# Dashboard, filtros, marcadores nacionales e informe PDF
# ======================================================

import os
import re
import glob
import base64
import unicodedata
from io import BytesIO
from datetime import date

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import folium
from streamlit_folium import st_folium

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image

# ======================================================
# CONFIGURACIÓN
# ======================================================

st.set_page_config(
    page_title="PUMI 2026 - Dashboard Nacional",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

LOGO_MINISTERIO = "Logo2.jpeg"
LOGO_PUMI = "logo_pumi.jpeg"
LOGO_FUERZA_PUBLICA = "Logo1.jpeg"
ARCHIVO_DELEGACIONES = "DELEGACIONES Y DISTRITOS.xlsx"

COLOR_AZUL = "#002B7F"
COLOR_AZUL_MEDIO = "#1E4FA3"
COLOR_AZUL_CLARO = "#DCE8FF"
COLOR_DORADO = "#B88A2A"
COLOR_VERDE = "#1F8A4C"
COLOR_NARANJA = "#F59E0B"
COLOR_ROJO = "#B42318"
COLOR_GRIS = "#2F3542"
COLOR_BLANCO = "#FFFFFF"

UMBRAL_CUMPLE = 0.50
UMBRAL_RIESGO = 0.45

MESES_META = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
]

MAPA_REGION_ARCHIVO = {
    "DR1 C": "Dirección Regional 1 - San José Central",
    "DR1 N": "Dirección Regional 1 - San José Norte",
    "DR1 S": "Dirección Regional 1 - San José Sur",
    "DR2": "Dirección Regional 2 - Alajuela",
    "DR3": "Dirección Regional 3 - Cartago",
    "DR4": "Dirección Regional 4 - Heredia",
    "DR5": "Dirección Regional 5 - Guanacaste",
    "DR6": "Dirección Regional 6 - Puntarenas",
    "DR7": "Dirección Regional 7 - Limón",
    "DR8": "Dirección Regional 8",
    "DR9": "Dirección Regional 9",
    "DR10": "Dirección Regional 10",
    "DR11": "Dirección Regional 11",
    "DR12": "Dirección Regional 12 - Caribe",
}

# Coordenadas aproximadas para ubicar las delegaciones en el mapa. No se muestran en tablas ni PDF.
COORDENADAS_REFERENCIA = {
    "CARMEN": [9.9365, -84.0750], "MERCED": [9.9386, -84.0828], "HOSPITAL": [9.9274, -84.0918],
    "CATEDRAL": [9.9289, -84.0740], "ZAPOTE": [9.9198, -84.0553], "SAN FRANCISCO": [9.9136, -84.0724],
    "URUCA": [9.9567, -84.1060], "MATA REDONDA": [9.9352, -84.1047], "PAVAS": [9.9488, -84.1342],
    "HATILLO": [9.9160, -84.1010], "SAN SEBASTIAN": [9.9121, -84.0909], "ESCAZU": [9.9180, -84.1399],
    "SANTA ANA": [9.9326, -84.1825], "ALAJUELITA": [9.9016, -84.1000], "VASQUEZ DE CORONADO": [9.9760, -84.0070],
    "CORONADO": [9.9760, -84.0070], "ACOSTA": [9.8003, -84.1604], "TIBAS": [9.9580, -84.0790],
    "MORAVIA": [9.9610, -84.0480], "MONTES DE OCA": [9.9369, -84.0500], "CURRIDABAT": [9.9136, -84.0405],
    "GOICOECHEA": [9.9480, -84.0430], "DESAMPARADOS": [9.8982, -84.0626], "ASERRI": [9.8587, -84.0917],
    "MORA": [9.9182, -84.2411], "PURISCAL": [9.8469, -84.3149], "TARRAZU": [9.6596, -84.0206],
    "DOTA": [9.6500, -83.9600], "LEON CORTES": [9.6830, -84.0500], "TURRUBARES": [9.9050, -84.4520],
    "ALAJUELA": [10.0162, -84.2116], "SAN RAMON": [10.0887, -84.4702], "GRECIA": [10.0739, -84.3112],
    "SAN MATEO": [9.9365, -84.5247], "ATENAS": [9.9787, -84.3801], "NARANJO": [10.0987, -84.3782],
    "PALMARES": [10.0567, -84.4370], "POAS": [10.0800, -84.2450], "OROTINA": [9.9111, -84.5230],
    "SAN CARLOS": [10.3290, -84.4310], "ZARCERO": [10.1852, -84.3900], "SARCHI": [10.0883, -84.3473],
    "UPALA": [10.8986, -85.0155], "LOS CHILES": [11.0350, -84.7130], "GUATUSO": [10.6667, -84.8167],
    "RIO CUARTO": [10.3410, -84.2140], "CARTAGO": [9.8644, -83.9194], "PARAISO": [9.8383, -83.8656],
    "LA UNION": [9.9084, -83.9886], "JIMENEZ": [9.9048, -83.6834], "TURRIALBA": [9.9050, -83.6830],
    "ALVARADO": [9.9333, -83.8000], "OREAMUNO": [9.9100, -83.9000], "EL GUARCO": [9.8472, -83.9460],
    "HEREDIA": [10.0024, -84.1165], "BARVA": [10.0208, -84.1233], "SANTO DOMINGO": [10.0639, -84.1547],
    "SANTA BARBARA": [10.0400, -84.1600], "SAN RAFAEL": [10.0138, -84.1002], "SAN ISIDRO": [10.0186, -84.0569],
    "BELEN": [9.9852, -84.1810], "FLORES": [10.0000, -84.1600], "SAN PABLO": [9.9953, -84.0966],
    "SARAPIQUI": [10.4522, -84.0166], "LIBERIA": [10.6350, -85.4377], "NICOYA": [10.1483, -85.4520],
    "SANTA CRUZ": [10.2600, -85.5850], "BAGACES": [10.5250, -85.2550], "CARRILLO": [10.4750, -85.5850],
    "CANAS": [10.4310, -85.0980], "ABANGARES": [10.2820, -84.9590], "TILARAN": [10.4670, -84.9670],
    "NANDAYURE": [9.9990, -85.2060], "LA CRUZ": [11.0730, -85.6320], "HOJANCHA": [10.0550, -85.4200],
    "PUNTARENAS": [9.9763, -84.8384], "CHOMES": [10.0950, -84.9250], "JUDAS": [10.0510, -84.8870],
    "ESPARZA": [9.9940, -84.6640], "BUENOS AIRES": [9.1667, -83.3333], "MONTES DE ORO": [10.0870, -84.7300],
    "OSA": [8.9590, -83.5230], "QUEPOS": [9.4319, -84.1617], "GOLFITO": [8.6390, -83.1660],
    "COTO BRUS": [8.8830, -82.9660], "PARRITA": [9.5200, -84.3200], "CORREDORES": [8.6420, -82.9460],
    "GARABITO": [9.6150, -84.6300], "LIMON": [9.9917, -83.0360], "POCOCI": [10.2150, -83.7870],
    "SIQUIRRES": [10.0970, -83.5060], "TALAMANCA": [9.6240, -82.8440], "MATINA": [10.0760, -83.2890],
    "GUACIMO": [10.2100, -83.6900], "PEREZ ZELEDON": [9.3540, -83.6340], "LOS SANTOS": [9.6550, -84.0300],
}

REGION_CENTRO = {
    "1": [9.93, -84.08], "2": [10.05, -84.32], "3": [9.87, -83.93], "4": [10.02, -84.12],
    "5": [10.45, -85.30], "6": [9.75, -84.70], "7": [10.10, -83.55], "8": [9.75, -84.20],
    "9": [9.40, -84.00], "10": [9.30, -83.40], "11": [10.80, -85.00], "12": [10.15, -83.50]
}

# ======================================================
# CSS
# ======================================================

st.markdown(f"""
<style>
.stApp {{
  background: radial-gradient(circle at top right, #DCE8FF 0, transparent 28%), linear-gradient(180deg,#FFFFFF 0%,#F5F8FC 48%,#FFFFFF 100%);
}}
section[data-testid="stSidebar"] {{
  background: linear-gradient(180deg,#FFFFFF 0%,#EAF1FF 50%,#FFFFFF 100%);
  border-right:5px solid {COLOR_DORADO};
}}
.encabezado-institucional {{
  background:white;border-radius:22px;padding:20px 34px;margin-bottom:24px;box-shadow:0 8px 22px rgba(0,0,0,.11);border-bottom:7px solid {COLOR_DORADO};
  display:grid;grid-template-columns:1fr 1fr 1fr;align-items:center;gap:20px;min-height:145px;
}}
.encabezado-institucional img {{max-height:115px;max-width:100%;object-fit:contain;}}
.titulo-principal {{
  background:linear-gradient(135deg,{COLOR_AZUL},#243B63,{COLOR_DORADO});padding:34px;border-radius:22px;text-align:center;color:white;margin-bottom:30px;
  box-shadow:0 8px 22px rgba(0,0,0,.28);border:3px solid white;
}}
.titulo-principal h1 {{font-size:44px;margin:0 0 8px 0;color:white;font-weight:900;letter-spacing:1px;text-shadow:1px 2px 5px rgba(0,0,0,.35);}}
.titulo-principal h3 {{font-size:23px;margin:0;color:white;font-weight:600;}}
.card-info {{background:white;border-radius:20px;padding:24px;box-shadow:0 6px 16px rgba(0,0,0,.12);border-left:9px solid {COLOR_AZUL};margin-bottom:18px;}}
.card-title {{color:{COLOR_AZUL};font-weight:900;font-size:28px;text-align:center;margin-bottom:10px;}}
.card-text {{color:{COLOR_GRIS};font-size:17px;line-height:1.65;text-align:justify;}}
.kpi-grid {{display:grid;grid-template-columns:repeat(6,minmax(135px,1fr));gap:18px;margin:18px 0 24px 0;}}
.kpi-card {{background:linear-gradient(145deg,#FFFFFF,#F3F6FF);border-radius:20px;padding:18px 16px;box-shadow:0 8px 22px rgba(0,0,0,.12);border-bottom:7px solid {COLOR_DORADO};min-height:125px;display:flex;flex-direction:column;justify-content:center;}}
.kpi-label {{font-size:15px;color:#4B5563;font-weight:800;line-height:1.2;margin-bottom:8px;}}
.kpi-value {{font-size:30px;color:#111827;font-weight:900;line-height:1.05;word-break:break-word;}}
.estado-cumple {{background:#DFF5E8;color:#064E3B;font-weight:800;}}
.estado-riesgo {{background:#FFF3D6;color:#7C2D12;font-weight:800;}}
.estado-critico {{background:#FFE2E2;color:#7F1D1D;font-weight:800;}}
.stButton > button, .stDownloadButton > button {{background:linear-gradient(90deg,{COLOR_AZUL},{COLOR_DORADO});color:white;border:none;border-radius:12px;font-weight:800;padding:.7rem 1.2rem;box-shadow:0 4px 10px rgba(0,0,0,.18);}}
div[data-baseweb="select"] > div {{background-color:white !important;border-radius:12px !important;border:1.8px solid #AFC3EE !important;}}
div[data-testid="stDataFrame"] {{border-radius:18px;overflow:hidden;box-shadow:0 5px 16px rgba(0,0,0,.13);background:white;}}
h1,h2,h3 {{color:{COLOR_GRIS};font-weight:900;}}
hr {{border:none;height:3px;background:linear-gradient(90deg,{COLOR_AZUL},white,{COLOR_DORADO});margin:24px 0;}}
</style>
""", unsafe_allow_html=True)

# ======================================================
# UTILIDADES
# ======================================================

def normalizar_texto(valor):
    if pd.isna(valor): return ""
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^A-Z0-9 ]", " ", texto)
    return " ".join(texto.split())


def limpiar_numero(valor):
    if pd.isna(valor) or valor == "": return 0.0
    if isinstance(valor, str): valor = valor.replace("%", "").replace(",", ".").strip()
    try: return float(valor)
    except Exception: return 0.0


def limpiar_nombre_archivo(valor):
    texto = normalizar_texto(valor)
    texto = re.sub(r"\s+", "_", texto)
    return texto if texto else "SIN_DATO"


def imagen_a_base64(ruta):
    if not os.path.exists(ruta): return ""
    try:
        with open(ruta, "rb") as f: return base64.b64encode(f.read()).decode()
    except Exception: return ""


def extraer_numero_region(valor):
    texto = normalizar_texto(valor)
    m = re.search(r"REGIONAL\s+(\d+)", texto) or re.search(r"\bDR\s*(\d+)", texto) or re.search(r"\bR\s*(\d+)", texto)
    return int(m.group(1)) if m else 999


def orden_region(valor):
    texto = normalizar_texto(valor)
    n = extraer_numero_region(valor)
    sub = 0
    if n == 1:
        if "CENTRAL" in texto or re.search(r"DR1\s*C", texto): sub = 0
        elif "NORTE" in texto or re.search(r"DR1\s*N", texto): sub = 1
        elif "SUR" in texto or re.search(r"DR1\s*S", texto): sub = 2
        else: sub = 9
    return (n, sub, texto)


def orden_delegacion(valor):
    texto = normalizar_texto(valor)
    m = re.search(r"\bD\s*(\d+)", texto)
    return (int(m.group(1)) if m else 9999, texto)


def clasificar_cumplimiento(pct):
    if pct >= UMBRAL_CUMPLE: return "Cumple"
    if pct >= UMBRAL_RIESGO: return "En riesgo"
    return "Crítico"


def color_estado(estado):
    return {"Cumple": COLOR_VERDE, "En riesgo": COLOR_NARANJA, "Crítico": COLOR_ROJO}.get(estado, COLOR_GRIS)


def obtener_region_desde_archivo(nombre_archivo):
    nombre = str(nombre_archivo).upper().replace("_", " ")
    nombre = re.sub(r"\s+", " ", nombre).strip()
    for clave in sorted(MAPA_REGION_ARCHIVO.keys(), key=len, reverse=True):
        if nombre.startswith(clave): return MAPA_REGION_ARCHIVO[clave]
    m = re.match(r"(DR\d+)", nombre)
    if m: return MAPA_REGION_ARCHIVO.get(m.group(1), m.group(1))
    return "Región no identificada"


def normalizar_columnas(cols):
    return [re.sub(r"\s+", " ", str(c).strip().upper()) for c in cols]


def detectar_coord_delegacion(delegacion, region=""):
    txt = normalizar_texto(delegacion)
    # Primero busca coincidencia completa por palabras relevantes.
    for clave, coord in COORDENADAS_REFERENCIA.items():
        if clave in txt:
            return coord
    # Si no encuentra, usa centro regional con leve desplazamiento estable para no montar todos los puntos.
    n = str(extraer_numero_region(region))
    base = REGION_CENTRO.get(n, [9.7489, -83.7534])
    numero = orden_delegacion(delegacion)[0]
    if numero == 9999: numero = abs(hash(txt)) % 50
    offset_lat = ((numero % 7) - 3) * 0.025
    offset_lon = (((numero // 7) % 7) - 3) * 0.025
    return [base[0] + offset_lat, base[1] + offset_lon]

# ======================================================
# CARGA DE DATOS
# ======================================================

@st.cache_data(show_spinner=False)
def cargar_catalogo_delegaciones():
    if not os.path.exists(ARCHIVO_DELEGACIONES):
        return pd.DataFrame(columns=["Dirección Regional", "Delegación"])
    try:
        df = pd.read_excel(ARCHIVO_DELEGACIONES)
        df.columns = [str(c).strip() for c in df.columns]
        col_region = next((c for c in df.columns if normalizar_texto(c) in ["DIRECCION REGIONAL", "REGION", "DIRECCION"]), None)
        col_deleg = next((c for c in df.columns if normalizar_texto(c) in ["DELEGACION", "DELEGACION POLICIAL", "DELEGACIONES"]), None)
        if not col_region or not col_deleg:
            return pd.DataFrame(columns=["Dirección Regional", "Delegación"])
        out = df[[col_region, col_deleg]].copy()
        out.columns = ["Dirección Regional", "Delegación"]
        out = out.fillna("").astype(str)
        out = out[(out["Delegación"].str.strip() != "")]
        out = out.drop_duplicates()
        return out
    except Exception:
        return pd.DataFrame(columns=["Dirección Regional", "Delegación"])


@st.cache_data(show_spinner=True)
def cargar_metas_regionales():
    patrones = ["DR* AVANCE JUNIO.xlsx", "DR*AVANCE JUNIO.xlsx", "DR* AVANCE JUNIO.xlsm", "DR*AVANCE JUNIO.xlsm"]
    archivos = []
    for patron in patrones:
        archivos.extend(glob.glob(patron))
        archivos.extend(glob.glob(os.path.join("/mnt/data", patron)))
    archivos = sorted(list(dict.fromkeys(archivos)), key=lambda r: orden_region(os.path.basename(r)))
    registros = []

    for ruta in archivos:
        nombre_archivo = os.path.basename(ruta)
        region = obtener_region_desde_archivo(nombre_archivo)
        try:
            xl = pd.ExcelFile(ruta)
        except Exception:
            continue

        for hoja in sorted(xl.sheet_names, key=orden_delegacion):
            delegacion = str(hoja).strip()
            if not delegacion or normalizar_texto(delegacion).startswith("TOTAL"):
                continue
            try:
                raw = pd.read_excel(ruta, sheet_name=hoja, header=None)
            except Exception:
                continue
            if raw.empty:
                continue

            fila_header = None
            for i in range(min(15, len(raw))):
                fila = " ".join(raw.iloc[i].astype(str).fillna("").str.upper().tolist())
                if "PROGRAMA" in fila and "META" in fila:
                    fila_header = i
                    break
            if fila_header is None:
                continue

            try:
                df = pd.read_excel(ruta, sheet_name=hoja, header=fila_header)
            except Exception:
                continue
            if df.empty or len(df.columns) < 4:
                continue
            df.columns = normalizar_columnas(df.columns)
            df = df.dropna(how="all")

            col_programa = df.columns[0]
            col_actividad = df.columns[1]
            col_meta = next((c for c in df.columns if "META" in c), None)
            col_avance = next((c for c in df.columns if "AVANCE" in c), None)
            if not col_meta:
                continue
            if not col_avance:
                df["AVANCE"] = 0
                col_avance = "AVANCE"
            df[col_programa] = df[col_programa].ffill()

            for _, row in df.iterrows():
                programa = str(row.get(col_programa, "")).strip()
                actividad = str(row.get(col_actividad, "")).strip()
                if not actividad or normalizar_texto(actividad) in ["NAN", "TOTAL", "TOTALES"]:
                    continue
                if not programa or normalizar_texto(programa) == "NAN":
                    continue
                meta = limpiar_numero(row.get(col_meta, 0))
                avance = limpiar_numero(row.get(col_avance, 0))
                pendiente = max(meta - avance, 0)
                pct = avance / meta if meta else 0
                lat, lon = detectar_coord_delegacion(delegacion, region)
                item = {
                    "Archivo meta": nombre_archivo,
                    "Dirección Regional": region,
                    "Delegación": delegacion,
                    "Programa": programa,
                    "Actividad": actividad,
                    "Meta": meta,
                    "Avance": avance,
                    "Pendiente": pendiente,
                    "% Cumplimiento": pct,
                    "Estado": clasificar_cumplimiento(pct),
                    "_lat": lat,
                    "_lon": lon,
                }
                for mes in MESES_META:
                    col_mes = next((c for c in df.columns if c == mes), None)
                    item[mes.title()] = limpiar_numero(row.get(col_mes, 0)) if col_mes else 0
                registros.append(item)

    columnas = ["Archivo meta", "Dirección Regional", "Delegación", "Programa", "Actividad", "Meta", "Avance", "Pendiente", "% Cumplimiento", "Estado", "_lat", "_lon"]
    if not registros:
        return pd.DataFrame(columns=columnas)
    df = pd.DataFrame(registros)
    df = df.sort_values(["Dirección Regional", "Delegación", "Programa", "Actividad"], key=lambda s: s.map(normalizar_texto)).reset_index(drop=True)
    return df

# ======================================================
# ENCABEZADOS Y TARJETAS
# ======================================================

def mostrar_logo_sidebar():
    logo = imagen_a_base64(LOGO_PUMI)
    if logo:
        st.sidebar.markdown(f"<div style='background:white;border-radius:16px;padding:12px;margin:8px 0 18px;text-align:center;box-shadow:0 4px 14px rgba(0,0,0,.10);'><img src='data:image/jpeg;base64,{logo}' style='width:100%;max-height:145px;object-fit:contain;'></div>", unsafe_allow_html=True)


def mostrar_encabezado():
    logos = []
    for ruta in [LOGO_MINISTERIO, LOGO_PUMI, LOGO_FUERZA_PUBLICA]:
        b64 = imagen_a_base64(ruta)
        logos.append(f"<img src='data:image/jpeg;base64,{b64}'>" if b64 else "")
    st.markdown(f"""
    <div class='encabezado-institucional'>
      <div style='text-align:left'>{logos[0]}</div>
      <div style='text-align:center'>{logos[1]}</div>
      <div style='text-align:right'>{logos[2]}</div>
    </div>
    """, unsafe_allow_html=True)


def mostrar_titulo():
    st.markdown("""
    <div class='titulo-principal'>
      <h1>📊 P.U.M.I. 2026 - Dashboard Nacional</h1>
      <h3>Visualización integral de metas, avances, cumplimiento y delegaciones críticas</h3>
    </div>
    """, unsafe_allow_html=True)


def tarjetas(items):
    html = ""
    for titulo, valor in items:
        html += f"<div class='kpi-card'><div class='kpi-label'>{titulo}</div><div class='kpi-value'>{valor}</div></div>"
    st.markdown(f"<div class='kpi-grid'>{html}</div>", unsafe_allow_html=True)

# ======================================================
# FILTROS Y RESÚMENES
# ======================================================

def aplicar_filtros(df, key="general"):
    if df.empty:
        return df, {}
    st.markdown("## Filtros")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        regiones = sorted(df["Dirección Regional"].dropna().unique().tolist(), key=orden_region)
        filtro_region = st.multiselect("Dirección Regional", regiones, key=f"{key}_region")
    base_deleg = df[df["Dirección Regional"].isin(filtro_region)] if filtro_region else df
    with c2:
        delegaciones = sorted(base_deleg["Delegación"].dropna().unique().tolist(), key=orden_delegacion)
        filtro_delegacion = st.multiselect("Delegación", delegaciones, key=f"{key}_delegacion")
    with c3:
        programas = sorted(df["Programa"].dropna().unique().tolist(), key=normalizar_texto)
        filtro_programa = st.multiselect("Programa", programas, key=f"{key}_programa")
    with c4:
        filtro_estado = st.multiselect("Cumplimiento", ["Cumple", "En riesgo", "Crítico"], key=f"{key}_estado")

    out = df.copy()
    if filtro_region: out = out[out["Dirección Regional"].isin(filtro_region)]
    if filtro_delegacion: out = out[out["Delegación"].isin(filtro_delegacion)]
    if filtro_programa: out = out[out["Programa"].isin(filtro_programa)]
    if filtro_estado: out = out[out["Estado"].isin(filtro_estado)]
    filtros = {"Dirección Regional": filtro_region, "Delegación": filtro_delegacion, "Programa": filtro_programa, "Cumplimiento": filtro_estado}
    st.session_state["filtros_pdf"] = filtros
    return out, filtros


def resumen_por_delegacion(df):
    if df.empty: return pd.DataFrame()
    res = df.groupby(["Dirección Regional", "Delegación"], as_index=False).agg(
        Actividades=("Actividad", "count"), Meta=("Meta", "sum"), Avance=("Avance", "sum"), Pendiente=("Pendiente", "sum"), _lat=("_lat", "first"), _lon=("_lon", "first")
    )
    res = res[(res["Meta"] > 0) | (res["Avance"] > 0) | (res["Pendiente"] > 0)]
    res["% Cumplimiento"] = res.apply(lambda x: x["Avance"] / x["Meta"] if x["Meta"] else 0, axis=1)
    res["Estado"] = res["% Cumplimiento"].apply(clasificar_cumplimiento)
    res = res.sort_values(["Dirección Regional", "Delegación"], key=lambda s: s.map(normalizar_texto)).reset_index(drop=True)
    return res


def resumen_por_region(df):
    if df.empty: return pd.DataFrame()
    res = df.groupby("Dirección Regional", as_index=False).agg(
        Delegaciones=("Delegación", "nunique"), Actividades=("Actividad", "count"), Meta=("Meta", "sum"), Avance=("Avance", "sum"), Pendiente=("Pendiente", "sum")
    )
    res = res[(res["Meta"] > 0) | (res["Avance"] > 0) | (res["Pendiente"] > 0)]
    res["% Cumplimiento"] = res.apply(lambda x: x["Avance"] / x["Meta"] if x["Meta"] else 0, axis=1)
    res["Estado"] = res["% Cumplimiento"].apply(clasificar_cumplimiento)
    return res.sort_values("Dirección Regional", key=lambda s: s.map(orden_region)).reset_index(drop=True)


def resumen_por_programa(df):
    if df.empty: return pd.DataFrame()
    res = df.groupby("Programa", as_index=False).agg(Actividades=("Actividad", "count"), Meta=("Meta", "sum"), Avance=("Avance", "sum"), Pendiente=("Pendiente", "sum"))
    res = res[(res["Meta"] > 0) | (res["Avance"] > 0)]
    res["% Cumplimiento"] = res.apply(lambda x: x["Avance"] / x["Meta"] if x["Meta"] else 0, axis=1)
    res["Estado"] = res["% Cumplimiento"].apply(clasificar_cumplimiento)
    return res.sort_values("% Cumplimiento", ascending=False).reset_index(drop=True)


def mostrar_metricas(df, catalogo):
    meta = df["Meta"].sum() if not df.empty else 0
    avance = df["Avance"].sum() if not df.empty else 0
    pendiente = max(meta - avance, 0)
    pct = avance / meta if meta else 0
    regiones = df["Dirección Regional"].nunique() if not df.empty else 0
    deleg_datos = df["Delegación"].nunique() if not df.empty else 0
    deleg_oficial = catalogo["Delegación"].nunique() if not catalogo.empty else deleg_datos
    tarjetas([
        ("Meta nacional", f"{meta:,.0f}"), ("Avance", f"{avance:,.0f}"), ("Pendiente", f"{pendiente:,.0f}"),
        ("% cumplimiento", f"{pct:.1%}"), ("Regiones", f"{regiones}"), ("Delegaciones", f"{deleg_datos} / {deleg_oficial}"),
    ])


def dataframe_visible(df):
    cols = ["Dirección Regional", "Delegación", "Programa", "Actividad", "Meta", "Avance", "Pendiente", "% Cumplimiento", "Estado"]
    out = df[[c for c in cols if c in df.columns]].copy()
    return out


def leyenda_cumplimiento():
    st.markdown(
        f"""
        <div class='status-legend'>
          <div class='status-pill' style='background:#E8F7EF;color:#065F46;'><span class='status-dot' style='background:{COLOR_VERDE};'></span>Cumple: 50% o más</div>
          <div class='status-pill' style='background:#FFF4DC;color:#92400E;'><span class='status-dot' style='background:{COLOR_NARANJA};'></span>En riesgo: 45% a 49.99%</div>
          <div class='status-pill' style='background:#FFE4E6;color:#991B1B;'><span class='status-dot' style='background:{COLOR_ROJO};'></span>Crítico: menor al 45%</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def estilo_tabla_cumplimiento(df):
    tabla = dataframe_visible(df).copy()

    # Guardamos el porcentaje real para pintar la fila y luego mostramos el valor
    # en formato institucional (ej. 45.1%) en lugar del decimal 0.451.
    pct_real = None
    if "% Cumplimiento" in tabla.columns:
        pct_real = pd.to_numeric(tabla["% Cumplimiento"], errors="coerce").fillna(0)

    def formato_numero_visible(v):
        try:
            return f"{float(v):,.0f}"
        except Exception:
            return str(v)

    if "% Cumplimiento" in tabla.columns:
        tabla["% Cumplimiento"] = pct_real.map(lambda x: f"{x:.1%}")

    for col in ["Meta", "Avance", "Pendiente", "Actividades", "Delegaciones"]:
        if col in tabla.columns:
            tabla[col] = tabla[col].map(formato_numero_visible)

    def obtener_color_estado(row):
        estado = str(row.get("Estado", ""))
        if estado == "Cumple":
            return "#ECFDF3", COLOR_VERDE, "#065F46"
        if estado == "En riesgo":
            return "#FFF7E6", COLOR_NARANJA, "#92400E"
        return "#FFF1F2", COLOR_ROJO, "#991B1B"

    def estilo_fila(row):
        bg, border, text = obtener_color_estado(row)
        estilos = [
            f"background-color:{bg}; color:#111827; border-bottom:1px solid #E5E7EB;"
            for _ in row
        ]
        if len(estilos) > 0:
            estilos[0] += f" border-left:8px solid {border};"
        return estilos

    def estilo_estado(v):
        estado = str(v)
        if estado == "Cumple":
            return f"background-color:{COLOR_VERDE}; color:white; font-weight:900; border-radius:14px; text-align:center;"
        if estado == "En riesgo":
            return f"background-color:{COLOR_NARANJA}; color:white; font-weight:900; border-radius:14px; text-align:center;"
        return f"background-color:{COLOR_ROJO}; color:white; font-weight:900; border-radius:14px; text-align:center;"

    def estilo_pct_visible(v):
        txt = str(v).replace('%','').strip()
        try:
            val = float(txt) / 100
        except Exception:
            val = 0
        color = COLOR_VERDE if val >= UMBRAL_CUMPLE else COLOR_NARANJA if val >= UMBRAL_RIESGO else COLOR_ROJO
        return f"font-weight:900; color:{color}; font-size:15px;"

    styler = tabla.style.apply(estilo_fila, axis=1)
    if "Estado" in tabla.columns:
        styler = styler.map(estilo_estado, subset=["Estado"])
    if "% Cumplimiento" in tabla.columns:
        styler = styler.map(estilo_pct_visible, subset=["% Cumplimiento"])

    styler = styler.set_table_styles([
        {"selector": "th", "props": [("background-color", COLOR_AZUL), ("color", "white"), ("font-weight", "900"), ("text-align", "center"), ("font-size", "14px")]},
        {"selector": "td", "props": [("font-size", "14px"), ("vertical-align", "middle"), ("padding", "8px 10px")]},
    ])
    return styler


def mostrar_tabla_cumplimiento(df, height=480):
    if df is None or df.empty:
        st.info("No hay datos para mostrar.")
        return
    leyenda_cumplimiento()
    st.dataframe(estilo_tabla_cumplimiento(df), use_container_width=True, hide_index=True, height=height)

# ======================================================
# GRÁFICOS
# ======================================================

def layout_fig(fig, height=520):
    fig.update_layout(
        title_x=0.5, paper_bgcolor="white", plot_bgcolor="white", height=height,
        font=dict(color="#374151", size=14), title_font=dict(color=COLOR_AZUL, size=23, family="Arial Black"),
        legend=dict(orientation="h", y=1.12, x=0.5, xanchor="center", bgcolor="rgba(255,255,255,.85)"),
        margin=dict(l=35, r=30, t=95, b=45),
        hoverlabel=dict(bgcolor="white", font_size=13, font_color="#111827")
    )
    fig.update_xaxes(gridcolor="#E6EAF2", zerolinecolor="#D7DCE5", title_font=dict(color=COLOR_AZUL, size=15))
    fig.update_yaxes(gridcolor="#E6EAF2", zerolinecolor="#D7DCE5", title_font=dict(color=COLOR_AZUL, size=15))
    return fig


def grafico_region(df):
    res = resumen_por_region(df)
    if res.empty:
        st.info("No hay datos para graficar.")
        return
    res = res.sort_values("% Cumplimiento", ascending=True)
    fig = px.bar(
        res,
        x="% Cumplimiento",
        y="Dirección Regional",
        orientation="h",
        color="Estado",
        color_discrete_map={"Cumple": COLOR_VERDE, "En riesgo": COLOR_NARANJA, "Crítico": COLOR_ROJO},
        text=res["% Cumplimiento"].map(lambda x: f"{x:.1%}"),
        title="Cumplimiento por Dirección Regional",
        hover_data={"Meta": ":,.0f", "Avance": ":,.0f", "Pendiente": ":,.0f", "% Cumplimiento": ":.1%"},
    )
    fig.add_vline(x=UMBRAL_CUMPLE, line_dash="dash", line_color=COLOR_AZUL, line_width=4)
    fig.add_annotation(x=UMBRAL_CUMPLE, y=1.03, xref="x", yref="paper", text="Referencia institucional 50%", showarrow=False, font=dict(color=COLOR_AZUL, size=14, family="Arial Black"))
    fig.update_traces(textposition="outside", marker_line_color="white", marker_line_width=1.6, cliponaxis=False)
    fig.update_xaxes(tickformat=".0%", range=[0, max(0.60, min(1.05, res["% Cumplimiento"].max() + 0.15))], title="% de cumplimiento")
    fig.update_yaxes(title="", automargin=True)
    st.plotly_chart(layout_fig(fig, max(520, 38 * len(res) + 160)), use_container_width=True)


def grafico_programa(df):
    res = resumen_por_programa(df)
    if res.empty:
        return
    res = res.sort_values("% Cumplimiento", ascending=True)
    fig = px.bar(
        res,
        x="% Cumplimiento",
        y="Programa",
        orientation="h",
        color="Estado",
        color_discrete_map={"Cumple": COLOR_VERDE, "En riesgo": COLOR_NARANJA, "Crítico": COLOR_ROJO},
        text=res["% Cumplimiento"].map(lambda x: f"{x:.1%}"),
        title="Cumplimiento por programa",
        hover_data={"Meta": ":,.0f", "Avance": ":,.0f", "Pendiente": ":,.0f", "% Cumplimiento": ":.1%"},
    )
    fig.add_vline(x=UMBRAL_CUMPLE, line_dash="dash", line_color=COLOR_AZUL, line_width=4)
    fig.add_annotation(x=UMBRAL_CUMPLE, y=1.04, xref="x", yref="paper", text="50%", showarrow=False, font=dict(color=COLOR_AZUL, size=14, family="Arial Black"))
    fig.update_traces(textposition="outside", marker_line_color="white", marker_line_width=1.6, cliponaxis=False)
    fig.update_xaxes(tickformat=".0%", range=[0, max(0.60, min(1.05, res["% Cumplimiento"].max() + 0.15))], title="% de cumplimiento")
    fig.update_yaxes(title="", automargin=True)
    st.plotly_chart(layout_fig(fig, 430), use_container_width=True)


def grafico_delegaciones_por_region(df, filtros):
    regiones = filtros.get("Dirección Regional", [])
    delegs = filtros.get("Delegación", [])
    if not regiones and not delegs:
        st.info("Seleccione una Dirección Regional o una delegación para visualizar el gráfico de delegaciones sin saturar la pantalla.")
        return

    res = resumen_por_delegacion(df)
    if res.empty:
        st.info("No hay delegaciones con datos para graficar.")
        return

    res = res.sort_values("% Cumplimiento", ascending=True)

    fig = px.bar(
        res,
        x="% Cumplimiento",
        y="Delegación",
        orientation="h",
        color="Estado",
        color_discrete_map={"Cumple": COLOR_VERDE, "En riesgo": COLOR_NARANJA, "Crítico": COLOR_ROJO},
        text=res["% Cumplimiento"].map(lambda x: f"{x:.1%}"),
        title="Cumplimiento por delegación",
        hover_data={"Dirección Regional": True, "Meta": ":,.0f", "Avance": ":,.0f", "Pendiente": ":,.0f", "% Cumplimiento": ":.1%"},
    )
    fig.add_vline(x=UMBRAL_CUMPLE, line_dash="dash", line_color=COLOR_AZUL, line_width=4)
    fig.add_annotation(x=UMBRAL_CUMPLE, y=1.035, xref="x", yref="paper", text="Límite 50%", showarrow=False, font=dict(color=COLOR_AZUL, size=14, family="Arial Black"))
    fig.update_traces(textposition="outside", marker_line_color="white", marker_line_width=1.6, cliponaxis=False)
    fig.update_xaxes(tickformat=".0%", range=[0, max(0.60, min(1.05, res["% Cumplimiento"].max() + 0.15))], title="% de cumplimiento")
    fig.update_yaxes(title="", automargin=True)
    st.plotly_chart(layout_fig(fig, max(480, 34 * len(res) + 150)), use_container_width=True)

    st.markdown("### Comparativo operativo de la selección")
    comparativo = res.sort_values("Pendiente", ascending=False).head(20) if len(res) > 20 else res.sort_values("Pendiente", ascending=False)
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(y=comparativo["Delegación"], x=comparativo["Meta"], name="Meta", orientation="h", marker=dict(color="#E8F0FF", line=dict(color=COLOR_AZUL, width=1.4))))
    fig2.add_trace(go.Bar(y=comparativo["Delegación"], x=comparativo["Avance"], name="Avance", orientation="h", marker_color=COLOR_AZUL))
    fig2.add_trace(go.Bar(y=comparativo["Delegación"], x=comparativo["Pendiente"], name="Pendiente", orientation="h", marker_color=COLOR_ROJO))
    fig2.update_layout(barmode="group", title="Meta, avance y pendiente por delegación", yaxis=dict(autorange="reversed"), xaxis_title="Cantidad")
    st.plotly_chart(layout_fig(fig2, max(480, 34 * len(comparativo) + 160)), use_container_width=True)

    top = res.sort_values("% Cumplimiento", ascending=False).head(2)
    bajo = res.sort_values("% Cumplimiento", ascending=True).head(2)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Top 2 mejor posicionadas")
        for i, (_, row) in enumerate(top.iterrows(), start=1):
            st.success(f"{i}. {row['Delegación']} - {row['% Cumplimiento']:.1%} | Meta {row['Meta']:,.0f} | Avance {row['Avance']:,.0f}")
    with c2:
        st.markdown("### Top 2 con menor cumplimiento")
        for i, (_, row) in enumerate(bajo.iterrows(), start=1):
            st.error(f"{i}. {row['Delegación']} - {row['% Cumplimiento']:.1%} | Pendiente {row['Pendiente']:,.0f}")


def grafico_brecha_pendiente(df, filtros):
    """Sustituye el pareto: muestra las brechas de cumplimiento más relevantes sin línea acumulada."""
    res = resumen_por_delegacion(df)
    if res.empty:
        return
    if not filtros.get("Dirección Regional") and not filtros.get("Delegación"):
        res = res.sort_values("Pendiente", ascending=False).head(12)
        titulo = "Delegaciones con mayor pendiente nacional"
    else:
        res = res.sort_values("Pendiente", ascending=False).head(15)
        titulo = "Pendiente por delegación filtrada"
    fig = px.bar(
        res.sort_values("Pendiente", ascending=True),
        x="Pendiente",
        y="Delegación",
        orientation="h",
        color="Estado",
        color_discrete_map={"Cumple": COLOR_VERDE, "En riesgo": COLOR_NARANJA, "Crítico": COLOR_ROJO},
        text=res.sort_values("Pendiente", ascending=True)["Pendiente"].map(lambda x: f"{x:,.0f}"),
        title=titulo,
        hover_data={"Dirección Regional": True, "Meta": ":,.0f", "Avance": ":,.0f", "% Cumplimiento": ":.1%"},
    )
    fig.update_traces(textposition="outside", marker_line_color="white", marker_line_width=1.6, cliponaxis=False)
    fig.update_xaxes(title="Pendiente")
    fig.update_yaxes(title="", automargin=True)
    st.plotly_chart(layout_fig(fig, max(480, 34 * len(res) + 160)), use_container_width=True)

# ======================================================
# MAPA DE MARCADORES
# ======================================================

def crear_mapa_marcadores(df):
    res = resumen_por_delegacion(df)
    mapa = folium.Map(location=[9.7489, -83.7534], zoom_start=7, tiles="CartoDB positron")
    if res.empty:
        return mapa
    for _, row in res.iterrows():
        estado = row["Estado"]
        color = "green" if estado == "Cumple" else "orange" if estado == "En riesgo" else "red"
        icono = "ok" if estado == "Cumple" else "warning-sign" if estado == "En riesgo" else "remove"
        popup = f"""
        <div style='font-family:Arial;width:310px;'>
          <h4 style='color:#002B7F;margin:0 0 8px 0;'>{row['Delegación']}</h4>
          <b>Dirección Regional:</b> {row['Dirección Regional']}<br>
          <b>Meta:</b> {row['Meta']:,.0f}<br>
          <b>Avance:</b> {row['Avance']:,.0f}<br>
          <b>Pendiente:</b> {row['Pendiente']:,.0f}<br>
          <b>Cumplimiento:</b> {row['% Cumplimiento']:.1%}<br>
          <b>Estado:</b> <span style='color:{color};font-weight:bold'>{estado}</span><br>
          <b>Actividades:</b> {row['Actividades']:,.0f}
        </div>
        """
        folium.Marker(
            location=[float(row["_lat"]), float(row["_lon"])],
            tooltip=f"{row['Delegación']} | {row['% Cumplimiento']:.1%} | {estado}",
            popup=folium.Popup(popup, max_width=340),
            icon=folium.Icon(color=color, icon=icono)
        ).add_to(mapa)
    try:
        bounds = [[res["_lat"].min(), res["_lon"].min()], [res["_lat"].max(), res["_lon"].max()]]
        mapa.fit_bounds(bounds)
    except Exception:
        pass
    return mapa

# ======================================================
# PDF
# ======================================================

def pdf_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TituloPUMI", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=22, alignment=TA_CENTER, textColor=colors.HexColor(COLOR_AZUL), spaceAfter=14))
    styles.add(ParagraphStyle(name="SubtituloPUMI", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=14, alignment=TA_LEFT, textColor=colors.HexColor(COLOR_AZUL), spaceBefore=8, spaceAfter=6))
    styles.add(ParagraphStyle(name="TextoPUMI", parent=styles["BodyText"], fontSize=9, leading=12))
    styles.add(ParagraphStyle(name="TablaPUMI", parent=styles["BodyText"], fontSize=7, leading=9))
    return styles


def p(texto, style):
    texto = "" if texto is None else str(texto)
    texto = texto.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return Paragraph(texto, style)


def logo_pdf(ruta, w=5*cm, h=1.8*cm):
    if os.path.exists(ruta):
        try: return Image(ruta, width=w, height=h)
        except Exception: return ""
    return ""


def generar_pdf(df, filtros):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=1.1*cm, leftMargin=1.1*cm, topMargin=1.2*cm, bottomMargin=1.1*cm)
    styles = pdf_styles()
    elementos = []
    logos = [[logo_pdf(LOGO_MINISTERIO), logo_pdf(LOGO_PUMI, 4*cm, 2*cm), logo_pdf(LOGO_FUERZA_PUBLICA)]]
    tabla_logos = Table(logos, colWidths=[8*cm, 8*cm, 8*cm])
    tabla_logos.setStyle(TableStyle([("ALIGN", (0,0), (-1,-1), "CENTER"), ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("BOX", (0,0), (-1,-1), 0.8, colors.HexColor(COLOR_AZUL)), ("TOPPADDING", (0,0), (-1,-1), 8), ("BOTTOMPADDING", (0,0), (-1,-1), 8)]))
    elementos += [tabla_logos, Spacer(1, .6*cm), p("INFORME NACIONAL DE CUMPLIMIENTO PUMI 2026", styles["TituloPUMI"]), Spacer(1, .3*cm)]
    meta = df["Meta"].sum() if not df.empty else 0
    avance = df["Avance"].sum() if not df.empty else 0
    pendiente = max(meta - avance, 0)
    pct = avance / meta if meta else 0
    datos = [["Indicador", "Resultado"], ["Meta nacional", f"{meta:,.0f}"], ["Avance", f"{avance:,.0f}"], ["Pendiente", f"{pendiente:,.0f}"], ["Cumplimiento", f"{pct:.1%}"], ["Fecha", date.today().strftime("%d/%m/%Y")]]
    tabla = Table(datos, colWidths=[7*cm, 10*cm])
    tabla.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor(COLOR_AZUL)), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), .4, colors.HexColor("#BFBFBF")), ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F4F7FB")])]))
    elementos += [tabla, Spacer(1, .5*cm), p("Filtros aplicados", styles["SubtituloPUMI"])]
    datos_f = [["Filtro", "Valor"]]
    for k, v in filtros.items(): datos_f.append([k, ", ".join(v) if v else "Todos"])
    tabla_f = Table(datos_f, colWidths=[6*cm, 17*cm])
    tabla_f.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor(COLOR_AZUL)), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), .35, colors.HexColor("#BFBFBF"))]))
    elementos.append(tabla_f)
    elementos.append(PageBreak())

    reg = resumen_por_region(df)
    elementos.append(p("Resumen por Dirección Regional", styles["SubtituloPUMI"]))
    datos_reg = [["Dirección Regional", "Delegaciones", "Meta", "Avance", "Pendiente", "%", "Estado"]]
    for _, r in reg.iterrows(): datos_reg.append([r["Dirección Regional"], r["Delegaciones"], f"{r['Meta']:,.0f}", f"{r['Avance']:,.0f}", f"{r['Pendiente']:,.0f}", f"{r['% Cumplimiento']:.1%}", r["Estado"]])
    tabla_reg = Table(datos_reg, colWidths=[6*cm, 2.3*cm, 2.4*cm, 2.4*cm, 2.4*cm, 2.0*cm, 2.4*cm], repeatRows=1)
    tabla_reg.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor(COLOR_AZUL)), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), .25, colors.HexColor("#BFBFBF")), ("FONTSIZE", (0,0), (-1,-1), 7), ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F4F7FB")])]))
    # Colores por estado
    for i, row in enumerate(datos_reg[1:], start=1):
        if row[-1] == "Cumple": bg = colors.HexColor("#DFF5E8")
        elif row[-1] == "En riesgo": bg = colors.HexColor("#FFF3D6")
        else: bg = colors.HexColor("#FFE2E2")
        tabla_reg.setStyle(TableStyle([("BACKGROUND", (0,i), (-1,i), bg)]))
    elementos.append(tabla_reg)
    elementos.append(PageBreak())

    deleg = resumen_por_delegacion(df)
    elementos.append(p("Detalle por delegación", styles["SubtituloPUMI"]))
    datos_del = [["Región", "Delegación", "Meta", "Avance", "Pendiente", "%", "Estado"]]
    for _, r in deleg.iterrows(): datos_del.append([r["Dirección Regional"], r["Delegación"], f"{r['Meta']:,.0f}", f"{r['Avance']:,.0f}", f"{r['Pendiente']:,.0f}", f"{r['% Cumplimiento']:.1%}", r["Estado"]])
    tabla_del = Table(datos_del, colWidths=[5.2*cm, 4.6*cm, 2.2*cm, 2.2*cm, 2.2*cm, 2.0*cm, 2.4*cm], repeatRows=1)
    tabla_del.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor(COLOR_AZUL)), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), .25, colors.HexColor("#BFBFBF")), ("FONTSIZE", (0,0), (-1,-1), 7)]))
    for i, row in enumerate(datos_del[1:], start=1):
        if row[-1] == "Cumple": bg = colors.HexColor("#DFF5E8")
        elif row[-1] == "En riesgo": bg = colors.HexColor("#FFF3D6")
        else: bg = colors.HexColor("#FFE2E2")
        tabla_del.setStyle(TableStyle([("BACKGROUND", (0,i), (-1,i), bg)]))
    elementos.append(tabla_del)
    doc.build(elementos)
    buffer.seek(0)
    return buffer.getvalue()

# ======================================================
# PÁGINAS
# ======================================================

def pagina_inicio(df, catalogo):
    st.markdown("""
    <div class='card-info'>
      <div class='card-title'>Dashboard Nacional PUMI 2026</div>
      <div class='card-text'>
        Esta aplicación consolida automáticamente los Excel regionales ubicados al mismo nivel del archivo app.py.
        Su objetivo es visualizar cumplimiento nacional, detectar regiones y delegaciones críticas, aplicar filtros,
        consultar mapas de marcadores e informes PDF según la información filtrada.
      </div>
    </div>
    """, unsafe_allow_html=True)
    mostrar_metricas(df, catalogo)


def pagina_dashboard(df, catalogo):
    filtrado, filtros = aplicar_filtros(df, key="dash")
    mostrar_metricas(filtrado, catalogo)
    st.markdown("---")
    grafico_region(filtrado)
    grafico_programa(filtrado)
    st.markdown("---")
    st.markdown("## Gráfico por delegación")
    grafico_delegaciones_por_region(filtrado, filtros)
    st.markdown("---")
    grafico_brecha_pendiente(filtrado, filtros)
    st.markdown("---")
    st.markdown("## Registros filtrados")
    mostrar_tabla_cumplimiento(filtrado, height=480)


def pagina_mapa(df, catalogo):
    filtrado, filtros = aplicar_filtros(df, key="mapa")
    mostrar_metricas(filtrado, catalogo)
    st.markdown("## Mapa nacional de cumplimiento")
    st.info("El mapa utiliza marcadores por delegación según cumplimiento. Al hacer clic en una marca se muestra meta, avance, pendiente, porcentaje y estado. Las coordenadas se usan únicamente de forma interna.")
    mapa = crear_mapa_marcadores(filtrado)
    st_folium(mapa, height=650, use_container_width=True, key="mapa_marcadores_nacional")
    st.markdown("### Detalle del mapa")
    mostrar_tabla_cumplimiento(resumen_por_delegacion(filtrado), height=420)


def pagina_analisis(df, catalogo):
    filtrado, filtros = aplicar_filtros(df, key="analisis")
    mostrar_metricas(filtrado, catalogo)
    deleg = resumen_por_delegacion(filtrado)
    reg = resumen_por_region(filtrado)
    prog = resumen_por_programa(filtrado)
    st.markdown("## Análisis estratégico automático")
    if deleg.empty:
        st.info("No hay datos para analizar.")
        return
    # Separación clara para no mezclar estados en las tablas ejecutivas.
    # Mejores: solo delegaciones que cumplen o están en riesgo.
    # Críticas: únicamente delegaciones con cumplimiento menor al 45%.
    deleg_no_criticas = deleg[deleg["Estado"].isin(["Cumple", "En riesgo"])].copy()
    deleg_criticas = deleg[deleg["Estado"] == "Crítico"].copy()

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Delegaciones con cumplimiento aceptable o cercano")
        if deleg_no_criticas.empty:
            st.info("No hay delegaciones en estado Cumple o En riesgo con los filtros aplicados.")
        else:
            mostrar_tabla_cumplimiento(
                deleg_no_criticas.sort_values("% Cumplimiento", ascending=False).head(10),
                height=360
            )
    with c2:
        st.markdown("### Delegaciones críticas")
        if deleg_criticas.empty:
            st.success("No hay delegaciones críticas con los filtros aplicados.")
        else:
            mostrar_tabla_cumplimiento(
                deleg_criticas.sort_values("% Cumplimiento", ascending=True).head(10),
                height=360
            )
    st.markdown("### Ranking por región")
    mostrar_tabla_cumplimiento(reg, height=420)
    st.markdown("### Ranking por programa")
    mostrar_tabla_cumplimiento(prog, height=320)


def pagina_pdf(df, catalogo):
    filtrado, filtros = aplicar_filtros(df, key="pdf")
    mostrar_metricas(filtrado, catalogo)
    st.markdown("## Generar informe PDF")
    st.info("El informe se genera con los datos que quedan después de aplicar los filtros.")
    if st.button("📄 Generar PDF nacional filtrado"):
        pdf = generar_pdf(filtrado, filtros)
        st.success("PDF generado correctamente.")
        st.download_button("⬇️ Descargar informe PDF", data=pdf, file_name=f"INFORME_NACIONAL_PUMI_{date.today().strftime('%Y%m%d')}.pdf", mime="application/pdf")

# ======================================================
# APP PRINCIPAL
# ======================================================

mostrar_logo_sidebar()
st.sidebar.markdown("## Dashboard Nacional")
menu = st.sidebar.radio("Menú", ["Inicio", "Dashboard Nacional", "Análisis estratégico", "Mapa nacional", "Informe PDF"])

mostrar_encabezado()
mostrar_titulo()

df_metas = cargar_metas_regionales()
catalogo = cargar_catalogo_delegaciones()

if df_metas.empty:
    st.warning("No se encontraron Excel regionales de metas al mismo nivel de la app.")
    st.info("Verifique que los archivos estén en el repositorio con nombres como DR2 AVANCE JUNIO.xlsx.")
else:
    # Nota interna de control visible solo como advertencia general, sin coordenadas.
    deleg_datos = df_metas["Delegación"].nunique()
    deleg_oficial = catalogo["Delegación"].nunique() if not catalogo.empty else deleg_datos
    if deleg_oficial and deleg_datos < deleg_oficial:
        st.sidebar.warning(f"Delegaciones con datos: {deleg_datos} de {deleg_oficial}. Revise si alguna hoja regional no trae metas o no fue leída.")

if menu == "Inicio":
    pagina_inicio(df_metas, catalogo)
elif menu == "Dashboard Nacional":
    pagina_dashboard(df_metas, catalogo)
elif menu == "Análisis estratégico":
    pagina_analisis(df_metas, catalogo)
elif menu == "Mapa nacional":
    pagina_mapa(df_metas, catalogo)
elif menu == "Informe PDF":
    pagina_pdf(df_metas, catalogo)
