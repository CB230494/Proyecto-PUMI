# ======================================================
# PUMI 2026 - DASHBOARD NACIONAL
# Visualización nacional de metas, avances, cumplimiento,
# mapa de calor y generación de informes PDF
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
from folium.plugins import HeatMap
from streamlit_folium import st_folium

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image

# ======================================================
# CONFIGURACIÓN GENERAL
# ======================================================

st.set_page_config(
    page_title="PUMI 2026 - Dashboard Nacional",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

LOGO_MINISTERIO = "Logo2.jpeg"
LOGO_PUMI = "logo_pumi.jpeg"
LOGO_FUERZA_PUBLICA = "Logo1.jpeg"

CARPETA_BASE = "."
CARPETA_METAS_ALTERNATIVA = "metas"

MESES = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
]

COLOR_AZUL = "#002B7F"
COLOR_DORADO = "#B88A2A"
COLOR_ROJO = "#B42318"
COLOR_NARANJA = "#F59E0B"
COLOR_VERDE = "#1F8A4C"
COLOR_GRIS_OSCURO = "#2F3542"
COLOR_BLANCO = "#FFFFFF"

UMBRAL_CUMPLE = 0.50
UMBRAL_RIESGO = 0.45

# Coordenadas aproximadas por cabecera/delegación cuando no hay coordenadas exactas.
# Se puede ampliar sin tocar el resto de la app.
COORDENADAS_DELEGACIONES = {
    "CURRIDABAT": [9.9147, -84.0348],
    "PAVAS": [9.9490, -84.1310],
    "HATILLO": [9.9145, -84.1000],
    "DESAMPARADOS": [9.8969, -84.0620],
    "ESCAZU": [9.9186, -84.1399],
    "SANTA ANA": [9.9326, -84.1825],
    "MORA": [9.9167, -84.2333],
    "PURISCAL": [9.8469, -84.3143],
    "TURRUBARES": [9.9089, -84.4711],
    "ACOSTA": [9.8000, -84.1667],
    "ALAJUELA": [10.0162, -84.2116],
    "SAN RAMON": [10.0880, -84.4702],
    "GRECIA": [10.0739, -84.3112],
    "SAN MATEO": [9.9365, -84.5247],
    "ATENAS": [9.9787, -84.3801],
    "NARANJO": [10.0987, -84.3782],
    "PALMARES": [10.0567, -84.4370],
    "POAS": [10.0800, -84.2450],
    "OROTINA": [9.9111, -84.5230],
    "SAN CARLOS": [10.3290, -84.4310],
    "ZARCERO": [10.1852, -84.3900],
    "SARCHI": [10.0883, -84.3473],
    "UPALA": [10.8986, -85.0155],
    "LOS CHILES": [11.0350, -84.7130],
    "GUATUSO": [10.6667, -84.8167],
    "RIO CUARTO": [10.3410, -84.2140],
    "CARTAGO": [9.8644, -83.9194],
    "PARAISO": [9.8383, -83.8656],
    "LA UNION": [9.9084, -83.9886],
    "JIMENEZ": [9.9048, -83.6834],
    "TURRIALBA": [9.9050, -83.6830],
    "ALVARADO": [9.9333, -83.8000],
    "OREAMUNO": [9.9100, -83.9000],
    "EL GUARCO": [9.8472, -83.9460],
    "HEREDIA": [10.0024, -84.1165],
    "BARVA": [10.0208, -84.1233],
    "SANTO DOMINGO": [10.0639, -84.1547],
    "SANTA BARBARA": [10.0400, -84.1600],
    "SAN RAFAEL": [10.0138, -84.1002],
    "SAN ISIDRO": [10.0186, -84.0569],
    "BELEN": [9.9852, -84.1810],
    "FLORES": [10.0000, -84.1600],
    "SAN PABLO": [9.9953, -84.0966],
    "SARAPIQUI": [10.4522, -84.0166],
    "LIBERIA": [10.6350, -85.4377],
    "NICOYA": [10.1483, -85.4520],
    "SANTA CRUZ": [10.2600, -85.5850],
    "BAGACES": [10.5250, -85.2550],
    "CARRILLO": [10.4750, -85.5850],
    "CANAS": [10.4310, -85.0980],
    "ABANGARES": [10.2820, -84.9590],
    "TILARAN": [10.4670, -84.9670],
    "NANDAYURE": [9.9990, -85.2060],
    "LA CRUZ": [11.0730, -85.6320],
    "HOJANCHA": [10.0550, -85.4200],
    "PUNTARENAS": [9.9763, -84.8384],
    "ESPARZA": [9.9940, -84.6640],
    "BUENOS AIRES": [9.1667, -83.3333],
    "MONTES DE ORO": [10.0870, -84.7300],
    "OSA": [8.9590, -83.5230],
    "QUEPOS": [9.4319, -84.1617],
    "GOLFITO": [8.6390, -83.1660],
    "COTO BRUS": [8.8830, -82.9660],
    "PARRITA": [9.5200, -84.3200],
    "CORREDORES": [8.6420, -82.9460],
    "GARABITO": [9.6150, -84.6300],
    "LIMON": [9.9917, -83.0360],
    "POCOCI": [10.2150, -83.7870],
    "SIQUIRRES": [10.0970, -83.5060],
    "TALAMANCA": [9.6240, -82.8440],
    "MATINA": [10.0760, -83.2890],
    "GUACIMO": [10.2100, -83.6900],
    "PEREZ ZELEDON": [9.3540, -83.6340],
    "CORONADO": [9.9760, -84.0070],
    "MORAVIA": [9.9610, -84.0480],
    "TIBAS": [9.9580, -84.0790],
    "GOICOECHEA": [9.9480, -84.0430],
}

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
    "DR12": "Dirección Regional 12",
}

# ======================================================
# UTILIDADES
# ======================================================

def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^A-Z0-9 ]", " ", texto)
    return " ".join(texto.split())


def limpiar_nombre_archivo(valor):
    texto = normalizar_texto(valor)
    texto = re.sub(r"\s+", "_", texto)
    return texto if texto else "SIN_DATO"


def imagen_a_base64(ruta):
    if not os.path.exists(ruta):
        return ""
    try:
        with open(ruta, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""


def limpiar_numero(valor):
    if pd.isna(valor) or valor == "":
        return 0.0
    if isinstance(valor, str):
        valor = valor.replace("%", "").replace(",", ".").strip()
    try:
        return float(valor)
    except Exception:
        return 0.0


def extraer_numero_region(valor):
    texto = normalizar_texto(valor)
    m = re.search(r"REGIONAL\s+(\d+)", texto)
    if m:
        return int(m.group(1))
    m = re.search(r"\bDR\s*(\d+)", texto)
    if m:
        return int(m.group(1))
    m = re.search(r"\bR\s*(\d+)", texto)
    if m:
        return int(m.group(1))
    return 999


def orden_region(valor):
    texto = normalizar_texto(valor)
    numero = extraer_numero_region(valor)
    sub = 0
    if numero == 1:
        if "CENTRAL" in texto:
            sub = 0
        elif "NORTE" in texto:
            sub = 1
        elif "SUR" in texto:
            sub = 2
        else:
            sub = 9
    return (numero, sub, texto)


def orden_delegacion(valor):
    texto = normalizar_texto(valor)
    m = re.search(r"\bD\s*(\d+)", texto)
    num = int(m.group(1)) if m else 9999
    return (num, texto)


def obtener_region_desde_archivo(nombre_archivo):
    nombre = os.path.basename(nombre_archivo).upper().replace("_", " ")
    nombre = re.sub(r"\s+", " ", nombre).strip()

    for clave in sorted(MAPA_REGION_ARCHIVO.keys(), key=len, reverse=True):
        if nombre.startswith(clave):
            return MAPA_REGION_ARCHIVO[clave]

    m = re.match(r"(DR\d+)", nombre)
    if m:
        return MAPA_REGION_ARCHIVO.get(m.group(1), m.group(1))

    return "Región no identificada"


def buscar_archivos_excel():
    patrones = [
        os.path.join(CARPETA_BASE, "DR*.xlsx"),
        os.path.join(CARPETA_BASE, "DR*.xlsm"),
        os.path.join(CARPETA_BASE, "DR*.xls"),
    ]
    if os.path.isdir(CARPETA_METAS_ALTERNATIVA):
        patrones += [
            os.path.join(CARPETA_METAS_ALTERNATIVA, "DR*.xlsx"),
            os.path.join(CARPETA_METAS_ALTERNATIVA, "DR*.xlsm"),
            os.path.join(CARPETA_METAS_ALTERNATIVA, "DR*.xls"),
        ]

    rutas = []
    for patron in patrones:
        rutas.extend(glob.glob(patron))

    limpias = []
    for r in rutas:
        nombre = os.path.basename(r).upper()
        if nombre.startswith("~$"):
            continue
        if "REGISTRO_PUMI" in nombre or "VERIFICACION" in nombre or "INFORME" in nombre:
            continue
        if r not in limpias:
            limpias.append(r)

    return sorted(limpias, key=lambda x: orden_region(os.path.basename(x)))


def clasificar_cumplimiento(porc):
    if porc >= UMBRAL_CUMPLE:
        return "Cumple"
    if porc >= UMBRAL_RIESGO:
        return "En riesgo"
    return "Crítico"


def color_estado(estado):
    return {
        "Cumple": COLOR_VERDE,
        "En riesgo": COLOR_NARANJA,
        "Crítico": COLOR_ROJO,
    }.get(estado, COLOR_GRIS_OSCURO)


def normalizar_columnas(cols):
    return [re.sub(r"\s+", " ", str(c).strip().upper()) for c in cols]


def ubicar_delegacion(delegacion):
    texto = normalizar_texto(delegacion)

    if texto in COORDENADAS_DELEGACIONES:
        return COORDENADAS_DELEGACIONES[texto]

    # Intenta coincidencia parcial.
    for clave, coord in COORDENADAS_DELEGACIONES.items():
        if clave in texto or texto in clave:
            return coord

    # Si no existe, ubica en centro nacional para no romper mapa.
    return [9.7489, -83.7534]

# ======================================================
# CARGA AUTOMÁTICA DE DATOS
# ======================================================

@st.cache_data(show_spinner=False)
def cargar_datos_nacionales():
    archivos = buscar_archivos_excel()
    registros = []

    for ruta in archivos:
        nombre_archivo = os.path.basename(ruta)
        region = obtener_region_desde_archivo(nombre_archivo)

        try:
            xl = pd.ExcelFile(ruta)
        except Exception:
            continue

        hojas = sorted(xl.sheet_names, key=orden_delegacion)

        for hoja in hojas:
            nombre_hoja = str(hoja).strip()
            if not nombre_hoja or nombre_hoja.upper().startswith("TOTAL"):
                continue

            try:
                raw = pd.read_excel(ruta, sheet_name=hoja, header=None)
            except Exception:
                continue

            if raw.empty:
                continue

            fila_header = None
            for i in range(min(15, len(raw))):
                fila_txt = " ".join(raw.iloc[i].astype(str).fillna("").str.upper().tolist())
                if "PROGRAMA" in fila_txt and "META" in fila_txt:
                    fila_header = i
                    break

            if fila_header is None:
                # Estructura usada en algunas versiones: primeras dos filas visuales.
                if raw.shape[1] >= 4:
                    df = raw.iloc[2:].copy()
                    df = df.rename(columns={0: "PROGRAMA", 1: "ACTIVIDAD", 2: "META", 3: "AVANCE"})
                else:
                    continue
            else:
                try:
                    df = pd.read_excel(ruta, sheet_name=hoja, header=fila_header)
                    df.columns = normalizar_columnas(df.columns)
                except Exception:
                    continue

            if df.empty or len(df.columns) < 4:
                continue

            columnas = list(df.columns)
            col_programa = columnas[0]
            col_actividad = columnas[1]
            col_meta = next((c for c in columnas if "META" in str(c).upper()), columnas[2])
            col_avance = next((c for c in columnas if "AVANCE" in str(c).upper()), columnas[3])

            df[col_programa] = df[col_programa].ffill()
            df = df.dropna(how="all")

            for _, row in df.iterrows():
                programa = str(row.get(col_programa, "")).strip()
                actividad = str(row.get(col_actividad, "")).strip()

                if not programa or programa.upper() == "NAN":
                    continue
                if not actividad or actividad.upper() in ["NAN", "TOTAL", "TOTALES"]:
                    continue

                meta = limpiar_numero(row.get(col_meta, 0))
                avance = limpiar_numero(row.get(col_avance, 0))

                meses_total = 0
                meses = {}
                for mes in MESES:
                    col_mes = next((c for c in columnas if str(c).upper().strip() == mes), None)
                    valor_mes = limpiar_numero(row.get(col_mes, 0)) if col_mes else 0
                    meses[mes.title()] = valor_mes
                    meses_total += valor_mes

                if meta == 0 and avance == 0 and meses_total == 0:
                    continue

                pendiente = max(meta - avance, 0)
                cumplimiento = avance / meta if meta else 0
                estado = clasificar_cumplimiento(cumplimiento)
                lat, lon = ubicar_delegacion(nombre_hoja)

                item = {
                    "Archivo": nombre_archivo,
                    "Dirección Regional": region,
                    "Delegación": nombre_hoja,
                    "Programa": programa,
                    "Actividad": actividad,
                    "Meta": meta,
                    "Avance": avance,
                    "Pendiente": pendiente,
                    "% Cumplimiento": cumplimiento,
                    "Estado": estado,
                    "Latitud": lat,
                    "Longitud": lon,
                    "Clave Región": normalizar_texto(region),
                    "Clave Delegación": normalizar_texto(nombre_hoja),
                    "Clave Programa": normalizar_texto(programa),
                }
                item.update(meses)
                registros.append(item)

    if not registros:
        return pd.DataFrame(columns=[
            "Archivo", "Dirección Regional", "Delegación", "Programa", "Actividad",
            "Meta", "Avance", "Pendiente", "% Cumplimiento", "Estado",
            "Latitud", "Longitud"
        ])

    df = pd.DataFrame(registros)
    df = df.drop_duplicates(subset=["Archivo", "Delegación", "Programa", "Actividad"], keep="first")
    df = df.sort_values(
        by=["Dirección Regional", "Delegación", "Programa", "Actividad"],
        key=lambda s: s.map(normalizar_texto)
    ).reset_index(drop=True)
    return df

# ======================================================
# ESTILOS
# ======================================================

st.markdown(f"""
<style>
.stApp {{
    background: radial-gradient(circle at top left, #ffffff 0, transparent 28%),
                radial-gradient(circle at top right, #dce8ff 0, transparent 30%),
                linear-gradient(180deg, #ffffff 0%, #f4f7fb 50%, #ffffff 100%);
}}
section[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, #ffffff 0%, #eaf1ff 45%, #ffffff 100%);
    border-right: 5px solid {COLOR_DORADO};
}}
.encabezado-institucional {{
    background:#fff; border-radius:22px; padding:22px 34px; margin-bottom:24px;
    box-shadow:0 8px 22px rgba(0,0,0,.11); border-bottom:7px solid {COLOR_DORADO};
    display:grid; grid-template-columns:1fr 1fr 1fr; align-items:center; gap:22px;
}}
.encabezado-institucional img {{ max-height:120px; max-width:100%; object-fit:contain; }}
.titulo-principal {{
    background:linear-gradient(135deg,{COLOR_AZUL},#243B63,{COLOR_DORADO});
    padding:34px; border-radius:22px; text-align:center; color:white; margin-bottom:30px;
    box-shadow:0 8px 22px rgba(0,0,0,.28); border:3px solid white;
}}
.titulo-principal h1 {{font-size:46px;margin:0 0 8px 0;color:white;font-weight:900;}}
.titulo-principal h3 {{font-size:24px;margin:0;color:white;font-weight:600;}}
.card-info {{
    background:#fff; border-radius:20px; padding:24px; margin-bottom:20px;
    box-shadow:0 6px 16px rgba(0,0,0,.13); border-left:9px solid {COLOR_AZUL};
}}
.subtitulo {{ color:{COLOR_AZUL}; font-weight:900; font-size:28px; text-align:center; margin-bottom:10px; }}
.texto {{ color:{COLOR_GRIS_OSCURO}; font-size:17px; line-height:1.6; text-align:center; }}
div[data-testid="stMetric"] {{
    background:linear-gradient(145deg,#fff 0%,#f3f6ff 100%); padding:18px; border-radius:18px;
    box-shadow:0 5px 15px rgba(0,0,0,.13); border-bottom:6px solid {COLOR_DORADO};
}}
.stButton > button, .stDownloadButton > button {{
    background:linear-gradient(90deg,{COLOR_AZUL},{COLOR_DORADO}); color:white; border-radius:12px;
    border:none; font-weight:800; padding:.7rem 1.2rem;
}}
div[data-testid="stDataFrame"] {{ border-radius:18px; overflow:hidden; box-shadow:0 5px 16px rgba(0,0,0,.13); }}
h1,h2,h3 {{ color:{COLOR_AZUL}; font-weight:900; }}
</style>
""", unsafe_allow_html=True)

# ======================================================
# UI BASE
# ======================================================

def mostrar_logo_sidebar():
    logo = imagen_a_base64(LOGO_PUMI)
    if logo:
        st.sidebar.markdown(
            f"<div style='background:white;border-radius:16px;padding:12px;margin:8px 0 18px;text-align:center;box-shadow:0 4px 14px rgba(0,0,0,.10);'><img src='data:image/jpeg;base64,{logo}' style='width:100%;max-height:145px;object-fit:contain;'></div>",
            unsafe_allow_html=True
        )


def mostrar_encabezado():
    logos = []
    for ruta in [LOGO_MINISTERIO, LOGO_PUMI, LOGO_FUERZA_PUBLICA]:
        b64 = imagen_a_base64(ruta)
        logos.append(f"<img src='data:image/jpeg;base64,{b64}'>" if b64 else "")
    st.markdown(
        f"""
        <div class='encabezado-institucional'>
            <div style='text-align:left'>{logos[0]}</div>
            <div style='text-align:center'>{logos[1]}</div>
            <div style='text-align:right'>{logos[2]}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def mostrar_titulo():
    st.markdown(
        """
        <div class='titulo-principal'>
            <h1>📊 P.U.M.I. 2026 - Dashboard Nacional</h1>
            <h3>Visualización integral de metas, avances, cumplimiento y zonas críticas</h3>
        </div>
        """,
        unsafe_allow_html=True
    )

# ======================================================
# FILTROS Y MÉTRICAS
# ======================================================

def aplicar_filtros(df, key="general"):
    if df.empty:
        return df

    st.markdown("### Filtros")
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        regiones = sorted(df["Dirección Regional"].dropna().unique().tolist(), key=orden_region)
        filtro_region = st.multiselect("Dirección Regional", regiones, key=f"{key}_region")

    base_delegaciones = df[df["Dirección Regional"].isin(filtro_region)] if filtro_region else df
    with c2:
        delegaciones = sorted(base_delegaciones["Delegación"].dropna().unique().tolist(), key=orden_delegacion)
        filtro_delegacion = st.multiselect("Delegación", delegaciones, key=f"{key}_delegacion")

    with c3:
        programas = sorted(df["Programa"].dropna().unique().tolist(), key=normalizar_texto)
        filtro_programa = st.multiselect("Programa", programas, key=f"{key}_programa")

    with c4:
        filtro_estado = st.multiselect("Cumplimiento", ["Cumple", "En riesgo", "Crítico"], key=f"{key}_estado")

    filtrado = df.copy()
    if filtro_region:
        filtrado = filtrado[filtrado["Dirección Regional"].isin(filtro_region)]
    if filtro_delegacion:
        filtrado = filtrado[filtrado["Delegación"].isin(filtro_delegacion)]
    if filtro_programa:
        filtrado = filtrado[filtrado["Programa"].isin(filtro_programa)]
    if filtro_estado:
        filtrado = filtrado[filtrado["Estado"].isin(filtro_estado)]

    st.session_state["filtros_pdf"] = {
        "Dirección Regional": filtro_region,
        "Delegación": filtro_delegacion,
        "Programa": filtro_programa,
        "Cumplimiento": filtro_estado,
    }
    return filtrado


def mostrar_metricas(df):
    meta = df["Meta"].sum() if not df.empty else 0
    avance = df["Avance"].sum() if not df.empty else 0
    pendiente = max(meta - avance, 0)
    cumplimiento = avance / meta if meta else 0
    regiones = df["Dirección Regional"].nunique() if not df.empty else 0
    delegaciones = df["Delegación"].nunique() if not df.empty else 0

    criticas = 0
    if not df.empty:
        por_deleg = resumen_por_delegacion(df)
        criticas = len(por_deleg[por_deleg["Estado"] == "Crítico"])

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Meta nacional", f"{meta:,.0f}")
    c2.metric("Avance", f"{avance:,.0f}")
    c3.metric("Pendiente", f"{pendiente:,.0f}")
    c4.metric("% cumplimiento", f"{cumplimiento:.1%}")
    c5.metric("Regiones", f"{regiones}")
    c6.metric("Delegaciones críticas", f"{criticas}")


def resumen_por_delegacion(df):
    if df.empty:
        return pd.DataFrame()
    res = df.groupby(["Dirección Regional", "Delegación", "Latitud", "Longitud"], as_index=False).agg(
        Actividades=("Actividad", "count"),
        Meta=("Meta", "sum"),
        Avance=("Avance", "sum"),
        Pendiente=("Pendiente", "sum"),
    )
    res["% Cumplimiento"] = res.apply(lambda x: x["Avance"] / x["Meta"] if x["Meta"] else 0, axis=1)
    res["Estado"] = res["% Cumplimiento"].apply(clasificar_cumplimiento)
    res = res.sort_values(["Dirección Regional", "Delegación"], key=lambda s: s.map(normalizar_texto))
    return res


def resumen_por_region(df):
    if df.empty:
        return pd.DataFrame()
    res = df.groupby("Dirección Regional", as_index=False).agg(
        Delegaciones=("Delegación", "nunique"),
        Actividades=("Actividad", "count"),
        Meta=("Meta", "sum"),
        Avance=("Avance", "sum"),
        Pendiente=("Pendiente", "sum"),
    )
    res["% Cumplimiento"] = res.apply(lambda x: x["Avance"] / x["Meta"] if x["Meta"] else 0, axis=1)
    res["Estado"] = res["% Cumplimiento"].apply(clasificar_cumplimiento)
    return res.sort_values("Dirección Regional", key=lambda s: s.map(orden_region))


def formato_por_estado(row):
    estado = row.get("Estado", "")
    if estado == "Cumple":
        return ["background-color:#DFF5E8;color:#064E3B;font-weight:700"] * len(row)
    if estado == "En riesgo":
        return ["background-color:#FFF3D6;color:#7C2D12;font-weight:700"] * len(row)
    if estado == "Crítico":
        return ["background-color:#FFE2E2;color:#7F1D1D;font-weight:700"] * len(row)
    return [""] * len(row)

# ======================================================
# GRÁFICOS
# ======================================================

def grafico_delegaciones_pareto(df):
    res = resumen_por_delegacion(df)
    if res.empty:
        st.info("No hay datos para graficar.")
        return

    res = res.sort_values("% Cumplimiento", ascending=False)

    fig = go.Figure()
    fig.add_trace(go.Bar(x=res["Delegación"], y=res["Meta"], name="Meta", marker_color="#DCE8FF", marker_line_color=COLOR_AZUL, marker_line_width=1))
    fig.add_trace(go.Bar(x=res["Delegación"], y=res["Avance"], name="Avance", marker_color=COLOR_AZUL))
    fig.add_trace(go.Bar(x=res["Delegación"], y=res["Pendiente"], name="Pendiente", marker_color=COLOR_ROJO))

    max_meta = max(res["Meta"].max(), 1)
    fig.add_trace(go.Scatter(
        x=res["Delegación"],
        y=[max_meta * UMBRAL_CUMPLE] * len(res),
        mode="lines",
        name="Línea 50%",
        line=dict(color=COLOR_DORADO, width=4, dash="dash")
    ))

    fig.update_layout(
        title="Meta, avance y pendiente por delegación con línea de referencia 50%",
        barmode="group",
        title_x=0.5,
        paper_bgcolor="white",
        plot_bgcolor="white",
        height=620,
        legend=dict(orientation="h", y=1.08, x=0.5, xanchor="center"),
        xaxis_tickangle=-45,
        yaxis_title="Cantidad",
    )
    st.plotly_chart(fig, use_container_width=True)

    top = res.sort_values("% Cumplimiento", ascending=False).head(2)
    bajo = res.sort_values("% Cumplimiento", ascending=True).head(2)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Top 2 mejor posicionadas")
        for i, (_, row) in enumerate(top.iterrows(), start=1):
            st.success(f"{i}. {row['Delegación']} - {row['% Cumplimiento']:.1%}")
    with c2:
        st.markdown("### Top 2 con menor cumplimiento")
        for i, (_, row) in enumerate(bajo.iterrows(), start=1):
            st.error(f"{i}. {row['Delegación']} - {row['% Cumplimiento']:.1%}")


def grafico_region(df):
    res = resumen_por_region(df)
    if res.empty:
        return
    fig = px.bar(
        res,
        x="% Cumplimiento",
        y="Dirección Regional",
        orientation="h",
        color="Estado",
        color_discrete_map={"Cumple": COLOR_VERDE, "En riesgo": COLOR_NARANJA, "Crítico": COLOR_ROJO},
        text=res["% Cumplimiento"].map(lambda x: f"{x:.1%}"),
        title="Cumplimiento por Dirección Regional"
    )
    fig.add_vline(x=0.50, line_dash="dash", line_color=COLOR_AZUL, annotation_text="50%")
    fig.update_layout(title_x=0.5, paper_bgcolor="white", plot_bgcolor="white", height=520)
    st.plotly_chart(fig, use_container_width=True)


def grafico_programa(df):
    if df.empty:
        return
    res = df.groupby("Programa", as_index=False).agg(Meta=("Meta", "sum"), Avance=("Avance", "sum"), Pendiente=("Pendiente", "sum"))
    res["% Cumplimiento"] = res.apply(lambda x: x["Avance"] / x["Meta"] if x["Meta"] else 0, axis=1)
    res["Estado"] = res["% Cumplimiento"].apply(clasificar_cumplimiento)
    fig = px.bar(
        res.sort_values("% Cumplimiento", ascending=False),
        x="Programa",
        y="% Cumplimiento",
        color="Estado",
        color_discrete_map={"Cumple": COLOR_VERDE, "En riesgo": COLOR_NARANJA, "Crítico": COLOR_ROJO},
        text=res["% Cumplimiento"].map(lambda x: f"{x:.1%}"),
        title="Cumplimiento por programa"
    )
    fig.add_hline(y=0.50, line_dash="dash", line_color=COLOR_AZUL, annotation_text="50%")
    fig.update_layout(title_x=0.5, paper_bgcolor="white", plot_bgcolor="white", height=460)
    st.plotly_chart(fig, use_container_width=True)


def grafico_pareto_pendiente(df):
    res = resumen_por_delegacion(df)
    if res.empty:
        return
    res = res.sort_values("Pendiente", ascending=False)
    total_pendiente = res["Pendiente"].sum()
    res["Acumulado"] = res["Pendiente"].cumsum()
    res["% Acumulado"] = res["Acumulado"] / total_pendiente if total_pendiente else 0

    fig = go.Figure()
    fig.add_trace(go.Bar(x=res["Delegación"], y=res["Pendiente"], name="Pendiente", marker_color=COLOR_ROJO))
    fig.add_trace(go.Scatter(x=res["Delegación"], y=res["% Acumulado"] * max(res["Pendiente"].max(), 1), name="% acumulado", mode="lines+markers", marker_color=COLOR_AZUL))
    fig.update_layout(
        title="Pareto nacional de pendiente por delegación",
        title_x=0.5,
        paper_bgcolor="white",
        plot_bgcolor="white",
        xaxis_tickangle=-45,
        height=560,
        yaxis_title="Pendiente"
    )
    st.plotly_chart(fig, use_container_width=True)

# ======================================================
# MAPAS
# ======================================================

def crear_mapa_calor(df):
    res = resumen_por_delegacion(df)
    mapa = folium.Map(location=[9.7489, -83.7534], zoom_start=7, tiles="CartoDB positron")

    if res.empty:
        return mapa

    puntos_calor = []
    for _, row in res.iterrows():
        lat = float(row["Latitud"])
        lon = float(row["Longitud"])
        criticidad = 1 - float(row["% Cumplimiento"])
        peso = max(criticidad, 0.05) * max(float(row["Meta"]), 1)
        puntos_calor.append([lat, lon, peso])

        estado = row["Estado"]
        color = "green" if estado == "Cumple" else "orange" if estado == "En riesgo" else "red"
        popup = f"""
        <div style='font-family:Arial;width:280px;'>
            <h4 style='color:#002B7F;margin-bottom:6px;'>{row['Delegación']}</h4>
            <b>Región:</b> {row['Dirección Regional']}<br>
            <b>Meta:</b> {row['Meta']:,.0f}<br>
            <b>Avance:</b> {row['Avance']:,.0f}<br>
            <b>Pendiente:</b> {row['Pendiente']:,.0f}<br>
            <b>Cumplimiento:</b> {row['% Cumplimiento']:.1%}<br>
            <b>Estado:</b> {estado}<br>
        </div>
        """
        folium.CircleMarker(
            location=[lat, lon], radius=8, color=color, fill=True, fill_color=color,
            fill_opacity=0.75, popup=folium.Popup(popup, max_width=320), tooltip=f"{row['Delegación']} - {row['% Cumplimiento']:.1%}"
        ).add_to(mapa)

    if puntos_calor:
        HeatMap(puntos_calor, radius=38, blur=28, max_zoom=9, min_opacity=0.35).add_to(mapa)

    return mapa

# ======================================================
# PDF
# ======================================================

def estilos_pdf():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TituloPUMI", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=22, alignment=TA_CENTER, textColor=colors.HexColor(COLOR_AZUL), spaceAfter=14))
    styles.add(ParagraphStyle(name="SubtituloPUMI", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=14, alignment=TA_LEFT, textColor=colors.HexColor(COLOR_AZUL), spaceBefore=8, spaceAfter=6))
    styles.add(ParagraphStyle(name="TextoPUMI", parent=styles["BodyText"], fontSize=9, leading=12))
    styles.add(ParagraphStyle(name="TablaPUMI", parent=styles["BodyText"], fontSize=7, leading=9))
    return styles


def p(texto, estilo):
    texto = "" if texto is None else str(texto)
    texto = texto.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return Paragraph(texto, estilo)


def logo_pdf(ruta, w=5*cm, h=1.8*cm):
    if os.path.exists(ruta):
        try:
            return Image(ruta, width=w, height=h)
        except Exception:
            return ""
    return ""


def generar_pdf(df, titulo="Informe Nacional PUMI 2026"):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=1.1*cm, leftMargin=1.1*cm, topMargin=1.2*cm, bottomMargin=1.1*cm)
    styles = estilos_pdf()
    elementos = []

    logos = [[logo_pdf(LOGO_MINISTERIO), logo_pdf(LOGO_PUMI, 4*cm, 2*cm), logo_pdf(LOGO_FUERZA_PUBLICA)]]
    tabla_logos = Table(logos, colWidths=[8*cm, 8*cm, 8*cm])
    tabla_logos.setStyle(TableStyle([("ALIGN", (0,0), (-1,-1), "CENTER"), ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("BOX", (0,0), (-1,-1), 0.8, colors.HexColor(COLOR_AZUL)), ("TOPPADDING", (0,0), (-1,-1), 8), ("BOTTOMPADDING", (0,0), (-1,-1), 8)]))
    elementos += [tabla_logos, Spacer(1, .6*cm), p(titulo, styles["TituloPUMI"]), p("Dashboard Nacional P.U.M.I. 2026", styles["TituloPUMI"]), Spacer(1, .3*cm)]

    meta = df["Meta"].sum() if not df.empty else 0
    avance = df["Avance"].sum() if not df.empty else 0
    pendiente = max(meta - avance, 0)
    cumplimiento = avance / meta if meta else 0
    filtros = st.session_state.get("filtros_pdf", {})

    datos = [
        ["Meta nacional", f"{meta:,.0f}"], ["Avance", f"{avance:,.0f}"], ["Pendiente", f"{pendiente:,.0f}"],
        ["Cumplimiento", f"{cumplimiento:.1%}"], ["Regiones", str(df['Dirección Regional'].nunique() if not df.empty else 0)], ["Delegaciones", str(df['Delegación'].nunique() if not df.empty else 0)], ["Fecha", date.today().strftime("%d/%m/%Y")]
    ]
    tabla = Table(datos, colWidths=[7*cm, 16*cm])
    tabla.setStyle(TableStyle([("BACKGROUND", (0,0), (0,-1), colors.HexColor(COLOR_AZUL)), ("TEXTCOLOR", (0,0), (0,-1), colors.white), ("FONTNAME", (0,0), (-1,-1), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), .4, colors.grey), ("ROWBACKGROUNDS", (1,0), (1,-1), [colors.white, colors.HexColor("#F4F7FB")])]))
    elementos += [tabla, PageBreak()]

    elementos.append(p("1. Filtros aplicados", styles["SubtituloPUMI"]))
    datos_filtros = [["Filtro", "Valor"]]
    for k, v in filtros.items():
        datos_filtros.append([k, ", ".join(v) if isinstance(v, list) and v else "Todos"])
    tabla_filtros = Table(datos_filtros, colWidths=[7*cm, 17*cm])
    tabla_filtros.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor(COLOR_AZUL)), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("GRID", (0,0), (-1,-1), .4, colors.grey), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold")]))
    elementos += [tabla_filtros, Spacer(1, .4*cm)]

    elementos.append(p("2. Resumen por Dirección Regional", styles["SubtituloPUMI"]))
    reg = resumen_por_region(df)
    datos_reg = [["Dirección Regional", "Delegaciones", "Meta", "Avance", "Pendiente", "%", "Estado"]]
    for _, row in reg.iterrows():
        datos_reg.append([row["Dirección Regional"], row["Delegaciones"], f"{row['Meta']:,.0f}", f"{row['Avance']:,.0f}", f"{row['Pendiente']:,.0f}", f"{row['% Cumplimiento']:.1%}", row["Estado"]])
    tabla_reg = Table(datos_reg, colWidths=[6*cm, 2.4*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2*cm, 2.5*cm], repeatRows=1)
    tabla_reg.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor(COLOR_AZUL)), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("GRID", (0,0), (-1,-1), .3, colors.grey), ("FONTSIZE", (0,0), (-1,-1), 7)]))
    elementos += [tabla_reg, PageBreak()]

    elementos.append(p("3. Detalle por delegación", styles["SubtituloPUMI"]))
    deleg = resumen_por_delegacion(df)
    datos_deleg = [["Región", "Delegación", "Meta", "Avance", "Pendiente", "%", "Estado"]]
    for _, row in deleg.iterrows():
        datos_deleg.append([row["Dirección Regional"], row["Delegación"], f"{row['Meta']:,.0f}", f"{row['Avance']:,.0f}", f"{row['Pendiente']:,.0f}", f"{row['% Cumplimiento']:.1%}", row["Estado"]])
    tabla_deleg = Table(datos_deleg, colWidths=[5.4*cm, 5*cm, 2.2*cm, 2.2*cm, 2.2*cm, 1.8*cm, 2.2*cm], repeatRows=1)
    estilo_tabla = [("BACKGROUND", (0,0), (-1,0), colors.HexColor(COLOR_AZUL)), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("GRID", (0,0), (-1,-1), .3, colors.grey), ("FONTSIZE", (0,0), (-1,-1), 7)]
    for i, row in enumerate(datos_deleg[1:], start=1):
        estado = row[-1]
        bg = colors.HexColor("#DFF5E8") if estado == "Cumple" else colors.HexColor("#FFF3D6") if estado == "En riesgo" else colors.HexColor("#FFE2E2")
        estilo_tabla.append(("BACKGROUND", (0, i), (-1, i), bg))
    tabla_deleg.setStyle(TableStyle(estilo_tabla))
    elementos.append(tabla_deleg)

    doc.build(elementos)
    buffer.seek(0)
    return buffer.getvalue()

# ======================================================
# FLUJO PRINCIPAL
# ======================================================

mostrar_logo_sidebar()
st.sidebar.markdown("## Dashboard Nacional PUMI")
menu = st.sidebar.radio("Menú", ["Inicio", "Dashboard Nacional", "Análisis Comparativo", "Mapa de calor", "Informe PDF"])

mostrar_encabezado()
mostrar_titulo()

df_nacional = cargar_datos_nacionales()

if df_nacional.empty:
    st.warning("No se encontraron archivos Excel regionales. Coloque los archivos DR*.xlsx al mismo nivel de esta app.py o en la carpeta metas.")
    st.stop()

if menu == "Inicio":
    st.markdown(
        """
        <div class='card-info'>
            <div class='subtitulo'>Visualización Nacional PUMI 2026</div>
            <div class='texto'>
            Esta aplicación consolida automáticamente los Excel regionales ubicados en el repositorio,
            calcula metas, avances, pendientes y niveles de cumplimiento por región, delegación, programa y actividad.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
    mostrar_metricas(df_nacional)
    st.markdown("### Archivos detectados")
    st.dataframe(pd.DataFrame({"Archivo": sorted(df_nacional["Archivo"].unique().tolist())}), use_container_width=True, hide_index=True)

elif menu == "Dashboard Nacional":
    df_filtrado = aplicar_filtros(df_nacional, key="dash")
    mostrar_metricas(df_filtrado)
    st.markdown("---")
    grafico_delegaciones_pareto(df_filtrado)
    st.markdown("### Tabla nacional por delegación")
    tabla = resumen_por_delegacion(df_filtrado)
    vista = tabla.copy()
    vista["% Cumplimiento"] = vista["% Cumplimiento"].map(lambda x: f"{x:.1%}")
    for col in ["Meta", "Avance", "Pendiente"]:
        vista[col] = vista[col].map(lambda x: f"{x:,.0f}")
    st.dataframe(vista.style.apply(formato_por_estado, axis=1), use_container_width=True, hide_index=True)

elif menu == "Análisis Comparativo":
    df_filtrado = aplicar_filtros(df_nacional, key="comparativo")
    mostrar_metricas(df_filtrado)
    st.markdown("---")
    grafico_region(df_filtrado)
    grafico_programa(df_filtrado)
    grafico_pareto_pendiente(df_filtrado)

elif menu == "Mapa de calor":
    df_filtrado = aplicar_filtros(df_nacional, key="mapa")
    mostrar_metricas(df_filtrado)
    st.markdown("### Mapa de calor de menor cumplimiento")
    st.info("El mapa de calor aumenta la intensidad donde existe mayor pendiente y menor porcentaje de cumplimiento según los filtros aplicados.")
    mapa = crear_mapa_calor(df_filtrado)
    st_folium(mapa, height=640, use_container_width=True, key="mapa_calor_nacional")

elif menu == "Informe PDF":
    df_filtrado = aplicar_filtros(df_nacional, key="pdf")
    mostrar_metricas(df_filtrado)
    st.markdown("### Generar informe PDF")
    st.info("El PDF se genera con la información filtrada, incluyendo colores por cumplimiento en el detalle por delegación.")
    if st.button("📄 Generar PDF nacional filtrado"):
        pdf = generar_pdf(df_filtrado)
        st.success("PDF generado correctamente.")
        st.download_button(
            "⬇️ Descargar informe PDF",
            data=pdf,
            file_name=f"INFORME_DASHBOARD_NACIONAL_PUMI_{date.today().strftime('%Y%m%d')}.pdf",
            mime="application/pdf"
        )

# Descarga global CSV/Excel simple desde sidebar.
st.sidebar.markdown("---")
st.sidebar.markdown("### Datos")
if st.sidebar.button("🔄 Recargar datos"):
    st.cache_data.clear()
    st.rerun()

salida = BytesIO()
with pd.ExcelWriter(salida, engine="openpyxl") as writer:
    df_nacional.to_excel(writer, index=False, sheet_name="DATOS_NACIONALES")
    resumen_por_region(df_nacional).to_excel(writer, index=False, sheet_name="RESUMEN_REGION")
    resumen_por_delegacion(df_nacional).to_excel(writer, index=False, sheet_name="RESUMEN_DELEGACION")
salida.seek(0)
st.sidebar.download_button(
    "⬇️ Descargar base consolidada",
    data=salida.getvalue(),
    file_name="BASE_CONSOLIDADA_DASHBOARD_NACIONAL_PUMI.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
