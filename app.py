# ======================================================
# PARTE 1 DE 5
# CONFIGURACIÓN GENERAL, LIBRERÍAS, CATÁLOGOS Y ESTILOS
# ======================================================

import streamlit as st
import pandas as pd
import gspread
import plotly.express as px
import unicodedata
import os

import folium
from streamlit_folium import st_folium
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from streamlit_js_eval import get_geolocation

from google.oauth2.service_account import Credentials
from datetime import datetime, date
from io import BytesIO
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image as RLImage, PageBreak
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY


# ======================================================
# CONFIGURACIÓN GENERAL DE LA APP
# ======================================================

st.set_page_config(
    page_title="PUMI 2026",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ======================================================
# ARCHIVOS BASE EN EL REPOSITORIO
# ======================================================

ARCHIVO_MEP = "BASE DE DATOS MEP 2025.xlsx"
ARCHIVO_DELEGACIONES = "DELEGACIONES Y DISTRITOS.xlsx"
ARCHIVO_DATOS_IMPORTANTES = "Datos Importantes.xlsx"


# ======================================================
# CLAVE ADMINISTRATIVA
# ======================================================

CLAVE_ADMIN = "DPPP23"


# ======================================================
# CONEXIÓN BASE GOOGLE SHEETS
# ======================================================

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1O-HNa1c4ppF0-ND7BUqKA2OzZDc1O0kTxZOMVDeA-GY/edit?gid=0#gid=0"
HOJA_REGISTRO = "REGISTRO_PUMI_2026"


# ======================================================
# ENCABEZADOS OFICIALES DE LA BASE DE DATOS
# ======================================================

ENCABEZADOS = [
    "ID",
    "Fecha Registro",
    "Fecha Actividad",
    "Dirección Regional",
    "Delegación",
    "Programa",
    "Actividad",
    "Provincia",
    "Cantón",
    "Distrito",
    "Tipo Lugar",
    "Lugar",
    "Centro Educativo",
    "Dirección Mapa",
    "Latitud",
    "Longitud",
    "Responsable",
    "Cantidad Participantes",
    "Instituciones Participantes",
    "Plan Estratégico Relacionado",
    "Evidencia",
    "Observaciones",
    "Estado Revisión",
    "Observación de Revisión",
    "Usuario Registra"
]


# ======================================================
# CATÁLOGOS BASE DE RESPALDO
# Si el Excel no carga, la app usa estos valores.
# ======================================================

PROGRAMAS = [
    "PSCC",
    "Política Local",
    "VIF",
    "DARE",
    "GREAT",
    "MPAS",
    "GREAT CAMP"
]

REGIONES = [
    "R1 San José Central",
    "R2 San José Norte",
    "R3 San José Sur",
    "R4 Alajuela",
    "R5 Cartago",
    "R6 Heredia",
    "R7 Chorotega",
    "R8 Puntarenas",
    "R9 Limón",
    "R10 Brunca",
    "R11 Chorotega Norte"
]

PROVINCIAS = [
    "San José",
    "Alajuela",
    "Cartago",
    "Heredia",
    "Guanacaste",
    "Puntarenas",
    "Limón"
]

ESTADOS_REVISION = [
    "Pendiente de revisión",
    "Aprobado",
    "Con observaciones",
    "Rechazado"
]

COLORES_PROGRAMA = {
    "PSCC": "blue",
    "Política Local": "green",
    "VIF": "red",
    "DARE": "purple",
    "GREAT": "orange",
    "MPAS": "darkblue",
    "GREAT CAMP": "cadetblue"
}


# ======================================================
# COLORES INSTITUCIONALES
# ======================================================

COLOR_AZUL = "#003366"
COLOR_AZUL_MEDIO = "#005B96"
COLOR_AZUL_CLARO = "#0B84C6"
COLOR_VERDE = "#1B7F3A"
COLOR_VERDE_CLARO = "#2EAD5F"
COLOR_DORADO = "#C9A227"
COLOR_GRIS = "#F4F6F8"
COLOR_GRIS_OSCURO = "#2F3542"
COLOR_BLANCO = "#FFFFFF"


# ======================================================
# ESTILOS VISUALES DE LA APP
# ======================================================

st.markdown(
    f"""
    <style>

    .stApp {{
        background:
            radial-gradient(circle at top left, #DFF3FF 0, transparent 32%),
            radial-gradient(circle at top right, #E4F8EA 0, transparent 28%),
            linear-gradient(180deg, #F4F8FB 0%, #EAF2F8 45%, #F7FAFC 100%);
    }}

    section[data-testid="stSidebar"] {{
        background: linear-gradient(180deg, #FFFFFF 0%, #EAF4FB 45%, #E6F5EC 100%);
        border-right: 5px solid {COLOR_AZUL};
        box-shadow: 4px 0px 12px rgba(0,0,0,0.10);
    }}

    .titulo-principal {{
        background: linear-gradient(135deg, {COLOR_AZUL}, {COLOR_VERDE});
        padding: 38px;
        border-radius: 22px;
        text-align: center;
        color: white;
        margin-bottom: 30px;
        box-shadow: 0px 8px 22px rgba(0,0,0,0.28);
        border: 2px solid rgba(255,255,255,0.35);
    }}

    .titulo-principal h1 {{
        font-size: 52px;
        margin-bottom: 8px;
        font-weight: 900;
        letter-spacing: 1.5px;
        text-shadow: 1px 2px 5px rgba(0,0,0,0.35);
    }}

    .titulo-principal h3 {{
        font-size: 27px;
        margin-top: 8px;
        font-weight: 600;
        color: #F4FFF8;
    }}

    .card-pumi {{
        background: linear-gradient(135deg, #FFFFFF 0%, #F0FAF4 100%);
        padding: 24px;
        border-radius: 20px;
        box-shadow: 0px 6px 16px rgba(0,0,0,0.13);
        border-left: 9px solid {COLOR_VERDE};
        margin-bottom: 20px;
    }}

    .card-azul {{
        background: linear-gradient(135deg, #FFFFFF 0%, #EEF7FF 100%);
        padding: 24px;
        border-radius: 20px;
        box-shadow: 0px 6px 16px rgba(0,0,0,0.13);
        border-left: 9px solid {COLOR_AZUL};
        margin-bottom: 20px;
    }}

    .card-dorado {{
        background: linear-gradient(135deg, #FFFFFF 0%, #FFF8DC 100%);
        padding: 24px;
        border-radius: 20px;
        box-shadow: 0px 6px 16px rgba(0,0,0,0.13);
        border-left: 9px solid {COLOR_DORADO};
        margin-bottom: 20px;
    }}

    .bloque-datos {{
        background: linear-gradient(135deg, #F0F7FF 0%, #FFFFFF 100%);
        padding: 18px;
        border-radius: 18px;
        border-left: 7px solid {COLOR_AZUL};
        margin-bottom: 18px;
    }}

    .bloque-territorio {{
        background: linear-gradient(135deg, #EFFAF3 0%, #FFFFFF 100%);
        padding: 18px;
        border-radius: 18px;
        border-left: 7px solid {COLOR_VERDE};
        margin-bottom: 18px;
    }}

    .bloque-actividad {{
        background: linear-gradient(135deg, #FFF8DC 0%, #FFFFFF 100%);
        padding: 18px;
        border-radius: 18px;
        border-left: 7px solid {COLOR_DORADO};
        margin-bottom: 18px;
    }}

    .bloque-mapa {{
        background: linear-gradient(135deg, #EAF7FF 0%, #FFFFFF 100%);
        padding: 18px;
        border-radius: 18px;
        border-left: 7px solid {COLOR_AZUL_CLARO};
        margin-bottom: 18px;
    }}

    .subtitulo-pumi {{
        color: {COLOR_AZUL};
        font-weight: 900;
        font-size: 30px;
        margin-bottom: 12px;
    }}

    .texto-pumi {{
        color: {COLOR_GRIS_OSCURO};
        font-size: 17px;
        line-height: 1.7;
    }}

    div[data-testid="stMetric"] {{
        background: linear-gradient(145deg, #FFFFFF 0%, #F0F7FB 100%);
        padding: 20px;
        border-radius: 20px;
        box-shadow: 0px 5px 15px rgba(0,0,0,0.13);
        border-bottom: 6px solid {COLOR_VERDE};
    }}

    div[data-baseweb="select"] > div {{
        background-color: #F7FBFF !important;
        border-radius: 12px !important;
        border: 1.8px solid #9FC5E8 !important;
    }}

    input {{
        background-color: #F8FFF9 !important;
        border-radius: 12px !important;
        border: 1.8px solid #A9D6B8 !important;
    }}

    textarea {{
        background-color: #FFFDF3 !important;
        border-radius: 12px !important;
        border: 1.8px solid #E1C75F !important;
    }}

    .stButton > button {{
        background: linear-gradient(90deg, {COLOR_AZUL}, {COLOR_VERDE});
        color: white;
        border-radius: 12px;
        border: none;
        font-weight: 800;
        padding: 0.7rem 1.2rem;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.18);
    }}

    .stDownloadButton > button {{
        background: linear-gradient(90deg, {COLOR_VERDE}, {COLOR_AZUL});
        color: white;
        border-radius: 12px;
        border: none;
        font-weight: 800;
        padding: 0.7rem 1.2rem;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.18);
    }}

    div[data-testid="stDataFrame"] {{
        border-radius: 18px;
        overflow: hidden;
        box-shadow: 0px 5px 16px rgba(0,0,0,0.13);
        background-color: white;
    }}

    iframe {{
        border-radius: 18px !important;
        box-shadow: 0px 5px 16px rgba(0,0,0,0.13);
    }}

    hr {{
        border: none;
        height: 2px;
        background: linear-gradient(90deg, {COLOR_AZUL}, {COLOR_VERDE}, {COLOR_DORADO});
        margin-top: 25px;
        margin-bottom: 25px;
    }}

    h1, h2, h3 {{
        color: {COLOR_AZUL};
        font-weight: 900;
    }}

    </style>
    """,
    unsafe_allow_html=True
)


# ======================================================
# FUNCIÓN PARA MOSTRAR LOGO
# ======================================================

def mostrar_logo():
    if os.path.exists("logo_pumi.jpeg"):
        logo = Image.open("logo_pumi.jpeg")
        st.sidebar.image(logo, use_container_width=True)

    elif os.path.exists("logo_pumi.jpg"):
        logo = Image.open("logo_pumi.jpg")
        st.sidebar.image(logo, use_container_width=True)

    elif os.path.exists("logo_pumi.png"):
        logo = Image.open("logo_pumi.png")
        st.sidebar.image(logo, use_container_width=True)

    else:
        st.sidebar.warning("Logo PUMI no encontrado.")
# ======================================================
# PARTE 2 DE 5
# CONEXIÓN GOOGLE SHEETS, BASES AUXILIARES, CRUD, MAPA Y FILTROS
# ======================================================

@st.cache_resource
def conectar_google_sheets():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    try:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=scopes
        )

        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(SPREADSHEET_URL)

        return spreadsheet

    except Exception as e:
        st.error("No se pudo conectar con Google Sheets.")
        st.exception(e)
        st.stop()


def obtener_hoja(nombre_hoja):
    try:
        spreadsheet = conectar_google_sheets()

        try:
            return spreadsheet.worksheet(nombre_hoja)

        except gspread.WorksheetNotFound:
            return spreadsheet.add_worksheet(
                title=nombre_hoja,
                rows=5000,
                cols=len(ENCABEZADOS)
            )

    except Exception as e:
        st.error("Error al conectar con la hoja de Google Sheets.")
        st.exception(e)
        st.stop()


def inicializar_hoja():
    hoja = obtener_hoja(HOJA_REGISTRO)

    try:
        datos = hoja.get_all_values()
    except Exception as e:
        st.error("No se pudieron leer los datos de la hoja.")
        st.exception(e)
        st.stop()

    if len(datos) == 0:
        hoja.append_row(ENCABEZADOS)
        return hoja

    primera_fila = datos[0]

    if primera_fila != ENCABEZADOS:
        hoja.insert_row(ENCABEZADOS, index=1)

    return hoja


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join([c for c in texto if not unicodedata.combining(c)])
    texto = " ".join(texto.split())

    return texto


def obtener_columna_por_nombre(df, posibles_nombres):
    columnas = list(df.columns)

    for posible in posibles_nombres:
        posible_norm = normalizar_texto(posible)

        for col in columnas:
            if normalizar_texto(col) == posible_norm:
                return col

    return None


# ======================================================
# BASE DATOS IMPORTANTES
# ======================================================

@st.cache_data
def cargar_datos_importantes():
    """
    Carga Datos Importantes.xlsx.

    Estructura esperada:
    Provincia, Cantón, Distritos, Dirección Regional,
    Delegación, Actividad Realizada, Programa.

    Las columnas vacías intermedias del Excel se ignoran.
    """

    columnas_base = [
        "Provincia",
        "Cantón",
        "Distrito",
        "Dirección Regional",
        "Delegación",
        "Actividad Realizada",
        "Programa"
    ]

    if not os.path.exists(ARCHIVO_DATOS_IMPORTANTES):
        return pd.DataFrame(columns=columnas_base)

    try:
        df_original = pd.read_excel(ARCHIVO_DATOS_IMPORTANTES)

        col_provincia = obtener_columna_por_nombre(
            df_original,
            ["Provincia"]
        )

        col_canton = obtener_columna_por_nombre(
            df_original,
            ["Cantón", "Canton"]
        )

        col_distrito = obtener_columna_por_nombre(
            df_original,
            ["Distrito", "Distritos"]
        )

        col_region = obtener_columna_por_nombre(
            df_original,
            ["Dirección Regional", "Direccion Regional"]
        )

        col_delegacion = obtener_columna_por_nombre(
            df_original,
            ["Delegación", "Delegacion"]
        )

        col_actividad = obtener_columna_por_nombre(
            df_original,
            ["Actividad Realizada", "Actividad"]
        )

        col_programa = obtener_columna_por_nombre(
            df_original,
            ["Programa"]
        )

        columnas_requeridas = [
            col_provincia,
            col_canton,
            col_distrito,
            col_region,
            col_delegacion,
            col_actividad,
            col_programa
        ]

        if not all(columnas_requeridas):
            st.warning(
                "El archivo Datos Importantes.xlsx no tiene todas las columnas requeridas."
            )
            return pd.DataFrame(columns=columnas_base)

        df = df_original[
            [
                col_provincia,
                col_canton,
                col_distrito,
                col_region,
                col_delegacion,
                col_actividad,
                col_programa
            ]
        ].copy()

        df.columns = columnas_base

        for col in columnas_base:
            df[col] = (
                df[col]
                .fillna("")
                .astype(str)
                .str.strip()
            )

        df = df.replace("nan", "")
        df = df.drop_duplicates()

        df["Provincia_Normalizada"] = df["Provincia"].apply(normalizar_texto)
        df["Cantón_Normalizado"] = df["Cantón"].apply(normalizar_texto)
        df["Distrito_Normalizado"] = df["Distrito"].apply(normalizar_texto)

        df["Región_Normalizada"] = df["Dirección Regional"].apply(
            normalizar_texto
        )

        df["Delegación_Normalizada"] = df["Delegación"].apply(
            normalizar_texto
        )

        df["Actividad_Normalizada"] = df["Actividad Realizada"].apply(
            normalizar_texto
        )

        df["Programa_Normalizado"] = df["Programa"].apply(
            normalizar_texto
        )

        return df

    except Exception as e:
        st.error(f"Error leyendo {ARCHIVO_DATOS_IMPORTANTES}")
        st.exception(e)
        return pd.DataFrame(columns=columnas_base)


def obtener_regiones_datos():
    df = cargar_datos_importantes()

    if df.empty:
        return REGIONES

    regiones = df["Dirección Regional"].dropna().unique().tolist()

    regiones = [
        x for x in regiones
        if str(x).strip() != ""
    ]

    return sorted(regiones)


def obtener_delegaciones_por_region(region):
    df = cargar_datos_importantes()

    if df.empty or not region:
        return obtener_delegaciones_unicas()

    region_norm = normalizar_texto(region)

    delegaciones = df[
        df["Región_Normalizada"] == region_norm
    ]["Delegación"].dropna().unique().tolist()

    delegaciones = [
        x for x in delegaciones
        if str(x).strip() != ""
    ]

    return sorted(delegaciones)


def obtener_actividades_por_delegacion(delegacion):
    df = cargar_datos_importantes()

    if df.empty or not delegacion:
        return []

    delegacion_norm = normalizar_texto(delegacion)

    actividades = df[
        df["Delegación_Normalizada"] == delegacion_norm
    ]["Actividad Realizada"].dropna().unique().tolist()

    actividades = [
        x for x in actividades
        if str(x).strip() != ""
    ]

    return sorted(actividades)


def obtener_programas_por_actividad(actividad):
    df = cargar_datos_importantes()

    if df.empty or not actividad:
        return PROGRAMAS

    actividad_norm = normalizar_texto(actividad)

    programas = df[
        df["Actividad_Normalizada"] == actividad_norm
    ]["Programa"].dropna().unique().tolist()

    programas = [
        x for x in programas
        if str(x).strip() != ""
    ]

    if not programas:
        return PROGRAMAS

    return sorted(programas)


def obtener_provincias_datos():
    df = cargar_datos_importantes()

    if df.empty:
        return PROVINCIAS

    provincias = df["Provincia"].dropna().unique().tolist()

    provincias = [
        x for x in provincias
        if str(x).strip() != ""
    ]

    return sorted(provincias)


def obtener_cantones_por_provincia(provincia):
    df = cargar_datos_importantes()

    if df.empty or not provincia:
        return []

    provincia_norm = normalizar_texto(provincia)

    cantones = df[
        df["Provincia_Normalizada"] == provincia_norm
    ]["Cantón"].dropna().unique().tolist()

    cantones = [
        x for x in cantones
        if str(x).strip() != ""
    ]

    return sorted(cantones)


def obtener_distritos_por_provincia_canton(provincia, canton):
    df = cargar_datos_importantes()

    if df.empty or not provincia or not canton:
        return []

    provincia_norm = normalizar_texto(provincia)
    canton_norm = normalizar_texto(canton)

    distritos = df[
        (df["Provincia_Normalizada"] == provincia_norm) &
        (df["Cantón_Normalizado"] == canton_norm)
    ]["Distrito"].dropna().unique().tolist()

    distritos = [
        x for x in distritos
        if str(x).strip() != ""
    ]

    return sorted(distritos)


# ======================================================
# BASE MEP
# ======================================================

@st.cache_data
def cargar_base_mep():
    if not os.path.exists(ARCHIVO_MEP):
        return pd.DataFrame(columns=["PROVINCIA", "NOMBRE"])

    df = pd.read_excel(ARCHIVO_MEP)
    df.columns = [str(c).strip().upper() for c in df.columns]

    if "PROVINCIA" not in df.columns or "NOMBRE" not in df.columns:
        return pd.DataFrame(columns=["PROVINCIA", "NOMBRE"])

    df = df[["PROVINCIA", "NOMBRE"]].copy()
    df = df.dropna(subset=["PROVINCIA", "NOMBRE"])

    df["PROVINCIA_NORMALIZADA"] = df["PROVINCIA"].apply(normalizar_texto)
    df["NOMBRE"] = df["NOMBRE"].astype(str).str.strip()

    df = df.drop_duplicates(subset=["PROVINCIA_NORMALIZADA", "NOMBRE"])

    return df


def obtener_centros_por_provincia(provincia):
    df = cargar_base_mep()

    if df.empty or not provincia:
        return []

    provincia_norm = normalizar_texto(provincia)

    centros = df[
        df["PROVINCIA_NORMALIZADA"] == provincia_norm
    ]["NOMBRE"].dropna().astype(str).drop_duplicates().sort_values().tolist()

    return centros


# ======================================================
# BASE DELEGACIONES Y DISTRITOS
# Se conserva como respaldo para delegaciones antiguas.
# Los distritos ahora salen de Datos Importantes.xlsx.
# ======================================================

@st.cache_data
def cargar_base_delegaciones():
    if not os.path.exists(ARCHIVO_DELEGACIONES):
        return pd.DataFrame(columns=["Delegacion", "Distrito"])

    df_original = pd.read_excel(ARCHIVO_DELEGACIONES)

    col_delegacion = obtener_columna_por_nombre(df_original, ["Delegacion", "Delegación"])
    col_distrito = obtener_columna_por_nombre(df_original, ["Distrito"])
    col_provincia = obtener_columna_por_nombre(df_original, ["Provincia"])
    col_canton = obtener_columna_por_nombre(df_original, ["Canton", "Cantón"])

    columnas = {}

    if col_delegacion:
        columnas[col_delegacion] = "Delegacion"

    if col_distrito:
        columnas[col_distrito] = "Distrito"

    if col_provincia:
        columnas[col_provincia] = "Provincia"

    if col_canton:
        columnas[col_canton] = "Cantón"

    if not col_delegacion and not col_distrito:
        return pd.DataFrame(columns=["Delegacion", "Distrito", "Provincia", "Cantón"])

    df = df_original[list(columnas.keys())].copy()
    df = df.rename(columns=columnas)

    for col in ["Delegacion", "Distrito", "Provincia", "Cantón"]:
        if col not in df.columns:
            df[col] = ""

        df[col] = df[col].astype(str).str.strip()

    df = df.replace("nan", "")

    df["Delegacion_Normalizada"] = df["Delegacion"].apply(normalizar_texto)
    df["Distrito_Normalizado"] = df["Distrito"].apply(normalizar_texto)
    df["Provincia_Normalizada"] = df["Provincia"].apply(normalizar_texto)
    df["Cantón_Normalizado"] = df["Cantón"].apply(normalizar_texto)

    df = df.drop_duplicates()

    return df


def obtener_delegaciones_unicas():
    df_datos = cargar_datos_importantes()

    if not df_datos.empty:
        delegaciones = df_datos["Delegación"].dropna().unique().tolist()

        delegaciones = [
            x for x in delegaciones
            if str(x).strip() != ""
        ]

        return sorted(delegaciones)

    df = cargar_base_delegaciones()

    if df.empty or "Delegacion" not in df.columns:
        return []

    df_tmp = df[["Delegacion", "Delegacion_Normalizada"]].drop_duplicates(
        subset=["Delegacion_Normalizada"]
    )

    return sorted(
        [
            x for x in df_tmp["Delegacion"].dropna().tolist()
            if x.strip() != ""
        ]
    )


def obtener_distritos_unicos():
    df_datos = cargar_datos_importantes()

    if not df_datos.empty:
        distritos = df_datos["Distrito"].dropna().unique().tolist()

        distritos = [
            x for x in distritos
            if str(x).strip() != ""
        ]

        return sorted(distritos)

    df = cargar_base_delegaciones()

    if df.empty:
        return []

    df_tmp = df[["Distrito", "Distrito_Normalizado"]].drop_duplicates(
        subset=["Distrito_Normalizado"]
    )

    return sorted(
        [
            x for x in df_tmp["Distrito"].dropna().tolist()
            if x.strip() != ""
        ]
    )


# ======================================================
# GEOREFERENCIA Y MAPA
# ======================================================

@st.cache_data(show_spinner=False)
def georreferenciar_direccion(direccion):
    if not direccion or str(direccion).strip() == "":
        return None, None, ""

    try:
        geolocator = Nominatim(user_agent="pumi_2026_streamlit_app")

        geocode = RateLimiter(
            geolocator.geocode,
            min_delay_seconds=1,
            max_retries=2,
            error_wait_seconds=2
        )

        consulta = f"{direccion}, Costa Rica"
        location = geocode(consulta)

        if location:
            return location.latitude, location.longitude, location.address

        return None, None, ""

    except Exception:
        return None, None, ""


def limpiar_coordenada(valor):
    try:
        if valor is None or str(valor).strip() == "":
            return None

        return float(str(valor).replace(",", ".").strip())

    except Exception:
        return None


def obtener_color_programa(programa):
    return COLORES_PROGRAMA.get(programa, "gray")


def preparar_dataframe_mapa(df):
    df_mapa = df.copy()

    if "Latitud" not in df_mapa.columns:
        df_mapa["Latitud"] = ""

    if "Longitud" not in df_mapa.columns:
        df_mapa["Longitud"] = ""

    df_mapa["Latitud_Num"] = df_mapa["Latitud"].apply(limpiar_coordenada)
    df_mapa["Longitud_Num"] = df_mapa["Longitud"].apply(limpiar_coordenada)

    df_mapa = df_mapa.dropna(subset=["Latitud_Num", "Longitud_Num"])

    return df_mapa


def crear_mapa_registros(df, zoom_start=8):
    df_mapa = preparar_dataframe_mapa(df)

    if df_mapa.empty:
        return folium.Map(
            location=[9.7489, -83.7534],
            zoom_start=7,
            tiles="OpenStreetMap"
        )

    centro_lat = df_mapa["Latitud_Num"].mean()
    centro_lon = df_mapa["Longitud_Num"].mean()

    mapa = folium.Map(
        location=[centro_lat, centro_lon],
        zoom_start=zoom_start,
        tiles="OpenStreetMap"
    )

    for _, row in df_mapa.iterrows():
        programa = str(row.get("Programa", ""))
        color = obtener_color_programa(programa)

        popup_html = f"""
        <div style="font-family: Arial; width: 260px;">
            <h4 style="color:#003366; margin-bottom:6px;">Registro PUMI #{row.get("ID", "")}</h4>
            <b>Fecha:</b> {row.get("Fecha Actividad", "")}<br>
            <b>Delegación:</b> {row.get("Delegación", "")}<br>
            <b>Programa:</b> {row.get("Programa", "")}<br>
            <b>Actividad:</b> {row.get("Actividad", "")}<br>
            <b>Provincia:</b> {row.get("Provincia", "")}<br>
            <b>Cantón:</b> {row.get("Cantón", "")}<br>
            <b>Distrito:</b> {row.get("Distrito", "")}<br>
            <b>Lugar:</b> {row.get("Lugar", "")}<br>
            <b>Estado:</b> {row.get("Estado Revisión", "")}<br>
            <b>Observación revisión:</b> {row.get("Observación de Revisión", "")}
        </div>
        """

        folium.Marker(
            location=[row["Latitud_Num"], row["Longitud_Num"]],
            popup=folium.Popup(popup_html, max_width=320),
            tooltip=f'{row.get("Programa", "")} - {row.get("Delegación", "")}',
            icon=folium.Icon(color=color, icon="info-sign")
        ).add_to(mapa)

    return mapa


def mostrar_mapa_registros(df, height=520, key="mapa_registros"):
    df_mapa = preparar_dataframe_mapa(df)

    if df_mapa.empty:
        st.info("No hay registros con coordenadas para mostrar en el mapa.")

    mapa = crear_mapa_registros(df)

    st_folium(
        mapa,
        width=1100,
        height=height,
        key=key
    )


# ======================================================
# DATOS PRINCIPALES GOOGLE SHEETS
# ======================================================

def cargar_datos():
    hoja = inicializar_hoja()

    try:
        datos = hoja.get_all_values()
    except Exception as e:
        st.error("No se pudieron cargar los datos desde Google Sheets.")
        st.exception(e)
        st.stop()

    if len(datos) <= 1:
        return pd.DataFrame(columns=ENCABEZADOS)

    filas = datos[1:]
    filas_limpias = []

    for fila in filas:
        fila_ajustada = fila[:len(ENCABEZADOS)]

        while len(fila_ajustada) < len(ENCABEZADOS):
            fila_ajustada.append("")

        if all(str(celda).strip() == "" for celda in fila_ajustada):
            continue

        if str(fila_ajustada[0]).strip().upper() == "ID":
            continue

        filas_limpias.append(fila_ajustada)

    df = pd.DataFrame(filas_limpias, columns=ENCABEZADOS)

    return df


def limpiar_encabezados_duplicados_en_sheet():
    hoja = inicializar_hoja()
    datos = hoja.get_all_values()

    filas_a_eliminar = []

    for i, fila in enumerate(datos[1:], start=2):
        if len(fila) > 0 and str(fila[0]).strip().upper() == "ID":
            filas_a_eliminar.append(i)

    for fila_num in reversed(filas_a_eliminar):
        hoja.delete_rows(fila_num)

    return len(filas_a_eliminar)


def generar_id_consecutivo():
    df = cargar_datos()

    if df.empty:
        return 1

    ids = pd.to_numeric(df["ID"], errors="coerce").dropna()

    if ids.empty:
        return 1

    return int(ids.max()) + 1


def guardar_registro(registro):
    hoja = inicializar_hoja()
    fila = [registro.get(col, "") for col in ENCABEZADOS]

    try:
        hoja.append_row(fila)
    except Exception as e:
        st.error("No se pudo guardar el registro en Google Sheets.")
        st.exception(e)
        st.stop()


def actualizar_registro_por_id(id_registro, nuevos_datos):
    hoja = inicializar_hoja()
    datos = hoja.get_all_values()

    ultima_columna = chr(64 + len(ENCABEZADOS))

    for i, fila in enumerate(datos[1:], start=2):
        if len(fila) > 0 and str(fila[0]) == str(id_registro):
            nueva_fila = [nuevos_datos.get(col, "") for col in ENCABEZADOS]
            rango = f"A{i}:{ultima_columna}{i}"

            try:
                hoja.update(rango, [nueva_fila])
                return True
            except Exception as e:
                st.error("No se pudo actualizar el registro.")
                st.exception(e)
                st.stop()

    return False


def eliminar_registro_por_id(id_registro):
    hoja = inicializar_hoja()
    datos = hoja.get_all_values()

    for i, fila in enumerate(datos[1:], start=2):
        if len(fila) > 0 and str(fila[0]) == str(id_registro):

            try:
                hoja.delete_rows(i)
                return True
            except Exception as e:
                st.error("No se pudo eliminar el registro.")
                st.exception(e)
                st.stop()

    return False


def convertir_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="REGISTRO_PUMI_2026"
        )

    return output.getvalue()


def limpiar_dataframe_para_metricas(df):
    df = df.copy()

    if "Cantidad Participantes" in df.columns:
        df["Cantidad Participantes"] = pd.to_numeric(
            df["Cantidad Participantes"],
            errors="coerce"
        ).fillna(0)

    if "Fecha Actividad" in df.columns:
        df["Fecha Actividad"] = pd.to_datetime(
            df["Fecha Actividad"],
            errors="coerce",
            dayfirst=True
        )

    if "Latitud" in df.columns:
        df["Latitud"] = df["Latitud"].astype(str)

    if "Longitud" in df.columns:
        df["Longitud"] = df["Longitud"].astype(str)

    return df
# ======================================================
# PARTE 3 DE 5
# INTERFAZ PRINCIPAL, LOGIN ADMIN, INICIO, REGISTRO, GPS, MAPA Y SEGUIMIENTO
# ======================================================

def crear_mapa_base(centro, zoom, tipo_mapa):
    mapa = folium.Map(location=centro, zoom_start=zoom, tiles=None)

    if tipo_mapa == "OpenStreetMap":
        folium.TileLayer("OpenStreetMap", name="OpenStreetMap").add_to(mapa)
    elif tipo_mapa == "Mapa claro":
        folium.TileLayer("CartoDB positron", name="Mapa claro").add_to(mapa)
    elif tipo_mapa == "Mapa oscuro":
        folium.TileLayer("CartoDB dark_matter", name="Mapa oscuro").add_to(mapa)
    elif tipo_mapa == "Topográfico":
        folium.TileLayer("OpenTopoMap", name="Topográfico").add_to(mapa)
    elif tipo_mapa == "Satélite":
        folium.TileLayer(
            tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
            attr="Esri World Imagery",
            name="Satélite"
        ).add_to(mapa)

    folium.LayerControl().add_to(mapa)
    return mapa


mostrar_logo()
st.sidebar.markdown("## Sistema PUMI 2026")

if "admin_autenticado" not in st.session_state:
    st.session_state.admin_autenticado = False

with st.sidebar.expander("🔒 Acceso Administrativo"):
    if not st.session_state.admin_autenticado:
        clave_ingresada = st.text_input("Clave administrativa", type="password")

        if st.button("Ingresar como administrador"):
            if clave_ingresada == CLAVE_ADMIN:
                st.session_state.admin_autenticado = True
                st.success("Acceso concedido.")
                st.rerun()
            else:
                st.error("Clave incorrecta.")
    else:
        st.success("Administrador autenticado")
        if st.button("Cerrar sesión administrativa"):
            st.session_state.admin_autenticado = False
            st.rerun()


opciones_menu = [
    "Inicio",
    "Registrar actividad",
    "Seguimiento de registros"
]

if st.session_state.admin_autenticado:
    opciones_menu.extend([
        "Consulta / edición administrativa",
        "Dashboard profesional",
        "Configuración"
    ])

menu = st.sidebar.radio("Menú principal", opciones_menu)

st.markdown(
    """
    <div class="titulo-principal">
        <h1>🛡️ P.U.M.I. 2026</h1>
        <h3>Proceso Unificado para el Manejo de la Información</h3>
    </div>
    """,
    unsafe_allow_html=True
)


# ======================================================
# INICIO
# ======================================================

if menu == "Inicio":

    st.markdown(
        """
        <div class="card-pumi">
            <div class="subtitulo-pumi">Bienvenido al Sistema PUMI 2026</div>
            <div class="texto-pumi">
                Esta aplicación permite registrar, consultar y dar seguimiento a las
                actividades desarrolladas dentro del Proceso Unificado para el Manejo
                de la Información.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    df = cargar_datos()
    df_metricas = limpiar_dataframe_para_metricas(df)

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Registros", len(df_metricas))
    col2.metric("Programas", df_metricas["Programa"].nunique() if not df_metricas.empty else 0)
    col3.metric("Delegaciones", df_metricas["Delegación"].nunique() if not df_metricas.empty else 0)
    col4.metric(
        "Participantes",
        int(df_metricas["Cantidad Participantes"].sum()) if not df_metricas.empty else 0
    )

    st.success("Conexión con Google Sheets activa.")


# ======================================================
# REGISTRAR ACTIVIDAD
# ======================================================

elif menu == "Registrar actividad":

    st.markdown("## Registrar nueva actividad preventiva")

    st.markdown(
        """
        <div class="card-azul">
            <div class="texto-pumi">
                Complete la información de la actividad. Los campos se despliegan
                de forma dependiente según la base Datos Importantes.xlsx.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    if "latitud_registro" not in st.session_state:
        st.session_state.latitud_registro = ""

    if "longitud_registro" not in st.session_state:
        st.session_state.longitud_registro = ""

    # ======================================================
    # DATOS GENERALES
    # ======================================================

    st.markdown(
        """
        <div class="bloque-datos">
            <b>Datos generales del registro</b>
        </div>
        """,
        unsafe_allow_html=True
    )

    regiones_lista = obtener_regiones_datos()

    col1, col2 = st.columns(2)

    with col1:
        fecha_actividad = st.date_input(
            "Fecha de la actividad",
            value=date.today(),
            format="DD/MM/YYYY"
        )

        direccion_regional = st.selectbox(
            "Dirección Regional",
            regiones_lista if regiones_lista else REGIONES
        )

        delegaciones_filtradas = obtener_delegaciones_por_region(direccion_regional)

        delegacion = st.selectbox(
            "Delegación",
            delegaciones_filtradas if delegaciones_filtradas else ["Sin datos disponibles"]
        )

        actividades_filtradas = obtener_actividades_por_delegacion(delegacion)

        actividad = st.selectbox(
            "Actividad realizada",
            actividades_filtradas if actividades_filtradas else ["Sin datos disponibles"]
        )

    with col2:
        programas_filtrados = obtener_programas_por_actividad(actividad)

        programa = st.selectbox(
            "Programa",
            programas_filtrados if programas_filtrados else PROGRAMAS
        )

        responsable = st.text_input("Funcionario responsable")
        usuario = st.text_input("Usuario que registra")

        cantidad = st.number_input(
            "Cantidad de participantes",
            min_value=0,
            step=1
        )

    # ======================================================
    # UBICACIÓN TERRITORIAL
    # ======================================================

    st.markdown(
        """
        <div class="bloque-territorio">
            <b>Ubicación territorial</b>
        </div>
        """,
        unsafe_allow_html=True
    )

    provincias_lista = obtener_provincias_datos()

    col3, col4, col5 = st.columns(3)

    with col3:
        provincia = st.selectbox(
            "Provincia",
            provincias_lista if provincias_lista else PROVINCIAS
        )

    with col4:
        cantones_filtrados = obtener_cantones_por_provincia(provincia)

        canton = st.selectbox(
            "Cantón",
            cantones_filtrados if cantones_filtrados else ["Sin datos disponibles"]
        )

    with col5:
        distritos_filtrados = obtener_distritos_por_provincia_canton(provincia, canton)

        distrito = st.selectbox(
            "Distrito",
            distritos_filtrados if distritos_filtrados else ["Sin datos disponibles"]
        )

    # ======================================================
    # LUGAR
    # ======================================================

    st.markdown(
        """
        <div class="bloque-actividad">
            <b>Lugar de realización de la actividad</b>
        </div>
        """,
        unsafe_allow_html=True
    )

    tipo_lugar = st.radio(
        "Seleccione el tipo de lugar",
        ["Centro educativo", "Otro lugar"],
        horizontal=True
    )

    lugar = ""
    centro_educativo = ""

    if tipo_lugar == "Centro educativo":
        centros = obtener_centros_por_provincia(provincia)

        if centros:
            centro_educativo = st.selectbox(
                "Centro educativo según base MEP 2025",
                centros
            )
            lugar = centro_educativo
        else:
            st.warning("No se encontraron centros educativos para la provincia seleccionada.")
            centro_educativo = st.text_input("Digite el centro educativo manualmente")
            lugar = centro_educativo
    else:
        lugar = st.text_input("Lugar donde se realizó la actividad")
        centro_educativo = ""

    # ======================================================
    # MAPA
    # ======================================================

    st.markdown(
        """
        <div class="bloque-mapa">
            <b>Georreferencia del lugar</b>
        </div>
        """,
        unsafe_allow_html=True
    )

    tipo_mapa_registro = st.selectbox(
        "Tipo de mapa",
        ["OpenStreetMap", "Mapa claro", "Mapa oscuro", "Topográfico", "Satélite"],
        key="tipo_mapa_registro"
    )

    direccion_mapa = st.text_input(
        "Dirección o referencia para ubicar en mapa",
        value=f"{lugar}, {distrito}, {canton}, {provincia}, Costa Rica"
    )

    metodo_ubicacion = st.radio(
        "Método para registrar ubicación",
        [
            "Buscar por nombre del lugar",
            "Usar GPS del dispositivo",
            "Ingresar coordenadas manualmente",
            "Marcar punto en el mapa",
            "No registrar ubicación"
        ],
        horizontal=False
    )

    if metodo_ubicacion == "Buscar por nombre del lugar":
        if st.button("Buscar ubicación en el mapa"):
            lat_busqueda, lon_busqueda, direccion_encontrada = georreferenciar_direccion(direccion_mapa)

            if lat_busqueda and lon_busqueda:
                st.session_state.latitud_registro = str(lat_busqueda)
                st.session_state.longitud_registro = str(lon_busqueda)
                st.success("Ubicación encontrada automáticamente.")
                st.caption(direccion_encontrada)
            else:
                st.warning("No se logró ubicar el lugar automáticamente. Puede usar la opción Marcar punto en el mapa.")

    elif metodo_ubicacion == "Usar GPS del dispositivo":
        st.info("El navegador puede solicitar permiso para acceder a la ubicación del dispositivo.")

        ubicacion_gps = get_geolocation()

        if ubicacion_gps and "coords" in ubicacion_gps:
            lat_gps = ubicacion_gps["coords"].get("latitude", "")
            lon_gps = ubicacion_gps["coords"].get("longitude", "")

            if lat_gps and lon_gps:
                st.session_state.latitud_registro = str(lat_gps)
                st.session_state.longitud_registro = str(lon_gps)
                st.success("Ubicación GPS obtenida correctamente.")
            else:
                st.warning("No se recibieron coordenadas válidas desde el GPS.")
        else:
            st.warning("No se obtuvo ubicación GPS. Puede usar la opción Marcar punto en el mapa.")

    elif metodo_ubicacion == "Ingresar coordenadas manualmente":
        col_lat, col_lon = st.columns(2)

        with col_lat:
            st.session_state.latitud_registro = st.text_input(
                "Latitud",
                value=st.session_state.latitud_registro,
                placeholder="Ejemplo: 9.9281"
            )

        with col_lon:
            st.session_state.longitud_registro = st.text_input(
                "Longitud",
                value=st.session_state.longitud_registro,
                placeholder="Ejemplo: -84.0907"
            )

    elif metodo_ubicacion == "Marcar punto en el mapa":
        st.info("Haga clic sobre el mapa para seleccionar la ubicación de la actividad.")

    elif metodo_ubicacion == "No registrar ubicación":
        st.session_state.latitud_registro = ""
        st.session_state.longitud_registro = ""
        st.info("El registro se guardará sin coordenadas.")

    lat_num = limpiar_coordenada(st.session_state.latitud_registro)
    lon_num = limpiar_coordenada(st.session_state.longitud_registro)

    if lat_num is not None and lon_num is not None:
        centro_mapa = [lat_num, lon_num]
        zoom_mapa = 15
    else:
        centro_mapa = [9.7489, -83.7534]
        zoom_mapa = 7

    mapa_seleccion = crear_mapa_base(
        centro=centro_mapa,
        zoom=zoom_mapa,
        tipo_mapa=tipo_mapa_registro
    )

    if lat_num is not None and lon_num is not None:
        folium.Marker(
            location=[lat_num, lon_num],
            popup="Ubicación seleccionada",
            icon=folium.Icon(
                color=obtener_color_programa(programa),
                icon="info-sign"
            )
        ).add_to(mapa_seleccion)

    resultado_mapa = st_folium(
        mapa_seleccion,
        width=1100,
        height=560,
        key=f"mapa_registro_{tipo_mapa_registro}_{metodo_ubicacion}"
    )

    if resultado_mapa and resultado_mapa.get("last_clicked"):
        lat_click = resultado_mapa["last_clicked"]["lat"]
        lon_click = resultado_mapa["last_clicked"]["lng"]

        st.session_state.latitud_registro = str(lat_click)
        st.session_state.longitud_registro = str(lon_click)

        st.success("Punto seleccionado en el mapa.")
        st.write("Latitud:", st.session_state.latitud_registro)
        st.write("Longitud:", st.session_state.longitud_registro)

    # ======================================================
    # INFORMACIÓN COMPLEMENTARIA
    # ======================================================

    st.markdown(
        """
        <div class="bloque-datos">
            <b>Información complementaria</b>
        </div>
        """,
        unsafe_allow_html=True
    )

    instituciones = st.text_area("Instituciones participantes")
    plan = st.text_input("Plan estratégico relacionado")
    evidencia = st.text_input("Enlace de evidencia")
    observaciones = st.text_area("Observaciones")

    if st.button("Guardar registro"):

        if (
            not delegacion
            or delegacion == "Sin datos disponibles"
            or actividad == "Sin datos disponibles"
            or not responsable
        ):
            st.warning("Debe completar al menos Delegación, Actividad realizada y Funcionario responsable.")

        else:
            nuevo_id = generar_id_consecutivo()

            registro = {
                "ID": nuevo_id,
                "Fecha Registro": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "Fecha Actividad": fecha_actividad.strftime("%d/%m/%Y"),
                "Dirección Regional": direccion_regional,
                "Delegación": delegacion,
                "Programa": programa,
                "Actividad": actividad,
                "Provincia": provincia,
                "Cantón": canton,
                "Distrito": distrito,
                "Tipo Lugar": tipo_lugar,
                "Lugar": lugar,
                "Centro Educativo": centro_educativo,
                "Dirección Mapa": direccion_mapa,
                "Latitud": st.session_state.latitud_registro,
                "Longitud": st.session_state.longitud_registro,
                "Responsable": responsable,
                "Cantidad Participantes": cantidad,
                "Instituciones Participantes": instituciones,
                "Plan Estratégico Relacionado": plan,
                "Evidencia": evidencia,
                "Observaciones": observaciones,
                "Estado Revisión": "Pendiente de revisión",
                "Observación de Revisión": "",
                "Usuario Registra": usuario
            }

            guardar_registro(registro)

            st.session_state.latitud_registro = ""
            st.session_state.longitud_registro = ""

            st.success(f"Registro guardado correctamente con el ID #{nuevo_id}.")


# ======================================================
# SEGUIMIENTO DE REGISTROS
# VISTA PARA USUARIOS
# ======================================================

elif menu == "Seguimiento de registros":

    st.markdown("## Seguimiento de registros")

    df = cargar_datos()

    if df.empty:
        st.info("No existen registros para consultar.")

    else:
        df_seguimiento = limpiar_dataframe_para_metricas(df)

        st.markdown("### Filtros de seguimiento")

        col1, col2, col3 = st.columns(3)

        with col1:
            filtro_id = st.text_input("Buscar por ID")

            filtro_delegacion = st.selectbox(
                "Delegación",
                ["Todas"] + sorted(df_seguimiento["Delegación"].dropna().astype(str).unique().tolist())
            )

        with col2:
            filtro_programa = st.selectbox(
                "Programa",
                ["Todos"] + sorted(df_seguimiento["Programa"].dropna().astype(str).unique().tolist())
            )

            filtro_estado = st.selectbox(
                "Estado de revisión",
                ["Todos"] + sorted(df_seguimiento["Estado Revisión"].dropna().astype(str).unique().tolist())
            )

        with col3:
            filtro_provincia = st.selectbox(
                "Provincia",
                ["Todas"] + sorted(df_seguimiento["Provincia"].dropna().astype(str).unique().tolist())
            )

            usar_fechas_seg = st.checkbox("Filtrar por rango de fechas", key="fechas_seguimiento")

        df_usuario = df_seguimiento.copy()

        if filtro_id:
            df_usuario = df_usuario[
                df_usuario["ID"].astype(str).str.contains(filtro_id, case=False, na=False)
            ]

        if filtro_delegacion != "Todas":
            df_usuario = df_usuario[df_usuario["Delegación"] == filtro_delegacion]

        if filtro_programa != "Todos":
            df_usuario = df_usuario[df_usuario["Programa"] == filtro_programa]

        if filtro_estado != "Todos":
            df_usuario = df_usuario[df_usuario["Estado Revisión"] == filtro_estado]

        if filtro_provincia != "Todas":
            df_usuario = df_usuario[df_usuario["Provincia"] == filtro_provincia]

        if usar_fechas_seg:
            fechas_validas = df_usuario["Fecha Actividad"].dropna()

            if not fechas_validas.empty:
                fecha_min = fechas_validas.min().date()
                fecha_max = fechas_validas.max().date()

                colf1, colf2 = st.columns(2)

                with colf1:
                    fecha_inicio_seg = st.date_input(
                        "Fecha inicial",
                        value=fecha_min,
                        format="DD/MM/YYYY",
                        key="fecha_inicio_seg"
                    )

                with colf2:
                    fecha_fin_seg = st.date_input(
                        "Fecha final",
                        value=fecha_max,
                        format="DD/MM/YYYY",
                        key="fecha_fin_seg"
                    )

                df_usuario = df_usuario[
                    (df_usuario["Fecha Actividad"].dt.date >= fecha_inicio_seg) &
                    (df_usuario["Fecha Actividad"].dt.date <= fecha_fin_seg)
                ]

        st.markdown("### Resumen del seguimiento")

        colm1, colm2, colm3, colm4 = st.columns(4)

        colm1.metric("Registros", len(df_usuario))
        colm2.metric("Aprobados", len(df_usuario[df_usuario["Estado Revisión"] == "Aprobado"]))
        colm3.metric("Con observaciones", len(df_usuario[df_usuario["Estado Revisión"] == "Con observaciones"]))
        colm4.metric("Rechazados", len(df_usuario[df_usuario["Estado Revisión"] == "Rechazado"]))

        st.markdown("### Mapa general de actividades")

        mostrar_mapa_registros(
            df_usuario,
            height=560,
            key="mapa_seguimiento_usuarios"
        )

        st.markdown("### Estado de los registros")

        df_mostrar_usuario = df_usuario.copy()

        if "Fecha Actividad" in df_mostrar_usuario.columns:
            df_mostrar_usuario["Fecha Actividad"] = df_mostrar_usuario["Fecha Actividad"].dt.strftime("%d/%m/%Y")

        columnas_usuario = [
            "ID",
            "Fecha Actividad",
            "Delegación",
            "Programa",
            "Actividad",
            "Provincia",
            "Cantón",
            "Distrito",
            "Lugar",
            "Estado Revisión",
            "Observación de Revisión"
        ]

        columnas_existentes = [
            col for col in columnas_usuario
            if col in df_mostrar_usuario.columns
        ]

        st.dataframe(
            df_mostrar_usuario[columnas_existentes],
            use_container_width=True,
            hide_index=True
        )

        excel_seguimiento = convertir_excel(df_mostrar_usuario[columnas_existentes])

        st.download_button(
            label="Descargar seguimiento en Excel",
            data=excel_seguimiento,
            file_name="seguimiento_registros_pumi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
# ======================================================
# PARTE 4 DE 5
# CONSULTA, EDICIÓN ADMINISTRATIVA, REVISIÓN, MAPA Y ELIMINACIÓN
# ======================================================

elif menu == "Consulta / edición administrativa":

    st.markdown("## Consulta y edición administrativa de registros PUMI")

    if not st.session_state.admin_autenticado:
        st.error("Debe ingresar la clave administrativa para acceder a esta sección.")
        st.stop()

    df = cargar_datos()

    if df.empty:
        st.info("Aún no hay registros guardados.")

    else:
        df_consulta = limpiar_dataframe_para_metricas(df)

        st.markdown(
            """
            <div class="card-dorado">
                <div class="texto-pumi">
                    Esta sección permite consultar, editar, eliminar, revisar y corregir
                    la ubicación geográfica de los registros.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        with st.expander("Herramientas de limpieza de base de datos"):
            if st.button("Eliminar encabezados duplicados de la hoja"):
                eliminados = limpiar_encabezados_duplicados_en_sheet()
                st.success(f"Filas de encabezado duplicadas eliminadas: {eliminados}")
                st.rerun()

        st.markdown("### Filtros administrativos")

        col1, col2, col3 = st.columns(3)

        with col1:
            filtro_id_admin = st.text_input("Buscar por ID")

            filtro_programa = st.selectbox(
                "Filtrar por programa",
                ["Todos"] + sorted(df_consulta["Programa"].dropna().astype(str).unique().tolist())
            )

            filtro_delegacion = st.selectbox(
                "Filtrar por delegación",
                ["Todas"] + sorted(df_consulta["Delegación"].dropna().astype(str).unique().tolist())
            )

        with col2:
            filtro_region = st.selectbox(
                "Filtrar por región",
                ["Todas"] + sorted(df_consulta["Dirección Regional"].dropna().astype(str).unique().tolist())
            )

            filtro_provincia = st.selectbox(
                "Filtrar por provincia",
                ["Todas"] + sorted(df_consulta["Provincia"].dropna().astype(str).unique().tolist())
            )

            filtro_distrito = st.selectbox(
                "Filtrar por distrito",
                ["Todos"] + sorted(df_consulta["Distrito"].dropna().astype(str).unique().tolist())
            )

        with col3:
            filtro_estado = st.selectbox(
                "Filtrar por estado de revisión",
                ["Todos"] + sorted(df_consulta["Estado Revisión"].dropna().astype(str).unique().tolist())
            )

            filtro_tipo_lugar = st.selectbox(
                "Filtrar por tipo de lugar",
                ["Todos"] + sorted(df_consulta["Tipo Lugar"].dropna().astype(str).unique().tolist())
            )

            usar_fechas = st.checkbox("Filtrar por rango de fechas")

        df_filtrado = df_consulta.copy()

        if filtro_id_admin:
            df_filtrado = df_filtrado[
                df_filtrado["ID"].astype(str).str.contains(
                    filtro_id_admin,
                    case=False,
                    na=False
                )
            ]

        if filtro_programa != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Programa"] == filtro_programa]

        if filtro_delegacion != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Delegación"] == filtro_delegacion]

        if filtro_region != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Dirección Regional"] == filtro_region]

        if filtro_provincia != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Provincia"] == filtro_provincia]

        if filtro_distrito != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Distrito"] == filtro_distrito]

        if filtro_estado != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Estado Revisión"] == filtro_estado]

        if filtro_tipo_lugar != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Tipo Lugar"] == filtro_tipo_lugar]

        if usar_fechas:
            fechas_validas = df_filtrado["Fecha Actividad"].dropna()

            if fechas_validas.empty:
                st.warning("No hay fechas válidas para aplicar el filtro.")
            else:
                fecha_min = fechas_validas.min().date()
                fecha_max = fechas_validas.max().date()

                colf1, colf2 = st.columns(2)

                with colf1:
                    fecha_inicio = st.date_input(
                        "Fecha inicial",
                        value=fecha_min,
                        format="DD/MM/YYYY",
                        key="fecha_inicio_admin"
                    )

                with colf2:
                    fecha_fin = st.date_input(
                        "Fecha final",
                        value=fecha_max,
                        format="DD/MM/YYYY",
                        key="fecha_fin_admin"
                    )

                df_filtrado = df_filtrado[
                    (df_filtrado["Fecha Actividad"].dt.date >= fecha_inicio) &
                    (df_filtrado["Fecha Actividad"].dt.date <= fecha_fin)
                ]

        st.markdown("### Resumen administrativo")

        colm1, colm2, colm3, colm4 = st.columns(4)

        colm1.metric("Registros filtrados", len(df_filtrado))
        colm2.metric("Aprobados", len(df_filtrado[df_filtrado["Estado Revisión"] == "Aprobado"]))
        colm3.metric("Con observaciones", len(df_filtrado[df_filtrado["Estado Revisión"] == "Con observaciones"]))
        colm4.metric("Rechazados", len(df_filtrado[df_filtrado["Estado Revisión"] == "Rechazado"]))

        st.markdown("### Mapa administrativo de registros filtrados")

        mostrar_mapa_registros(
            df_filtrado,
            height=520,
            key="mapa_admin_registros"
        )

        st.markdown("### Registros encontrados")

        df_mostrar = df_filtrado.copy()

        if "Fecha Actividad" in df_mostrar.columns:
            df_mostrar["Fecha Actividad"] = df_mostrar["Fecha Actividad"].dt.strftime("%d/%m/%Y")

        st.dataframe(
            df_mostrar,
            use_container_width=True,
            hide_index=True
        )

        excel = convertir_excel(df_mostrar)

        st.download_button(
            label="Descargar registros administrativos en Excel",
            data=excel,
            file_name="registro_pumi_administrativo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("---")
        st.markdown("## Editar, revisar, georreferenciar o eliminar registro")

        if df_filtrado.empty:
            st.info("No hay registros filtrados para editar, revisar o eliminar.")

        else:
            ids = df_filtrado["ID"].astype(str).tolist()

            id_seleccionado = st.selectbox(
                "Seleccione el ID del registro",
                ids
            )

            registro_actual = df[
                df["ID"].astype(str) == str(id_seleccionado)
            ].iloc[0].to_dict()

            accion = st.radio(
                "Acción administrativa",
                [
                    "Editar registro completo",
                    "Actualizar revisión",
                    "Corregir ubicación en mapa",
                    "Eliminar registro"
                ],
                horizontal=True
            )

            # ==================================================
            # ACTUALIZAR REVISIÓN
            # ==================================================

            if accion == "Actualizar revisión":

                st.markdown(
                    """
                    <div class="bloque-actividad">
                        <b>Revisión administrativa del registro</b>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                with st.form("form_actualizar_revision"):

                    st.write(f"Registro seleccionado: ID #{registro_actual.get('ID', '')}")
                    st.write(f"Delegación: {registro_actual.get('Delegación', '')}")
                    st.write(f"Programa: {registro_actual.get('Programa', '')}")
                    st.write(f"Actividad: {registro_actual.get('Actividad', '')}")

                    estado_actual = registro_actual.get("Estado Revisión", "Pendiente de revisión")

                    nuevo_estado = st.selectbox(
                        "Estado de revisión",
                        ESTADOS_REVISION,
                        index=ESTADOS_REVISION.index(estado_actual)
                        if estado_actual in ESTADOS_REVISION else 0
                    )

                    observacion_revision = st.text_area(
                        "Observación de Revisión",
                        value=registro_actual.get("Observación de Revisión", ""),
                        help="Esta observación será visible para el usuario."
                    )

                    guardar_revision = st.form_submit_button("Guardar revisión")

                    if guardar_revision:
                        nuevos_datos = registro_actual.copy()
                        nuevos_datos["Estado Revisión"] = nuevo_estado
                        nuevos_datos["Observación de Revisión"] = observacion_revision

                        actualizado = actualizar_registro_por_id(
                            id_seleccionado,
                            nuevos_datos
                        )

                        if actualizado:
                            st.success("Revisión actualizada correctamente.")
                            st.rerun()
                        else:
                            st.error("No se encontró el registro para actualizar.")

            # ==================================================
            # CORREGIR UBICACIÓN EN MAPA
            # ==================================================

            elif accion == "Corregir ubicación en mapa":

                st.markdown(
                    """
                    <div class="bloque-mapa">
                        <b>Corrección de ubicación geográfica</b>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                if "lat_admin" not in st.session_state:
                    st.session_state.lat_admin = registro_actual.get("Latitud", "")

                if "lon_admin" not in st.session_state:
                    st.session_state.lon_admin = registro_actual.get("Longitud", "")

                tipo_mapa_admin = st.selectbox(
                    "Tipo de mapa",
                    ["OpenStreetMap", "Mapa claro", "Mapa oscuro", "Topográfico", "Satélite"],
                    key="tipo_mapa_admin_correccion"
                )

                direccion_actual = registro_actual.get("Dirección Mapa", "")

                direccion_mapa_admin = st.text_input(
                    "Dirección o referencia para ubicar en mapa",
                    value=direccion_actual if direccion_actual else f'{registro_actual.get("Lugar", "")}, {registro_actual.get("Distrito", "")}, {registro_actual.get("Cantón", "")}, {registro_actual.get("Provincia", "")}, Costa Rica'
                )

                metodo_admin = st.radio(
                    "Método para corregir ubicación",
                    [
                        "Buscar por nombre del lugar",
                        "Usar GPS del dispositivo",
                        "Ingresar coordenadas manualmente",
                        "Marcar punto en el mapa"
                    ],
                    horizontal=False
                )

                if metodo_admin == "Buscar por nombre del lugar":
                    if st.button("Buscar ubicación"):
                        lat_busqueda, lon_busqueda, direccion_encontrada = georreferenciar_direccion(direccion_mapa_admin)

                        if lat_busqueda and lon_busqueda:
                            st.session_state.lat_admin = str(lat_busqueda)
                            st.session_state.lon_admin = str(lon_busqueda)
                            st.session_state.dir_admin = direccion_encontrada
                            st.success("Ubicación encontrada.")
                            st.caption(direccion_encontrada)
                        else:
                            st.warning("No se logró ubicar el lugar automáticamente.")

                elif metodo_admin == "Usar GPS del dispositivo":
                    st.info("El navegador puede solicitar permiso para acceder a la ubicación.")

                    ubicacion_gps = get_geolocation()

                    if ubicacion_gps and "coords" in ubicacion_gps:
                        lat_gps = ubicacion_gps["coords"].get("latitude", "")
                        lon_gps = ubicacion_gps["coords"].get("longitude", "")

                        if lat_gps and lon_gps:
                            st.session_state.lat_admin = str(lat_gps)
                            st.session_state.lon_admin = str(lon_gps)
                            st.success("Ubicación GPS obtenida correctamente.")
                        else:
                            st.warning("No se recibieron coordenadas válidas desde el GPS.")
                    else:
                        st.warning("No se obtuvo ubicación GPS.")

                elif metodo_admin == "Ingresar coordenadas manualmente":
                    col_lat, col_lon = st.columns(2)

                    with col_lat:
                        st.session_state.lat_admin = st.text_input(
                            "Latitud",
                            value=str(st.session_state.lat_admin)
                        )

                    with col_lon:
                        st.session_state.lon_admin = st.text_input(
                            "Longitud",
                            value=str(st.session_state.lon_admin)
                        )

                elif metodo_admin == "Marcar punto en el mapa":
                    st.info("Haga clic sobre el mapa para seleccionar la ubicación correcta.")

                lat_num = limpiar_coordenada(st.session_state.lat_admin)
                lon_num = limpiar_coordenada(st.session_state.lon_admin)

                if lat_num is not None and lon_num is not None:
                    centro_mapa_admin = [lat_num, lon_num]
                    zoom_admin = 15
                else:
                    centro_mapa_admin = [9.7489, -83.7534]
                    zoom_admin = 7

                mapa_preview_admin = crear_mapa_base(
                    centro=centro_mapa_admin,
                    zoom=zoom_admin,
                    tipo_mapa=tipo_mapa_admin
                )

                if lat_num is not None and lon_num is not None:
                    folium.Marker(
                        location=[lat_num, lon_num],
                        popup=f"Registro PUMI #{registro_actual.get('ID', '')}",
                        icon=folium.Icon(
                            color=obtener_color_programa(registro_actual.get("Programa", "")),
                            icon="info-sign"
                        )
                    ).add_to(mapa_preview_admin)

                resultado_admin = st_folium(
                    mapa_preview_admin,
                    width=1100,
                    height=520,
                    key=f"mapa_correccion_admin_{tipo_mapa_admin}_{metodo_admin}"
                )

                if resultado_admin and resultado_admin.get("last_clicked"):
                    st.session_state.lat_admin = str(resultado_admin["last_clicked"]["lat"])
                    st.session_state.lon_admin = str(resultado_admin["last_clicked"]["lng"])

                    st.success("Punto seleccionado en el mapa.")
                    st.write("Latitud:", st.session_state.lat_admin)
                    st.write("Longitud:", st.session_state.lon_admin)

                if st.button("Guardar ubicación corregida"):

                    lat_final = limpiar_coordenada(st.session_state.lat_admin)
                    lon_final = limpiar_coordenada(st.session_state.lon_admin)

                    if lat_final is None or lon_final is None:
                        st.warning("Debe indicar coordenadas válidas antes de guardar.")
                    else:
                        nuevos_datos = registro_actual.copy()
                        nuevos_datos["Dirección Mapa"] = direccion_mapa_admin
                        nuevos_datos["Latitud"] = str(lat_final)
                        nuevos_datos["Longitud"] = str(lon_final)

                        actualizado = actualizar_registro_por_id(
                            id_seleccionado,
                            nuevos_datos
                        )

                        if actualizado:
                            st.success("Ubicación actualizada correctamente.")
                            st.rerun()
                        else:
                            st.error("No se encontró el registro para actualizar.")

            # ==================================================
            # EDITAR REGISTRO COMPLETO
            # ==================================================

            elif accion == "Editar registro completo":

                st.markdown(
                    """
                    <div class="bloque-datos">
                        <b>Datos generales del registro</b>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                fecha_registro = st.text_input(
                    "Fecha Registro",
                    registro_actual.get("Fecha Registro", "")
                )

                fecha_actividad_actual = pd.to_datetime(
                    registro_actual.get("Fecha Actividad", ""),
                    errors="coerce",
                    dayfirst=True
                )

                if pd.isna(fecha_actividad_actual):
                    fecha_actividad_actual = date.today()
                else:
                    fecha_actividad_actual = fecha_actividad_actual.date()

                col1, col2 = st.columns(2)

                with col1:
                    fecha_actividad_editada = st.date_input(
                        "Fecha de la actividad",
                        value=fecha_actividad_actual,
                        format="DD/MM/YYYY"
                    )

                    regiones_lista = obtener_regiones_datos()

                    region_actual = registro_actual.get("Dirección Regional", "")
                    opciones_region = regiones_lista if regiones_lista else REGIONES

                    if region_actual and region_actual not in opciones_region:
                        opciones_region = [region_actual] + opciones_region

                    nueva_region = st.selectbox(
                        "Dirección Regional",
                        opciones_region,
                        index=opciones_region.index(region_actual)
                        if region_actual in opciones_region else 0
                    )

                    delegaciones_filtradas = obtener_delegaciones_por_region(nueva_region)

                    delegacion_actual = registro_actual.get("Delegación", "")
                    opciones_delegacion = delegaciones_filtradas if delegaciones_filtradas else obtener_delegaciones_unicas()

                    if delegacion_actual and delegacion_actual not in opciones_delegacion:
                        opciones_delegacion = [delegacion_actual] + opciones_delegacion

                    nueva_delegacion = st.selectbox(
                        "Delegación",
                        opciones_delegacion if opciones_delegacion else ["Sin datos disponibles"],
                        index=opciones_delegacion.index(delegacion_actual)
                        if delegacion_actual in opciones_delegacion else 0
                    )

                    actividades_filtradas = obtener_actividades_por_delegacion(nueva_delegacion)

                    actividad_actual = registro_actual.get("Actividad", "")
                    opciones_actividad = actividades_filtradas if actividades_filtradas else [actividad_actual]

                    if actividad_actual and actividad_actual not in opciones_actividad:
                        opciones_actividad = [actividad_actual] + opciones_actividad

                    nueva_actividad = st.selectbox(
                        "Actividad realizada",
                        opciones_actividad if opciones_actividad else ["Sin datos disponibles"],
                        index=opciones_actividad.index(actividad_actual)
                        if actividad_actual in opciones_actividad else 0
                    )

                with col2:
                    programas_filtrados = obtener_programas_por_actividad(nueva_actividad)

                    programa_actual = registro_actual.get("Programa", "")
                    opciones_programa = programas_filtrados if programas_filtrados else PROGRAMAS

                    if programa_actual and programa_actual not in opciones_programa:
                        opciones_programa = [programa_actual] + opciones_programa

                    nuevo_programa = st.selectbox(
                        "Programa",
                        opciones_programa,
                        index=opciones_programa.index(programa_actual)
                        if programa_actual in opciones_programa else 0
                    )

                    nuevo_responsable = st.text_input(
                        "Responsable",
                        registro_actual.get("Responsable", "")
                    )

                    nuevo_usuario = st.text_input(
                        "Usuario Registra",
                        registro_actual.get("Usuario Registra", "")
                    )

                    cantidad_actual = pd.to_numeric(
                        registro_actual.get("Cantidad Participantes", 0),
                        errors="coerce"
                    )

                    if pd.isna(cantidad_actual):
                        cantidad_actual = 0

                    nueva_cantidad = st.number_input(
                        "Cantidad Participantes",
                        min_value=0,
                        step=1,
                        value=int(cantidad_actual)
                    )

                st.markdown(
                    """
                    <div class="bloque-territorio">
                        <b>Ubicación territorial</b>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                provincias_lista = obtener_provincias_datos()

                provincia_actual = registro_actual.get("Provincia", "")
                opciones_provincia = provincias_lista if provincias_lista else PROVINCIAS

                if provincia_actual and provincia_actual not in opciones_provincia:
                    opciones_provincia = [provincia_actual] + opciones_provincia

                col3, col4, col5 = st.columns(3)

                with col3:
                    nueva_provincia = st.selectbox(
                        "Provincia",
                        opciones_provincia,
                        index=opciones_provincia.index(provincia_actual)
                        if provincia_actual in opciones_provincia else 0
                    )

                with col4:
                    cantones_filtrados = obtener_cantones_por_provincia(nueva_provincia)

                    canton_actual = registro_actual.get("Cantón", "")
                    opciones_canton = cantones_filtrados if cantones_filtrados else [canton_actual]

                    if canton_actual and canton_actual not in opciones_canton:
                        opciones_canton = [canton_actual] + opciones_canton

                    nuevo_canton = st.selectbox(
                        "Cantón",
                        opciones_canton if opciones_canton else ["Sin datos disponibles"],
                        index=opciones_canton.index(canton_actual)
                        if canton_actual in opciones_canton else 0
                    )

                with col5:
                    distritos_filtrados = obtener_distritos_por_provincia_canton(
                        nueva_provincia,
                        nuevo_canton
                    )

                    distrito_actual = registro_actual.get("Distrito", "")
                    opciones_distrito = distritos_filtrados if distritos_filtrados else [distrito_actual]

                    if distrito_actual and distrito_actual not in opciones_distrito:
                        opciones_distrito = [distrito_actual] + opciones_distrito

                    nuevo_distrito = st.selectbox(
                        "Distrito",
                        opciones_distrito if opciones_distrito else ["Sin datos disponibles"],
                        index=opciones_distrito.index(distrito_actual)
                        if distrito_actual in opciones_distrito else 0
                    )

                st.markdown(
                    """
                    <div class="bloque-actividad">
                        <b>Lugar de realización de la actividad</b>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                tipo_actual = registro_actual.get("Tipo Lugar", "Otro lugar")

                if tipo_actual not in ["Centro educativo", "Otro lugar"]:
                    tipo_actual = "Otro lugar"

                tipo_lugar_editado = st.radio(
                    "Seleccione el tipo de lugar",
                    ["Centro educativo", "Otro lugar"],
                    index=["Centro educativo", "Otro lugar"].index(tipo_actual),
                    horizontal=True
                )

                lugar_editado = ""
                centro_editado = ""

                if tipo_lugar_editado == "Centro educativo":

                    centros = obtener_centros_por_provincia(nueva_provincia)
                    centro_actual = registro_actual.get("Centro Educativo", "")

                    if centros:
                        opciones_centros = centros

                        if centro_actual and centro_actual not in opciones_centros:
                            opciones_centros = [centro_actual] + opciones_centros

                        centro_editado = st.selectbox(
                            "Centro educativo según base MEP 2025",
                            opciones_centros,
                            index=opciones_centros.index(centro_actual)
                            if centro_actual in opciones_centros else 0
                        )
                    else:
                        centro_editado = st.text_input(
                            "Digite el centro educativo manualmente",
                            centro_actual
                        )

                    lugar_editado = centro_editado

                else:
                    lugar_editado = st.text_input(
                        "Lugar donde se realizó la actividad",
                        registro_actual.get("Lugar", "")
                    )
                    centro_editado = ""

                st.markdown(
                    """
                    <div class="bloque-mapa">
                        <b>Datos de georreferencia</b>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                nueva_direccion_mapa = st.text_input(
                    "Dirección Mapa",
                    registro_actual.get("Dirección Mapa", "")
                )

                nueva_latitud = st.text_input(
                    "Latitud",
                    registro_actual.get("Latitud", "")
                )

                nueva_longitud = st.text_input(
                    "Longitud",
                    registro_actual.get("Longitud", "")
                )

                st.markdown(
                    """
                    <div class="bloque-datos">
                        <b>Información complementaria y revisión</b>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                nuevas_instituciones = st.text_area(
                    "Instituciones Participantes",
                    registro_actual.get("Instituciones Participantes", "")
                )

                nuevo_plan = st.text_input(
                    "Plan Estratégico Relacionado",
                    registro_actual.get("Plan Estratégico Relacionado", "")
                )

                nueva_evidencia = st.text_input(
                    "Evidencia",
                    registro_actual.get("Evidencia", "")
                )

                nuevas_observaciones = st.text_area(
                    "Observaciones del registro",
                    registro_actual.get("Observaciones", "")
                )

                nuevo_estado = st.selectbox(
                    "Estado Revisión",
                    ESTADOS_REVISION,
                    index=ESTADOS_REVISION.index(registro_actual.get("Estado Revisión", ""))
                    if registro_actual.get("Estado Revisión", "") in ESTADOS_REVISION else 0
                )

                nueva_observacion_revision = st.text_area(
                    "Observación de Revisión",
                    registro_actual.get("Observación de Revisión", "")
                )

                if st.button("Guardar cambios completos"):

                    nuevos_datos = {
                        "ID": registro_actual.get("ID", ""),
                        "Fecha Registro": fecha_registro,
                        "Fecha Actividad": fecha_actividad_editada.strftime("%d/%m/%Y"),
                        "Dirección Regional": nueva_region,
                        "Delegación": nueva_delegacion,
                        "Programa": nuevo_programa,
                        "Actividad": nueva_actividad,
                        "Provincia": nueva_provincia,
                        "Cantón": nuevo_canton,
                        "Distrito": nuevo_distrito,
                        "Tipo Lugar": tipo_lugar_editado,
                        "Lugar": lugar_editado,
                        "Centro Educativo": centro_editado,
                        "Dirección Mapa": nueva_direccion_mapa,
                        "Latitud": nueva_latitud,
                        "Longitud": nueva_longitud,
                        "Responsable": nuevo_responsable,
                        "Cantidad Participantes": nueva_cantidad,
                        "Instituciones Participantes": nuevas_instituciones,
                        "Plan Estratégico Relacionado": nuevo_plan,
                        "Evidencia": nueva_evidencia,
                        "Observaciones": nuevas_observaciones,
                        "Estado Revisión": nuevo_estado,
                        "Observación de Revisión": nueva_observacion_revision,
                        "Usuario Registra": nuevo_usuario
                    }

                    actualizado = actualizar_registro_por_id(
                        id_seleccionado,
                        nuevos_datos
                    )

                    if actualizado:
                        st.success("Registro actualizado correctamente.")
                        st.rerun()
                    else:
                        st.error("No se encontró el registro para actualizar.")

            # ==================================================
            # ELIMINAR REGISTRO
            # ==================================================

            else:
                st.warning("Esta acción eliminará el registro seleccionado de forma permanente.")

                confirmar = st.checkbox(
                    "Confirmo que deseo eliminar este registro"
                )

                if st.button("Eliminar registro"):
                    if confirmar:
                        eliminado = eliminar_registro_por_id(id_seleccionado)

                        if eliminado:
                            st.success("Registro eliminado correctamente.")
                            st.rerun()
                        else:
                            st.error("No se encontró el registro para eliminar.")
                    else:
                        st.warning("Debe marcar la confirmación antes de eliminar.")
# ======================================================
# PARTE 5 DE 5
# DASHBOARD PROFESIONAL, MAPA ADMINISTRATIVO, INFORME PDF Y CONFIGURACIÓN
# ======================================================

elif menu == "Dashboard profesional":

    st.markdown("## Dashboard profesional PUMI 2026")

    if not st.session_state.admin_autenticado:
        st.error("Debe ingresar la clave administrativa para acceder a esta sección.")
        st.stop()

    df = cargar_datos()

    if df.empty:
        st.info("Aún no hay datos para generar dashboard.")

    else:
        df_dash = limpiar_dataframe_para_metricas(df)

        st.markdown(
            """
            <div class="card-azul">
                <div class="texto-pumi">
                    Panel administrativo para analizar registros por programa, delegación,
                    provincia, estado de revisión, participantes, comportamiento mensual
                    y ubicación geográfica de las actividades registradas.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        # ======================================================
        # FILTROS DASHBOARD
        # ======================================================

        st.markdown("### Filtros del dashboard")

        col1, col2, col3 = st.columns(3)

        with col1:
            filtro_programa_dash = st.selectbox(
                "Programa",
                ["Todos"] + sorted(df_dash["Programa"].dropna().astype(str).unique().tolist()),
                key="dash_programa"
            )

            filtro_delegacion_dash = st.selectbox(
                "Delegación",
                ["Todas"] + sorted(df_dash["Delegación"].dropna().astype(str).unique().tolist()),
                key="dash_delegacion"
            )

        with col2:
            filtro_region_dash = st.selectbox(
                "Región",
                ["Todas"] + sorted(df_dash["Dirección Regional"].dropna().astype(str).unique().tolist()),
                key="dash_region"
            )

            filtro_provincia_dash = st.selectbox(
                "Provincia",
                ["Todas"] + sorted(df_dash["Provincia"].dropna().astype(str).unique().tolist()),
                key="dash_provincia"
            )

        with col3:
            filtro_estado_dash = st.selectbox(
                "Estado de revisión",
                ["Todos"] + sorted(df_dash["Estado Revisión"].dropna().astype(str).unique().tolist()),
                key="dash_estado"
            )

            usar_fechas_dash = st.checkbox(
                "Filtrar por rango de fechas",
                key="dash_fechas"
            )

        df_filtrado = df_dash.copy()

        if filtro_programa_dash != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Programa"] == filtro_programa_dash]

        if filtro_delegacion_dash != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Delegación"] == filtro_delegacion_dash]

        if filtro_region_dash != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Dirección Regional"] == filtro_region_dash]

        if filtro_provincia_dash != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Provincia"] == filtro_provincia_dash]

        if filtro_estado_dash != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Estado Revisión"] == filtro_estado_dash]

        if usar_fechas_dash:
            fechas_validas = df_filtrado["Fecha Actividad"].dropna()

            if fechas_validas.empty:
                st.warning("No hay fechas válidas para aplicar el filtro.")
            else:
                fecha_min = fechas_validas.min().date()
                fecha_max = fechas_validas.max().date()

                colf1, colf2 = st.columns(2)

                with colf1:
                    fecha_inicio_dash = st.date_input(
                        "Fecha inicial",
                        value=fecha_min,
                        format="DD/MM/YYYY",
                        key="fecha_inicio_dash"
                    )

                with colf2:
                    fecha_fin_dash = st.date_input(
                        "Fecha final",
                        value=fecha_max,
                        format="DD/MM/YYYY",
                        key="fecha_fin_dash"
                    )

                df_filtrado = df_filtrado[
                    (df_filtrado["Fecha Actividad"].dt.date >= fecha_inicio_dash) &
                    (df_filtrado["Fecha Actividad"].dt.date <= fecha_fin_dash)
                ]

        # ======================================================
        # MÉTRICAS
        # ======================================================

        st.markdown("### Indicadores generales")

        colm1, colm2, colm3, colm4 = st.columns(4)

        colm1.metric("Registros", len(df_filtrado))
        colm2.metric("Programas", df_filtrado["Programa"].nunique())
        colm3.metric("Delegaciones", df_filtrado["Delegación"].nunique())
        colm4.metric("Participantes", int(df_filtrado["Cantidad Participantes"].sum()))

        colm5, colm6, colm7, colm8 = st.columns(4)

        colm5.metric("Aprobados", len(df_filtrado[df_filtrado["Estado Revisión"] == "Aprobado"]))
        colm6.metric("Pendientes", len(df_filtrado[df_filtrado["Estado Revisión"] == "Pendiente de revisión"]))
        colm7.metric("Con observaciones", len(df_filtrado[df_filtrado["Estado Revisión"] == "Con observaciones"]))
        colm8.metric("Rechazados", len(df_filtrado[df_filtrado["Estado Revisión"] == "Rechazado"]))

        st.markdown("---")

        if df_filtrado.empty:
            st.warning("No hay datos con los filtros seleccionados.")

        else:

            # ======================================================
            # MAPA GENERAL
            # ======================================================

            st.markdown("### Mapa general de actividades filtradas")

            mostrar_mapa_registros(
                df_filtrado,
                height=560,
                key="mapa_dashboard_admin"
            )

            st.markdown("---")

            # ======================================================
            # GRÁFICOS
            # ======================================================

            colg1, colg2 = st.columns(2)

            with colg1:
                st.markdown("### Actividades por programa")

                data_programa = df_filtrado.groupby("Programa", as_index=False)["ID"].count()

                fig_programa = px.bar(
                    data_programa,
                    x="Programa",
                    y="ID",
                    text="ID",
                    labels={"ID": "Cantidad de actividades"},
                    color="Programa",
                    template="plotly_white"
                )

                fig_programa.update_layout(showlegend=False)
                st.plotly_chart(fig_programa, use_container_width=True)

            with colg2:
                st.markdown("### Participantes por programa")

                data_participantes = df_filtrado.groupby(
                    "Programa",
                    as_index=False
                )["Cantidad Participantes"].sum()

                fig_participantes = px.bar(
                    data_participantes,
                    x="Programa",
                    y="Cantidad Participantes",
                    text="Cantidad Participantes",
                    color="Programa",
                    template="plotly_white"
                )

                fig_participantes.update_layout(showlegend=False)
                st.plotly_chart(fig_participantes, use_container_width=True)

            colg3, colg4 = st.columns(2)

            with colg3:
                st.markdown("### Estado de revisión")

                fig_estado = px.pie(
                    df_filtrado,
                    names="Estado Revisión",
                    hole=0.45,
                    template="plotly_white"
                )

                st.plotly_chart(fig_estado, use_container_width=True)

            with colg4:
                st.markdown("### Actividades por provincia")

                data_provincia = df_filtrado.groupby("Provincia", as_index=False)["ID"].count()

                fig_provincia = px.bar(
                    data_provincia,
                    x="Provincia",
                    y="ID",
                    text="ID",
                    labels={"ID": "Cantidad de actividades"},
                    color="Provincia",
                    template="plotly_white"
                )

                fig_provincia.update_layout(showlegend=False)
                st.plotly_chart(fig_provincia, use_container_width=True)

            colg5, colg6 = st.columns(2)

            with colg5:
                st.markdown("### Ranking de delegaciones")

                data_delegacion = df_filtrado.groupby(
                    "Delegación",
                    as_index=False
                )["ID"].count().sort_values("ID", ascending=False).head(15)

                fig_delegacion = px.bar(
                    data_delegacion,
                    x="ID",
                    y="Delegación",
                    orientation="h",
                    text="ID",
                    labels={"ID": "Cantidad de actividades"},
                    color="ID",
                    template="plotly_white"
                )

                st.plotly_chart(fig_delegacion, use_container_width=True)

            with colg6:
                st.markdown("### Registros por tipo de lugar")

                if "Tipo Lugar" in df_filtrado.columns:
                    fig_tipo_lugar = px.pie(
                        df_filtrado,
                        names="Tipo Lugar",
                        hole=0.35,
                        template="plotly_white"
                    )

                    st.plotly_chart(fig_tipo_lugar, use_container_width=True)
                else:
                    st.info("No existe la columna Tipo Lugar.")

            st.markdown("### Evolución mensual de actividades")

            df_mensual = df_filtrado.dropna(subset=["Fecha Actividad"]).copy()

            if not df_mensual.empty:
                df_mensual["Mes"] = df_mensual["Fecha Actividad"].dt.to_period("M").astype(str)

                data_mensual = df_mensual.groupby("Mes", as_index=False)["ID"].count()

                fig_mensual = px.line(
                    data_mensual,
                    x="Mes",
                    y="ID",
                    markers=True,
                    text="ID",
                    labels={"ID": "Cantidad de actividades"},
                    template="plotly_white"
                )

                st.plotly_chart(fig_mensual, use_container_width=True)
            else:
                st.info("No hay fechas válidas para generar evolución mensual.")

            # ======================================================
            # TABLA DE REGISTROS
            # ======================================================

            st.markdown("### Registros analizados")

            df_mostrar = df_filtrado.copy()

            if "Fecha Actividad" in df_mostrar.columns:
                df_mostrar["Fecha Actividad"] = df_mostrar["Fecha Actividad"].dt.strftime("%d/%m/%Y")

            st.dataframe(
                df_mostrar,
                use_container_width=True,
                hide_index=True
            )

            # ======================================================
            # INFORME PDF
            # ======================================================

            def generar_pdf_profesional(df_pdf):
                buffer = BytesIO()

                doc = SimpleDocTemplate(
                    buffer,
                    pagesize=letter,
                    rightMargin=40,
                    leftMargin=40,
                    topMargin=40,
                    bottomMargin=40
                )

                elementos = []
                estilos = getSampleStyleSheet()

                titulo = ParagraphStyle(
                    "TituloPUMI",
                    parent=estilos["Title"],
                    alignment=TA_CENTER,
                    textColor=colors.HexColor(COLOR_AZUL),
                    fontSize=22,
                    spaceAfter=20
                )

                subtitulo = ParagraphStyle(
                    "SubtituloPUMI",
                    parent=estilos["Heading2"],
                    textColor=colors.HexColor(COLOR_VERDE),
                    fontSize=15,
                    spaceAfter=10
                )

                texto = ParagraphStyle(
                    "TextoPUMI",
                    parent=estilos["Normal"],
                    alignment=TA_JUSTIFY,
                    fontSize=10,
                    leading=14
                )

                logo_path = None

                for posible_logo in ["logo_pumi.jpeg", "logo_pumi.jpg", "logo_pumi.png"]:
                    if os.path.exists(posible_logo):
                        logo_path = posible_logo
                        break

                if logo_path:
                    elementos.append(RLImage(logo_path, width=120, height=120))

                elementos.append(Paragraph("INFORME ADMINISTRATIVO PUMI 2026", titulo))
                elementos.append(Paragraph("Proceso Unificado para el Manejo de la Información", subtitulo))
                elementos.append(Spacer(1, 12))

                introduccion = """
                El presente informe consolida la información registrada en el Sistema PUMI 2026,
                correspondiente a las actividades preventivas desarrolladas por los Programas
                Policiales Preventivos. La información permite observar cantidad de registros,
                programas atendidos, delegaciones participantes, población alcanzada, estado de
                revisión, observaciones administrativas y ubicación geográfica registrada.
                """

                elementos.append(Paragraph(introduccion, texto))
                elementos.append(Spacer(1, 18))

                total_registros = len(df_pdf)
                total_programas = df_pdf["Programa"].nunique()
                total_delegaciones = df_pdf["Delegación"].nunique()
                total_participantes = int(df_pdf["Cantidad Participantes"].sum())
                total_aprobados = len(df_pdf[df_pdf["Estado Revisión"] == "Aprobado"])
                total_pendientes = len(df_pdf[df_pdf["Estado Revisión"] == "Pendiente de revisión"])
                total_observaciones = len(df_pdf[df_pdf["Estado Revisión"] == "Con observaciones"])
                total_rechazados = len(df_pdf[df_pdf["Estado Revisión"] == "Rechazado"])
                total_georreferenciados = len(preparar_dataframe_mapa(df_pdf))

                resumen = [
                    ["Indicador", "Resultado"],
                    ["Total de registros", total_registros],
                    ["Programas registrados", total_programas],
                    ["Delegaciones registradas", total_delegaciones],
                    ["Participantes atendidos", total_participantes],
                    ["Registros georreferenciados", total_georreferenciados],
                    ["Registros aprobados", total_aprobados],
                    ["Pendientes de revisión", total_pendientes],
                    ["Con observaciones", total_observaciones],
                    ["Rechazados", total_rechazados]
                ]

                tabla_resumen = Table(resumen, colWidths=[260, 180])
                tabla_resumen.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(COLOR_AZUL)),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("ALIGN", (1, 1), (-1, -1), "CENTER")
                ]))

                elementos.append(Paragraph("Resumen general", subtitulo))
                elementos.append(tabla_resumen)
                elementos.append(Spacer(1, 18))

                elementos.append(Paragraph("Distribución por programa", subtitulo))

                tabla_programas = [["Programa", "Actividades", "Participantes"]]

                resumen_programa = df_pdf.groupby("Programa").agg({
                    "ID": "count",
                    "Cantidad Participantes": "sum"
                }).reset_index()

                for _, row in resumen_programa.iterrows():
                    tabla_programas.append([
                        row["Programa"],
                        int(row["ID"]),
                        int(row["Cantidad Participantes"])
                    ])

                tabla_prog = Table(tabla_programas, colWidths=[220, 110, 130])
                tabla_prog.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(COLOR_VERDE)),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("ALIGN", (1, 1), (-1, -1), "CENTER")
                ]))

                elementos.append(tabla_prog)
                elementos.append(Spacer(1, 18))

                elementos.append(Paragraph("Detalle de registros", subtitulo))

                columnas_detalle = [
                    "ID",
                    "Fecha Actividad",
                    "Delegación",
                    "Programa",
                    "Provincia",
                    "Distrito",
                    "Lugar",
                    "Cantidad Participantes",
                    "Estado Revisión",
                    "Observación de Revisión"
                ]

                columnas_detalle = [
                    col for col in columnas_detalle
                    if col in df_pdf.columns
                ]

                detalle = [columnas_detalle]

                df_detalle = df_pdf.copy()

                if "Fecha Actividad" in df_detalle.columns:
                    df_detalle["Fecha Actividad"] = df_detalle["Fecha Actividad"].dt.strftime("%d/%m/%Y")

                for _, row in df_detalle[columnas_detalle].head(30).iterrows():
                    detalle.append([str(row[col]) for col in columnas_detalle])

                tabla_detalle = Table(detalle, repeatRows=1)
                tabla_detalle.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(COLOR_AZUL_CLARO)),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 6)
                ]))

                elementos.append(tabla_detalle)
                elementos.append(Spacer(1, 15))

                cierre = """
                Informe generado automáticamente por el Sistema PUMI 2026.
                La información presentada corresponde a los registros disponibles
                en la base de datos conectada a Google Sheets al momento de la descarga.
                """

                elementos.append(Paragraph(cierre, texto))

                doc.build(elementos)
                buffer.seek(0)

                return buffer.getvalue()

            pdf = generar_pdf_profesional(df_filtrado)

            st.download_button(
                label="Descargar informe administrativo en PDF",
                data=pdf,
                file_name="informe_administrativo_pumi_2026.pdf",
                mime="application/pdf"
            )


# ======================================================
# CONFIGURACIÓN
# ======================================================

elif menu == "Configuración":

    st.markdown("## Configuración del sistema")

    if not st.session_state.admin_autenticado:
        st.error("Debe ingresar la clave administrativa para acceder a esta sección.")
        st.stop()

    st.markdown(
        """
        <div class="card-pumi">
            <div class="texto-pumi">
                En este apartado se verifica la conexión con Google Sheets,
                la existencia de la hoja principal, la estructura de encabezados,
                la disponibilidad de las bases auxiliares y los campos de georreferencia.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    try:
        spreadsheet = conectar_google_sheets()
        hoja = inicializar_hoja()

        st.success("Conexión exitosa con Google Sheets.")
        st.write("Archivo conectado:", spreadsheet.title)
        st.write("Hoja principal:", HOJA_REGISTRO)
        st.write("Total de columnas esperadas:", len(ENCABEZADOS))

        st.markdown("### Bases auxiliares")

        df_mep = cargar_base_mep()
        df_delegaciones = cargar_base_delegaciones()

        st.write("Registros base MEP:", len(df_mep))
        st.write("Registros base delegaciones/distritos:", len(df_delegaciones))

        st.markdown("### Georreferencia")

        df_config = cargar_datos()
        df_mapa_config = preparar_dataframe_mapa(df_config)

        st.write("Registros totales:", len(df_config))
        st.write("Registros con coordenadas válidas:", len(df_mapa_config))
        st.write("Registros sin coordenadas:", len(df_config) - len(df_mapa_config))

        st.markdown("### Limpieza de encabezados duplicados")

        if st.button("Eliminar encabezados duplicados"):
            eliminados = limpiar_encabezados_duplicados_en_sheet()
            st.success(f"Filas eliminadas: {eliminados}")
            st.rerun()

        st.markdown("### Vista previa de la base de datos")

        st.dataframe(
            df_config.head(10),
            use_container_width=True,
            hide_index=True
        )

    except Exception as e:
        st.error("Error en la conexión o configuración.")
        st.exception(e)
