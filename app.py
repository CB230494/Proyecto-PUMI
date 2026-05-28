# ======================================================
# PARTE 1 DE 12
# CONFIGURACIÓN GENERAL, LIBRERÍAS, CLAVES, CATÁLOGOS Y ESTILOS
# ======================================================

import streamlit as st
import pandas as pd
import gspread
import plotly.express as px
import unicodedata
import os
import base64

import folium
from streamlit_folium import st_folium
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from streamlit_js_eval import get_geolocation

from google.oauth2.service_account import Credentials
from datetime import datetime, date, time
from io import BytesIO
from PIL import Image

from reportlab.lib.pagesizes import letter
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
    PageBreak,
    KeepTogether
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY


# ======================================================
# CONFIGURACIÓN GENERAL DE STREAMLIT
# ======================================================

st.set_page_config(
    page_title="PUMI 2026",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ======================================================
# ARCHIVOS BASE
# ======================================================

ARCHIVO_MEP = "BASE DE DATOS MEP 2025.xlsx"
ARCHIVO_DELEGACIONES = "DELEGACIONES Y DISTRITOS.xlsx"
ARCHIVO_DATOS_IMPORTANTES = "Datos Importantes.xlsx"


# ======================================================
# LOGOS INSTITUCIONALES
# ======================================================

LOGO_MINISTERIO = "Logo2.jpeg"
LOGO_PUMI = "logo_pumi.jpeg"
LOGO_FUERZA_PUBLICA = "Logo1.jpeg"


# ======================================================
# CONEXIÓN GOOGLE SHEETS
# ======================================================

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1O-HNa1c4ppF0-ND7BUqKA2OzZDc1O0kTxZOMVDeA-GY/edit?gid=0#gid=0"
HOJA_REGISTRO = "REGISTRO_PUMI_2026"


# ======================================================
# CLAVES ADMINISTRATIVAS
# ======================================================

CLAVES_ADMINISTRATIVAS = {
    "DPPP23": {"perfil": "DPPP", "programa": "TODOS", "programas_permitidos": "TODOS", "descripcion": "Administrador general"},
    "DPPP2026": {"perfil": "DPPP", "programa": "TODOS", "programas_permitidos": "TODOS", "descripcion": "Administrador general"},

    "DARE23": {"perfil": "DARE", "programa": "DARE", "programas_permitidos": ["DARE"], "descripcion": "Administrador programa DARE"},
    "DARE2026": {"perfil": "DARE", "programa": "DARE", "programas_permitidos": ["DARE"], "descripcion": "Administrador programa DARE"},

    "GREAT23": {"perfil": "GREAT", "programa": "GREAT", "programas_permitidos": ["GREAT", "GREAT CAMP"], "descripcion": "Administrador programa GREAT"},
    "GREAT2026": {"perfil": "GREAT", "programa": "GREAT", "programas_permitidos": ["GREAT", "GREAT CAMP"], "descripcion": "Administrador programa GREAT"},

    "MPAS23": {"perfil": "MPAS", "programa": "MPAS", "programas_permitidos": ["MPAS"], "descripcion": "Administrador programa MPAS"},
    "MPAS2026": {"perfil": "MPAS", "programa": "MPAS", "programas_permitidos": ["MPAS"], "descripcion": "Administrador programa MPAS"},

    "PSCC23": {"perfil": "PSCC", "programa": "PSCC", "programas_permitidos": ["PSCC"], "descripcion": "Administrador programa PSCC"},
    "PSCC2026": {"perfil": "PSCC", "programa": "PSCC", "programas_permitidos": ["PSCC"], "descripcion": "Administrador programa PSCC"},

    "VIF23": {"perfil": "VIF", "programa": "VIF", "programas_permitidos": ["VIF"], "descripcion": "Administrador programa VIF"},
    "VIF2026": {"perfil": "VIF", "programa": "VIF", "programas_permitidos": ["VIF"], "descripcion": "Administrador programa VIF"},

    "POLITICA23": {"perfil": "POLITICA PUBLICA", "programa": "Política Pública", "programas_permitidos": ["Política Pública"], "descripcion": "Administrador programa Política Pública"},
    "POLITICA2026": {"perfil": "POLITICA PUBLICA", "programa": "Política Pública", "programas_permitidos": ["Política Pública"], "descripcion": "Administrador programa Política Pública"}
}


# ======================================================
# CLAVES DE CONSULTA PARA USUARIOS POR PROGRAMA
# ======================================================

CLAVES_USUARIO_PROGRAMA = {
    "DPPP": {"claves": ["DPPP23", "DPPP2026"], "programas_permitidos": "TODOS"},
    "DARE": {"claves": ["DARE23", "DARE2026", "DAREUSUARIO"], "programas_permitidos": ["DARE"]},
    "GREAT": {"claves": ["GREAT23", "GREAT2026", "GREATUSUARIO"], "programas_permitidos": ["GREAT", "GREAT CAMP"]},
    "MPAS": {"claves": ["MPAS23", "MPAS2026", "MPASUSUARIO"], "programas_permitidos": ["MPAS"]},
    "PSCC": {"claves": ["PSCC23", "PSCC2026", "PSCCUSUARIO"], "programas_permitidos": ["PSCC"]},
    "VIF": {"claves": ["VIF23", "VIF2026", "VIFUSUARIO"], "programas_permitidos": ["VIF"]},
    "Política Pública": {"claves": ["POLITICA23", "POLITICA2026", "POLITICAUSUARIO"], "programas_permitidos": ["Política Pública"]}
}


# ======================================================
# ENCABEZADOS OFICIALES
# ORDEN ACTUALIZADO:
# Dirección Regional → Delegación → Responde a → Programa → Actividad
# ======================================================

ENCABEZADOS = [
    "ID", "Fecha Registro", "Fecha Actividad", "Hora Actividad",
    "Dirección Regional", "Delegación", "Responde a", "Programa", "Actividad",
    "Provincia", "Cantón", "Distrito", "Tipo Lugar", "Lugar",
    "Centro Educativo", "Código Presupuestario", "Dirección Mapa",
    "Latitud", "Longitud", "Responsable", "Cantidad Participantes",
    "Cantidad Hombres", "Cantidad Mujeres", "Edad 10 a 18",
    "Edad 19 a 30", "Edad 31 a 45", "Edad 46 en adelante",
    "Instituciones Participantes", "Plan Estratégico Relacionado",
    "Número de Referencia", "Número de Expediente Referencia",
    "Observaciones", "Estado Revisión", "Observación de Revisión",
    "Usuario Registra"
]


# ======================================================
# CATÁLOGOS BASE
# ======================================================

PROGRAMAS = ["DARE", "GREAT", "MPAS", "PSCC", "VIF", "Política Pública"]

REGIONES = [
    "R1 San José Central", "R2 San José Norte", "R3 San José Sur",
    "R4 Alajuela", "R5 Cartago", "R6 Heredia", "R7 Chorotega",
    "R8 Puntarenas", "R9 Limón", "R10 Brunca",
    "R11 Chorotega Norte", "R12"
]

PROVINCIAS = ["San José", "Alajuela", "Cartago", "Heredia", "Guanacaste", "Puntarenas", "Limón"]

RANGOS_EDAD = ["Edad 10 a 18", "Edad 19 a 30", "Edad 31 a 45", "Edad 46 en adelante"]

ESTADOS_REVISION = ["Pendiente de revisión", "Aprobado", "Con observaciones", "Rechazado"]


# ======================================================
# COLORES INSTITUCIONALES
# ======================================================

COLOR_AZUL = "#002B7F"
COLOR_AZUL_MEDIO = "#1E4FA3"
COLOR_AZUL_CLARO = "#DCE8FF"

COLOR_DORADO = "#B88A2A"
COLOR_DORADO_OSCURO = "#8A6418"
COLOR_DORADO_CLARO = "#F4E6C1"

COLOR_BLANCO = "#FFFFFF"
COLOR_GRIS = "#F4F6F8"
COLOR_GRIS_OSCURO = "#2F3542"

COLOR_ROJO = COLOR_DORADO
COLOR_ROJO_OSCURO = COLOR_DORADO_OSCURO
COLOR_ROJO_CLARO = COLOR_DORADO_CLARO

COLOR_VERDE = COLOR_DORADO
COLOR_VERDE_CLARO = COLOR_DORADO_CLARO


COLORES_PROGRAMA = {
    "DARE": "purple",
    "GREAT": "orange",
    "GREAT CAMP": "orange",
    "MPAS": "darkblue",
    "PSCC": "blue",
    "VIF": "orange",
    "Política Pública": "green"
}


# ======================================================
# FUNCIÓN PARA CONVERTIR IMÁGENES A BASE64
# ======================================================

def imagen_a_base64(ruta):
    if not os.path.exists(ruta):
        return ""

    try:
        with open(ruta, "rb") as archivo:
            return base64.b64encode(archivo.read()).decode()
    except Exception:
        return ""


# ======================================================
# ESTILOS CSS
# ======================================================

st.markdown(
    f"""
    <style>

    .stApp {{
        background:
            radial-gradient(circle at top left, #FFFFFF 0, transparent 28%),
            radial-gradient(circle at top right, #DCE8FF 0, transparent 30%),
            linear-gradient(180deg, #FFFFFF 0%, #F4F7FB 48%, #FFFFFF 100%);
    }}

    section[data-testid="stSidebar"] {{
        background: linear-gradient(180deg, #FFFFFF 0%, #EAF1FF 45%, #FFFFFF 100%);
        border-right: 5px solid {COLOR_DORADO};
        box-shadow: 4px 0px 12px rgba(0,0,0,0.10);
    }}

    .sidebar-logo-pumi {{
        background: #FFFFFF;
        border-radius: 16px;
        padding: 12px;
        margin: 8px 0px 18px 0px;
        box-shadow: 0px 4px 14px rgba(0,0,0,0.10);
        text-align: center;
    }}

    .sidebar-logo-pumi img {{
        width: 100%;
        max-height: 145px;
        object-fit: contain;
    }}

    .encabezado-institucional {{
        background: #FFFFFF;
        border-radius: 22px;
        padding: 22px 34px;
        margin-bottom: 24px;
        box-shadow: 0px 8px 22px rgba(0,0,0,0.11);
        border-bottom: 7px solid {COLOR_DORADO};
        display: grid;
        grid-template-columns: 1fr 1fr 1fr;
        align-items: center;
        gap: 22px;
        min-height: 155px;
    }}

    .encabezado-logo-izq,
    .encabezado-logo-centro,
    .encabezado-logo-der {{
        height: 125px;
        display: flex;
        align-items: center;
    }}

    .encabezado-logo-izq {{
        justify-content: flex-start;
    }}

    .encabezado-logo-centro {{
        justify-content: center;
    }}

    .encabezado-logo-der {{
        justify-content: flex-end;
    }}

    .logo-ministerio-app {{
        width: 350px;
        height: 100px;
        object-fit: contain;
    }}

    .logo-pumi-app {{
        width: 300px;
        height: 100px;
        object-fit: contain;
    }}

    .logo-fp-app {{
        width: 400px;
        height: 150px;
        object-fit: contain;
    }}

    .titulo-principal {{
        background: linear-gradient(135deg, {COLOR_AZUL}, #243B63, {COLOR_DORADO});
        padding: 38px;
        border-radius: 22px;
        text-align: center;
        color: white;
        margin-bottom: 30px;
        box-shadow: 0px 8px 22px rgba(0,0,0,0.28);
        border: 3px solid {COLOR_BLANCO};
    }}

    .titulo-principal h1 {{
        font-size: 52px;
        margin-bottom: 8px;
        font-weight: 900;
        letter-spacing: 1.5px;
        color: white;
        text-shadow: 1px 2px 5px rgba(0,0,0,0.35);
    }}

    .titulo-principal h3 {{
        font-size: 27px;
        margin-top: 8px;
        font-weight: 600;
        color: white;
    }}

    .card-pumi, .card-azul, .card-dorado {{
        padding: 24px;
        border-radius: 20px;
        box-shadow: 0px 6px 16px rgba(0,0,0,0.13);
        margin-bottom: 20px;
        background: linear-gradient(135deg, #FFFFFF 0%, #F8FAFF 100%);
    }}

    .card-pumi {{
        border-left: 9px solid {COLOR_AZUL};
    }}

    .card-azul {{
        border-left: 9px solid {COLOR_DORADO};
    }}

    .card-dorado {{
        border-left: 9px solid {COLOR_DORADO};
    }}

    .bloque-datos {{
        background: linear-gradient(135deg, #FFFFFF 0%, #EEF4FF 100%);
        padding: 18px;
        border-radius: 18px;
        border-left: 7px solid {COLOR_AZUL};
        margin-bottom: 18px;
    }}

    .bloque-territorio {{
        background: linear-gradient(135deg, #FFFFFF 0%, #FFF8E8 100%);
        padding: 18px;
        border-radius: 18px;
        border-left: 7px solid {COLOR_DORADO};
        margin-bottom: 18px;
    }}

    .bloque-actividad {{
        background: linear-gradient(135deg, #FFFFFF 0%, #EEF4FF 100%);
        padding: 18px;
        border-radius: 18px;
        border-left: 7px solid {COLOR_AZUL};
        margin-bottom: 18px;
    }}

    .bloque-mapa {{
        background: linear-gradient(135deg, #FFFFFF 0%, #FFF8E8 100%);
        padding: 18px;
        border-radius: 18px;
        border-left: 7px solid {COLOR_DORADO};
        margin-bottom: 18px;
    }}

    .subtitulo-pumi {{
        color: {COLOR_AZUL};
        font-weight: 900;
        font-size: 30px;
        margin-bottom: 12px;
        text-align: center;
    }}

    .texto-pumi {{
        color: {COLOR_GRIS_OSCURO};
        font-size: 17px;
        line-height: 1.7;
        text-align: justify;
    }}

    div[data-testid="stMetric"] {{
        background: linear-gradient(145deg, #FFFFFF 0%, #F3F6FF 100%);
        padding: 20px;
        border-radius: 20px;
        box-shadow: 0px 5px 15px rgba(0,0,0,0.13);
        border-bottom: 6px solid {COLOR_DORADO};
    }}

    div[data-baseweb="select"] > div {{
        background-color: #FFFFFF !important;
        border-radius: 12px !important;
        border: 1.8px solid #AFC3EE !important;
    }}

    input {{
        background-color: #FFFFFF !important;
        border-radius: 12px !important;
        border: 1.8px solid #AFC3EE !important;
    }}

    textarea {{
        background-color: #FFFFFF !important;
        border-radius: 12px !important;
        border: 1.8px solid {COLOR_DORADO_CLARO} !important;
    }}

    .stButton > button {{
        background: linear-gradient(90deg, {COLOR_AZUL}, {COLOR_DORADO});
        color: white;
        border-radius: 12px;
        border: none;
        font-weight: 800;
        padding: 0.7rem 1.2rem;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.18);
    }}

    .stDownloadButton > button {{
        background: linear-gradient(90deg, {COLOR_DORADO}, {COLOR_AZUL});
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
        width: 100% !important;
    }}

    hr {{
        border: none;
        height: 3px;
        background: linear-gradient(90deg, {COLOR_AZUL}, {COLOR_BLANCO}, {COLOR_DORADO});
        margin-top: 25px;
        margin-bottom: 25px;
    }}

    h1, h2, h3 {{
        color: {COLOR_AZUL};
        font-weight: 900;
    }}

    @media (max-width: 900px) {{
        .encabezado-institucional {{
            grid-template-columns: 1fr;
            text-align: center;
            min-height: auto;
        }}

        .encabezado-logo-izq,
        .encabezado-logo-centro,
        .encabezado-logo-der {{
            justify-content: center;
            height: auto;
        }}

        .logo-ministerio-app {{
            width: 300px;
            height: 80px;
        }}

        .logo-pumi-app {{
            width: 250px;
            height: 80px;
        }}

        .logo-fp-app {{
            width: 350px;
            height: 100px;
        }}
    }}

    </style>
    """,
    unsafe_allow_html=True
)


# ======================================================
# LOGO EN SIDEBAR
# SOLO PUMI
# ======================================================

def mostrar_logo():
    logo_pumi_b64 = imagen_a_base64(LOGO_PUMI)

    if logo_pumi_b64:
        st.sidebar.markdown(
            f"""
            <div class="sidebar-logo-pumi">
                <img src="data:image/jpeg;base64,{logo_pumi_b64}">
            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        st.sidebar.warning("Logo PUMI no encontrado.")


# ======================================================
# ENCABEZADO INSTITUCIONAL SUPERIOR
# IZQUIERDA: MINISTERIO | CENTRO: PUMI | DERECHA: FUERZA PÚBLICA
# ======================================================

def mostrar_encabezado_institucional():
    logo_min_b64 = imagen_a_base64(LOGO_MINISTERIO)
    logo_pumi_b64 = imagen_a_base64(LOGO_PUMI)
    logo_fp_b64 = imagen_a_base64(LOGO_FUERZA_PUBLICA)

    img_min = (
        f'<img class="logo-ministerio-app" src="data:image/jpeg;base64,{logo_min_b64}">'
        if logo_min_b64 else ""
    )

    img_pumi = (
        f'<img class="logo-pumi-app" src="data:image/jpeg;base64,{logo_pumi_b64}">'
        if logo_pumi_b64 else ""
    )

    img_fp = (
        f'<img class="logo-fp-app" src="data:image/jpeg;base64,{logo_fp_b64}">'
        if logo_fp_b64 else ""
    )

    st.markdown(
        f"""
        <div class="encabezado-institucional">
            <div class="encabezado-logo-izq">
                {img_min}
            </div>
            <div class="encabezado-logo-centro">
                {img_pumi}
            </div>
            <div class="encabezado-logo-der">
                {img_fp}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ======================================================
# COMPATIBILIDAD
# ======================================================

def mostrar_logos_encabezado():
    mostrar_encabezado_institucional()
# ======================================================
# PARTE 2 DE 12
# CONEXIÓN GOOGLE SHEETS, FUNCIONES BASE Y DATOS IMPORTANTES
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
            hoja_nueva = spreadsheet.add_worksheet(
                title=nombre_hoja,
                rows=5000,
                cols=len(ENCABEZADOS)
            )
            hoja_nueva.append_row(ENCABEZADOS)
            return hoja_nueva

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

    encabezados_actuales = datos[0]

    if encabezados_actuales == ENCABEZADOS:
        return hoja

    encabezados_normalizados = [
        normalizar_texto(col) for col in encabezados_actuales
    ]

    encabezados_oficiales_normalizados = [
        normalizar_texto(col) for col in ENCABEZADOS
    ]

    # Si la hoja ya tiene encabezados, pero le falta "Responde a",
    # se actualiza únicamente la fila 1 con el orden oficial.
    # No borra datos existentes.
    if "ID" in encabezados_actuales:
        try:
            hoja.update(
                values=[ENCABEZADOS],
                range_name="A1"
            )
        except Exception as e:
            st.error("No se pudieron actualizar los encabezados oficiales.")
            st.exception(e)
            st.stop()

    # Si la primera fila no parece encabezado, inserta encabezados arriba.
    elif encabezados_normalizados != encabezados_oficiales_normalizados:
        try:
            hoja.insert_row(ENCABEZADOS, index=1)
        except Exception as e:
            st.error("No se pudieron insertar los encabezados oficiales.")
            st.exception(e)
            st.stop()

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


def ordenar_regiones_numericamente(lista_regiones):
    def extraer_numero_region(region):
        texto = str(region)
        numero = ""

        for caracter in texto:
            if caracter.isdigit():
                numero += caracter
            elif numero:
                break

        if numero:
            return int(numero)

        return 999

    return sorted(lista_regiones, key=extraer_numero_region)


def porcentaje(valor, total):
    try:
        valor = float(valor)
        total = float(total)

        if total == 0:
            return "0.0%"

        return f"{(valor / total) * 100:.1f}%"

    except Exception:
        return "0.0%"


def aplicar_filtro_perfil_admin(df):
    if df.empty:
        return df

    if "admin_perfil" not in st.session_state:
        return df

    perfil = st.session_state.get("admin_perfil", {})
    programas_permitidos = perfil.get("programas_permitidos", "TODOS")

    if programas_permitidos == "TODOS":
        return df

    if "Programa" not in df.columns:
        return df

    return df[df["Programa"].isin(programas_permitidos)]


# ======================================================
# FUNCIONES PARA RESPONDE A
# Maneja celdas divididas por barra inclinada /
# ======================================================

def separar_responde_a(valor):
    if pd.isna(valor):
        return []

    texto = str(valor).strip()

    if texto == "" or texto.lower() == "nan":
        return []

    partes = texto.split("/")

    return [
        parte.strip()
        for parte in partes
        if parte.strip() != ""
    ]


def fila_responde_a_contiene(valor_celda, responde_a_seleccionado):
    opciones = separar_responde_a(valor_celda)

    opciones_norm = [
        normalizar_texto(opcion)
        for opcion in opciones
    ]

    return normalizar_texto(responde_a_seleccionado) in opciones_norm


# ======================================================
# BASE DATOS IMPORTANTES
# Provincia - Cantón - Distrito
# Dirección Regional - Delegación
# Responde a - Programa - Actividad Realizada
# ======================================================

@st.cache_data
def cargar_datos_importantes():
    columnas_base = [
        "Provincia",
        "Cantón",
        "Distrito",
        "Dirección Regional",
        "Delegación",
        "Responde a",
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

        col_responde_a = obtener_columna_por_nombre(
            df_original,
            ["Responde a", "Responde a:", "Responde", "Responda a", "Responda a:"]
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
            col_responde_a,
            col_actividad,
            col_programa
        ]

        if not all(columnas_requeridas):
            st.warning(
                "El archivo Datos Importantes.xlsx no tiene todas las columnas requeridas. "
                "Debe incluir: Provincia, Cantón, Distrito, Dirección Regional, Delegación, "
                "Responde a, Actividad Realizada y Programa."
            )
            return pd.DataFrame(columns=columnas_base)

        df = df_original[
            [
                col_provincia,
                col_canton,
                col_distrito,
                col_region,
                col_delegacion,
                col_responde_a,
                col_actividad,
                col_programa
            ]
        ].copy()

        df.columns = columnas_base

        for col in columnas_base:
            df[col] = df[col].fillna("").astype(str).str.strip()

        df = df.replace("nan", "")
        df = df.drop_duplicates()

        df["Provincia_Normalizada"] = df["Provincia"].apply(normalizar_texto)
        df["Cantón_Normalizado"] = df["Cantón"].apply(normalizar_texto)
        df["Distrito_Normalizado"] = df["Distrito"].apply(normalizar_texto)
        df["Región_Normalizada"] = df["Dirección Regional"].apply(normalizar_texto)
        df["Delegación_Normalizada"] = df["Delegación"].apply(normalizar_texto)
        df["Responde_a_Normalizado"] = df["Responde a"].apply(normalizar_texto)
        df["Actividad_Normalizada"] = df["Actividad Realizada"].apply(normalizar_texto)
        df["Programa_Normalizado"] = df["Programa"].apply(normalizar_texto)

        # Se crea una versión expandida para los casos donde Responde a viene con "/".
        # Ejemplo: Plan A / Plan B queda como dos filas lógicas.
        filas_expandidas = []

        for _, fila in df.iterrows():
            opciones_responde = separar_responde_a(fila["Responde a"])

            if not opciones_responde:
                opciones_responde = [fila["Responde a"]]

            for opcion in opciones_responde:
                nueva_fila = fila.copy()
                nueva_fila["Responde a"] = opcion
                nueva_fila["Responde_a_Normalizado"] = normalizar_texto(opcion)
                filas_expandidas.append(nueva_fila)

        if filas_expandidas:
            df = pd.DataFrame(filas_expandidas)
            df = df.drop_duplicates()

        return df

    except Exception as e:
        st.error(f"Error leyendo {ARCHIVO_DATOS_IMPORTANTES}")
        st.exception(e)
        return pd.DataFrame(columns=columnas_base)


def obtener_regiones_datos():
    df = cargar_datos_importantes()

    if df.empty:
        return ordenar_regiones_numericamente(REGIONES)

    regiones = df["Dirección Regional"].dropna().unique().tolist()
    regiones = [x for x in regiones if str(x).strip() != ""]

    return ordenar_regiones_numericamente(regiones)


def obtener_delegaciones_por_region(region):
    df = cargar_datos_importantes()

    if df.empty or not region:
        return obtener_delegaciones_unicas()

    region_norm = normalizar_texto(region)

    delegaciones = df[
        df["Región_Normalizada"] == region_norm
    ]["Delegación"].dropna().unique().tolist()

    delegaciones = [x for x in delegaciones if str(x).strip() != ""]

    return sorted(delegaciones)


# ======================================================
# NUEVA CASCADA:
# Responde a → Programa → Actividad realizada
# ======================================================

def obtener_responde_a_datos():
    df = cargar_datos_importantes()

    if df.empty or "Responde a" not in df.columns:
        return []

    responde_a = df["Responde a"].dropna().unique().tolist()
    responde_a = [x for x in responde_a if str(x).strip() != ""]

    return sorted(responde_a)


def obtener_programas_por_responde_a(responde_a):
    df = cargar_datos_importantes()

    if df.empty or not responde_a:
        return obtener_programas_datos()

    responde_norm = normalizar_texto(responde_a)

    programas = df[
        df["Responde_a_Normalizado"] == responde_norm
    ]["Programa"].dropna().unique().tolist()

    programas = [x for x in programas if str(x).strip() != ""]

    if not programas:
        return []

    programas_limpios = []

    for programa in programas:
        if normalizar_texto(programa) == "GREAT CAMP":
            if "GREAT" not in programas_limpios:
                programas_limpios.append("GREAT")
        elif programa not in programas_limpios:
            programas_limpios.append(programa)

    orden_oficial = [p for p in PROGRAMAS if p in programas_limpios]
    extras = sorted([p for p in programas_limpios if p not in orden_oficial])

    return orden_oficial + extras


def obtener_actividades_por_responde_a_programa(responde_a, programa):
    df = cargar_datos_importantes()

    if df.empty or not responde_a or not programa:
        return []

    responde_norm = normalizar_texto(responde_a)
    programa_norm = normalizar_texto(programa)

    if programa_norm == "GREAT":
        actividades = df[
            (df["Responde_a_Normalizado"] == responde_norm) &
            (df["Programa_Normalizado"].isin(["GREAT", "GREAT CAMP"]))
        ]["Actividad Realizada"].dropna().unique().tolist()
    else:
        actividades = df[
            (df["Responde_a_Normalizado"] == responde_norm) &
            (df["Programa_Normalizado"] == programa_norm)
        ]["Actividad Realizada"].dropna().unique().tolist()

    actividades = [x for x in actividades if str(x).strip() != ""]

    return sorted(actividades)


# ======================================================
# FUNCIONES ANTERIORES CONSERVADAS COMO RESPALDO
# ======================================================

def obtener_programas_datos():
    df = cargar_datos_importantes()

    if df.empty:
        return PROGRAMAS

    programas = df["Programa"].dropna().unique().tolist()
    programas = [x for x in programas if str(x).strip() != ""]

    if not programas:
        return PROGRAMAS

    programas_limpios = []

    for programa in programas:
        if normalizar_texto(programa) == "GREAT CAMP":
            if "GREAT" not in programas_limpios:
                programas_limpios.append("GREAT")
        elif programa not in programas_limpios:
            programas_limpios.append(programa)

    orden_oficial = [p for p in PROGRAMAS if p in programas_limpios]
    extras = sorted([p for p in programas_limpios if p not in orden_oficial])

    return orden_oficial + extras


def obtener_actividades_por_programa(programa):
    df = cargar_datos_importantes()

    if df.empty or not programa:
        return []

    programa_norm = normalizar_texto(programa)

    if programa_norm == "GREAT":
        actividades = df[
            df["Programa_Normalizado"].isin(["GREAT", "GREAT CAMP"])
        ]["Actividad Realizada"].dropna().unique().tolist()
    else:
        actividades = df[
            df["Programa_Normalizado"] == programa_norm
        ]["Actividad Realizada"].dropna().unique().tolist()

    actividades = [x for x in actividades if str(x).strip() != ""]

    return sorted(actividades)


def obtener_provincias_datos():
    df = cargar_datos_importantes()

    if df.empty:
        return PROVINCIAS

    provincias = df["Provincia"].dropna().unique().tolist()
    provincias = [x for x in provincias if str(x).strip() != ""]

    return sorted(provincias)


def obtener_cantones_por_provincia(provincia):
    df = cargar_datos_importantes()

    if df.empty or not provincia:
        return []

    provincia_norm = normalizar_texto(provincia)

    cantones = df[
        df["Provincia_Normalizada"] == provincia_norm
    ]["Cantón"].dropna().unique().tolist()

    cantones = [x for x in cantones if str(x).strip() != ""]

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

    distritos = [x for x in distritos if str(x).strip() != ""]

    return sorted(distritos)
# ======================================================
# PARTE 3 DE 12
# BASE MEP, BASE DELEGACIONES, MAPA INTELIGENTE Y CRUD PRINCIPAL
# ======================================================

# ======================================================
# BASE MEP
# Lee la base de centros educativos y obtiene:
# Provincia, nombre del centro educativo y código presupuestario.
# ======================================================

@st.cache_data
def cargar_base_mep():
    columnas_base = [
        "PROVINCIA",
        "NOMBRE",
        "CODIGO_PRESUPUESTARIO"
    ]

    if not os.path.exists(ARCHIVO_MEP):
        return pd.DataFrame(columns=columnas_base)

    try:
        df_original = pd.read_excel(ARCHIVO_MEP)

        col_provincia = obtener_columna_por_nombre(
            df_original,
            ["PROVINCIA", "Provincia"]
        )

        col_nombre = obtener_columna_por_nombre(
            df_original,
            [
                "NOMBRE",
                "Nombre",
                "CENTRO EDUCATIVO",
                "Centro Educativo",
                "INSTITUCION",
                "Institución"
            ]
        )

        col_codigo = obtener_columna_por_nombre(
            df_original,
            [
                "CODIGO PRESUPUESTARIO",
                "CÓDIGO PRESUPUESTARIO",
                "Codigo Presupuestario",
                "Código Presupuestario",
                "CODIGO",
                "CÓDIGO"
            ]
        )

        if not col_provincia or not col_nombre:
            return pd.DataFrame(columns=columnas_base)

        if not col_codigo:
            df_original["CODIGO_PRESUPUESTARIO_TMP"] = ""
            col_codigo = "CODIGO_PRESUPUESTARIO_TMP"

        df = df_original[
            [
                col_provincia,
                col_nombre,
                col_codigo
            ]
        ].copy()

        df.columns = columnas_base

        for col in columnas_base:
            df[col] = df[col].fillna("").astype(str).str.strip()

        df = df.dropna(subset=["PROVINCIA", "NOMBRE"])

        df["PROVINCIA_NORMALIZADA"] = df["PROVINCIA"].apply(normalizar_texto)
        df["NOMBRE_NORMALIZADO"] = df["NOMBRE"].apply(normalizar_texto)

        df["CENTRO_MOSTRAR"] = df.apply(
            lambda row: f"{row['NOMBRE']} - Código: {row['CODIGO_PRESUPUESTARIO']}"
            if row["CODIGO_PRESUPUESTARIO"].strip() != ""
            else row["NOMBRE"],
            axis=1
        )

        df = df.drop_duplicates(
            subset=[
                "PROVINCIA_NORMALIZADA",
                "NOMBRE_NORMALIZADO",
                "CODIGO_PRESUPUESTARIO"
            ]
        )

        return df

    except Exception as e:
        st.error(f"Error leyendo {ARCHIVO_MEP}")
        st.exception(e)
        return pd.DataFrame(columns=columnas_base)


def obtener_centros_por_provincia(provincia):
    df = cargar_base_mep()

    if df.empty or not provincia:
        return []

    provincia_norm = normalizar_texto(provincia)

    centros = df[
        df["PROVINCIA_NORMALIZADA"] == provincia_norm
    ]["CENTRO_MOSTRAR"].dropna().astype(str).drop_duplicates().sort_values().tolist()

    return centros


def obtener_datos_centro_educativo(centro_mostrar):
    df = cargar_base_mep()

    if df.empty or not centro_mostrar:
        return "", ""

    fila = df[df["CENTRO_MOSTRAR"] == centro_mostrar]

    if fila.empty:
        return centro_mostrar, ""

    nombre = fila.iloc[0].get("NOMBRE", "")
    codigo = fila.iloc[0].get("CODIGO_PRESUPUESTARIO", "")

    return nombre, codigo


# ======================================================
# BASE DELEGACIONES Y DISTRITOS
# Se conserva como respaldo.
# La lógica principal de provincia, cantón y distrito
# sale de Datos Importantes.xlsx.
# ======================================================

@st.cache_data
def cargar_base_delegaciones():
    if not os.path.exists(ARCHIVO_DELEGACIONES):
        return pd.DataFrame(columns=["Delegacion", "Distrito"])

    try:
        df_original = pd.read_excel(ARCHIVO_DELEGACIONES)

        col_delegacion = obtener_columna_por_nombre(
            df_original,
            ["Delegacion", "Delegación"]
        )

        col_distrito = obtener_columna_por_nombre(
            df_original,
            ["Distrito"]
        )

        col_provincia = obtener_columna_por_nombre(
            df_original,
            ["Provincia"]
        )

        col_canton = obtener_columna_por_nombre(
            df_original,
            ["Canton", "Cantón"]
        )

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
            return pd.DataFrame(
                columns=[
                    "Delegacion",
                    "Distrito",
                    "Provincia",
                    "Cantón"
                ]
            )

        df = df_original[list(columnas.keys())].copy()
        df = df.rename(columns=columnas)

        for col in ["Delegacion", "Distrito", "Provincia", "Cantón"]:
            if col not in df.columns:
                df[col] = ""

            df[col] = df[col].fillna("").astype(str).str.strip()

        df = df.replace("nan", "")

        df["Delegacion_Normalizada"] = df["Delegacion"].apply(normalizar_texto)
        df["Distrito_Normalizado"] = df["Distrito"].apply(normalizar_texto)
        df["Provincia_Normalizada"] = df["Provincia"].apply(normalizar_texto)
        df["Cantón_Normalizado"] = df["Cantón"].apply(normalizar_texto)

        df = df.drop_duplicates()

        return df

    except Exception as e:
        st.error(f"Error leyendo {ARCHIVO_DELEGACIONES}")
        st.exception(e)
        return pd.DataFrame(columns=["Delegacion", "Distrito"])


def obtener_delegaciones_unicas():
    df_datos = cargar_datos_importantes()

    if not df_datos.empty:
        delegaciones = df_datos["Delegación"].dropna().unique().tolist()
        delegaciones = [x for x in delegaciones if str(x).strip() != ""]
        return sorted(delegaciones)

    df = cargar_base_delegaciones()

    if df.empty or "Delegacion" not in df.columns:
        return []

    df_tmp = df[
        ["Delegacion", "Delegacion_Normalizada"]
    ].drop_duplicates(
        subset=["Delegacion_Normalizada"]
    )

    return sorted(
        [
            x for x in df_tmp["Delegacion"].dropna().tolist()
            if str(x).strip() != ""
        ]
    )


def obtener_distritos_unicos():
    df_datos = cargar_datos_importantes()

    if not df_datos.empty:
        distritos = df_datos["Distrito"].dropna().unique().tolist()
        distritos = [x for x in distritos if str(x).strip() != ""]
        return sorted(distritos)

    df = cargar_base_delegaciones()

    if df.empty:
        return []

    df_tmp = df[
        ["Distrito", "Distrito_Normalizado"]
    ].drop_duplicates(
        subset=["Distrito_Normalizado"]
    )

    return sorted(
        [
            x for x in df_tmp["Distrito"].dropna().tolist()
            if str(x).strip() != ""
        ]
    )


# ======================================================
# GEOREFERENCIA Y MAPA INTELIGENTE
# ======================================================

CENTROS_PROVINCIA = {
    "San José": [9.9281, -84.0907],
    "Alajuela": [10.0162, -84.2116],
    "Cartago": [9.8644, -83.9194],
    "Heredia": [10.0024, -84.1165],
    "Guanacaste": [10.6267, -85.4437],
    "Puntarenas": [9.9763, -84.8384],
    "Limón": [9.9917, -83.0360]
}


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

    df_mapa = df_mapa.dropna(
        subset=[
            "Latitud_Num",
            "Longitud_Num"
        ]
    )

    return df_mapa


def obtener_centro_mapa_por_provincia(df):
    if df.empty or "Provincia" not in df.columns:
        return [9.7489, -83.7534], 7

    provincias = df["Provincia"].dropna().astype(str).unique().tolist()
    provincias = [p for p in provincias if p.strip() != ""]

    if len(provincias) == 1:
        provincia = provincias[0]

        if provincia in CENTROS_PROVINCIA:
            return CENTROS_PROVINCIA[provincia], 10

    return [9.7489, -83.7534], 7


def crear_mapa_registros(df, zoom_start=8):
    df_mapa = preparar_dataframe_mapa(df)

    if df_mapa.empty:
        centro, zoom = obtener_centro_mapa_por_provincia(df)

        mapa = folium.Map(
            location=centro,
            zoom_start=zoom,
            tiles="OpenStreetMap"
        )

        return mapa

    total_puntos = len(df_mapa)

    if total_puntos == 1:
        centro_lat = float(df_mapa.iloc[0]["Latitud_Num"])
        centro_lon = float(df_mapa.iloc[0]["Longitud_Num"])
        zoom_mapa = 16

    else:
        centro_lat = float(df_mapa["Latitud_Num"].mean())
        centro_lon = float(df_mapa["Longitud_Num"].mean())
        zoom_mapa = zoom_start

    mapa = folium.Map(
        location=[centro_lat, centro_lon],
        zoom_start=zoom_mapa,
        tiles="OpenStreetMap"
    )

    coordenadas = []

    for _, row in df_mapa.iterrows():
        programa = str(row.get("Programa", ""))
        color = obtener_color_programa(programa)

        lat = float(row["Latitud_Num"])
        lon = float(row["Longitud_Num"])
        coordenadas.append([lat, lon])

        popup_html = f"""
        <div style="font-family: Arial; width: 280px;">
            <h4 style="color:#002B7F; margin-bottom:6px;">
                Registro PUMI #{row.get("ID", "")}
            </h4>
            <b>Fecha:</b> {row.get("Fecha Actividad", "")}<br>
            <b>Hora:</b> {row.get("Hora Actividad", "")}<br>
            <b>Delegación:</b> {row.get("Delegación", "")}<br>
            <b>Programa:</b> {row.get("Programa", "")}<br>
            <b>Actividad:</b> {row.get("Actividad", "")}<br>
            <b>Provincia:</b> {row.get("Provincia", "")}<br>
            <b>Cantón:</b> {row.get("Cantón", "")}<br>
            <b>Distrito:</b> {row.get("Distrito", "")}<br>
            <b>Lugar:</b> {row.get("Lugar", "")}<br>
            <b>Participantes:</b> {row.get("Cantidad Participantes", "")}<br>
            <b>Hombres:</b> {row.get("Cantidad Hombres", "")}<br>
            <b>Mujeres:</b> {row.get("Cantidad Mujeres", "")}<br>
            <b>Referencia:</b> {row.get("Número de Referencia", "")}<br>
            <b>Expediente:</b> {row.get("Número de Expediente Referencia", "")}<br>
            <b>Estado:</b> {row.get("Estado Revisión", "")}<br>
        </div>
        """

        folium.Marker(
            location=[lat, lon],
            popup=folium.Popup(popup_html, max_width=330),
            tooltip=f'{row.get("Programa", "")} - {row.get("Delegación", "")}',
            icon=folium.Icon(
                color=color,
                icon="info-sign"
            )
        ).add_to(mapa)

    if total_puntos == 1:
        folium.Circle(
            location=coordenadas[0],
            radius=1200,
            color=COLOR_AZUL,
            fill=True,
            fill_color=COLOR_ROJO,
            fill_opacity=0.10,
            weight=2,
            popup="Área de referencia del registro filtrado"
        ).add_to(mapa)

    elif total_puntos > 1:
        min_lat = df_mapa["Latitud_Num"].min()
        max_lat = df_mapa["Latitud_Num"].max()
        min_lon = df_mapa["Longitud_Num"].min()
        max_lon = df_mapa["Longitud_Num"].max()

        margen_lat = max((max_lat - min_lat) * 0.25, 0.02)
        margen_lon = max((max_lon - min_lon) * 0.25, 0.02)

        bounds = [
            [min_lat - margen_lat, min_lon - margen_lon],
            [max_lat + margen_lat, max_lon + margen_lon]
        ]

        mapa.fit_bounds(bounds)

        radio = max(
            1500,
            int(max(max_lat - min_lat, max_lon - min_lon) * 111000 / 2)
        )

        folium.Circle(
            location=[centro_lat, centro_lon],
            radius=radio,
            color=COLOR_AZUL,
            fill=True,
            fill_color=COLOR_ROJO,
            fill_opacity=0.08,
            weight=2,
            popup=f"Área aproximada de {total_puntos} registros filtrados"
        ).add_to(mapa)

    return mapa


def mostrar_mapa_registros(df, height=560, key="mapa_registros"):
    df_mapa = preparar_dataframe_mapa(df)

    if df_mapa.empty:
        st.info(
            "No hay registros con coordenadas para mostrar. "
            "Se mostrará la ubicación aproximada según provincia si está disponible."
        )

    mapa = crear_mapa_registros(df)

    st_folium(
        mapa,
        height=height,
        use_container_width=True,
        key=key
    )


# ======================================================
# DATOS PRINCIPALES GOOGLE SHEETS
# Lee registros nuevos y antiguos sin correr columnas.
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

    encabezados_hoja = datos[0]
    filas = datos[1:]
    registros = []

    def parece_hora(valor):
        valor = str(valor).strip()

        if ":" not in valor:
            return False

        partes = valor.split(":")

        if len(partes) != 2:
            return False

        return partes[0].isdigit() and partes[1].isdigit()

    for fila in filas:
        if all(str(celda).strip() == "" for celda in fila):
            continue

        if len(fila) > 0 and str(fila[0]).strip().upper() == "ID":
            continue

        registro = {}

        # Registro nuevo: ya tiene Hora Actividad en la posición correcta.
        if len(fila) > 3 and parece_hora(fila[3]):
            fila_ajustada = fila[:len(ENCABEZADOS)]

            while len(fila_ajustada) < len(ENCABEZADOS):
                fila_ajustada.append("")

            registro = dict(zip(ENCABEZADOS, fila_ajustada))

        # Registro antiguo o encabezados desordenados:
        # se intenta leer por nombre de columna.
        else:
            for col in ENCABEZADOS:
                if col in encabezados_hoja:
                    idx = encabezados_hoja.index(col)

                    if idx < len(fila):
                        registro[col] = fila[idx]
                    else:
                        registro[col] = ""
                else:
                    registro[col] = ""

            if "Hora Actividad" not in encabezados_hoja:
                registro["Hora Actividad"] = ""

        registros.append(registro)

    df = pd.DataFrame(registros, columns=ENCABEZADOS)
    df = aplicar_filtro_perfil_admin(df)

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

    ids = pd.to_numeric(
        df["ID"],
        errors="coerce"
    ).dropna()

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

    for i, fila in enumerate(datos[1:], start=2):
        if len(fila) > 0 and str(fila[0]) == str(id_registro):
            nueva_fila = [
                nuevos_datos.get(col, "")
                for col in ENCABEZADOS
            ]

            try:
                hoja.update(
                    values=[nueva_fila],
                    range_name=f"A{i}"
                )
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

    columnas_numericas = [
        "Cantidad Participantes",
        "Cantidad Hombres",
        "Cantidad Mujeres",
        "Edad 10 a 18",
        "Edad 19 a 30",
        "Edad 31 a 45",
        "Edad 46 en adelante"
    ]

    for col in columnas_numericas:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col],
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
# PARTE 4 DE 12
# INTERFAZ PRINCIPAL, LOGIN ADMINISTRATIVO, MENÚ E INICIO
# ======================================================

# ======================================================
# FUNCIÓN PARA CREAR MAPA BASE
# ======================================================

def crear_mapa_base(centro, zoom, tipo_mapa):
    mapa = folium.Map(
        location=centro,
        zoom_start=zoom,
        tiles=None
    )

    if tipo_mapa == "OpenStreetMap":
        folium.TileLayer(
            "OpenStreetMap",
            name="OpenStreetMap"
        ).add_to(mapa)

    elif tipo_mapa == "Mapa claro":
        folium.TileLayer(
            "CartoDB positron",
            name="Mapa claro"
        ).add_to(mapa)

    elif tipo_mapa == "Mapa oscuro":
        folium.TileLayer(
            "CartoDB dark_matter",
            name="Mapa oscuro"
        ).add_to(mapa)

    elif tipo_mapa == "Topográfico":
        folium.TileLayer(
            "OpenTopoMap",
            name="Topográfico"
        ).add_to(mapa)

    elif tipo_mapa == "Satélite":
        folium.TileLayer(
            tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
            attr="Esri World Imagery",
            name="Satélite"
        ).add_to(mapa)

    folium.LayerControl().add_to(mapa)

    return mapa


# ======================================================
# SIDEBAR, LOGOS Y ACCESO ADMINISTRATIVO
# ======================================================

mostrar_logo()

st.sidebar.markdown("## Sistema PUMI 2026")

if "admin_autenticado" not in st.session_state:
    st.session_state.admin_autenticado = False

if "admin_perfil" not in st.session_state:
    st.session_state.admin_perfil = {}


with st.sidebar.expander("🔒 Acceso Administrativo"):
    if not st.session_state.admin_autenticado:

        clave_ingresada = st.text_input(
            "Clave administrativa",
            type="password"
        )

        if st.button("Ingresar como administrador"):
            clave_normalizada = str(clave_ingresada).strip().upper()

            if clave_normalizada in CLAVES_ADMINISTRATIVAS:
                st.session_state.admin_autenticado = True
                st.session_state.admin_perfil = CLAVES_ADMINISTRATIVAS[
                    clave_normalizada
                ]

                st.success("Acceso concedido.")
                st.rerun()

            else:
                st.error("Clave incorrecta.")

    else:
        perfil_actual = st.session_state.get("admin_perfil", {})

        st.success(
            f"Administrador autenticado: "
            f"{perfil_actual.get('descripcion', 'Perfil administrativo')}"
        )

        if st.button("Cerrar sesión administrativa"):
            st.session_state.admin_autenticado = False
            st.session_state.admin_perfil = {}
            st.rerun()


# ======================================================
# MENÚ PRINCIPAL
# ======================================================

opciones_menu = [
    "Inicio",
    "Registrar actividad",
    "Seguimiento de registros"
]

if st.session_state.admin_autenticado:
    opciones_menu.extend(
        [
            "Consulta / edición administrativa",
            "Dashboard profesional",
            "Configuración"
        ]
    )

menu = st.sidebar.radio(
    "Menú principal",
    opciones_menu
)


# ======================================================
# ENCABEZADO INSTITUCIONAL SUPERIOR
# ======================================================

mostrar_encabezado_institucional()


# ======================================================
# ENCABEZADO PRINCIPAL
# ======================================================

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
# Página limpia.
# Solo muestra mensaje institucional de bienvenida.
# No muestra métricas ni registros.
# ======================================================

if menu == "Inicio":

    st.markdown(
        """
        <div class="card-pumi">
            <div class="subtitulo-pumi">
                Bienvenido al Sistema PUMI 2026
            </div>
            <div class="texto-pumi">
                Esta aplicación permite registrar, consultar y dar seguimiento
                a las actividades desarrolladas dentro del Proceso Unificado
                para el Manejo de la Información.
                <br><br>
                El sistema facilita la trazabilidad de los registros,
                la revisión administrativa, la consulta por programa,
                el análisis territorial y la generación de reportes
                institucionales.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
# ======================================================
# PARTE 5 DE 12
# REGISTRO DE ACTIVIDADES
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

        hora_actividad = st.time_input(
            "Hora de la actividad",
            value=time(8, 0)
        )

        direccion_regional = st.selectbox(
            "Dirección Regional",
            regiones_lista if regiones_lista else REGIONES
        )

        delegaciones_filtradas = obtener_delegaciones_por_region(
            direccion_regional
        )

        delegacion = st.selectbox(
            "Delegación",
            delegaciones_filtradas
            if delegaciones_filtradas
            else ["Sin datos disponibles"]
        )

    with col2:
        responde_a_lista = obtener_responde_a_datos()

        responde_a = st.selectbox(
            "Responde a",
            responde_a_lista
            if responde_a_lista
            else ["Sin datos disponibles"]
        )

        programas_lista = obtener_programas_por_responde_a(responde_a)

        programa = st.selectbox(
            "Programa",
            programas_lista
            if programas_lista
            else ["Sin datos disponibles"]
        )

        actividades_filtradas = obtener_actividades_por_responde_a_programa(
            responde_a,
            programa
        )

        actividad = st.selectbox(
            "Actividad realizada",
            actividades_filtradas
            if actividades_filtradas
            else ["Sin datos disponibles"]
        )

        responsable = st.text_input("Funcionario responsable")

        usuario = st.text_input("Usuario que registra")


    # ======================================================
    # PARTICIPANTES
    # ======================================================

    st.markdown(
        """
        <div class="bloque-datos">
            <b>Participación registrada</b>
        </div>
        """,
        unsafe_allow_html=True
    )

    colp1, colp2, colp3 = st.columns(3)

    with colp1:
        cantidad = st.number_input(
            "Cantidad total de participantes",
            min_value=0,
            step=1
        )

    with colp2:
        cantidad_hombres = st.number_input(
            "Cantidad de hombres",
            min_value=0,
            step=1
        )

    with colp3:
        cantidad_mujeres = st.number_input(
            "Cantidad de mujeres",
            min_value=0,
            step=1
        )

    st.markdown("### Rangos de edad")

    colr1, colr2, colr3, colr4 = st.columns(4)

    with colr1:
        edad_10_18 = st.number_input(
            "Edad 10 a 18",
            min_value=0,
            step=1
        )

    with colr2:
        edad_19_30 = st.number_input(
            "Edad 19 a 30",
            min_value=0,
            step=1
        )

    with colr3:
        edad_31_45 = st.number_input(
            "Edad 31 a 45",
            min_value=0,
            step=1
        )

    with colr4:
        edad_46_mas = st.number_input(
            "Edad 46 en adelante",
            min_value=0,
            step=1
        )

    suma_sexo = cantidad_hombres + cantidad_mujeres
    suma_edades = edad_10_18 + edad_19_30 + edad_31_45 + edad_46_mas

    if cantidad > 0:
        if suma_sexo != cantidad:
            st.warning(
                f"La suma de hombres y mujeres ({suma_sexo}) no coincide "
                f"con la cantidad total de participantes ({cantidad})."
            )

        if suma_edades != cantidad:
            st.warning(
                f"La suma de los rangos de edad ({suma_edades}) no coincide "
                f"con la cantidad total de participantes ({cantidad})."
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
            cantones_filtrados
            if cantones_filtrados
            else ["Sin datos disponibles"]
        )

    with col5:
        distritos_filtrados = obtener_distritos_por_provincia_canton(
            provincia,
            canton
        )

        distrito = st.selectbox(
            "Distrito",
            distritos_filtrados
            if distritos_filtrados
            else ["Sin datos disponibles"]
        )


    # ======================================================
    # LUGAR DE REALIZACIÓN
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
    codigo_presupuestario = ""

    if tipo_lugar == "Centro educativo":
        centros = obtener_centros_por_provincia(provincia)

        if centros:
            centro_mostrar = st.selectbox(
                "Centro educativo según base MEP 2025",
                centros
            )

            centro_educativo, codigo_presupuestario = obtener_datos_centro_educativo(
                centro_mostrar
            )

            st.info(f"Centro educativo: {centro_educativo}")
            st.info(
                "Código presupuestario: "
                f"{codigo_presupuestario if codigo_presupuestario else 'No disponible'}"
            )

            lugar = centro_educativo

        else:
            st.warning(
                "No se encontraron centros educativos para la provincia seleccionada."
            )

            centro_educativo = st.text_input(
                "Digite el centro educativo manualmente"
            )

            codigo_presupuestario = st.text_input(
                "Código presupuestario"
            )

            lugar = centro_educativo

    else:
        lugar = st.text_input("Lugar donde se realizó la actividad")
        centro_educativo = ""
        codigo_presupuestario = ""


    # ======================================================
    # MAPA Y GEOREFERENCIA
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
        [
            "OpenStreetMap",
            "Mapa claro",
            "Mapa oscuro",
            "Topográfico",
            "Satélite"
        ],
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
            lat_busqueda, lon_busqueda, direccion_encontrada = georreferenciar_direccion(
                direccion_mapa
            )

            if lat_busqueda and lon_busqueda:
                st.session_state.latitud_registro = str(lat_busqueda)
                st.session_state.longitud_registro = str(lon_busqueda)

                st.success("Ubicación encontrada automáticamente.")
                st.caption(direccion_encontrada)
            else:
                st.warning(
                    "No se logró ubicar el lugar automáticamente. "
                    "Puede usar la opción Marcar punto en el mapa."
                )

    elif metodo_ubicacion == "Usar GPS del dispositivo":
        st.info(
            "El navegador puede solicitar permiso para acceder a la ubicación del dispositivo."
        )

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
            st.warning(
                "No se obtuvo ubicación GPS. Puede usar la opción Marcar punto en el mapa."
            )

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
        st.info(
            "Haga clic sobre el mapa para seleccionar la ubicación de la actividad."
        )

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
        height=650,
        use_container_width=True,
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

    instituciones = st.text_area(
        "Instituciones participantes",
        max_chars=8000,
        height=130
    )

    plan = st.text_area(
        "Plan estratégico relacionado",
        max_chars=8000,
        height=130
    )

    numero_referencia = st.text_area(
        "Número de Referencia",
        max_chars=8000,
        height=100
    )

    numero_expediente_referencia = st.text_area(
        "Número de Expediente Referencia",
        max_chars=8000,
        height=100
    )

    observaciones = st.text_area(
        "Observaciones",
        max_chars=8000,
        height=170
    )


    # ======================================================
    # GUARDAR REGISTRO
    # ======================================================

    if st.button("Guardar registro"):

        if (
            not delegacion
            or delegacion == "Sin datos disponibles"
            or responde_a == "Sin datos disponibles"
            or programa == "Sin datos disponibles"
            or actividad == "Sin datos disponibles"
            or not responsable
        ):
            st.warning(
                "Debe completar al menos Delegación, Responde a, Programa, "
                "Actividad realizada y Funcionario responsable."
            )

        elif cantidad > 0 and suma_sexo != cantidad:
            st.warning(
                "La suma de hombres y mujeres debe coincidir con la cantidad total de participantes."
            )

        elif cantidad > 0 and suma_edades != cantidad:
            st.warning(
                "La suma de los rangos de edad debe coincidir con la cantidad total de participantes."
            )

        else:
            nuevo_id = generar_id_consecutivo()

            registro = {
                "ID": nuevo_id,
                "Fecha Registro": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "Fecha Actividad": fecha_actividad.strftime("%d/%m/%Y"),
                "Hora Actividad": hora_actividad.strftime("%H:%M"),
                "Dirección Regional": direccion_regional,
                "Delegación": delegacion,
                "Responde a": responde_a,
                "Programa": programa,
                "Actividad": actividad,
                "Provincia": provincia,
                "Cantón": canton,
                "Distrito": distrito,
                "Tipo Lugar": tipo_lugar,
                "Lugar": lugar,
                "Centro Educativo": centro_educativo,
                "Código Presupuestario": codigo_presupuestario,
                "Dirección Mapa": direccion_mapa,
                "Latitud": st.session_state.latitud_registro,
                "Longitud": st.session_state.longitud_registro,
                "Responsable": responsable,
                "Cantidad Participantes": cantidad,
                "Cantidad Hombres": cantidad_hombres,
                "Cantidad Mujeres": cantidad_mujeres,
                "Edad 10 a 18": edad_10_18,
                "Edad 19 a 30": edad_19_30,
                "Edad 31 a 45": edad_31_45,
                "Edad 46 en adelante": edad_46_mas,
                "Instituciones Participantes": instituciones,
                "Plan Estratégico Relacionado": plan,
                "Número de Referencia": numero_referencia,
                "Número de Expediente Referencia": numero_expediente_referencia,
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
# PARTE 6 DE 12
# SEGUIMIENTO DE REGISTROS CON CLAVE POR PROGRAMA
# ======================================================

elif menu == "Seguimiento de registros":

    st.markdown("## Seguimiento de registros")

    st.markdown(
        """
        <div class="card-azul">
            <div class="texto-pumi">
                Para consultar registros, seleccione el programa e ingrese la clave correspondiente.
                Las claves por programa permiten visualizar únicamente los registros autorizados.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    programa_consulta = st.selectbox(
        "Seleccione el programa a consultar",
        ["Seleccione"] + PROGRAMAS + ["Todos"]
    )

    clave_usuario = st.text_input(
        "Clave de consulta",
        type="password"
    )

    acceso_valido = False
    programas_permitidos_usuario = []

    if programa_consulta != "Seleccione" and clave_usuario:
        clave_normalizada = str(clave_usuario).strip().upper()

        claves_dppp = [
            c.upper()
            for c in CLAVES_USUARIO_PROGRAMA["DPPP"]["claves"]
        ]

        if clave_normalizada in claves_dppp:
            acceso_valido = True
            programas_permitidos_usuario = "TODOS"

        elif programa_consulta != "Todos":
            datos_programa = CLAVES_USUARIO_PROGRAMA.get(programa_consulta, {})
            claves_validas = [
                c.upper()
                for c in datos_programa.get("claves", [])
            ]

            if clave_normalizada in claves_validas:
                acceso_valido = True
                programas_permitidos_usuario = datos_programa.get(
                    "programas_permitidos",
                    []
                )

        if not acceso_valido:
            st.error("Clave incorrecta para el programa seleccionado.")

    if programa_consulta == "Seleccione":
        st.info("Seleccione un programa para continuar.")

    elif not clave_usuario:
        st.warning("Ingrese la clave de consulta para ver los registros.")

    elif acceso_valido:

        df = cargar_datos()

        if df.empty:
            st.info("No existen registros para consultar.")

        else:
            df_seguimiento = limpiar_dataframe_para_metricas(df)

            if "Responde a" not in df_seguimiento.columns:
                df_seguimiento["Responde a"] = ""

            if programas_permitidos_usuario != "TODOS":
                df_seguimiento = df_seguimiento[
                    df_seguimiento["Programa"].isin(programas_permitidos_usuario)
                ]

            elif programa_consulta != "Todos":
                if programa_consulta == "GREAT":
                    df_seguimiento = df_seguimiento[
                        df_seguimiento["Programa"].isin(["GREAT", "GREAT CAMP"])
                    ]
                else:
                    df_seguimiento = df_seguimiento[
                        df_seguimiento["Programa"] == programa_consulta
                    ]

            st.success("Acceso concedido.")

            # ======================================================
            # FILTROS DE SEGUIMIENTO
            # ======================================================

            st.markdown("### Filtros de seguimiento")

            col1, col2, col3 = st.columns(3)

            with col1:
                filtro_id = st.text_input("Buscar por ID")

                filtro_responde_a = st.selectbox(
                    "Responde a",
                    ["Todos"] + sorted(
                        df_seguimiento["Responde a"].dropna().astype(str).unique().tolist()
                    )
                )

            with col2:
                filtro_delegacion = st.selectbox(
                    "Delegación",
                    ["Todas"] + sorted(
                        df_seguimiento["Delegación"].dropna().astype(str).unique().tolist()
                    )
                )

                filtro_estado = st.selectbox(
                    "Estado de revisión",
                    ["Todos"] + sorted(
                        df_seguimiento["Estado Revisión"].dropna().astype(str).unique().tolist()
                    )
                )

            with col3:
                filtro_provincia = st.selectbox(
                    "Provincia",
                    ["Todas"] + sorted(
                        df_seguimiento["Provincia"].dropna().astype(str).unique().tolist()
                    )
                )

                filtro_distrito = st.selectbox(
                    "Distrito",
                    ["Todos"] + sorted(
                        df_seguimiento["Distrito"].dropna().astype(str).unique().tolist()
                    )
                )

                usar_fechas_seg = st.checkbox(
                    "Filtrar por rango de fechas",
                    key="fechas_seguimiento"
                )

            df_usuario = df_seguimiento.copy()

            if filtro_id:
                df_usuario = df_usuario[
                    df_usuario["ID"].astype(str).str.contains(
                        filtro_id,
                        case=False,
                        na=False
                    )
                ]

            if filtro_responde_a != "Todos":
                df_usuario = df_usuario[
                    df_usuario["Responde a"] == filtro_responde_a
                ]

            if filtro_delegacion != "Todas":
                df_usuario = df_usuario[
                    df_usuario["Delegación"] == filtro_delegacion
                ]

            if filtro_estado != "Todos":
                df_usuario = df_usuario[
                    df_usuario["Estado Revisión"] == filtro_estado
                ]

            if filtro_provincia != "Todas":
                df_usuario = df_usuario[
                    df_usuario["Provincia"] == filtro_provincia
                ]

            if filtro_distrito != "Todos":
                df_usuario = df_usuario[
                    df_usuario["Distrito"] == filtro_distrito
                ]

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
                        (df_usuario["Fecha Actividad"].dt.date >= fecha_inicio_seg)
                        &
                        (df_usuario["Fecha Actividad"].dt.date <= fecha_fin_seg)
                    ]

            # ======================================================
            # RESUMEN DEL SEGUIMIENTO
            # ======================================================

            st.markdown("### Resumen del seguimiento")

            total_participantes = (
                int(df_usuario["Cantidad Participantes"].sum())
                if not df_usuario.empty else 0
            )

            total_hombres = (
                int(df_usuario["Cantidad Hombres"].sum())
                if "Cantidad Hombres" in df_usuario.columns else 0
            )

            total_mujeres = (
                int(df_usuario["Cantidad Mujeres"].sum())
                if "Cantidad Mujeres" in df_usuario.columns else 0
            )

            colm1, colm2, colm3, colm4 = st.columns(4)

            colm1.metric("Registros", len(df_usuario))
            colm2.metric("Participantes", total_participantes)
            colm3.metric(
                "Hombres",
                f"{total_hombres} ({porcentaje(total_hombres, total_participantes)})"
            )
            colm4.metric(
                "Mujeres",
                f"{total_mujeres} ({porcentaje(total_mujeres, total_participantes)})"
            )

            colm5, colm6, colm7, colm8 = st.columns(4)

            colm5.metric(
                "Aprobados",
                len(df_usuario[df_usuario["Estado Revisión"] == "Aprobado"])
            )

            colm6.metric(
                "Con observaciones",
                len(df_usuario[df_usuario["Estado Revisión"] == "Con observaciones"])
            )

            colm7.metric(
                "Rechazados",
                len(df_usuario[df_usuario["Estado Revisión"] == "Rechazado"])
            )

            colm8.metric(
                "Pendientes",
                len(df_usuario[df_usuario["Estado Revisión"] == "Pendiente de revisión"])
            )

            # ======================================================
            # MAPA DE SEGUIMIENTO
            # ======================================================

            st.markdown("### Mapa general de actividades")

            mostrar_mapa_registros(
                df_usuario,
                height=650,
                key="mapa_seguimiento_usuarios"
            )

            # ======================================================
            # TABLA DE REGISTROS
            # ======================================================

            st.markdown("### Estado de los registros")

            df_mostrar_usuario = df_usuario.copy()

            if "Fecha Actividad" in df_mostrar_usuario.columns:
                df_mostrar_usuario["Fecha Actividad"] = df_mostrar_usuario[
                    "Fecha Actividad"
                ].dt.strftime("%d/%m/%Y")

            columnas_usuario = [
                "ID",
                "Fecha Actividad",
                "Hora Actividad",
                "Delegación",
                "Responde a",
                "Programa",
                "Actividad",
                "Provincia",
                "Cantón",
                "Distrito",
                "Lugar",
                "Centro Educativo",
                "Código Presupuestario",
                "Cantidad Participantes",
                "Cantidad Hombres",
                "Cantidad Mujeres",
                "Edad 10 a 18",
                "Edad 19 a 30",
                "Edad 31 a 45",
                "Edad 46 en adelante",
                "Número de Referencia",
                "Número de Expediente Referencia",
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

            excel_seguimiento = convertir_excel(
                df_mostrar_usuario[columnas_existentes]
            )

            st.download_button(
                label="Descargar seguimiento en Excel",
                data=excel_seguimiento,
                file_name="seguimiento_registros_pumi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
# ======================================================
# PARTE 7 DE 12
# CONSULTA ADMINISTRATIVA, FILTROS, MAPA Y TABLA GENERAL
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

        if "Responde a" not in df_consulta.columns:
            df_consulta["Responde a"] = ""

        st.markdown(
            """
            <div class="card-dorado">
                <div class="texto-pumi">
                    Esta sección permite consultar, filtrar, revisar, editar,
                    corregir ubicación geográfica y eliminar registros.
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

        # ======================================================
        # FILTROS ADMINISTRATIVOS
        # ======================================================

        st.markdown("### Filtros administrativos")

        col1, col2, col3 = st.columns(3)

        with col1:
            filtro_id_admin = st.text_input("Buscar por ID")

            filtro_responde_a_admin = st.selectbox(
                "Filtrar por Responde a",
                ["Todos"] + sorted(
                    df_consulta["Responde a"].dropna().astype(str).unique().tolist()
                )
            )

            filtro_programa = st.selectbox(
                "Filtrar por programa",
                ["Todos"] + sorted(
                    df_consulta["Programa"].dropna().astype(str).unique().tolist()
                )
            )

        with col2:
            filtro_delegacion = st.selectbox(
                "Filtrar por delegación",
                ["Todas"] + sorted(
                    df_consulta["Delegación"].dropna().astype(str).unique().tolist()
                )
            )

            filtro_region = st.selectbox(
                "Filtrar por región",
                ["Todas"] + ordenar_regiones_numericamente(
                    sorted(
                        df_consulta["Dirección Regional"]
                        .dropna()
                        .astype(str)
                        .unique()
                        .tolist()
                    )
                )
            )

            filtro_provincia = st.selectbox(
                "Filtrar por provincia",
                ["Todas"] + sorted(
                    df_consulta["Provincia"].dropna().astype(str).unique().tolist()
                )
            )

        with col3:
            filtro_distrito = st.selectbox(
                "Filtrar por distrito",
                ["Todos"] + sorted(
                    df_consulta["Distrito"].dropna().astype(str).unique().tolist()
                )
            )

            filtro_estado = st.selectbox(
                "Filtrar por estado de revisión",
                ["Todos"] + sorted(
                    df_consulta["Estado Revisión"].dropna().astype(str).unique().tolist()
                )
            )

            filtro_tipo_lugar = st.selectbox(
                "Filtrar por tipo de lugar",
                ["Todos"] + sorted(
                    df_consulta["Tipo Lugar"].dropna().astype(str).unique().tolist()
                )
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

        if filtro_responde_a_admin != "Todos":
            df_filtrado = df_filtrado[
                df_filtrado["Responde a"] == filtro_responde_a_admin
            ]

        if filtro_programa != "Todos":
            df_filtrado = df_filtrado[
                df_filtrado["Programa"] == filtro_programa
            ]

        if filtro_delegacion != "Todas":
            df_filtrado = df_filtrado[
                df_filtrado["Delegación"] == filtro_delegacion
            ]

        if filtro_region != "Todas":
            df_filtrado = df_filtrado[
                df_filtrado["Dirección Regional"] == filtro_region
            ]

        if filtro_provincia != "Todas":
            df_filtrado = df_filtrado[
                df_filtrado["Provincia"] == filtro_provincia
            ]

        if filtro_distrito != "Todos":
            df_filtrado = df_filtrado[
                df_filtrado["Distrito"] == filtro_distrito
            ]

        if filtro_estado != "Todos":
            df_filtrado = df_filtrado[
                df_filtrado["Estado Revisión"] == filtro_estado
            ]

        if filtro_tipo_lugar != "Todos":
            df_filtrado = df_filtrado[
                df_filtrado["Tipo Lugar"] == filtro_tipo_lugar
            ]

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
                    (df_filtrado["Fecha Actividad"].dt.date >= fecha_inicio)
                    &
                    (df_filtrado["Fecha Actividad"].dt.date <= fecha_fin)
                ]

        # ======================================================
        # RESUMEN ADMINISTRATIVO
        # ======================================================

        st.markdown("### Resumen administrativo")

        total_registros = len(df_filtrado)

        total_participantes = (
            int(df_filtrado["Cantidad Participantes"].sum())
            if not df_filtrado.empty else 0
        )

        total_hombres = (
            int(df_filtrado["Cantidad Hombres"].sum())
            if "Cantidad Hombres" in df_filtrado.columns else 0
        )

        total_mujeres = (
            int(df_filtrado["Cantidad Mujeres"].sum())
            if "Cantidad Mujeres" in df_filtrado.columns else 0
        )

        total_aprobados = len(
            df_filtrado[df_filtrado["Estado Revisión"] == "Aprobado"]
        )

        total_observaciones = len(
            df_filtrado[df_filtrado["Estado Revisión"] == "Con observaciones"]
        )

        total_rechazados = len(
            df_filtrado[df_filtrado["Estado Revisión"] == "Rechazado"]
        )

        total_pendientes = len(
            df_filtrado[df_filtrado["Estado Revisión"] == "Pendiente de revisión"]
        )

        colm1, colm2, colm3, colm4 = st.columns(4)

        colm1.metric("Registros filtrados", total_registros)
        colm2.metric("Participantes", total_participantes)
        colm3.metric(
            "Hombres",
            f"{total_hombres} ({porcentaje(total_hombres, total_participantes)})"
        )
        colm4.metric(
            "Mujeres",
            f"{total_mujeres} ({porcentaje(total_mujeres, total_participantes)})"
        )

        colm5, colm6, colm7, colm8 = st.columns(4)

        colm5.metric(
            "Aprobados",
            f"{total_aprobados} ({porcentaje(total_aprobados, total_registros)})"
        )

        colm6.metric(
            "Con observaciones",
            f"{total_observaciones} ({porcentaje(total_observaciones, total_registros)})"
        )

        colm7.metric(
            "Rechazados",
            f"{total_rechazados} ({porcentaje(total_rechazados, total_registros)})"
        )

        colm8.metric(
            "Pendientes",
            f"{total_pendientes} ({porcentaje(total_pendientes, total_registros)})"
        )

        # ======================================================
        # MAPA ADMINISTRATIVO
        # ======================================================

        st.markdown("### Mapa administrativo de registros filtrados")

        mostrar_mapa_registros(
            df_filtrado,
            height=650,
            key="mapa_admin_registros"
        )

        # ======================================================
        # TABLA ADMINISTRATIVA
        # ======================================================

        st.markdown("### Registros encontrados")

        df_mostrar = df_filtrado.copy()

        if "Fecha Actividad" in df_mostrar.columns:
            df_mostrar["Fecha Actividad"] = df_mostrar["Fecha Actividad"].dt.strftime(
                "%d/%m/%Y"
            )

        columnas_admin = [
            "ID",
            "Fecha Registro",
            "Fecha Actividad",
            "Hora Actividad",
            "Dirección Regional",
            "Delegación",
            "Responde a",
            "Programa",
            "Actividad",
            "Provincia",
            "Cantón",
            "Distrito",
            "Tipo Lugar",
            "Lugar",
            "Centro Educativo",
            "Código Presupuestario",
            "Dirección Mapa",
            "Latitud",
            "Longitud",
            "Responsable",
            "Cantidad Participantes",
            "Cantidad Hombres",
            "Cantidad Mujeres",
            "Edad 10 a 18",
            "Edad 19 a 30",
            "Edad 31 a 45",
            "Edad 46 en adelante",
            "Instituciones Participantes",
            "Plan Estratégico Relacionado",
            "Número de Referencia",
            "Número de Expediente Referencia",
            "Observaciones",
            "Estado Revisión",
            "Observación de Revisión",
            "Usuario Registra"
        ]

        columnas_admin_existentes = [
            col for col in columnas_admin
            if col in df_mostrar.columns
        ]

        st.dataframe(
            df_mostrar[columnas_admin_existentes],
            use_container_width=True,
            hide_index=True
        )

        excel = convertir_excel(df_mostrar[columnas_admin_existentes])

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
            # ======================================================
# PARTE 8 DE 12
# SELECCIÓN DE REGISTRO Y ACCIONES ADMINISTRATIVAS
# ACTUALIZAR REVISIÓN Y CORREGIR UBICACIÓN
# ======================================================

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

                    estado_actual = registro_actual.get(
                        "Estado Revisión",
                        "Pendiente de revisión"
                    )

                    nuevo_estado = st.selectbox(
                        "Estado de revisión",
                        ESTADOS_REVISION,
                        index=ESTADOS_REVISION.index(estado_actual)
                        if estado_actual in ESTADOS_REVISION else 0
                    )

                    observacion_revision = st.text_area(
                        "Observación de Revisión",
                        value=registro_actual.get("Observación de Revisión", ""),
                        max_chars=8000,
                        height=160,
                        help="Esta observación será visible para el usuario."
                    )

                    guardar_revision = st.form_submit_button(
                        "Guardar revisión"
                    )

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
                    [
                        "OpenStreetMap",
                        "Mapa claro",
                        "Mapa oscuro",
                        "Topográfico",
                        "Satélite"
                    ],
                    key="tipo_mapa_admin_correccion"
                )

                direccion_actual = registro_actual.get("Dirección Mapa", "")

                direccion_mapa_admin = st.text_input(
                    "Dirección o referencia para ubicar en mapa",
                    value=direccion_actual
                    if direccion_actual
                    else (
                        f'{registro_actual.get("Lugar", "")}, '
                        f'{registro_actual.get("Distrito", "")}, '
                        f'{registro_actual.get("Cantón", "")}, '
                        f'{registro_actual.get("Provincia", "")}, Costa Rica'
                    )
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
                        lat_busqueda, lon_busqueda, direccion_encontrada = georreferenciar_direccion(
                            direccion_mapa_admin
                        )

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
                    st.info(
                        "Haga clic sobre el mapa para seleccionar la ubicación correcta."
                    )

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
                            color=obtener_color_programa(
                                registro_actual.get("Programa", "")
                            ),
                            icon="info-sign"
                        )
                    ).add_to(mapa_preview_admin)

                resultado_admin = st_folium(
                    mapa_preview_admin,
                    height=650,
                    use_container_width=True,
                    key=f"mapa_correccion_admin_{tipo_mapa_admin}_{metodo_admin}"
                )

                if resultado_admin and resultado_admin.get("last_clicked"):
                    st.session_state.lat_admin = str(
                        resultado_admin["last_clicked"]["lat"]
                    )
                    st.session_state.lon_admin = str(
                        resultado_admin["last_clicked"]["lng"]
                    )

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
# ======================================================
# PARTE 9 DE 12
# EDICIÓN ADMINISTRATIVA COMPLETA Y ELIMINACIÓN
# ======================================================

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

                hora_actual_texto = str(
                    registro_actual.get("Hora Actividad", "08:00")
                ).strip()

                try:
                    hora_actual = datetime.strptime(
                        hora_actual_texto,
                        "%H:%M"
                    ).time()
                except Exception:
                    hora_actual = time(8, 0)

                col1, col2 = st.columns(2)

                with col1:
                    fecha_actividad_editada = st.date_input(
                        "Fecha de la actividad",
                        value=fecha_actividad_actual,
                        format="DD/MM/YYYY"
                    )

                    hora_actividad_editada = st.time_input(
                        "Hora de la actividad",
                        value=hora_actual
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

                    delegaciones_filtradas = obtener_delegaciones_por_region(
                        nueva_region
                    )

                    delegacion_actual = registro_actual.get("Delegación", "")
                    opciones_delegacion = (
                        delegaciones_filtradas
                        if delegaciones_filtradas
                        else obtener_delegaciones_unicas()
                    )

                    if delegacion_actual and delegacion_actual not in opciones_delegacion:
                        opciones_delegacion = [delegacion_actual] + opciones_delegacion

                    nueva_delegacion = st.selectbox(
                        "Delegación",
                        opciones_delegacion
                        if opciones_delegacion
                        else ["Sin datos disponibles"],
                        index=opciones_delegacion.index(delegacion_actual)
                        if delegacion_actual in opciones_delegacion else 0
                    )

                with col2:
                    responde_a_lista = obtener_responde_a_datos()

                    responde_a_actual = registro_actual.get("Responde a", "")
                    opciones_responde_a = (
                        responde_a_lista
                        if responde_a_lista
                        else [responde_a_actual]
                    )

                    if responde_a_actual and responde_a_actual not in opciones_responde_a:
                        opciones_responde_a = [responde_a_actual] + opciones_responde_a

                    nuevo_responde_a = st.selectbox(
                        "Responde a",
                        opciones_responde_a
                        if opciones_responde_a
                        else ["Sin datos disponibles"],
                        index=opciones_responde_a.index(responde_a_actual)
                        if responde_a_actual in opciones_responde_a else 0
                    )

                    programas_filtrados = obtener_programas_por_responde_a(
                        nuevo_responde_a
                    )

                    programa_actual = registro_actual.get("Programa", "")
                    opciones_programa = (
                        programas_filtrados
                        if programas_filtrados
                        else obtener_programas_datos()
                    )

                    if programa_actual and programa_actual not in opciones_programa:
                        opciones_programa = [programa_actual] + opciones_programa

                    nuevo_programa = st.selectbox(
                        "Programa",
                        opciones_programa
                        if opciones_programa
                        else ["Sin datos disponibles"],
                        index=opciones_programa.index(programa_actual)
                        if programa_actual in opciones_programa else 0
                    )

                    actividades_filtradas = obtener_actividades_por_responde_a_programa(
                        nuevo_responde_a,
                        nuevo_programa
                    )

                    actividad_actual = registro_actual.get("Actividad", "")
                    opciones_actividad = (
                        actividades_filtradas
                        if actividades_filtradas
                        else [actividad_actual]
                    )

                    if actividad_actual and actividad_actual not in opciones_actividad:
                        opciones_actividad = [actividad_actual] + opciones_actividad

                    nueva_actividad = st.selectbox(
                        "Actividad realizada",
                        opciones_actividad
                        if opciones_actividad
                        else ["Sin datos disponibles"],
                        index=opciones_actividad.index(actividad_actual)
                        if actividad_actual in opciones_actividad else 0
                    )

                    nuevo_responsable = st.text_input(
                        "Responsable",
                        registro_actual.get("Responsable", "")
                    )

                    nuevo_usuario = st.text_input(
                        "Usuario Registra",
                        registro_actual.get("Usuario Registra", "")
                    )

                st.markdown(
                    """
                    <div class="bloque-datos">
                        <b>Participación registrada</b>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                cantidad_actual = pd.to_numeric(
                    registro_actual.get("Cantidad Participantes", 0),
                    errors="coerce"
                )

                hombres_actual = pd.to_numeric(
                    registro_actual.get("Cantidad Hombres", 0),
                    errors="coerce"
                )

                mujeres_actual = pd.to_numeric(
                    registro_actual.get("Cantidad Mujeres", 0),
                    errors="coerce"
                )

                edad_10_18_actual = pd.to_numeric(
                    registro_actual.get("Edad 10 a 18", 0),
                    errors="coerce"
                )

                edad_19_30_actual = pd.to_numeric(
                    registro_actual.get("Edad 19 a 30", 0),
                    errors="coerce"
                )

                edad_31_45_actual = pd.to_numeric(
                    registro_actual.get("Edad 31 a 45", 0),
                    errors="coerce"
                )

                edad_46_actual = pd.to_numeric(
                    registro_actual.get("Edad 46 en adelante", 0),
                    errors="coerce"
                )

                if pd.isna(cantidad_actual):
                    cantidad_actual = 0
                if pd.isna(hombres_actual):
                    hombres_actual = 0
                if pd.isna(mujeres_actual):
                    mujeres_actual = 0
                if pd.isna(edad_10_18_actual):
                    edad_10_18_actual = 0
                if pd.isna(edad_19_30_actual):
                    edad_19_30_actual = 0
                if pd.isna(edad_31_45_actual):
                    edad_31_45_actual = 0
                if pd.isna(edad_46_actual):
                    edad_46_actual = 0

                colp1, colp2, colp3 = st.columns(3)

                with colp1:
                    nueva_cantidad = st.number_input(
                        "Cantidad Participantes",
                        min_value=0,
                        step=1,
                        value=int(cantidad_actual)
                    )

                with colp2:
                    nueva_cantidad_hombres = st.number_input(
                        "Cantidad Hombres",
                        min_value=0,
                        step=1,
                        value=int(hombres_actual)
                    )

                with colp3:
                    nueva_cantidad_mujeres = st.number_input(
                        "Cantidad Mujeres",
                        min_value=0,
                        step=1,
                        value=int(mujeres_actual)
                    )

                st.markdown("### Rangos de edad")

                colr1, colr2, colr3, colr4 = st.columns(4)

                with colr1:
                    nueva_edad_10_18 = st.number_input(
                        "Edad 10 a 18",
                        min_value=0,
                        step=1,
                        value=int(edad_10_18_actual)
                    )

                with colr2:
                    nueva_edad_19_30 = st.number_input(
                        "Edad 19 a 30",
                        min_value=0,
                        step=1,
                        value=int(edad_19_30_actual)
                    )

                with colr3:
                    nueva_edad_31_45 = st.number_input(
                        "Edad 31 a 45",
                        min_value=0,
                        step=1,
                        value=int(edad_31_45_actual)
                    )

                with colr4:
                    nueva_edad_46 = st.number_input(
                        "Edad 46 en adelante",
                        min_value=0,
                        step=1,
                        value=int(edad_46_actual)
                    )

                suma_sexo_editada = (
                    nueva_cantidad_hombres
                    + nueva_cantidad_mujeres
                )

                suma_edades_editada = (
                    nueva_edad_10_18
                    + nueva_edad_19_30
                    + nueva_edad_31_45
                    + nueva_edad_46
                )

                if nueva_cantidad > 0:
                    if suma_sexo_editada != nueva_cantidad:
                        st.warning(
                            f"La suma de hombres y mujeres ({suma_sexo_editada}) no coincide "
                            f"con la cantidad total ({nueva_cantidad})."
                        )

                    if suma_edades_editada != nueva_cantidad:
                        st.warning(
                            f"La suma de rangos de edad ({suma_edades_editada}) no coincide "
                            f"con la cantidad total ({nueva_cantidad})."
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
                    cantones_filtrados = obtener_cantones_por_provincia(
                        nueva_provincia
                    )

                    canton_actual = registro_actual.get("Cantón", "")
                    opciones_canton = (
                        cantones_filtrados
                        if cantones_filtrados
                        else [canton_actual]
                    )

                    if canton_actual and canton_actual not in opciones_canton:
                        opciones_canton = [canton_actual] + opciones_canton

                    nuevo_canton = st.selectbox(
                        "Cantón",
                        opciones_canton
                        if opciones_canton
                        else ["Sin datos disponibles"],
                        index=opciones_canton.index(canton_actual)
                        if canton_actual in opciones_canton else 0
                    )

                with col5:
                    distritos_filtrados = obtener_distritos_por_provincia_canton(
                        nueva_provincia,
                        nuevo_canton
                    )

                    distrito_actual = registro_actual.get("Distrito", "")
                    opciones_distrito = (
                        distritos_filtrados
                        if distritos_filtrados
                        else [distrito_actual]
                    )

                    if distrito_actual and distrito_actual not in opciones_distrito:
                        opciones_distrito = [distrito_actual] + opciones_distrito

                    nuevo_distrito = st.selectbox(
                        "Distrito",
                        opciones_distrito
                        if opciones_distrito
                        else ["Sin datos disponibles"],
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
                codigo_presupuestario_editado = registro_actual.get(
                    "Código Presupuestario",
                    ""
                )

                if tipo_lugar_editado == "Centro educativo":

                    centros = obtener_centros_por_provincia(nueva_provincia)
                    centro_actual = registro_actual.get("Centro Educativo", "")

                    if centros:
                        opciones_centros = centros
                        centro_mostrar_actual = centro_actual

                        for centro_opcion in opciones_centros:
                            nombre_tmp, codigo_tmp = obtener_datos_centro_educativo(
                                centro_opcion
                            )

                            if normalizar_texto(nombre_tmp) == normalizar_texto(
                                centro_actual
                            ):
                                centro_mostrar_actual = centro_opcion
                                break

                        if centro_mostrar_actual and centro_mostrar_actual not in opciones_centros:
                            opciones_centros = [centro_mostrar_actual] + opciones_centros

                        centro_seleccionado = st.selectbox(
                            "Centro educativo según base MEP 2025",
                            opciones_centros,
                            index=opciones_centros.index(centro_mostrar_actual)
                            if centro_mostrar_actual in opciones_centros else 0
                        )

                        centro_editado, codigo_presupuestario_editado = obtener_datos_centro_educativo(
                            centro_seleccionado
                        )

                        st.info(f"Centro educativo: {centro_editado}")
                        st.info(
                            "Código presupuestario: "
                            f"{codigo_presupuestario_editado if codigo_presupuestario_editado else 'No disponible'}"
                        )

                    else:
                        centro_editado = st.text_input(
                            "Digite el centro educativo manualmente",
                            centro_actual
                        )

                        codigo_presupuestario_editado = st.text_input(
                            "Código Presupuestario",
                            registro_actual.get("Código Presupuestario", "")
                        )

                    lugar_editado = centro_editado

                else:
                    lugar_editado = st.text_input(
                        "Lugar donde se realizó la actividad",
                        registro_actual.get("Lugar", "")
                    )
                    centro_editado = ""
                    codigo_presupuestario_editado = ""

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
                    registro_actual.get("Instituciones Participantes", ""),
                    max_chars=8000,
                    height=130
                )

                nuevo_plan = st.text_area(
                    "Plan Estratégico Relacionado",
                    registro_actual.get("Plan Estratégico Relacionado", ""),
                    max_chars=8000,
                    height=130
                )

                nuevo_numero_referencia = st.text_area(
                    "Número de Referencia",
                    registro_actual.get("Número de Referencia", ""),
                    max_chars=8000,
                    height=100
                )

                nuevo_numero_expediente = st.text_area(
                    "Número de Expediente Referencia",
                    registro_actual.get("Número de Expediente Referencia", ""),
                    max_chars=8000,
                    height=100
                )

                nuevas_observaciones = st.text_area(
                    "Observaciones del registro",
                    registro_actual.get("Observaciones", ""),
                    max_chars=8000,
                    height=160
                )

                nuevo_estado = st.selectbox(
                    "Estado Revisión",
                    ESTADOS_REVISION,
                    index=ESTADOS_REVISION.index(
                        registro_actual.get("Estado Revisión", "")
                    )
                    if registro_actual.get("Estado Revisión", "") in ESTADOS_REVISION else 0
                )

                nueva_observacion_revision = st.text_area(
                    "Observación de Revisión",
                    registro_actual.get("Observación de Revisión", ""),
                    max_chars=8000,
                    height=160
                )

                if st.button("Guardar cambios completos"):

                    if (
                        nueva_delegacion == "Sin datos disponibles"
                        or nuevo_responde_a == "Sin datos disponibles"
                        or nuevo_programa == "Sin datos disponibles"
                        or nueva_actividad == "Sin datos disponibles"
                    ):
                        st.warning(
                            "Debe completar Delegación, Responde a, Programa y Actividad realizada."
                        )

                    elif nueva_cantidad > 0 and suma_sexo_editada != nueva_cantidad:
                        st.warning(
                            "La suma de hombres y mujeres debe coincidir con la cantidad total de participantes."
                        )

                    elif nueva_cantidad > 0 and suma_edades_editada != nueva_cantidad:
                        st.warning(
                            "La suma de los rangos de edad debe coincidir con la cantidad total de participantes."
                        )

                    else:
                        nuevos_datos = {
                            "ID": registro_actual.get("ID", ""),
                            "Fecha Registro": fecha_registro,
                            "Fecha Actividad": fecha_actividad_editada.strftime("%d/%m/%Y"),
                            "Hora Actividad": hora_actividad_editada.strftime("%H:%M"),
                            "Dirección Regional": nueva_region,
                            "Delegación": nueva_delegacion,
                            "Responde a": nuevo_responde_a,
                            "Programa": nuevo_programa,
                            "Actividad": nueva_actividad,
                            "Provincia": nueva_provincia,
                            "Cantón": nuevo_canton,
                            "Distrito": nuevo_distrito,
                            "Tipo Lugar": tipo_lugar_editado,
                            "Lugar": lugar_editado,
                            "Centro Educativo": centro_editado,
                            "Código Presupuestario": codigo_presupuestario_editado,
                            "Dirección Mapa": nueva_direccion_mapa,
                            "Latitud": nueva_latitud,
                            "Longitud": nueva_longitud,
                            "Responsable": nuevo_responsable,
                            "Cantidad Participantes": nueva_cantidad,
                            "Cantidad Hombres": nueva_cantidad_hombres,
                            "Cantidad Mujeres": nueva_cantidad_mujeres,
                            "Edad 10 a 18": nueva_edad_10_18,
                            "Edad 19 a 30": nueva_edad_19_30,
                            "Edad 31 a 45": nueva_edad_31_45,
                            "Edad 46 en adelante": nueva_edad_46,
                            "Instituciones Participantes": nuevas_instituciones,
                            "Plan Estratégico Relacionado": nuevo_plan,
                            "Número de Referencia": nuevo_numero_referencia,
                            "Número de Expediente Referencia": nuevo_numero_expediente,
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
# PARTE 10 DE 12
# DASHBOARD PROFESIONAL: FILTROS, MÉTRICAS, MAPA Y GRÁFICOS
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

        if "Responde a" not in df_dash.columns:
            df_dash["Responde a"] = ""

        st.markdown(
            """
            <div class="card-azul">
                <div class="texto-pumi">
                    Panel administrativo para analizar registros por responde a, programa,
                    delegación, provincia, estado de revisión, participantes, sexo,
                    rangos de edad, comportamiento mensual y ubicación geográfica.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        perfil_actual = st.session_state.get("admin_perfil", {})

        if perfil_actual:
            st.info(
                f"Perfil activo: {perfil_actual.get('descripcion', 'Perfil administrativo')}"
            )

        # ======================================================
        # FILTROS DASHBOARD
        # ======================================================

        st.markdown("### Filtros del dashboard")

        col1, col2, col3 = st.columns(3)

        with col1:
            filtro_responde_a_dash = st.selectbox(
                "Responde a",
                ["Todos"] + sorted(
                    df_dash["Responde a"].dropna().astype(str).unique().tolist()
                ),
                key="dash_responde_a"
            )

            filtro_programa_dash = st.selectbox(
                "Programa",
                ["Todos"] + sorted(
                    df_dash["Programa"].dropna().astype(str).unique().tolist()
                ),
                key="dash_programa"
            )

        with col2:
            filtro_delegacion_dash = st.selectbox(
                "Delegación",
                ["Todas"] + sorted(
                    df_dash["Delegación"].dropna().astype(str).unique().tolist()
                ),
                key="dash_delegacion"
            )

            filtro_region_dash = st.selectbox(
                "Región",
                ["Todas"] + ordenar_regiones_numericamente(
                    sorted(
                        df_dash["Dirección Regional"]
                        .dropna()
                        .astype(str)
                        .unique()
                        .tolist()
                    )
                ),
                key="dash_region"
            )

        with col3:
            filtro_provincia_dash = st.selectbox(
                "Provincia",
                ["Todas"] + sorted(
                    df_dash["Provincia"].dropna().astype(str).unique().tolist()
                ),
                key="dash_provincia"
            )

            filtro_estado_dash = st.selectbox(
                "Estado de revisión",
                ["Todos"] + sorted(
                    df_dash["Estado Revisión"].dropna().astype(str).unique().tolist()
                ),
                key="dash_estado"
            )

            usar_fechas_dash = st.checkbox(
                "Filtrar por rango de fechas",
                key="dash_fechas"
            )

        df_filtrado = df_dash.copy()

        if filtro_responde_a_dash != "Todos":
            df_filtrado = df_filtrado[
                df_filtrado["Responde a"] == filtro_responde_a_dash
            ]

        if filtro_programa_dash != "Todos":
            df_filtrado = df_filtrado[
                df_filtrado["Programa"] == filtro_programa_dash
            ]

        if filtro_delegacion_dash != "Todas":
            df_filtrado = df_filtrado[
                df_filtrado["Delegación"] == filtro_delegacion_dash
            ]

        if filtro_region_dash != "Todas":
            df_filtrado = df_filtrado[
                df_filtrado["Dirección Regional"] == filtro_region_dash
            ]

        if filtro_provincia_dash != "Todas":
            df_filtrado = df_filtrado[
                df_filtrado["Provincia"] == filtro_provincia_dash
            ]

        if filtro_estado_dash != "Todos":
            df_filtrado = df_filtrado[
                df_filtrado["Estado Revisión"] == filtro_estado_dash
            ]

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
                    (df_filtrado["Fecha Actividad"].dt.date >= fecha_inicio_dash)
                    &
                    (df_filtrado["Fecha Actividad"].dt.date <= fecha_fin_dash)
                ]

        # ======================================================
        # MÉTRICAS
        # ======================================================

        st.markdown("### Indicadores generales")

        total_registros = len(df_filtrado)
        total_responde_a = df_filtrado["Responde a"].nunique()
        total_programas = df_filtrado["Programa"].nunique()
        total_delegaciones = df_filtrado["Delegación"].nunique()

        total_participantes = (
            int(df_filtrado["Cantidad Participantes"].sum())
            if not df_filtrado.empty else 0
        )

        total_hombres = (
            int(df_filtrado["Cantidad Hombres"].sum())
            if "Cantidad Hombres" in df_filtrado.columns else 0
        )

        total_mujeres = (
            int(df_filtrado["Cantidad Mujeres"].sum())
            if "Cantidad Mujeres" in df_filtrado.columns else 0
        )

        edad_10_18 = (
            int(df_filtrado["Edad 10 a 18"].sum())
            if "Edad 10 a 18" in df_filtrado.columns else 0
        )

        edad_19_30 = (
            int(df_filtrado["Edad 19 a 30"].sum())
            if "Edad 19 a 30" in df_filtrado.columns else 0
        )

        edad_31_45 = (
            int(df_filtrado["Edad 31 a 45"].sum())
            if "Edad 31 a 45" in df_filtrado.columns else 0
        )

        edad_46 = (
            int(df_filtrado["Edad 46 en adelante"].sum())
            if "Edad 46 en adelante" in df_filtrado.columns else 0
        )

        total_aprobados = len(
            df_filtrado[df_filtrado["Estado Revisión"] == "Aprobado"]
        )

        total_pendientes = len(
            df_filtrado[df_filtrado["Estado Revisión"] == "Pendiente de revisión"]
        )

        total_observaciones = len(
            df_filtrado[df_filtrado["Estado Revisión"] == "Con observaciones"]
        )

        total_rechazados = len(
            df_filtrado[df_filtrado["Estado Revisión"] == "Rechazado"]
        )

        colm1, colm2, colm3, colm4 = st.columns(4)

        colm1.metric("Registros", total_registros)
        colm2.metric("Responde a", total_responde_a)
        colm3.metric("Programas", total_programas)
        colm4.metric("Delegaciones", total_delegaciones)

        colm5, colm6, colm7, colm8 = st.columns(4)

        colm5.metric("Participantes", total_participantes)

        colm6.metric(
            "Hombres",
            f"{total_hombres} ({porcentaje(total_hombres, total_participantes)})"
        )

        colm7.metric(
            "Mujeres",
            f"{total_mujeres} ({porcentaje(total_mujeres, total_participantes)})"
        )

        colm8.metric(
            "Edad 10 a 18",
            f"{edad_10_18} ({porcentaje(edad_10_18, total_participantes)})"
        )

        colm9, colm10, colm11, colm12 = st.columns(4)

        colm9.metric(
            "Edad 19 a 30",
            f"{edad_19_30} ({porcentaje(edad_19_30, total_participantes)})"
        )

        colm10.metric(
            "Edad 31 a 45",
            f"{edad_31_45} ({porcentaje(edad_31_45, total_participantes)})"
        )

        colm11.metric(
            "Edad 46 en adelante",
            f"{edad_46} ({porcentaje(edad_46, total_participantes)})"
        )

        colm12.metric(
            "Aprobados",
            f"{total_aprobados} ({porcentaje(total_aprobados, total_registros)})"
        )

        colm13, colm14, colm15 = st.columns(3)

        colm13.metric(
            "Pendientes",
            f"{total_pendientes} ({porcentaje(total_pendientes, total_registros)})"
        )

        colm14.metric(
            "Con observaciones",
            f"{total_observaciones} ({porcentaje(total_observaciones, total_registros)})"
        )

        colm15.metric(
            "Rechazados",
            f"{total_rechazados} ({porcentaje(total_rechazados, total_registros)})"
        )

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
                height=650,
                key="mapa_dashboard_admin"
            )

            st.markdown("---")

            # ======================================================
            # GRÁFICOS
            # ======================================================

            colg1, colg2 = st.columns(2)

            with colg1:
                st.markdown("### Actividades por Responde a")

                data_responde_a = df_filtrado.groupby(
                    "Responde a",
                    as_index=False
                )["ID"].count()

                data_responde_a["Porcentaje"] = data_responde_a["ID"].apply(
                    lambda x: porcentaje(x, data_responde_a["ID"].sum())
                )

                data_responde_a["Etiqueta"] = data_responde_a.apply(
                    lambda row: f"{int(row['ID'])} ({row['Porcentaje']})",
                    axis=1
                )

                fig_responde_a = px.bar(
                    data_responde_a,
                    x="Responde a",
                    y="ID",
                    text="Etiqueta",
                    labels={"ID": "Cantidad de actividades"},
                    color="Responde a",
                    template="plotly_white"
                )

                fig_responde_a.update_layout(
                    showlegend=False,
                    title_x=0.5,
                    xaxis_tickangle=-35
                )

                st.plotly_chart(
                    fig_responde_a,
                    use_container_width=True
                )

            with colg2:
                st.markdown("### Actividades por programa")

                data_programa = df_filtrado.groupby(
                    "Programa",
                    as_index=False
                )["ID"].count()

                data_programa["Porcentaje"] = data_programa["ID"].apply(
                    lambda x: porcentaje(x, data_programa["ID"].sum())
                )

                data_programa["Etiqueta"] = data_programa.apply(
                    lambda row: f"{int(row['ID'])} ({row['Porcentaje']})",
                    axis=1
                )

                fig_programa = px.bar(
                    data_programa,
                    x="Programa",
                    y="ID",
                    text="Etiqueta",
                    labels={"ID": "Cantidad de actividades"},
                    color="Programa",
                    template="plotly_white"
                )

                fig_programa.update_layout(
                    showlegend=False,
                    title_x=0.5
                )

                st.plotly_chart(
                    fig_programa,
                    use_container_width=True
                )

            colg3, colg4 = st.columns(2)

            with colg3:
                st.markdown("### Participantes por programa")

                data_participantes = df_filtrado.groupby(
                    "Programa",
                    as_index=False
                )["Cantidad Participantes"].sum()

                data_participantes["Porcentaje"] = data_participantes[
                    "Cantidad Participantes"
                ].apply(
                    lambda x: porcentaje(
                        x,
                        data_participantes["Cantidad Participantes"].sum()
                    )
                )

                data_participantes["Etiqueta"] = data_participantes.apply(
                    lambda row: (
                        f"{int(row['Cantidad Participantes'])} "
                        f"({row['Porcentaje']})"
                    ),
                    axis=1
                )

                fig_participantes = px.bar(
                    data_participantes,
                    x="Programa",
                    y="Cantidad Participantes",
                    text="Etiqueta",
                    color="Programa",
                    template="plotly_white"
                )

                fig_participantes.update_layout(
                    showlegend=False,
                    title_x=0.5
                )

                st.plotly_chart(
                    fig_participantes,
                    use_container_width=True
                )

            with colg4:
                st.markdown("### Estado de revisión")

                fig_estado = px.pie(
                    df_filtrado,
                    names="Estado Revisión",
                    hole=0.45,
                    template="plotly_white"
                )

                fig_estado.update_traces(
                    textinfo="label+value+percent"
                )

                st.plotly_chart(
                    fig_estado,
                    use_container_width=True
                )

            st.markdown("### Actividades por provincia")

            data_provincia = df_filtrado.groupby(
                "Provincia",
                as_index=False
            )["ID"].count()

            data_provincia["Porcentaje"] = data_provincia["ID"].apply(
                lambda x: porcentaje(x, data_provincia["ID"].sum())
            )

            data_provincia["Etiqueta"] = data_provincia.apply(
                lambda row: f"{int(row['ID'])} ({row['Porcentaje']})",
                axis=1
            )

            fig_provincia = px.bar(
                data_provincia,
                x="Provincia",
                y="ID",
                text="Etiqueta",
                labels={"ID": "Cantidad de actividades"},
                color="Provincia",
                template="plotly_white"
            )

            fig_provincia.update_layout(
                showlegend=False,
                title_x=0.5
            )

            st.plotly_chart(
                fig_provincia,
                use_container_width=True
            )
# ======================================================
# PARTE 11 DE 12
# DASHBOARD PROFESIONAL: GRÁFICOS COMPLEMENTARIOS, TABLA E INFORME PDF
# ======================================================

            colg5, colg6 = st.columns(2)

            with colg5:
                st.markdown("### Ranking de delegaciones")

                data_delegacion = df_filtrado.groupby(
                    "Delegación",
                    as_index=False
                )["ID"].count().sort_values(
                    "ID",
                    ascending=False
                ).head(15)

                data_delegacion["Porcentaje"] = data_delegacion["ID"].apply(
                    lambda x: porcentaje(x, data_delegacion["ID"].sum())
                )

                data_delegacion["Etiqueta"] = data_delegacion.apply(
                    lambda row: f"{int(row['ID'])} ({row['Porcentaje']})",
                    axis=1
                )

                fig_delegacion = px.bar(
                    data_delegacion,
                    x="ID",
                    y="Delegación",
                    orientation="h",
                    text="Etiqueta",
                    labels={"ID": "Cantidad de actividades"},
                    color="ID",
                    template="plotly_white"
                )

                fig_delegacion.update_layout(title_x=0.5)

                st.plotly_chart(
                    fig_delegacion,
                    use_container_width=True
                )

            with colg6:
                st.markdown("### Registros por tipo de lugar")

                if "Tipo Lugar" in df_filtrado.columns:
                    fig_tipo_lugar = px.pie(
                        df_filtrado,
                        names="Tipo Lugar",
                        hole=0.35,
                        template="plotly_white"
                    )

                    fig_tipo_lugar.update_traces(
                        textinfo="label+value+percent"
                    )

                    st.plotly_chart(
                        fig_tipo_lugar,
                        use_container_width=True
                    )
                else:
                    st.info("No existe la columna Tipo Lugar.")

            colg7, colg8 = st.columns(2)

            with colg7:
                st.markdown("### Distribución por sexo")

                data_sexo = pd.DataFrame(
                    {
                        "Sexo": ["Hombres", "Mujeres"],
                        "Cantidad": [total_hombres, total_mujeres]
                    }
                )

                if data_sexo["Cantidad"].sum() > 0:
                    fig_sexo = px.pie(
                        data_sexo,
                        names="Sexo",
                        values="Cantidad",
                        hole=0.40,
                        template="plotly_white"
                    )

                    fig_sexo.update_traces(
                        textinfo="label+value+percent"
                    )

                    st.plotly_chart(
                        fig_sexo,
                        use_container_width=True
                    )
                else:
                    st.info("No hay datos de sexo registrados.")

            with colg8:
                st.markdown("### Distribución por rangos de edad")

                data_edades = pd.DataFrame(
                    {
                        "Rango de edad": [
                            "10 a 18",
                            "19 a 30",
                            "31 a 45",
                            "46 en adelante"
                        ],
                        "Cantidad": [
                            edad_10_18,
                            edad_19_30,
                            edad_31_45,
                            edad_46
                        ]
                    }
                )

                if data_edades["Cantidad"].sum() > 0:
                    data_edades["Porcentaje"] = data_edades["Cantidad"].apply(
                        lambda x: porcentaje(x, data_edades["Cantidad"].sum())
                    )

                    data_edades["Etiqueta"] = data_edades.apply(
                        lambda row: f"{int(row['Cantidad'])} ({row['Porcentaje']})",
                        axis=1
                    )

                    fig_edades = px.bar(
                        data_edades,
                        x="Rango de edad",
                        y="Cantidad",
                        text="Etiqueta",
                        color="Rango de edad",
                        template="plotly_white"
                    )

                    fig_edades.update_layout(
                        showlegend=False,
                        title_x=0.5
                    )

                    st.plotly_chart(
                        fig_edades,
                        use_container_width=True
                    )
                else:
                    st.info("No hay datos de rangos de edad registrados.")

            # ======================================================
            # EVOLUCIÓN MENSUAL
            # ======================================================

            st.markdown("### Evolución mensual de actividades")

            df_mensual = df_filtrado.dropna(
                subset=["Fecha Actividad"]
            ).copy()

            if not df_mensual.empty:
                df_mensual["Mes"] = df_mensual[
                    "Fecha Actividad"
                ].dt.to_period("M").astype(str)

                data_mensual = df_mensual.groupby(
                    "Mes",
                    as_index=False
                )["ID"].count()

                data_mensual["Porcentaje"] = data_mensual["ID"].apply(
                    lambda x: porcentaje(x, data_mensual["ID"].sum())
                )

                data_mensual["Etiqueta"] = data_mensual.apply(
                    lambda row: f"{int(row['ID'])} ({row['Porcentaje']})",
                    axis=1
                )

                fig_mensual = px.line(
                    data_mensual,
                    x="Mes",
                    y="ID",
                    markers=True,
                    text="Etiqueta",
                    labels={"ID": "Cantidad de actividades"},
                    template="plotly_white"
                )

                fig_mensual.update_layout(title_x=0.5)

                st.plotly_chart(
                    fig_mensual,
                    use_container_width=True
                )
            else:
                st.info("No hay fechas válidas para generar evolución mensual.")

            # ======================================================
            # TABLA DE REGISTROS EN DASHBOARD
            # ======================================================

            st.markdown("### Registros analizados")

            df_mostrar = df_filtrado.copy()

            if "Fecha Actividad" in df_mostrar.columns:
                df_mostrar["Fecha Actividad"] = df_mostrar[
                    "Fecha Actividad"
                ].dt.strftime("%d/%m/%Y")

            columnas_dashboard = [
                "ID",
                "Fecha Registro",
                "Fecha Actividad",
                "Hora Actividad",
                "Dirección Regional",
                "Delegación",
                "Responde a",
                "Programa",
                "Actividad",
                "Provincia",
                "Cantón",
                "Distrito",
                "Tipo Lugar",
                "Lugar",
                "Centro Educativo",
                "Código Presupuestario",
                "Dirección Mapa",
                "Latitud",
                "Longitud",
                "Responsable",
                "Cantidad Participantes",
                "Cantidad Hombres",
                "Cantidad Mujeres",
                "Edad 10 a 18",
                "Edad 19 a 30",
                "Edad 31 a 45",
                "Edad 46 en adelante",
                "Instituciones Participantes",
                "Plan Estratégico Relacionado",
                "Número de Referencia",
                "Número de Expediente Referencia",
                "Observaciones",
                "Estado Revisión",
                "Observación de Revisión",
                "Usuario Registra"
            ]

            columnas_dashboard_existentes = [
                col for col in columnas_dashboard
                if col in df_mostrar.columns
            ]

            st.dataframe(
                df_mostrar[columnas_dashboard_existentes],
                use_container_width=True,
                hide_index=True
            )

            # ======================================================
            # CONFIGURACIÓN DEL INFORME PDF
            # ======================================================

            st.markdown("---")
            st.markdown("## Generar informe PDF")

            col_pdf1, col_pdf2 = st.columns(2)

            with col_pdf1:
                dirigido_a = st.text_input(
                    "Informe dirigido a",
                    placeholder="Ejemplo: Delegación Policial D88-Limón o Dirección Regional R9 Limón"
                )

                fecha_informe = st.date_input(
                    "Fecha del informe",
                    value=date.today(),
                    format="DD/MM/YYYY",
                    key="fecha_informe_pdf"
                )

            with col_pdf2:
                hora_informe = st.time_input(
                    "Hora del informe",
                    value=datetime.now().time().replace(second=0, microsecond=0),
                    key="hora_informe_pdf"
                )

                incluir_detalle = st.checkbox(
                    "Incluir detalle de registros en formato ficha",
                    value=True
                )

            graficos_pdf = st.multiselect(
                "Seleccione los gráficos que desea incluir en el informe",
                [
                    "Actividades por Responde a",
                    "Actividades por programa",
                    "Participantes por programa",
                    "Estado de revisión",
                    "Actividades por provincia",
                    "Ranking de delegaciones",
                    "Tipo de lugar",
                    "Distribución por sexo",
                    "Distribución por rangos de edad",
                    "Evolución mensual"
                ],
                default=[
                    "Actividades por Responde a",
                    "Actividades por programa",
                    "Participantes por programa",
                    "Estado de revisión",
                    "Distribución por sexo",
                    "Distribución por rangos de edad"
                ]
            )

            # ======================================================
            # FUNCIONES AUXILIARES PARA PDF
            # ======================================================

            def obtener_logo_existente(ruta):
                if ruta and os.path.exists(ruta):
                    return ruta
                return None


            def agregar_marco_y_paginacion(canvas, doc):
                canvas.saveState()

                width, height = letter

                canvas.setStrokeColor(colors.HexColor(COLOR_AZUL))
                canvas.setLineWidth(2)
                canvas.line(42, height - 52, width - 42, height - 52)

                canvas.setStrokeColor(colors.HexColor(COLOR_DORADO))
                canvas.setLineWidth(1.5)
                canvas.line(42, height - 57, width - 42, height - 57)

                canvas.setStrokeColor(colors.HexColor(COLOR_AZUL))
                canvas.setLineWidth(2)
                canvas.rect(
                    24,
                    24,
                    width - 48,
                    height - 96
                )

                canvas.setStrokeColor(colors.HexColor(COLOR_DORADO))
                canvas.setLineWidth(1)
                canvas.rect(
                    30,
                    30,
                    width - 60,
                    height - 108
                )

                canvas.setStrokeColor(colors.HexColor(COLOR_AZUL))
                canvas.setLineWidth(1)
                canvas.line(42, 43, 205, 43)

                canvas.setStrokeColor(colors.HexColor(COLOR_DORADO))
                canvas.setLineWidth(1)
                canvas.line(42, 38, 205, 38)

                canvas.setStrokeColor(colors.HexColor(COLOR_AZUL))
                canvas.setLineWidth(1)
                canvas.line(width - 205, 43, width - 42, 43)

                canvas.setStrokeColor(colors.HexColor(COLOR_DORADO))
                canvas.setLineWidth(1)
                canvas.line(width - 205, 38, width - 42, 38)

                canvas.setFont("Helvetica", 8)
                canvas.setFillColor(colors.HexColor(COLOR_GRIS_OSCURO))
                canvas.drawCentredString(
                    width / 2,
                    36,
                    f"Página {doc.page}"
                )

                canvas.restoreState()


            def crear_tabla_estilizada(data, col_widths=None, header_color=None, font_size=7):
                if header_color is None:
                    header_color = COLOR_AZUL

                tabla = Table(
                    data,
                    colWidths=col_widths,
                    repeatRows=1
                )

                tabla.setStyle(
                    TableStyle(
                        [
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(header_color)),
                            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                            ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
                            ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("FONTSIZE", (0, 0), (-1, -1), font_size),
                            ("LEFTPADDING", (0, 0), (-1, -1), 4),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                            ("TOPPADDING", (0, 0), (-1, -1), 5),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                        ]
                    )
                )

                return tabla


            def agregar_imagen_plotly_al_pdf(
                elementos_pdf,
                figura,
                titulo_grafico,
                estilo_subtitulo,
                estilo_texto
            ):
                try:
                    img_buffer = BytesIO()

                    figura.update_layout(
                        title=None,
                        margin=dict(l=40, r=40, t=40, b=40),
                        paper_bgcolor="white",
                        plot_bgcolor="white"
                    )

                    figura.write_image(
                        img_buffer,
                        format="png",
                        width=900,
                        height=500,
                        scale=2
                    )

                    img_buffer.seek(0)

                    imagen = RLImage(
                        img_buffer,
                        width=460,
                        height=250
                    )

                    imagen.hAlign = "CENTER"

                    elementos_pdf.append(
                        KeepTogether(
                            [
                                Paragraph(titulo_grafico, estilo_subtitulo),
                                imagen,
                                Spacer(1, 12)
                            ]
                        )
                    )

                except Exception:
                    elementos_pdf.append(
                        KeepTogether(
                            [
                                Paragraph(titulo_grafico, estilo_subtitulo),
                                Paragraph(
                                    "No fue posible insertar el gráfico como imagen. "
                                    "Agregue kaleido==0.2.1 al archivo requirements.txt y vuelva a desplegar la aplicación.",
                                    estilo_texto
                                ),
                                Spacer(1, 10)
                            ]
                        )
                    )


            def crear_tabla_logos_pdf(logo_min, logo_pumi, logo_fp):
                tabla_logos_data = [
                    [
                        RLImage(logo_min, width=100, height=40) if logo_min else "",
                        RLImage(logo_pumi, width=135, height=40) if logo_pumi else "",
                        RLImage(logo_fp, width=170, height=60) if logo_fp else ""
                    ]
                ]

                tabla_logos = Table(
                    tabla_logos_data,
                    colWidths=[165, 150, 185]
                )

                tabla_logos.setStyle(
                    TableStyle(
                        [
                            ("ALIGN", (0, 0), (0, 0), "LEFT"),
                            ("ALIGN", (1, 0), (1, 0), "CENTER"),
                            ("ALIGN", (2, 0), (2, 0), "RIGHT"),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("BOX", (0, 0), (-1, -1), 0, colors.white),
                            ("LEFTPADDING", (0, 0), (-1, -1), 0),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                            ("TOPPADDING", (0, 0), (-1, -1), 2),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                        ]
                    )
                )

                return tabla_logos

            # ======================================================
            # GENERACIÓN DEL PDF PROFESIONAL
            # ======================================================

            def generar_pdf_profesional(
                df_pdf,
                dirigido_a_pdf,
                fecha_pdf,
                hora_pdf,
                graficos_seleccionados,
                incluir_detalle_pdf
            ):
                buffer = BytesIO()

                doc = SimpleDocTemplate(
                    buffer,
                    pagesize=letter,
                    rightMargin=48,
                    leftMargin=48,
                    topMargin=78,
                    bottomMargin=62
                )

                elementos = []
                estilos = getSampleStyleSheet()

                titulo = ParagraphStyle(
                    "TituloPUMI",
                    parent=estilos["Title"],
                    alignment=TA_CENTER,
                    textColor=colors.HexColor(COLOR_AZUL),
                    fontSize=21,
                    leading=25,
                    spaceAfter=12
                )

                subtitulo = ParagraphStyle(
                    "SubtituloPUMI",
                    parent=estilos["Heading2"],
                    alignment=TA_CENTER,
                    textColor=colors.HexColor(COLOR_DORADO),
                    fontSize=14,
                    leading=18,
                    spaceBefore=8,
                    spaceAfter=10
                )

                texto = ParagraphStyle(
                    "TextoPUMI",
                    parent=estilos["Normal"],
                    alignment=TA_JUSTIFY,
                    fontSize=9,
                    leading=13
                )

                texto_centrado = ParagraphStyle(
                    "TextoCentradoPUMI",
                    parent=estilos["Normal"],
                    alignment=TA_CENTER,
                    fontSize=9,
                    leading=13
                )

                estilo_celda = ParagraphStyle(
                    "CeldaTablaPDF",
                    parent=estilos["Normal"],
                    fontSize=7,
                    leading=9,
                    alignment=TA_JUSTIFY,
                    wordWrap="CJK"
                )

                estilo_campo = ParagraphStyle(
                    "CampoFichaPDF",
                    parent=estilos["Normal"],
                    fontSize=7,
                    leading=9,
                    alignment=TA_JUSTIFY,
                    textColor=colors.HexColor(COLOR_AZUL),
                    wordWrap="CJK"
                )

                logo_min = obtener_logo_existente(LOGO_MINISTERIO)
                logo_pumi = obtener_logo_existente(LOGO_PUMI)
                logo_fp = obtener_logo_existente(LOGO_FUERZA_PUBLICA)

                elementos.append(
                    crear_tabla_logos_pdf(
                        logo_min,
                        logo_pumi,
                        logo_fp
                    )
                )

                elementos.append(Spacer(1, 24))

                elementos.append(
                    Paragraph(
                        "INFORME ADMINISTRATIVO PUMI 2026",
                        titulo
                    )
                )

                elementos.append(
                    Paragraph(
                        "Proceso Unificado para el Manejo de la Información",
                        subtitulo
                    )
                )

                if dirigido_a_pdf:
                    elementos.append(
                        Paragraph(
                            f"<b>Dirigido a:</b> {dirigido_a_pdf}",
                            texto_centrado
                        )
                    )

                elementos.append(
                    Paragraph(
                        f"<b>Fecha y hora del informe:</b> "
                        f"{fecha_pdf.strftime('%d/%m/%Y')} {hora_pdf.strftime('%H:%M')}",
                        texto_centrado
                    )
                )

                elementos.append(Spacer(1, 14))

                introduccion = """
                El presente informe consolida la información registrada en el Sistema PUMI 2026,
                correspondiente a las actividades preventivas desarrolladas por los Programas
                Policiales Preventivos. La información permite observar registros, responde a,
                programas, delegaciones, población alcanzada, distribución por sexo, rangos de edad,
                estado de revisión, observaciones administrativas y ubicación geográfica.
                """

                elementos.append(Paragraph(introduccion, texto))
                elementos.append(Spacer(1, 16))

                total_registros_pdf = len(df_pdf)
                total_responde_a_pdf = df_pdf["Responde a"].nunique() if "Responde a" in df_pdf.columns else 0
                total_programas_pdf = df_pdf["Programa"].nunique()
                total_delegaciones_pdf = df_pdf["Delegación"].nunique()

                total_participantes_pdf = (
                    int(df_pdf["Cantidad Participantes"].sum())
                    if "Cantidad Participantes" in df_pdf.columns
                    else 0
                )

                total_hombres_pdf = (
                    int(df_pdf["Cantidad Hombres"].sum())
                    if "Cantidad Hombres" in df_pdf.columns
                    else 0
                )

                total_mujeres_pdf = (
                    int(df_pdf["Cantidad Mujeres"].sum())
                    if "Cantidad Mujeres" in df_pdf.columns
                    else 0
                )

                edad_10_18_pdf = (
                    int(df_pdf["Edad 10 a 18"].sum())
                    if "Edad 10 a 18" in df_pdf.columns
                    else 0
                )

                edad_19_30_pdf = (
                    int(df_pdf["Edad 19 a 30"].sum())
                    if "Edad 19 a 30" in df_pdf.columns
                    else 0
                )

                edad_31_45_pdf = (
                    int(df_pdf["Edad 31 a 45"].sum())
                    if "Edad 31 a 45" in df_pdf.columns
                    else 0
                )

                edad_46_pdf = (
                    int(df_pdf["Edad 46 en adelante"].sum())
                    if "Edad 46 en adelante" in df_pdf.columns
                    else 0
                )

                total_aprobados_pdf = len(
                    df_pdf[df_pdf["Estado Revisión"] == "Aprobado"]
                )

                total_pendientes_pdf = len(
                    df_pdf[df_pdf["Estado Revisión"] == "Pendiente de revisión"]
                )

                total_observaciones_pdf = len(
                    df_pdf[df_pdf["Estado Revisión"] == "Con observaciones"]
                )

                total_rechazados_pdf = len(
                    df_pdf[df_pdf["Estado Revisión"] == "Rechazado"]
                )

                total_georreferenciados_pdf = len(
                    preparar_dataframe_mapa(df_pdf)
                )

                resumen = [
                    ["Indicador", "Resultado"],
                    ["Total de registros", f"{total_registros_pdf}"],
                    ["Responde a registrados", f"{total_responde_a_pdf}"],
                    ["Programas registrados", f"{total_programas_pdf}"],
                    ["Delegaciones registradas", f"{total_delegaciones_pdf}"],
                    ["Participantes atendidos", f"{total_participantes_pdf}"],
                    ["Hombres", f"{total_hombres_pdf} ({porcentaje(total_hombres_pdf, total_participantes_pdf)})"],
                    ["Mujeres", f"{total_mujeres_pdf} ({porcentaje(total_mujeres_pdf, total_participantes_pdf)})"],
                    ["Edad 10 a 18", f"{edad_10_18_pdf} ({porcentaje(edad_10_18_pdf, total_participantes_pdf)})"],
                    ["Edad 19 a 30", f"{edad_19_30_pdf} ({porcentaje(edad_19_30_pdf, total_participantes_pdf)})"],
                    ["Edad 31 a 45", f"{edad_31_45_pdf} ({porcentaje(edad_31_45_pdf, total_participantes_pdf)})"],
                    ["Edad 46 en adelante", f"{edad_46_pdf} ({porcentaje(edad_46_pdf, total_participantes_pdf)})"],
                    ["Registros georreferenciados", f"{total_georreferenciados_pdf} ({porcentaje(total_georreferenciados_pdf, total_registros_pdf)})"],
                    ["Registros aprobados", f"{total_aprobados_pdf} ({porcentaje(total_aprobados_pdf, total_registros_pdf)})"],
                    ["Pendientes de revisión", f"{total_pendientes_pdf} ({porcentaje(total_pendientes_pdf, total_registros_pdf)})"],
                    ["Con observaciones", f"{total_observaciones_pdf} ({porcentaje(total_observaciones_pdf, total_registros_pdf)})"],
                    ["Rechazados", f"{total_rechazados_pdf} ({porcentaje(total_rechazados_pdf, total_registros_pdf)})"]
                ]

                tabla_resumen = crear_tabla_estilizada(
                    resumen,
                    col_widths=[260, 190],
                    header_color=COLOR_AZUL,
                    font_size=8
                )

                elementos.append(
                    KeepTogether(
                        [
                            Paragraph("Resumen general", subtitulo),
                            tabla_resumen
                        ]
                    )
                )

                elementos.append(Spacer(1, 18))

                if "Responde a" in df_pdf.columns:
                    resumen_responde_a = df_pdf.groupby("Responde a").agg(
                        {
                            "ID": "count",
                            "Cantidad Participantes": "sum"
                        }
                    ).reset_index()

                    tabla_responde_a = [
                        [
                            "Responde a",
                            "Actividades",
                            "% Actividades",
                            "Participantes",
                            "% Participantes"
                        ]
                    ]

                    total_actividades_responde = resumen_responde_a["ID"].sum()
                    total_participantes_responde = resumen_responde_a[
                        "Cantidad Participantes"
                    ].sum()

                    for _, row in resumen_responde_a.iterrows():
                        tabla_responde_a.append(
                            [
                                str(row["Responde a"]),
                                int(row["ID"]),
                                porcentaje(row["ID"], total_actividades_responde),
                                int(row["Cantidad Participantes"]),
                                porcentaje(
                                    row["Cantidad Participantes"],
                                    total_participantes_responde
                                )
                            ]
                        )

                    tabla_resp = crear_tabla_estilizada(
                        tabla_responde_a,
                        col_widths=[180, 65, 70, 75, 75],
                        header_color=COLOR_AZUL,
                        font_size=6
                    )

                    elementos.append(
                        KeepTogether(
                            [
                                Paragraph("Distribución por Responde a", subtitulo),
                                tabla_resp
                            ]
                        )
                    )

                    elementos.append(Spacer(1, 18))

                resumen_programa = df_pdf.groupby("Programa").agg(
                    {
                        "ID": "count",
                        "Cantidad Participantes": "sum"
                    }
                ).reset_index()

                tabla_programas = [
                    [
                        "Programa",
                        "Actividades",
                        "% Actividades",
                        "Participantes",
                        "% Participantes"
                    ]
                ]

                total_actividades_programa = resumen_programa["ID"].sum()
                total_participantes_programa = resumen_programa[
                    "Cantidad Participantes"
                ].sum()

                for _, row in resumen_programa.iterrows():
                    tabla_programas.append(
                        [
                            str(row["Programa"]),
                            int(row["ID"]),
                            porcentaje(row["ID"], total_actividades_programa),
                            int(row["Cantidad Participantes"]),
                            porcentaje(
                                row["Cantidad Participantes"],
                                total_participantes_programa
                            )
                        ]
                    )

                tabla_prog = crear_tabla_estilizada(
                    tabla_programas,
                    col_widths=[125, 75, 80, 85, 85],
                    header_color=COLOR_DORADO,
                    font_size=7
                )

                elementos.append(
                    KeepTogether(
                        [
                            Paragraph("Distribución por programa", subtitulo),
                            tabla_prog
                        ]
                    )
                )

                elementos.append(Spacer(1, 18))

                if graficos_seleccionados:
                    elementos.append(PageBreak())
                    elementos.append(
                        Paragraph("Gráficos seleccionados", titulo)
                    )

                    if "Actividades por Responde a" in graficos_seleccionados:
                        agregar_imagen_plotly_al_pdf(
                            elementos,
                            fig_responde_a,
                            "Actividades por Responde a",
                            subtitulo,
                            texto
                        )

                    if "Actividades por programa" in graficos_seleccionados:
                        agregar_imagen_plotly_al_pdf(
                            elementos,
                            fig_programa,
                            "Actividades por programa",
                            subtitulo,
                            texto
                        )

                    if "Participantes por programa" in graficos_seleccionados:
                        agregar_imagen_plotly_al_pdf(
                            elementos,
                            fig_participantes,
                            "Participantes por programa",
                            subtitulo,
                            texto
                        )

                    if "Estado de revisión" in graficos_seleccionados:
                        agregar_imagen_plotly_al_pdf(
                            elementos,
                            fig_estado,
                            "Estado de revisión",
                            subtitulo,
                            texto
                        )

                    if "Actividades por provincia" in graficos_seleccionados:
                        agregar_imagen_plotly_al_pdf(
                            elementos,
                            fig_provincia,
                            "Actividades por provincia",
                            subtitulo,
                            texto
                        )

                    if "Ranking de delegaciones" in graficos_seleccionados:
                        agregar_imagen_plotly_al_pdf(
                            elementos,
                            fig_delegacion,
                            "Ranking de delegaciones",
                            subtitulo,
                            texto
                        )

                    if (
                        "Tipo de lugar" in graficos_seleccionados
                        and "Tipo Lugar" in df_pdf.columns
                    ):
                        agregar_imagen_plotly_al_pdf(
                            elementos,
                            fig_tipo_lugar,
                            "Registros por tipo de lugar",
                            subtitulo,
                            texto
                        )

                    if (
                        "Distribución por sexo" in graficos_seleccionados
                        and data_sexo["Cantidad"].sum() > 0
                    ):
                        agregar_imagen_plotly_al_pdf(
                            elementos,
                            fig_sexo,
                            "Distribución por sexo",
                            subtitulo,
                            texto
                        )

                    if (
                        "Distribución por rangos de edad" in graficos_seleccionados
                        and data_edades["Cantidad"].sum() > 0
                    ):
                        agregar_imagen_plotly_al_pdf(
                            elementos,
                            fig_edades,
                            "Distribución por rangos de edad",
                            subtitulo,
                            texto
                        )

                    if (
                        "Evolución mensual" in graficos_seleccionados
                        and not df_mensual.empty
                    ):
                        agregar_imagen_plotly_al_pdf(
                            elementos,
                            fig_mensual,
                            "Evolución mensual de actividades",
                            subtitulo,
                            texto
                        )

                # ==================================================
                # DETALLE DE REGISTROS EN FORMATO FICHA
                # ==================================================

                if incluir_detalle_pdf:
                    elementos.append(PageBreak())
                    elementos.append(
                        Paragraph("Detalle de registros", titulo)
                    )

                    df_detalle = df_pdf.copy()

                    if "Fecha Actividad" in df_detalle.columns:
                        df_detalle["Fecha Actividad"] = df_detalle[
                            "Fecha Actividad"
                        ].dt.strftime("%d/%m/%Y")

                    columnas_ficha = [
                        "ID",
                        "Fecha Actividad",
                        "Hora Actividad",
                        "Dirección Regional",
                        "Delegación",
                        "Responde a",
                        "Programa",
                        "Actividad",
                        "Provincia",
                        "Cantón",
                        "Distrito",
                        "Lugar",
                        "Cantidad Participantes",
                        "Estado Revisión"
                    ]

                    columnas_ficha = [
                        col for col in columnas_ficha
                        if col in df_detalle.columns
                    ]

                    for _, row in df_detalle[columnas_ficha].head(20).iterrows():
                        datos_ficha = [
                            [
                                Paragraph("<b>Campo</b>", estilo_campo),
                                Paragraph("<b>Detalle</b>", estilo_campo)
                            ]
                        ]

                        for col in columnas_ficha:
                            valor = str(row[col]).replace("\n", " ").strip()

                            datos_ficha.append(
                                [
                                    Paragraph(f"<b>{col}</b>", estilo_campo),
                                    Paragraph(valor, estilo_celda)
                                ]
                            )

                        tabla_ficha = crear_tabla_estilizada(
                            datos_ficha,
                            col_widths=[145, 330],
                            header_color=COLOR_AZUL,
                            font_size=7
                        )

                        tabla_ficha.setStyle(
                            TableStyle(
                                [
                                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                    ("BACKGROUND", (0, 1), (0, -1), colors.HexColor(COLOR_AZUL_CLARO)),
                                    ("TEXTCOLOR", (0, 1), (0, -1), colors.HexColor(COLOR_AZUL)),
                                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
                                ]
                            )
                        )

                        elementos.append(
                            KeepTogether(
                                [
                                    tabla_ficha,
                                    Spacer(1, 12)
                                ]
                            )
                        )

                    if len(df_detalle) > 20:
                        elementos.append(Spacer(1, 10))
                        elementos.append(
                            Paragraph(
                                f"Nota: se muestran los primeros 20 registros de un total de {len(df_detalle)} registros filtrados.",
                                texto
                            )
                        )

                # ==================================================
                # PÁGINA FINAL
                # ==================================================

                elementos.append(PageBreak())
                elementos.append(Spacer(1, 120))

                elementos.append(
                    crear_tabla_logos_pdf(
                        logo_min,
                        logo_pumi,
                        logo_fp
                    )
                )

                elementos.append(Spacer(1, 28))

                elementos.append(
                    Paragraph(
                        "Sistema PUMI 2026",
                        titulo
                    )
                )

                elementos.append(
                    Paragraph(
                        "Informe generado automáticamente.",
                        texto_centrado
                    )
                )

                elementos.append(Spacer(1, 20))

                elementos.append(
                    Paragraph(
                        "Elaborado por la Estrategia Integral de Prevención para la Seguridad Pública.",
                        texto_centrado
                    )
                )

                doc.build(
                    elementos,
                    onFirstPage=agregar_marco_y_paginacion,
                    onLaterPages=agregar_marco_y_paginacion
                )

                buffer.seek(0)
                return buffer.getvalue()


            pdf = generar_pdf_profesional(
                df_filtrado,
                dirigido_a,
                fecha_informe,
                hora_informe,
                graficos_pdf,
                incluir_detalle
            )

            st.download_button(
                label="Descargar informe administrativo en PDF",
                data=pdf,
                file_name="informe_administrativo_pumi_2026.pdf",
                mime="application/pdf"
            )
            # ======================================================
# PARTE 12 DE 12
# CONFIGURACIÓN, DIAGNÓSTICO, RESPALDOS Y CIERRE DEL SISTEMA
# ======================================================

elif menu == "Configuración":

    st.markdown("## Configuración y diagnóstico del sistema")

    if not st.session_state.admin_autenticado:
        st.error("Debe ingresar la clave administrativa para acceder a esta sección.")
        st.stop()

    st.markdown(
        """
        <div class="card-pumi">
            <div class="texto-pumi">
                Esta sección permite verificar el estado general del sistema,
                revisar la conexión con Google Sheets, validar los archivos base,
                consultar la estructura de columnas y descargar respaldos generales.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    # ======================================================
    # ESTADO DE CONEXIÓN
    # ======================================================

    st.markdown("### Estado de conexión")

    try:
        spreadsheet = conectar_google_sheets()
        hoja = inicializar_hoja()

        st.success("Conexión activa con Google Sheets.")
        st.write("Archivo conectado:", spreadsheet.title)
        st.write("Hoja principal:", HOJA_REGISTRO)

    except Exception as e:
        st.error("Error en la conexión con Google Sheets.")
        st.exception(e)
        st.stop()

    # ======================================================
    # VALIDACIÓN DE ARCHIVOS BASE
    # ======================================================

    st.markdown("### Archivos base del sistema")

    archivos_base = {
        "Base MEP": ARCHIVO_MEP,
        "Delegaciones y Distritos": ARCHIVO_DELEGACIONES,
        "Datos Importantes": ARCHIVO_DATOS_IMPORTANTES
    }

    for nombre, archivo in archivos_base.items():
        if os.path.exists(archivo):
            st.success(f"{nombre}: encontrado ({archivo})")
        else:
            st.error(f"{nombre}: no encontrado ({archivo})")

    # ======================================================
    # VALIDACIÓN DE DATOS CARGADOS
    # ======================================================

    st.markdown("### Validación de bases auxiliares")

    df_mep = cargar_base_mep()
    df_delegaciones = cargar_base_delegaciones()
    df_importantes = cargar_datos_importantes()

    colv1, colv2, colv3 = st.columns(3)

    colv1.metric("Registros MEP", len(df_mep))
    colv2.metric("Registros Delegaciones", len(df_delegaciones))
    colv3.metric("Registros Datos Importantes", len(df_importantes))

    with st.expander("Ver columnas detectadas en bases auxiliares"):
        st.write("Columnas Base MEP:")
        st.write(list(df_mep.columns))

        st.write("Columnas Delegaciones:")
        st.write(list(df_delegaciones.columns))

        st.write("Columnas Datos Importantes:")
        st.write(list(df_importantes.columns))

    # ======================================================
    # VALIDACIÓN DE BASE PRINCIPAL
    # ======================================================

    st.markdown("### Base principal de registros")

    df_config = cargar_datos()
    df_config_limpio = limpiar_dataframe_para_metricas(df_config)
    df_mapa_config = preparar_dataframe_mapa(df_config_limpio)

    colb1, colb2, colb3, colb4 = st.columns(4)

    colb1.metric("Registros totales", len(df_config))
    colb2.metric("Columnas esperadas", len(ENCABEZADOS))
    colb3.metric("Con coordenadas", len(df_mapa_config))
    colb4.metric("Sin coordenadas", len(df_config) - len(df_mapa_config))

    st.markdown("### Validación de encabezados oficiales")

    encabezados_actuales = hoja.row_values(1)

    faltantes = [col for col in ENCABEZADOS if col not in encabezados_actuales]
    sobrantes = [col for col in encabezados_actuales if col not in ENCABEZADOS]

    if not faltantes and not sobrantes:
        st.success("Los encabezados de la hoja coinciden con la estructura oficial.")
    else:
        if faltantes:
            st.error("Columnas faltantes:")
            st.write(faltantes)

        if sobrantes:
            st.warning("Columnas adicionales encontradas:")
            st.write(sobrantes)

    with st.expander("Ver encabezados oficiales"):
        st.write(ENCABEZADOS)

    # ======================================================
    # HERRAMIENTAS DE LIMPIEZA
    # ======================================================

    st.markdown("### Herramientas de mantenimiento")

    if st.button("Eliminar encabezados duplicados"):
        eliminados = limpiar_encabezados_duplicados_en_sheet()
        st.success(f"Filas de encabezado duplicadas eliminadas: {eliminados}")
        st.rerun()

    # ======================================================
    # RESPALDOS
    # ======================================================

    st.markdown("### Respaldos de información")

    if df_config.empty:
        st.info("No existen registros para respaldar.")
    else:
        excel_completo = convertir_excel(df_config)

        st.download_button(
            label="Descargar respaldo completo en Excel",
            data=excel_completo,
            file_name="respaldo_completo_pumi_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        df_aprobados = df_config[
            df_config["Estado Revisión"] == "Aprobado"
        ]

        df_pendientes = df_config[
            df_config["Estado Revisión"] == "Pendiente de revisión"
        ]

        colr1, colr2 = st.columns(2)

        with colr1:
            excel_aprobados = convertir_excel(df_aprobados)

            st.download_button(
                label="Descargar registros aprobados",
                data=excel_aprobados,
                file_name="registros_aprobados_pumi_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with colr2:
            excel_pendientes = convertir_excel(df_pendientes)

            st.download_button(
                label="Descargar registros pendientes",
                data=excel_pendientes,
                file_name="registros_pendientes_pumi_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # ======================================================
    # VISTA PREVIA
    # ======================================================

    st.markdown("### Vista previa de registros")

    if df_config.empty:
        st.info("No hay registros disponibles.")
    else:
        st.dataframe(
            df_config.head(20),
            use_container_width=True,
            hide_index=True
        )

    # ======================================================
    # INFORMACIÓN DEL SISTEMA
    # ======================================================

    st.markdown("### Información del sistema")

    st.markdown(
        f"""
        <div class="card-azul">
            <div class="texto-pumi">
                <b>Sistema:</b> PUMI 2026<br>
                <b>Nombre:</b> Proceso Unificado para el Manejo de la Información<br>
                <b>Fecha de revisión:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}<br>
                <b>Paleta institucional:</b> Azul, blanco y dorado institucional.<br>
                <b>Estado:</b> Sistema operativo.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("---")

    st.markdown(
        """
        <div style="text-align:center; color:#002B7F; font-weight:800;">
            Sistema PUMI 2026 · Proceso Unificado para el Manejo de la Información
        </div>
        """,
        unsafe_allow_html=True
    )


        
