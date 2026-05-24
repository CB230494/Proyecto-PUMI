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
# CATÁLOGOS BASE
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
# CONEXIÓN GOOGLE SHEETS, BASES AUXILIARES, CRUD, MAPA Y GEOREFERENCIA
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
        st.warning(
            "Revise que el archivo secrets.toml esté correctamente configurado "
            "y que la hoja esté compartida con el correo del service account."
        )
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
        st.warning(
            "Revise: permisos del service account, nombre de pestaña, "
            "Google Sheets API y Google Drive API."
        )
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


@st.cache_data
def cargar_base_delegaciones():
    if not os.path.exists(ARCHIVO_DELEGACIONES):
        return pd.DataFrame(columns=["Delegacion", "Distrito"])

    df = pd.read_excel(ARCHIVO_DELEGACIONES)

    col_delegacion = None
    col_distrito = None

    for col in df.columns:
        col_norm = normalizar_texto(col)

        if col_norm == "DELEGACION":
            col_delegacion = col

        if col_norm == "DISTRITO":
            col_distrito = col

    if col_delegacion is None or col_distrito is None:
        return pd.DataFrame(columns=["Delegacion", "Distrito"])

    df = df[[col_delegacion, col_distrito]].copy()
    df.columns = ["Delegacion", "Distrito"]

    df = df.dropna(subset=["Delegacion", "Distrito"])
    df["Delegacion"] = df["Delegacion"].astype(str).str.strip()
    df["Distrito"] = df["Distrito"].astype(str).str.strip()

    df["Delegacion_Normalizada"] = df["Delegacion"].apply(normalizar_texto)
    df["Distrito_Normalizado"] = df["Distrito"].apply(normalizar_texto)

    df = df.drop_duplicates(
        subset=["Delegacion_Normalizada", "Distrito_Normalizado"]
    )

    return df


def obtener_delegaciones_unicas():
    df = cargar_base_delegaciones()

    if df.empty:
        return []

    df_tmp = df[["Delegacion", "Delegacion_Normalizada"]].drop_duplicates(
        subset=["Delegacion_Normalizada"]
    )

    return sorted(df_tmp["Delegacion"].dropna().tolist())


def obtener_distritos_unicos():
    df = cargar_base_delegaciones()

    if df.empty:
        return []

    df_tmp = df[["Distrito", "Distrito_Normalizado"]].drop_duplicates(
        subset=["Distrito_Normalizado"]
    )

    return sorted(df_tmp["Distrito"].dropna().tolist())


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
# FUNCIONES DE GEOREFERENCIA Y MAPA
# ======================================================

@st.cache_data(show_spinner=False)
def georreferenciar_direccion(direccion):
    """
    Convierte un nombre de lugar o dirección en coordenadas.
    Usa Nominatim/OpenStreetMap.
    """

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
    """
    Convierte una coordenada a número.
    Si no es válida devuelve None.
    """

    try:
        if valor is None or str(valor).strip() == "":
            return None

        return float(str(valor).replace(",", ".").strip())

    except Exception:
        return None


def obtener_color_programa(programa):
    """
    Devuelve color de marcador según programa.
    """

    return COLORES_PROGRAMA.get(programa, "gray")


def preparar_dataframe_mapa(df):
    """
    Limpia latitud y longitud para mostrar registros en el mapa.
    """

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
    """
    Crea mapa Folium con todos los registros que tengan coordenadas.
    Cada programa se muestra con color diferente.
    """

    df_mapa = preparar_dataframe_mapa(df)

    if df_mapa.empty:
        mapa = folium.Map(
            location=[9.7489, -83.7534],
            zoom_start=7,
            tiles="OpenStreetMap"
        )
        return mapa

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
    """
    Muestra el mapa en Streamlit.
    """

    df_mapa = preparar_dataframe_mapa(df)

    if df_mapa.empty:
        st.info("No hay registros con coordenadas para mostrar en el mapa.")

    mapa = crear_mapa_registros(df)

    st_folium(
        mapa,
        width=None,
        height=height,
        key=key
    )


# ======================================================
# DATOS PRINCIPALES
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
# INTERFAZ PRINCIPAL, LOGIN ADMIN, INICIO, REGISTRO, GPS Y SEGUIMIENTO
# ======================================================

mostrar_logo()

st.sidebar.markdown("## Sistema PUMI 2026")


# ======================================================
# ACCESO ADMINISTRATIVO
# ======================================================

if "admin_autenticado" not in st.session_state:
    st.session_state.admin_autenticado = False

with st.sidebar.expander("🔒 Acceso Administrativo"):

    if not st.session_state.admin_autenticado:

        clave_ingresada = st.text_input(
            "Clave administrativa",
            type="password"
        )

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


# ======================================================
# MENÚ PRINCIPAL
# ======================================================

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

menu = st.sidebar.radio(
    "Menú principal",
    opciones_menu
)


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
# PANTALLA DE INICIO
# ======================================================

if menu == "Inicio":

    st.markdown(
        """
        <div class="card-pumi">
            <div class="subtitulo-pumi">Bienvenido al Sistema PUMI 2026</div>
            <div class="texto-pumi">
                Esta aplicación permite registrar, consultar y dar seguimiento a las
                actividades desarrolladas dentro del Proceso Unificado para el Manejo
                de la Información. La información se almacena en Google Sheets como
                base de datos institucional.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    try:
        df = cargar_datos()
        df_metricas = limpiar_dataframe_para_metricas(df)

        total_registros = len(df_metricas)
        total_programas = df_metricas["Programa"].nunique() if not df_metricas.empty else 0
        total_delegaciones = df_metricas["Delegación"].nunique() if not df_metricas.empty else 0
        total_participantes = int(df_metricas["Cantidad Participantes"].sum()) if not df_metricas.empty else 0

        col1, col2, col3, col4 = st.columns(4)

        col1.metric("Registros", total_registros)
        col2.metric("Programas", total_programas)
        col3.metric("Delegaciones", total_delegaciones)
        col4.metric("Participantes", total_participantes)

        st.success("Conexión con Google Sheets activa.")

        if st.session_state.admin_autenticado:
            st.info("Modo administrador activo.")
        else:
            st.info("Modo usuario activo. Puede registrar actividades y consultar el seguimiento.")

    except Exception as e:
        st.error("No se pudo conectar con Google Sheets.")
        st.exception(e)


# ======================================================
# FORMULARIO DE REGISTRO
# ======================================================

elif menu == "Registrar actividad":

    st.markdown("## Registrar nueva actividad preventiva")

    st.markdown(
        """
        <div class="card-azul">
            <div class="texto-pumi">
                Complete la información de la actividad preventiva realizada.
                El sistema asignará automáticamente un ID consecutivo simple.
                Además, puede registrar la ubicación por GPS, coordenadas manuales
                o búsqueda por nombre del lugar.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    delegaciones_lista = obtener_delegaciones_unicas()
    distritos_lista = obtener_distritos_unicos()

    if not delegaciones_lista:
        st.warning("No se encontró la base DELEGACIONES Y DISTRITOS.xlsx o no contiene la columna Delegacion.")

    if not distritos_lista:
        st.warning("No se encontró la base DELEGACIONES Y DISTRITOS.xlsx o no contiene la columna Distrito.")

    with st.form("form_registro_pumi"):

        st.markdown(
            """
            <div class="bloque-datos">
                <b>Datos generales del registro</b>
            </div>
            """,
            unsafe_allow_html=True
        )

        col1, col2 = st.columns(2)

        with col1:
            fecha_actividad = st.date_input(
                "Fecha de la actividad",
                value=date.today(),
                format="DD/MM/YYYY"
            )

            direccion_regional = st.selectbox(
                "Dirección Regional",
                REGIONES
            )

            delegacion = st.selectbox(
                "Delegación",
                delegaciones_lista if delegaciones_lista else ["Sin datos disponibles"]
            )

            programa = st.selectbox(
                "Programa",
                PROGRAMAS
            )

        with col2:
            actividad = st.text_input(
                "Actividad realizada"
            )

            responsable = st.text_input(
                "Funcionario responsable"
            )

            usuario = st.text_input(
                "Usuario que registra"
            )

            cantidad = st.number_input(
                "Cantidad de participantes",
                min_value=0,
                step=1
            )

        st.markdown(
            """
            <div class="bloque-territorio">
                <b>Ubicación territorial</b>
            </div>
            """,
            unsafe_allow_html=True
        )

        col3, col4, col5 = st.columns(3)

        with col3:
            provincia = st.selectbox(
                "Provincia",
                PROVINCIAS
            )

        with col4:
            canton = st.text_input(
                "Cantón"
            )

        with col5:
            distrito = st.selectbox(
                "Distrito",
                distritos_lista if distritos_lista else ["Sin datos disponibles"]
            )

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
                centro_educativo = st.text_input(
                    "Digite el centro educativo manualmente"
                )
                lugar = centro_educativo

        else:
            lugar = st.text_input(
                "Lugar donde se realizó la actividad"
            )
            centro_educativo = ""

        st.markdown(
            """
            <div class="bloque-mapa">
                <b>Georreferencia del lugar</b>
            </div>
            """,
            unsafe_allow_html=True
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
                "No registrar ubicación"
            ],
            horizontal=False
        )

        latitud = ""
        longitud = ""

        if metodo_ubicacion == "Buscar por nombre del lugar":
            lat_busqueda, lon_busqueda, direccion_encontrada = georreferenciar_direccion(direccion_mapa)

            if lat_busqueda and lon_busqueda:
                latitud = str(lat_busqueda)
                longitud = str(lon_busqueda)
                st.success("Ubicación encontrada automáticamente.")
                st.caption(direccion_encontrada)

                mapa_preview = folium.Map(
                    location=[lat_busqueda, lon_busqueda],
                    zoom_start=15
                )

                folium.Marker(
                    location=[lat_busqueda, lon_busqueda],
                    popup="Ubicación encontrada",
                    icon=folium.Icon(color=obtener_color_programa(programa), icon="info-sign")
                ).add_to(mapa_preview)

                st_folium(
                    mapa_preview,
                    height=350,
                    key="mapa_preview_busqueda"
                )

            else:
                st.warning("No se logró ubicar el lugar automáticamente. Puede ingresar coordenadas manualmente.")

        elif metodo_ubicacion == "Usar GPS del dispositivo":
            st.info("El navegador puede solicitar permiso para acceder a la ubicación del dispositivo.")

            ubicacion_gps = get_geolocation()

            if ubicacion_gps and "coords" in ubicacion_gps:
                latitud = str(ubicacion_gps["coords"].get("latitude", ""))
                longitud = str(ubicacion_gps["coords"].get("longitude", ""))

                if latitud and longitud:
                    st.success("Ubicación GPS obtenida correctamente.")

                    mapa_preview = folium.Map(
                        location=[float(latitud), float(longitud)],
                        zoom_start=16
                    )

                    folium.Marker(
                        location=[float(latitud), float(longitud)],
                        popup="Ubicación GPS del dispositivo",
                        icon=folium.Icon(color=obtener_color_programa(programa), icon="info-sign")
                    ).add_to(mapa_preview)

                    st_folium(
                        mapa_preview,
                        height=350,
                        key="mapa_preview_gps"
                    )
                else:
                    st.warning("No se recibieron coordenadas válidas desde el GPS.")
            else:
                st.warning("No se obtuvo ubicación GPS. Revise permisos del navegador o use otro método.")

        elif metodo_ubicacion == "Ingresar coordenadas manualmente":
            col_lat, col_lon = st.columns(2)

            with col_lat:
                latitud = st.text_input(
                    "Latitud",
                    placeholder="Ejemplo: 9.9281"
                )

            with col_lon:
                longitud = st.text_input(
                    "Longitud",
                    placeholder="Ejemplo: -84.0907"
                )

            lat_num = limpiar_coordenada(latitud)
            lon_num = limpiar_coordenada(longitud)

            if lat_num is not None and lon_num is not None:
                mapa_preview = folium.Map(
                    location=[lat_num, lon_num],
                    zoom_start=15
                )

                folium.Marker(
                    location=[lat_num, lon_num],
                    popup="Ubicación manual",
                    icon=folium.Icon(color=obtener_color_programa(programa), icon="info-sign")
                ).add_to(mapa_preview)

                st_folium(
                    mapa_preview,
                    height=350,
                    key="mapa_preview_manual"
                )

        else:
            st.info("El registro se guardará sin coordenadas de mapa.")

        st.markdown(
            """
            <div class="bloque-datos">
                <b>Información complementaria</b>
            </div>
            """,
            unsafe_allow_html=True
        )

        instituciones = st.text_area(
            "Instituciones participantes"
        )

        plan = st.text_input(
            "Plan estratégico relacionado"
        )

        evidencia = st.text_input(
            "Enlace de evidencia"
        )

        observaciones = st.text_area(
            "Observaciones"
        )

        guardar = st.form_submit_button(
            "Guardar registro"
        )

        if guardar:

            if not delegacion or delegacion == "Sin datos disponibles" or not actividad or not responsable:
                st.warning(
                    "Debe completar al menos Delegación, Actividad realizada y Funcionario responsable."
                )

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
                    "Latitud": latitud,
                    "Longitud": longitud,
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

                st.success(
                    f"Registro guardado correctamente con el ID #{nuevo_id}."
                )


# ======================================================
# SEGUIMIENTO DE REGISTROS
# VISTA PARA USUARIOS
# ======================================================

elif menu == "Seguimiento de registros":

    st.markdown("## Seguimiento de registros")

    st.markdown(
        """
        <div class="card-dorado">
            <div class="texto-pumi">
                En esta sección puede consultar el estado de revisión de los registros
                ingresados, visualizar observaciones administrativas y ver el mapa general
                de actividades registradas.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

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

            if fechas_validas.empty:
                st.warning("No hay fechas válidas para aplicar el filtro.")

            else:
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
            height=520,
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

        delegaciones_lista = obtener_delegaciones_unicas()
        distritos_lista = obtener_distritos_unicos()

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
                df_filtrado["ID"].astype(str).str.contains(filtro_id_admin, case=False, na=False)
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

                        actualizado = actualizar_registro_por_id(id_seleccionado, nuevos_datos)

                        if actualizado:
                            st.success("Revisión actualizada correctamente.")
                            st.rerun()
                        else:
                            st.error("No se encontró el registro para actualizar.")

            elif accion == "Corregir ubicación en mapa":

                st.markdown(
                    """
                    <div class="bloque-mapa">
                        <b>Corrección de ubicación geográfica</b>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                st.info(
                    "Puede buscar nuevamente por dirección, usar GPS del dispositivo "
                    "o ingresar coordenadas manualmente."
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
                        "Ingresar coordenadas manualmente"
                    ],
                    horizontal=False
                )

                latitud_nueva = registro_actual.get("Latitud", "")
                longitud_nueva = registro_actual.get("Longitud", "")

                if metodo_admin == "Buscar por nombre del lugar":

                    if st.button("Buscar ubicación"):
                        lat_busqueda, lon_busqueda, direccion_encontrada = georreferenciar_direccion(direccion_mapa_admin)

                        if lat_busqueda and lon_busqueda:
                            st.session_state["lat_admin"] = str(lat_busqueda)
                            st.session_state["lon_admin"] = str(lon_busqueda)
                            st.session_state["dir_admin"] = direccion_encontrada
                            st.success("Ubicación encontrada.")
                        else:
                            st.warning("No se logró ubicar el lugar automáticamente.")

                    latitud_nueva = st.session_state.get("lat_admin", latitud_nueva)
                    longitud_nueva = st.session_state.get("lon_admin", longitud_nueva)

                    if st.session_state.get("dir_admin", ""):
                        st.caption(st.session_state["dir_admin"])

                elif metodo_admin == "Usar GPS del dispositivo":

                    st.info("El navegador puede solicitar permiso para acceder a la ubicación.")

                    ubicacion_gps = get_geolocation()

                    if ubicacion_gps and "coords" in ubicacion_gps:
                        latitud_nueva = str(ubicacion_gps["coords"].get("latitude", ""))
                        longitud_nueva = str(ubicacion_gps["coords"].get("longitude", ""))
                        st.success("Ubicación GPS obtenida correctamente.")
                    else:
                        st.warning("No se obtuvo ubicación GPS.")

                else:
                    col_lat, col_lon = st.columns(2)

                    with col_lat:
                        latitud_nueva = st.text_input(
                            "Latitud",
                            value=str(latitud_nueva)
                        )

                    with col_lon:
                        longitud_nueva = st.text_input(
                            "Longitud",
                            value=str(longitud_nueva)
                        )

                lat_num = limpiar_coordenada(latitud_nueva)
                lon_num = limpiar_coordenada(longitud_nueva)

                if lat_num is not None and lon_num is not None:
                    mapa_preview_admin = folium.Map(
                        location=[lat_num, lon_num],
                        zoom_start=15
                    )

                    folium.Marker(
                        location=[lat_num, lon_num],
                        popup=f"Registro PUMI #{registro_actual.get('ID', '')}",
                        icon=folium.Icon(
                            color=obtener_color_programa(registro_actual.get("Programa", "")),
                            icon="info-sign"
                        )
                    ).add_to(mapa_preview_admin)

                    st_folium(
                        mapa_preview_admin,
                        height=400,
                        key="mapa_correccion_admin"
                    )

                if st.button("Guardar ubicación corregida"):

                    if lat_num is None or lon_num is None:
                        st.warning("Debe indicar coordenadas válidas antes de guardar.")
                    else:
                        nuevos_datos = registro_actual.copy()
                        nuevos_datos["Dirección Mapa"] = direccion_mapa_admin
                        nuevos_datos["Latitud"] = str(lat_num)
                        nuevos_datos["Longitud"] = str(lon_num)

                        actualizado = actualizar_registro_por_id(
                            id_seleccionado,
                            nuevos_datos
                        )

                        if actualizado:
                            st.success("Ubicación actualizada correctamente.")
                            st.rerun()
                        else:
                            st.error("No se encontró el registro para actualizar.")

            elif accion == "Editar registro completo":

                with st.form("form_editar_registro"):

                    st.markdown(
                        """
                        <div class="bloque-datos">
                            <b>Datos generales del registro</b>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                    col1, col2 = st.columns(2)

                    with col1:
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

                        fecha_actividad_editada = st.date_input(
                            "Fecha de la actividad",
                            value=fecha_actividad_actual,
                            format="DD/MM/YYYY"
                        )

                        nueva_region = st.selectbox(
                            "Dirección Regional",
                            REGIONES,
                            index=REGIONES.index(registro_actual.get("Dirección Regional", ""))
                            if registro_actual.get("Dirección Regional", "") in REGIONES else 0
                        )

                        delegacion_actual = registro_actual.get("Delegación", "")

                        if delegacion_actual and delegacion_actual not in delegaciones_lista:
                            opciones_delegacion = [delegacion_actual] + delegaciones_lista
                        else:
                            opciones_delegacion = delegaciones_lista if delegaciones_lista else [delegacion_actual]

                        nueva_delegacion = st.selectbox(
                            "Delegación",
                            opciones_delegacion,
                            index=opciones_delegacion.index(delegacion_actual)
                            if delegacion_actual in opciones_delegacion else 0
                        )

                        nuevo_programa = st.selectbox(
                            "Programa",
                            PROGRAMAS,
                            index=PROGRAMAS.index(registro_actual.get("Programa", ""))
                            if registro_actual.get("Programa", "") in PROGRAMAS else 0
                        )

                    with col2:
                        nueva_actividad = st.text_input(
                            "Actividad realizada",
                            registro_actual.get("Actividad", "")
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

                    col3, col4, col5 = st.columns(3)

                    with col3:
                        nueva_provincia = st.selectbox(
                            "Provincia",
                            PROVINCIAS,
                            index=PROVINCIAS.index(registro_actual.get("Provincia", ""))
                            if registro_actual.get("Provincia", "") in PROVINCIAS else 0
                        )

                    with col4:
                        nuevo_canton = st.text_input(
                            "Cantón",
                            registro_actual.get("Cantón", "")
                        )

                    with col5:
                        distrito_actual = registro_actual.get("Distrito", "")

                        if distrito_actual and distrito_actual not in distritos_lista:
                            opciones_distrito = [distrito_actual] + distritos_lista
                        else:
                            opciones_distrito = distritos_lista if distritos_lista else [distrito_actual]

                        nuevo_distrito = st.selectbox(
                            "Distrito",
                            opciones_distrito,
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
                            if centro_actual and centro_actual not in centros:
                                opciones_centros = [centro_actual] + centros
                            else:
                                opciones_centros = centros

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

                    actualizar = st.form_submit_button("Guardar cambios completos")

                    if actualizar:
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
