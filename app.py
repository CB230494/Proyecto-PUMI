# ======================================================
# PARTE 1 DE 10
# LIBRERÍAS, CONFIGURACIÓN GENERAL, ARCHIVOS BASE,
# ENCABEZADOS, CATÁLOGOS Y SESSION STATE
# ======================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata
import os
import base64
import folium
import re

from streamlit_folium import st_folium
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from streamlit_js_eval import get_geolocation
from datetime import datetime, date, time
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection


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
# ARCHIVOS BASE LOCALES
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
# NOMBRE DE HOJA PARA EXPORTAR / IMPORTAR EXCEL
# ======================================================

NOMBRE_HOJA_EXCEL = "REGISTRO_PUMI_2026"


# ======================================================
# ENCABEZADOS OFICIALES
# ======================================================

ENCABEZADOS = [
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
    "Número de Referencia",
    "Número de Expediente Referencia",
    "Observaciones",
    "Usuario Registra"
]


# ======================================================
# CATÁLOGOS BASE
# ======================================================

PROGRAMAS = [
    "DARE",
    "GREAT",
    "MPAS",
    "PSCC",
    "VIF",
    "Política Pública"
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
    "R11 Chorotega Norte",
    "R12"
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

RANGOS_EDAD = [
    "Edad 10 a 18",
    "Edad 19 a 30",
    "Edad 31 a 45",
    "Edad 46 en adelante"
]


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
# CENTROS GENERALES PARA MAPA POR PROVINCIA
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


# ======================================================
# SESSION STATE GENERAL
# Aquí se guarda temporalmente la información mientras
# la app está abierta.
# ======================================================

def inicializar_session_state():
    if "registros_pumi" not in st.session_state:
        st.session_state.registros_pumi = pd.DataFrame(columns=ENCABEZADOS)

    if "latitud_registro" not in st.session_state:
        st.session_state.latitud_registro = ""

    if "longitud_registro" not in st.session_state:
        st.session_state.longitud_registro = ""

    if "direccion_encontrada_mapa" not in st.session_state:
        st.session_state.direccion_encontrada_mapa = ""

    if "metodo_ubicacion_actual" not in st.session_state:
        st.session_state.metodo_ubicacion_actual = ""

    if "archivo_cargado_nombre" not in st.session_state:
        st.session_state.archivo_cargado_nombre = ""

    if "ultima_direccion_regional" not in st.session_state:
        st.session_state.ultima_direccion_regional = ""

    if "ultima_delegacion" not in st.session_state:
        st.session_state.ultima_delegacion = ""


inicializar_session_state()
# ======================================================
# PARTE 2 DE 10
# ESTILOS CSS, LOGOS Y ENCABEZADO INSTITUCIONAL
# ======================================================

# ======================================================
# FUNCIÓN PARA CONVERTIR IMÁGENES A BASE64
# Se usa para mostrar logos dentro del HTML de Streamlit.
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
# ESTILOS CSS GENERALES DE LA APP
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

    .card-pumi,
    .card-azul,
    .card-dorado {{
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


def mostrar_logos_encabezado():
    mostrar_encabezado_institucional()


# ======================================================
# TÍTULO PRINCIPAL DE LA APP
# ======================================================

def mostrar_titulo_principal():
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
# PARTE 3 DE 10
# FUNCIONES GENERALES, NORMALIZACIÓN, ORDENAMIENTO
# Y MANEJO DE REGISTROS EN SESSION STATE
# ======================================================


# ======================================================
# NORMALIZAR TEXTO
# Convierte texto a mayúsculas, elimina tildes y limpia espacios.
# Sirve para comparar columnas aunque tengan diferencias de escritura.
# ======================================================

def normalizar_texto(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        [
            caracter
            for caracter in texto
            if not unicodedata.combining(caracter)
        ]
    )
    texto = " ".join(texto.split())

    return texto


# ======================================================
# BUSCAR COLUMNA POR NOMBRE
# Permite encontrar columnas aunque vengan con tildes,
# mayúsculas, minúsculas o pequeñas variaciones.
# ======================================================

def obtener_columna_por_nombre(df, posibles_nombres):
    columnas = list(df.columns)

    for posible in posibles_nombres:
        posible_norm = normalizar_texto(posible)

        for col in columnas:
            if normalizar_texto(col) == posible_norm:
                return col

    return None


# ======================================================
# ORDENAR REGIONES NUMÉRICAMENTE
# Ordena R1, R2, R3... sin que R10 quede antes de R2.
# ======================================================

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


# ======================================================
# SEPARAR CAMPO "RESPONDE A"
# Cuando una actividad responde a dos opciones separadas por "/",
# la función las divide para poder filtrarlas correctamente.
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


# ======================================================
# PORCENTAJE EN FORMATO TEXTO
# ======================================================

def porcentaje(valor, total):
    try:
        valor = float(valor)
        total = float(total)

        if total == 0:
            return "0.0%"

        return f"{(valor / total) * 100:.1f}%"

    except Exception:
        return "0.0%"


# ======================================================
# LIMPIAR TEXTO PARA NOMBRE DE ARCHIVO
# Evita caracteres inválidos en Windows o en navegadores.
# ======================================================

def limpiar_nombre_archivo(valor):
    texto = str(valor).strip()

    if texto == "" or texto.lower() == "nan":
        return "SIN_DATO"

    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        [
            caracter
            for caracter in texto
            if not unicodedata.combining(caracter)
        ]
    )

    texto = texto.upper()
    texto = re.sub(r"[^\w\s-]", "", texto)
    texto = re.sub(r"\s+", "_", texto)
    texto = texto.replace("__", "_")
    texto = texto.strip("_")

    if texto == "":
        return "SIN_DATO"

    return texto


# ======================================================
# GENERAR NOMBRE AUTOMÁTICO PARA EXCEL
# Usa Dirección Regional + Delegación.
# Ejemplo:
# REGISTRO_PUMI_2026_R1_SAN_JOSE_CENTRAL_DELEGACION_PAVAS.xlsx
# ======================================================

def generar_nombre_excel(direccion_regional="", delegacion="", filtrado=False):
    direccion = limpiar_nombre_archivo(direccion_regional)
    deleg = limpiar_nombre_archivo(delegacion)

    prefijo = "REGISTRO_PUMI_2026"

    if filtrado:
        prefijo = "REGISTRO_PUMI_2026_FILTRADO"

    if direccion == "SIN_DATO" and deleg == "SIN_DATO":
        return f"{prefijo}.xlsx"

    return f"{prefijo}_{direccion}_{deleg}.xlsx"


# ======================================================
# OBTENER NOMBRE DE EXCEL SEGÚN DATOS DEL DATAFRAME
# Si hay una sola Dirección Regional y una sola Delegación,
# usa esos nombres. Si hay varias, usa "VARIAS".
# ======================================================

def generar_nombre_excel_desde_df(df, filtrado=False):
    if df is None or df.empty:
        return generar_nombre_excel(filtrado=filtrado)

    direccion = "VARIAS_DIRECCIONES"
    delegacion = "VARIAS_DELEGACIONES"

    if "Dirección Regional" in df.columns:
        regiones = (
            df["Dirección Regional"]
            .replace("", pd.NA)
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        if len(regiones) == 1:
            direccion = regiones[0]

    if "Delegación" in df.columns:
        delegaciones = (
            df["Delegación"]
            .replace("", pd.NA)
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        if len(delegaciones) == 1:
            delegacion = delegaciones[0]

    return generar_nombre_excel(
        direccion_regional=direccion,
        delegacion=delegacion,
        filtrado=filtrado
    )


# ======================================================
# PREPARAR DATAFRAME IMPORTADO
# Acomoda cualquier Excel cargado para que tenga las columnas oficiales.
# ======================================================

def preparar_dataframe_importado(df):
    df = df.copy()

    columnas_a_eliminar = [
        "Fecha Registro",
        "Plan Estratégico Relacionado"
    ]

    for col in columnas_a_eliminar:
        if col in df.columns:
            df = df.drop(columns=[col])

    for col in ENCABEZADOS:
        if col not in df.columns:
            df[col] = ""

    df = df[ENCABEZADOS]

    for col in ENCABEZADOS:
        df[col] = df[col].fillna("").astype(str)

    columnas_numericas = [
        "ID",
        "Cantidad Participantes",
        "Cantidad Hombres",
        "Cantidad Mujeres",
        "Edad 10 a 18",
        "Edad 19 a 30",
        "Edad 31 a 45",
        "Edad 46 en adelante"
    ]

    for col in columnas_numericas:
        df[col] = (
            pd.to_numeric(df[col], errors="coerce")
            .fillna(0)
            .astype(int)
        )

    return df


# ======================================================
# CARGAR EXCEL EXISTENTE
# Permite continuar trabajando sobre un archivo exportado antes.
# ======================================================

def cargar_excel_existente(archivo_excel):
    try:
        df = pd.read_excel(
            archivo_excel,
            sheet_name=NOMBRE_HOJA_EXCEL
        )
    except Exception:
        try:
            df = pd.read_excel(archivo_excel)
        except Exception as e:
            st.error("No se pudo leer el archivo Excel cargado.")
            st.exception(e)
            return pd.DataFrame(columns=ENCABEZADOS)

    return preparar_dataframe_importado(df)


# ======================================================
# GENERAR ID CONSECUTIVO
# ======================================================

def generar_id_consecutivo():
    df = st.session_state.registros_pumi

    if df.empty:
        return 1

    ids = pd.to_numeric(df["ID"], errors="coerce").dropna()

    if ids.empty:
        return 1

    return int(ids.max()) + 1


# ======================================================
# AGREGAR REGISTRO A SESSION STATE
# ======================================================

def agregar_registro_session(registro):
    nuevo_df = pd.DataFrame(
        [registro],
        columns=ENCABEZADOS
    )

    st.session_state.registros_pumi = pd.concat(
        [
            st.session_state.registros_pumi,
            nuevo_df
        ],
        ignore_index=True
    )


# ======================================================
# ELIMINAR REGISTRO POR ID
# ======================================================

def eliminar_registro_session(id_registro):
    df = st.session_state.registros_pumi.copy()

    df = df[df["ID"].astype(str) != str(id_registro)]

    st.session_state.registros_pumi = df.reset_index(drop=True)


# ======================================================
# ACTUALIZAR REGISTRO POR ID
# ======================================================

def actualizar_registro_session(id_registro, nuevos_datos):
    df = st.session_state.registros_pumi.copy()

    mascara = df["ID"].astype(str) == str(id_registro)

    if not mascara.any():
        return False

    for col in ENCABEZADOS:
        df.loc[mascara, col] = nuevos_datos.get(col, "")

    st.session_state.registros_pumi = df.reset_index(drop=True)

    return True


# ======================================================
# LIMPIAR DATAFRAME PARA MÉTRICAS Y DASHBOARD
# Convierte columnas numéricas y fechas para poder graficar.
# ======================================================

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
# PARTE 4 DE 10
# CARGA DE CATÁLOGOS:
# DATOS IMPORTANTES, BASE MEP Y DELEGACIONES
# ======================================================


# ======================================================
# CARGA DE DATOS IMPORTANTES
# Contiene:
# - Provincia
# - Cantón
# - Distrito
# - Dirección Regional
# - Delegación
# - Responde a
# - Actividad Realizada
# - Programa
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

        col_provincia = obtener_columna_por_nombre(df_original, ["Provincia"])
        col_canton = obtener_columna_por_nombre(df_original, ["Cantón", "Canton"])
        col_distrito = obtener_columna_por_nombre(df_original, ["Distrito", "Distritos"])

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
            [
                "Responde a",
                "Responde a:",
                "Responde",
                "Responda a",
                "Responda a:"
            ]
        )

        col_actividad = obtener_columna_por_nombre(
            df_original,
            [
                "Actividad Realizada",
                "Actividad"
            ]
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


# ======================================================
# OBTENER REGIONES DESDE DATOS IMPORTANTES
# ======================================================

def obtener_regiones_datos():
    df = cargar_datos_importantes()

    if df.empty:
        return ordenar_regiones_numericamente(REGIONES)

    regiones = (
        df["Dirección Regional"]
        .dropna()
        .unique()
        .tolist()
    )

    regiones = [
        x for x in regiones
        if str(x).strip() != ""
    ]

    return ordenar_regiones_numericamente(regiones)


# ======================================================
# OBTENER DELEGACIONES SEGÚN DIRECCIÓN REGIONAL
# ======================================================

def obtener_delegaciones_por_region(region):
    df = cargar_datos_importantes()

    if df.empty or not region:
        return obtener_delegaciones_unicas()

    region_norm = normalizar_texto(region)

    delegaciones = (
        df[df["Región_Normalizada"] == region_norm]["Delegación"]
        .dropna()
        .unique()
        .tolist()
    )

    delegaciones = [
        x for x in delegaciones
        if str(x).strip() != ""
    ]

    return sorted(delegaciones)


# ======================================================
# OBTENER RESPONDE A
# ======================================================

def obtener_responde_a_datos():
    df = cargar_datos_importantes()

    if df.empty or "Responde a" not in df.columns:
        return []

    responde_a = (
        df["Responde a"]
        .dropna()
        .unique()
        .tolist()
    )

    responde_a = [
        x for x in responde_a
        if str(x).strip() != ""
    ]

    return sorted(responde_a)


# ======================================================
# OBTENER PROGRAMAS SEGÚN RESPONDE A
# ======================================================

def obtener_programas_por_responde_a(responde_a):
    df = cargar_datos_importantes()

    if df.empty or not responde_a:
        return obtener_programas_datos()

    responde_norm = normalizar_texto(responde_a)

    programas = (
        df[df["Responde_a_Normalizado"] == responde_norm]["Programa"]
        .dropna()
        .unique()
        .tolist()
    )

    programas = [
        x for x in programas
        if str(x).strip() != ""
    ]

    if not programas:
        return []

    programas_limpios = []

    for programa in programas:
        if normalizar_texto(programa) == "GREAT CAMP":
            if "GREAT" not in programas_limpios:
                programas_limpios.append("GREAT")
        elif programa not in programas_limpios:
            programas_limpios.append(programa)

    orden_oficial = [
        p for p in PROGRAMAS
        if p in programas_limpios
    ]

    extras = sorted(
        [
            p for p in programas_limpios
            if p not in orden_oficial
        ]
    )

    return orden_oficial + extras


# ======================================================
# OBTENER ACTIVIDADES SEGÚN RESPONDE A Y PROGRAMA
# ======================================================

def obtener_actividades_por_responde_a_programa(responde_a, programa):
    df = cargar_datos_importantes()

    if df.empty or not responde_a or not programa:
        return []

    responde_norm = normalizar_texto(responde_a)
    programa_norm = normalizar_texto(programa)

    if programa_norm == "GREAT":
        actividades = (
            df[
                (df["Responde_a_Normalizado"] == responde_norm) &
                (df["Programa_Normalizado"].isin(["GREAT", "GREAT CAMP"]))
            ]["Actividad Realizada"]
            .dropna()
            .unique()
            .tolist()
        )
    else:
        actividades = (
            df[
                (df["Responde_a_Normalizado"] == responde_norm) &
                (df["Programa_Normalizado"] == programa_norm)
            ]["Actividad Realizada"]
            .dropna()
            .unique()
            .tolist()
        )

    actividades = [
        x for x in actividades
        if str(x).strip() != ""
    ]

    return sorted(actividades)


# ======================================================
# OBTENER PROGRAMAS GENERALES
# ======================================================

def obtener_programas_datos():
    df = cargar_datos_importantes()

    if df.empty:
        return PROGRAMAS

    programas = (
        df["Programa"]
        .dropna()
        .unique()
        .tolist()
    )

    programas = [
        x for x in programas
        if str(x).strip() != ""
    ]

    if not programas:
        return PROGRAMAS

    programas_limpios = []

    for programa in programas:
        if normalizar_texto(programa) == "GREAT CAMP":
            if "GREAT" not in programas_limpios:
                programas_limpios.append("GREAT")
        elif programa not in programas_limpios:
            programas_limpios.append(programa)

    orden_oficial = [
        p for p in PROGRAMAS
        if p in programas_limpios
    ]

    extras = sorted(
        [
            p for p in programas_limpios
            if p not in orden_oficial
        ]
    )

    return orden_oficial + extras


# ======================================================
# OBTENER ACTIVIDADES SEGÚN PROGRAMA
# ======================================================

def obtener_actividades_por_programa(programa):
    df = cargar_datos_importantes()

    if df.empty or not programa:
        return []

    programa_norm = normalizar_texto(programa)

    if programa_norm == "GREAT":
        actividades = (
            df[
                df["Programa_Normalizado"].isin(["GREAT", "GREAT CAMP"])
            ]["Actividad Realizada"]
            .dropna()
            .unique()
            .tolist()
        )
    else:
        actividades = (
            df[
                df["Programa_Normalizado"] == programa_norm
            ]["Actividad Realizada"]
            .dropna()
            .unique()
            .tolist()
        )

    actividades = [
        x for x in actividades
        if str(x).strip() != ""
    ]

    return sorted(actividades)


# ======================================================
# OBTENER PROVINCIAS
# ======================================================

def obtener_provincias_datos():
    df = cargar_datos_importantes()

    if df.empty:
        return PROVINCIAS

    provincias = (
        df["Provincia"]
        .dropna()
        .unique()
        .tolist()
    )

    provincias = [
        x for x in provincias
        if str(x).strip() != ""
    ]

    return sorted(provincias)


# ======================================================
# OBTENER CANTONES POR PROVINCIA
# ======================================================

def obtener_cantones_por_provincia(provincia):
    df = cargar_datos_importantes()

    if df.empty or not provincia:
        return []

    provincia_norm = normalizar_texto(provincia)

    cantones = (
        df[df["Provincia_Normalizada"] == provincia_norm]["Cantón"]
        .dropna()
        .unique()
        .tolist()
    )

    cantones = [
        x for x in cantones
        if str(x).strip() != ""
    ]

    return sorted(cantones)


# ======================================================
# OBTENER DISTRITOS POR PROVINCIA Y CANTÓN
# ======================================================

def obtener_distritos_por_provincia_canton(provincia, canton):
    df = cargar_datos_importantes()

    if df.empty or not provincia or not canton:
        return []

    provincia_norm = normalizar_texto(provincia)
    canton_norm = normalizar_texto(canton)

    distritos = (
        df[
            (df["Provincia_Normalizada"] == provincia_norm) &
            (df["Cantón_Normalizado"] == canton_norm)
        ]["Distrito"]
        .dropna()
        .unique()
        .tolist()
    )

    distritos = [
        x for x in distritos
        if str(x).strip() != ""
    ]

    return sorted(distritos)


# ======================================================
# BASE MEP
# Carga centros educativos y código presupuestario.
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


# ======================================================
# OBTENER CENTROS EDUCATIVOS POR PROVINCIA
# ======================================================

def obtener_centros_por_provincia(provincia):
    df = cargar_base_mep()

    if df.empty or not provincia:
        return []

    provincia_norm = normalizar_texto(provincia)

    centros = (
        df[df["PROVINCIA_NORMALIZADA"] == provincia_norm]["CENTRO_MOSTRAR"]
        .dropna()
        .astype(str)
        .drop_duplicates()
        .sort_values()
        .tolist()
    )

    return centros


# ======================================================
# OBTENER DATOS DEL CENTRO EDUCATIVO
# ======================================================

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
# BASE DELEGACIONES Y DISTRITOS COMO RESPALDO
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


# ======================================================
# OBTENER TODAS LAS DELEGACIONES ÚNICAS
# ======================================================

def obtener_delegaciones_unicas():
    df_datos = cargar_datos_importantes()

    if not df_datos.empty:
        delegaciones = (
            df_datos["Delegación"]
            .dropna()
            .unique()
            .tolist()
        )

        delegaciones = [
            x for x in delegaciones
            if str(x).strip() != ""
        ]

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
# ======================================================
# PARTE 5 DE 10
# MAPA, GEOREFERENCIA, GPS, COORDENADAS MANUALES
# Y MARCADOR VISIBLE SIEMPRE
# ======================================================


# ======================================================
# GEOREFERENCIAR DIRECCIÓN O LUGAR
# Se conserva por compatibilidad, pero ya no se usa
# en el formulario para evitar llenar "Dirección Mapa".
# ======================================================

@st.cache_data(show_spinner=False)
def georreferenciar_direccion(direccion):
    if not direccion or str(direccion).strip() == "":
        return None, None, ""

    try:
        geolocator = Nominatim(
            user_agent="pumi_2026_streamlit_app"
        )

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


# ======================================================
# LIMPIAR COORDENADAS
# ======================================================

def limpiar_coordenada(valor):
    try:
        if valor is None or str(valor).strip() == "":
            return None

        return float(str(valor).replace(",", ".").strip())

    except Exception:
        return None


# ======================================================
# COLOR DE MARCADOR SEGÚN PROGRAMA
# ======================================================

def obtener_color_programa(programa):
    return COLORES_PROGRAMA.get(programa, "gray")


# ======================================================
# CREAR MAPA BASE
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

    elif tipo_mapa == "Topográfico":
        folium.TileLayer(
            "OpenTopoMap",
            name="Topográfico"
        ).add_to(mapa)

    elif tipo_mapa == "Satélite":
        folium.TileLayer(
            tiles=(
                "https://server.arcgisonline.com/ArcGIS/rest/services/"
                "World_Imagery/MapServer/tile/{z}/{y}/{x}"
            ),
            attr="Esri World Imagery",
            name="Satélite"
        ).add_to(mapa)

    folium.LayerControl().add_to(mapa)

    return mapa


# ======================================================
# OBTENER CENTRO DEL MAPA DE REGISTRO
# ======================================================

def obtener_centro_mapa_registro(provincia):
    lat_actual = limpiar_coordenada(
        st.session_state.get("latitud_registro", "")
    )

    lon_actual = limpiar_coordenada(
        st.session_state.get("longitud_registro", "")
    )

    if lat_actual is not None and lon_actual is not None:
        return [lat_actual, lon_actual], 16

    centro_provincia = CENTROS_PROVINCIA.get(
        provincia,
        [9.7489, -83.7534]
    )

    return centro_provincia, 10


# ======================================================
# CREAR MAPA DE REGISTRO INDIVIDUAL
# ======================================================

def crear_mapa_registro_individual(
    provincia,
    tipo_mapa,
    lugar="",
    delegacion="",
    programa="",
    permitir_click=True
):
    centro_mapa, zoom_mapa = obtener_centro_mapa_registro(provincia)

    mapa = crear_mapa_base(
        centro_mapa,
        zoom_mapa,
        tipo_mapa
    )

    lat_actual = limpiar_coordenada(
        st.session_state.get("latitud_registro", "")
    )

    lon_actual = limpiar_coordenada(
        st.session_state.get("longitud_registro", "")
    )

    if lat_actual is not None and lon_actual is not None:
        direccion_encontrada = st.session_state.get(
            "direccion_encontrada_mapa",
            ""
        )

        popup_html = f"""
        <div style="font-family: Arial; width: 280px;">
            <h4 style="color:#002B7F; margin-bottom:6px;">
                Punto seleccionado para registro
            </h4>
            <b>Delegación:</b> {delegacion}<br>
            <b>Programa:</b> {programa}<br>
            <b>Lugar:</b> {lugar}<br>
            <b>Referencia:</b> {direccion_encontrada}<br>
            <b>Latitud:</b> {lat_actual}<br>
            <b>Longitud:</b> {lon_actual}<br>
        </div>
        """

        folium.Marker(
            location=[lat_actual, lon_actual],
            popup=folium.Popup(popup_html, max_width=330),
            tooltip="Punto registrado",
            icon=folium.Icon(
                color="red",
                icon="map-marker"
            )
        ).add_to(mapa)

        folium.Circle(
            location=[lat_actual, lon_actual],
            radius=180,
            color=COLOR_AZUL,
            fill=True,
            fill_color=COLOR_DORADO,
            fill_opacity=0.18,
            weight=2
        ).add_to(mapa)

    return mapa


# ======================================================
# MOSTRAR MAPA INTERACTIVO DE REGISTRO
# ======================================================

def mostrar_mapa_registro_interactivo(
    provincia,
    tipo_mapa,
    lugar="",
    delegacion="",
    programa="",
    key="mapa_registro_general"
):
    mapa_registro = crear_mapa_registro_individual(
        provincia=provincia,
        tipo_mapa=tipo_mapa,
        lugar=lugar,
        delegacion=delegacion,
        programa=programa
    )

    resultado_mapa = st_folium(
        mapa_registro,
        height=520,
        use_container_width=True,
        key=key
    )

    if resultado_mapa and resultado_mapa.get("last_clicked"):
        lat_click = resultado_mapa["last_clicked"]["lat"]
        lon_click = resultado_mapa["last_clicked"]["lng"]

        st.session_state.latitud_registro = str(lat_click)
        st.session_state.longitud_registro = str(lon_click)
        st.session_state.direccion_encontrada_mapa = "Punto marcado manualmente en el mapa"

        st.success(
            f"Punto seleccionado en el mapa: {lat_click:.6f}, {lon_click:.6f}"
        )

        st.rerun()


# ======================================================
# BLOQUE COMPLETO DE GEOREFERENCIA PARA EL FORMULARIO
# ======================================================

def bloque_georreferencia_formulario(
    provincia,
    canton,
    distrito,
    lugar,
    delegacion,
    programa
):
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
            "Topográfico",
            "Satélite"
        ],
        key="tipo_mapa_registro"
    )

    metodo_ubicacion = st.radio(
        "Método para registrar ubicación",
        [
            "Usar GPS del dispositivo",
            "Ingresar coordenadas manualmente",
            "Marcar punto en el mapa",
            "No registrar ubicación"
        ],
        horizontal=False,
        key="metodo_ubicacion_registro"
    )

    st.session_state.metodo_ubicacion_actual = metodo_ubicacion

    # ==================================================
    # USAR GPS DEL DISPOSITIVO
    # ==================================================

    if metodo_ubicacion == "Usar GPS del dispositivo":
        st.info(
            "El navegador puede solicitar permiso para acceder a la ubicación "
            "del dispositivo. Cuando se obtenga la ubicación, el punto se verá "
            "marcado en el mapa."
        )

        ubicacion_gps = get_geolocation()

        if ubicacion_gps and "coords" in ubicacion_gps:
            lat_gps = ubicacion_gps["coords"].get("latitude", "")
            lon_gps = ubicacion_gps["coords"].get("longitude", "")

            if lat_gps and lon_gps:
                st.session_state.latitud_registro = str(lat_gps)
                st.session_state.longitud_registro = str(lon_gps)
                st.session_state.direccion_encontrada_mapa = "Ubicación obtenida por GPS del dispositivo"

                st.success("Ubicación GPS obtenida correctamente.")
            else:
                st.warning("No se recibieron coordenadas válidas desde el GPS.")
        else:
            st.warning(
                "No se obtuvo ubicación GPS todavía. "
                "Revise que el navegador tenga permisos de ubicación."
            )

    # ==================================================
    # INGRESAR COORDENADAS MANUALMENTE
    # ==================================================

    elif metodo_ubicacion == "Ingresar coordenadas manualmente":
        st.info(
            "Ingrese latitud y longitud. El mapa se actualizará mostrando "
            "el punto registrado."
        )

        col_lat, col_lon = st.columns(2)

        with col_lat:
            lat_manual = st.text_input(
                "Latitud",
                value=st.session_state.latitud_registro,
                key="lat_manual_registro"
            )

        with col_lon:
            lon_manual = st.text_input(
                "Longitud",
                value=st.session_state.longitud_registro,
                key="lon_manual_registro"
            )

        st.session_state.latitud_registro = lat_manual
        st.session_state.longitud_registro = lon_manual
        st.session_state.direccion_encontrada_mapa = "Coordenadas ingresadas manualmente"

    # ==================================================
    # MARCAR PUNTO EN MAPA
    # ==================================================

    elif metodo_ubicacion == "Marcar punto en el mapa":
        st.info(
            "Haga clic sobre el mapa para seleccionar el punto exacto. "
            "El marcador quedará visible después de seleccionar la ubicación."
        )

    # ==================================================
    # NO REGISTRAR UBICACIÓN
    # ==================================================

    elif metodo_ubicacion == "No registrar ubicación":
        st.session_state.latitud_registro = ""
        st.session_state.longitud_registro = ""
        st.session_state.direccion_encontrada_mapa = ""

        st.warning("No se registrará ubicación para este registro.")

    # ==================================================
    # MOSTRAR MAPA SIEMPRE
    # ==================================================

    st.markdown("### Vista del punto en el mapa")

    mostrar_mapa_registro_interactivo(
        provincia=provincia,
        tipo_mapa=tipo_mapa_registro,
        lugar=lugar,
        delegacion=delegacion,
        programa=programa,
        key="mapa_registro_visible_siempre"
    )

    # ==================================================
    # MOSTRAR COORDENADAS REGISTRADAS
    # ==================================================

    col_geo1, col_geo2 = st.columns(2)

    with col_geo1:
        st.text_input(
            "Latitud registrada",
            value=st.session_state.latitud_registro,
            disabled=True,
            key="latitud_registrada_visible"
        )

    with col_geo2:
        st.text_input(
            "Longitud registrada",
            value=st.session_state.longitud_registro,
            disabled=True,
            key="longitud_registrada_visible"
        )

    if st.session_state.get("direccion_encontrada_mapa", ""):
        st.info(
            f"Referencia del punto: {st.session_state.direccion_encontrada_mapa}"
        )

    return ""


# ======================================================
# PREPARAR DATAFRAME PARA MAPA GENERAL
# ======================================================

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


# ======================================================
# CENTRAR MAPA GENERAL POR PROVINCIA
# ======================================================

def obtener_centro_mapa_por_provincia(df):
    if df.empty or "Provincia" not in df.columns:
        return [9.7489, -83.7534], 7

    provincias = (
        df["Provincia"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    provincias = [
        p for p in provincias
        if p.strip() != ""
    ]

    if len(provincias) == 1:
        provincia = provincias[0]

        if provincia in CENTROS_PROVINCIA:
            return CENTROS_PROVINCIA[provincia], 10

    return [9.7489, -83.7534], 7


# ======================================================
# CREAR MAPA GENERAL DE REGISTROS
# ======================================================

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
            <b>Dirección Regional:</b> {row.get("Dirección Regional", "")}<br>
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
            fill_color=COLOR_DORADO,
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
            fill_color=COLOR_DORADO,
            fill_opacity=0.08,
            weight=2,
            popup=f"Área aproximada de {total_puntos} registros filtrados"
        ).add_to(mapa)

    return mapa


# ======================================================
# MOSTRAR MAPA GENERAL DE REGISTROS
# ======================================================

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
# PARTE 6 DE 10
# EXPORTACIÓN A EXCEL CON FORMATO INSTITUCIONAL,
# MARCA DE AGUA, PROTECCIÓN TOTAL Y NOMBRE AUTOMÁTICO
# ======================================================


# ======================================================
# CONVERTIR DATAFRAME A EXCEL
# Esta función genera el archivo Excel en memoria.
# Mantiene los encabezados oficiales, pero elimina
# "Dirección Mapa" del Excel final para evitar confusión.
# Protege la hoja y la estructura del libro.
# Contraseña de desbloqueo: DPPP23
# ======================================================

def convertir_excel(df):
    output = BytesIO()

    df_exportar = df.copy()

    columnas_a_eliminar = [
        "Fecha Registro",
        "Plan Estratégico Relacionado",
        "Dirección Mapa"
    ]

    for col in columnas_a_eliminar:
        if col in df_exportar.columns:
            df_exportar = df_exportar.drop(columns=[col])

    encabezados_excel = [
        col for col in ENCABEZADOS
        if col != "Dirección Mapa"
    ]

    for col in encabezados_excel:
        if col not in df_exportar.columns:
            df_exportar[col] = ""

    df_exportar = df_exportar[encabezados_excel]

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_exportar.to_excel(
            writer,
            index=False,
            sheet_name=NOMBRE_HOJA_EXCEL
        )

    output.seek(0)

    workbook = load_workbook(output)
    worksheet = workbook[NOMBRE_HOJA_EXCEL]

    # ==================================================
    # FORMATO DEL ENCABEZADO
    # ==================================================

    fill_header = PatternFill(
        start_color="002B7F",
        end_color="002B7F",
        fill_type="solid"
    )

    font_header = Font(
        color="FFFFFF",
        bold=True
    )

    border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF")
    )

    for cell in worksheet[1]:
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.border = border

    # ==================================================
    # FORMATO DEL CUERPO DE LA TABLA
    # ==================================================

    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )
            cell.border = border

    # ==================================================
    # AJUSTE AUTOMÁTICO DE ANCHO DE COLUMNAS
    # ==================================================

    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter

        for cell in column_cells:
            try:
                valor = str(cell.value) if cell.value is not None else ""

                if len(valor) > max_length:
                    max_length = len(valor)

            except Exception:
                pass

        adjusted_width = min(
            max(max_length + 2, 14),
            45
        )

        worksheet.column_dimensions[column_letter].width = adjusted_width

    # ==================================================
    # FIJAR ENCABEZADO
    # ==================================================

    worksheet.freeze_panes = "A2"

    # ==================================================
    # MARCA DE AGUA EN IMPRESIÓN / DISEÑO DE PÁGINA
    # ==================================================

    worksheet.oddHeader.center.text = "Dirección de Programas Preventivos Policiales"
    worksheet.oddHeader.center.size = 24
    worksheet.oddHeader.center.font = "Arial,Bold"
    worksheet.oddHeader.center.color = "D9D9D9"

    worksheet.evenHeader.center.text = "Dirección de Programas Preventivos Policiales"
    worksheet.evenHeader.center.size = 24
    worksheet.evenHeader.center.font = "Arial,Bold"
    worksheet.evenHeader.center.color = "D9D9D9"

    # ==================================================
    # CONFIGURACIÓN DE IMPRESIÓN
    # ==================================================

    worksheet.page_setup.orientation = "landscape"
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True

    worksheet.page_margins.left = 0.25
    worksheet.page_margins.right = 0.25
    worksheet.page_margins.top = 0.75
    worksheet.page_margins.bottom = 0.75

    # ==================================================
    # PROTEGER TODAS LAS CELDAS
    # ==================================================

    for fila in worksheet.iter_rows():
        for celda in fila:
            celda.protection = Protection(
                locked=True
            )

    # ==================================================
    # PROTEGER HOJA COMPLETA
    # Contraseña: DPPP23
    # ==================================================

    worksheet.protection.sheet = True
    worksheet.protection.password = "DPPP23"

    worksheet.protection.formatCells = False
    worksheet.protection.formatColumns = False
    worksheet.protection.formatRows = False
    worksheet.protection.insertColumns = False
    worksheet.protection.insertRows = False
    worksheet.protection.insertHyperlinks = False
    worksheet.protection.deleteColumns = False
    worksheet.protection.deleteRows = False
    worksheet.protection.sort = False
    worksheet.protection.autoFilter = False
    worksheet.protection.pivotTables = False
    worksheet.protection.selectLockedCells = True
    worksheet.protection.selectUnlockedCells = False

    # ==================================================
    # PROTEGER ESTRUCTURA DEL LIBRO
    # Evita agregar, eliminar, mover o renombrar hojas
    # Contraseña: DPPP23
    # ==================================================

    workbook.security.lockStructure = True
    workbook.security.workbookPassword = "DPPP23"

    final_output = BytesIO()
    workbook.save(final_output)
    final_output.seek(0)

    return final_output.getvalue()


# ======================================================
# BOTÓN DE DESCARGA ESTANDARIZADO
# ======================================================

def boton_descargar_excel(
    df,
    texto_boton="⬇️ Descargar Excel",
    filtrado=False,
    key=None
):
    if df is None or df.empty:
        st.info("No hay registros para descargar.")
        return

    nombre_archivo = generar_nombre_excel_desde_df(
        df,
        filtrado=filtrado
    )

    st.download_button(
        texto_boton,
        data=convertir_excel(df),
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key
    )


# ======================================================
# BOTÓN DE DESCARGA SEGÚN SELECCIÓN ACTUAL DEL FORMULARIO
# ======================================================

def boton_descargar_excel_formulario(
    df,
    direccion_regional,
    delegacion,
    texto_boton="⬇️ Descargar Excel actualizado",
    key=None
):
    if df is None or df.empty:
        st.info("No hay registros para descargar.")
        return

    nombre_archivo = generar_nombre_excel(
        direccion_regional=direccion_regional,
        delegacion=delegacion,
        filtrado=False
    )

    st.download_button(
        texto_boton,
        data=convertir_excel(df),
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key
    )
    # ======================================================
# PARTE 7 DE 10
# SIDEBAR, CARGA DE EXCEL EXISTENTE, MENÚ PRINCIPAL
# Y ENCABEZADOS VISUALES DE LA APP
# ======================================================


# ======================================================
# MOSTRAR LOGO EN SIDEBAR
# ======================================================

mostrar_logo()


# ======================================================
# TÍTULO DEL SIDEBAR
# ======================================================

st.sidebar.markdown("## Sistema PUMI 2026")

st.sidebar.markdown(
    """
    <div style="
        background:#FFFFFF;
        padding:14px;
        border-radius:14px;
        border-left:6px solid #B88A2A;
        box-shadow:0px 4px 12px rgba(0,0,0,0.10);
        margin-bottom:15px;
        font-size:14px;
    ">
    Modo formulario local con exportación a Excel.
    </div>
    """,
    unsafe_allow_html=True
)


# ======================================================
# CARGA DE EXCEL EXISTENTE
# Permite subir un Excel generado antes para continuar
# agregando, editando o eliminando registros.
# ======================================================

archivo_excel_cargado = st.sidebar.file_uploader(
    "Subir Excel existente para continuar registros",
    type=["xlsx"]
)

if archivo_excel_cargado is not None:
    if st.session_state.archivo_cargado_nombre != archivo_excel_cargado.name:
        df_cargado = cargar_excel_existente(archivo_excel_cargado)

        if not df_cargado.empty:
            st.session_state.registros_pumi = df_cargado
            st.session_state.archivo_cargado_nombre = archivo_excel_cargado.name
            st.sidebar.success("Excel cargado correctamente.")
        else:
            st.sidebar.warning("El Excel cargado no contiene registros válidos.")


# ======================================================
# MENÚ PRINCIPAL
# ======================================================

menu = st.sidebar.radio(
    "Menú principal",
    [
        "Inicio",
        "Registrar actividad",
        "Registros cargados",
        "Dashboard"
    ]
)


# ======================================================
# ENCABEZADO INSTITUCIONAL SUPERIOR
# ======================================================

mostrar_encabezado_institucional()


# ======================================================
# TÍTULO PRINCIPAL
# ======================================================

mostrar_titulo_principal()
# ======================================================
# PARTE 8 DE 10
# PANTALLA DE INICIO Y PRIMERA MITAD DEL FORMULARIO:
# DATOS GENERALES, ACTIVIDAD Y PARTICIPANTES
# ======================================================


# ======================================================
# INICIO
# ======================================================

if menu == "Inicio":

    st.markdown(
        """
        <div class="card-pumi">
            <div class="subtitulo-pumi">
                Bienvenido al Sistema PUMI 2026
            </div>
            <div class="texto-pumi">
                Esta aplicación permite llenar registros de actividades,
                conservarlos temporalmente durante la sesión y descargarlos
                en un archivo Excel.
                <br><br>
                También permite subir nuevamente un Excel generado anteriormente
                para continuar agregando o actualizando información sin necesidad
                de conexión a una base de datos.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    df_actual = st.session_state.registros_pumi

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Registros cargados", len(df_actual))

    with col2:
        total_participantes = 0

        if not df_actual.empty and "Cantidad Participantes" in df_actual.columns:
            total_participantes = pd.to_numeric(
                df_actual["Cantidad Participantes"],
                errors="coerce"
            ).fillna(0).sum()

        st.metric("Participantes", int(total_participantes))

    with col3:
        programas = 0

        if not df_actual.empty and "Programa" in df_actual.columns:
            programas = (
                df_actual["Programa"]
                .replace("", pd.NA)
                .dropna()
                .nunique()
            )

        st.metric("Programas registrados", programas)

    st.markdown("---")

    if not df_actual.empty:
        boton_descargar_excel(
            df_actual,
            texto_boton="⬇️ Descargar Excel actual",
            filtrado=False,
            key="descarga_inicio"
        )


# ======================================================
# REGISTRAR ACTIVIDAD
# PRIMERA MITAD DEL FORMULARIO
# ======================================================

elif menu == "Registrar actividad":

    st.markdown("## Registrar nueva actividad preventiva")

    st.markdown(
        """
        <div class="card-azul">
            <div class="texto-pumi">
                Complete la información de la actividad. Cada registro se agregará
                a la tabla temporal y luego podrá descargarse en Excel.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


    # ==================================================
    # DATOS GENERALES DEL REGISTRO
    # ==================================================

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

        st.session_state.ultima_direccion_regional = direccion_regional

        delegaciones_filtradas = obtener_delegaciones_por_region(
            direccion_regional
        )

        delegacion = st.selectbox(
            "Delegación",
            delegaciones_filtradas
            if delegaciones_filtradas
            else ["Sin datos disponibles"]
        )

        st.session_state.ultima_delegacion = delegacion

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


    # ==================================================
    # PARTICIPANTES
    # ==================================================

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

    suma_edades = (
        edad_10_18 +
        edad_19_30 +
        edad_31_45 +
        edad_46_mas
    )

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
# PARTE 9 DE 10
# SEGUNDA MITAD DEL FORMULARIO:
# UBICACIÓN TERRITORIAL, LUGAR, GEOREFERENCIA,
# INFORMACIÓN COMPLEMENTARIA, GUARDAR Y DESCARGAR
# ======================================================


    # ==================================================
    # UBICACIÓN TERRITORIAL
    # ==================================================

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


    # ==================================================
    # LUGAR DE REALIZACIÓN
    # ==================================================

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


    # ==================================================
    # MAPA Y GEOREFERENCIA
    # La función ya no devuelve ni guarda Dirección Mapa.
    # Esa columna queda siempre vacía y no sale en Excel.
    # ==================================================

    bloque_georreferencia_formulario(
        provincia=provincia,
        canton=canton,
        distrito=distrito,
        lugar=lugar,
        delegacion=delegacion,
        programa=programa
    )


    # ==================================================
    # INFORMACIÓN COMPLEMENTARIA
    # ==================================================

    st.markdown(
        """
        <div class="bloque-datos">
            <b>Información complementaria</b>
        </div>
        """,
        unsafe_allow_html=True
    )

    instituciones_participantes = st.text_area(
        "Instituciones participantes"
    )

    numero_referencia = st.text_input(
        "Número de referencia"
    )

    numero_expediente = st.text_input(
        "Número de expediente referencia"
    )

    observaciones = st.text_area(
        "Observaciones"
    )


    # ==================================================
    # GUARDAR REGISTRO
    # ==================================================

    st.markdown("---")

    if st.button("💾 Agregar registro al Excel temporal"):

        if cantidad > 0 and suma_sexo != cantidad:
            st.error(
                "No se puede agregar el registro porque la suma de hombres "
                "y mujeres no coincide con el total de participantes."
            )
            st.stop()

        if cantidad > 0 and suma_edades != cantidad:
            st.error(
                "No se puede agregar el registro porque la suma de rangos de edad "
                "no coincide con el total de participantes."
            )
            st.stop()

        nuevo_id = generar_id_consecutivo()

        registro = {
            "ID": nuevo_id,
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
            "Dirección Mapa": "",
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
            "Instituciones Participantes": instituciones_participantes,
            "Número de Referencia": numero_referencia,
            "Número de Expediente Referencia": numero_expediente,
            "Observaciones": observaciones,
            "Usuario Registra": usuario
        }

        agregar_registro_session(registro)

        st.success(
            "Registro agregado correctamente. Puede seguir agregando más registros "
            "o descargar el Excel desde esta misma pantalla."
        )

        st.session_state.latitud_registro = ""
        st.session_state.longitud_registro = ""
        st.session_state.direccion_encontrada_mapa = ""

        st.rerun()


    # ==================================================
    # DESCARGA RÁPIDA DESDE EL FORMULARIO
    # ==================================================

    df_actual = st.session_state.registros_pumi

    if not df_actual.empty:
        st.markdown("### Descargar registros actuales")

        boton_descargar_excel_formulario(
            df=df_actual,
            direccion_regional=direccion_regional,
            delegacion=delegacion,
            texto_boton="⬇️ Descargar Excel actualizado",
            key="descarga_formulario"
        )

        st.markdown("### Vista previa de registros cargados")

        st.dataframe(
            df_actual,
            use_container_width=True,
            hide_index=True
        )
        # ======================================================
# PARTE 10 DE 10
# REGISTROS CARGADOS, EDICIÓN, ELIMINACIÓN,
# DASHBOARD, MAPA GENERAL Y DESCARGA GLOBAL
# ======================================================


elif menu == "Registros cargados":

    st.markdown("## Registros cargados en la sesión")

    df_actual = st.session_state.registros_pumi.copy()

    if df_actual.empty:
        st.info("Aún no hay registros cargados o agregados.")
    else:
        st.markdown(
            """
            <div class="card-dorado">
                <div class="texto-pumi">
                    En esta sección puede revisar los registros agregados,
                    eliminar registros específicos, editar información y descargar
                    el Excel actualizado.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.dataframe(
            df_actual,
            use_container_width=True,
            hide_index=True
        )

        st.markdown("### Descargar Excel actualizado")

        boton_descargar_excel(
            df_actual,
            texto_boton="⬇️ Descargar Excel",
            filtrado=False,
            key="descarga_registros"
        )

        st.markdown("---")
        st.markdown("### Eliminar registro")

        ids_disponibles = df_actual["ID"].astype(str).tolist()

        id_eliminar = st.selectbox(
            "Seleccione el ID del registro que desea eliminar",
            ids_disponibles
        )

        if st.button("🗑️ Eliminar registro seleccionado"):
            eliminar_registro_session(id_eliminar)
            st.success("Registro eliminado correctamente.")
            st.rerun()

        st.markdown("---")
        st.markdown("### Editar registro existente")

        id_editar = st.selectbox(
            "Seleccione el ID del registro que desea editar",
            ids_disponibles,
            key="id_editar_registro"
        )

        fila_editar = df_actual[
            df_actual["ID"].astype(str) == str(id_editar)
        ]

        if not fila_editar.empty:
            fila = fila_editar.iloc[0].to_dict()

            with st.expander("Abrir formulario de edición", expanded=False):

                col1, col2 = st.columns(2)

                with col1:
                    fecha_actividad_edit = st.text_input(
                        "Fecha Actividad",
                        value=str(fila.get("Fecha Actividad", ""))
                    )

                    hora_actividad_edit = st.text_input(
                        "Hora Actividad",
                        value=str(fila.get("Hora Actividad", ""))
                    )

                    direccion_regional_edit = st.text_input(
                        "Dirección Regional",
                        value=str(fila.get("Dirección Regional", ""))
                    )

                    delegacion_edit = st.text_input(
                        "Delegación",
                        value=str(fila.get("Delegación", ""))
                    )

                    responde_a_edit = st.text_input(
                        "Responde a",
                        value=str(fila.get("Responde a", ""))
                    )

                    programa_edit = st.text_input(
                        "Programa",
                        value=str(fila.get("Programa", ""))
                    )

                    actividad_edit = st.text_area(
                        "Actividad",
                        value=str(fila.get("Actividad", ""))
                    )

                    provincia_edit = st.text_input(
                        "Provincia",
                        value=str(fila.get("Provincia", ""))
                    )

                    canton_edit = st.text_input(
                        "Cantón",
                        value=str(fila.get("Cantón", ""))
                    )

                    distrito_edit = st.text_input(
                        "Distrito",
                        value=str(fila.get("Distrito", ""))
                    )

                    tipo_lugar_edit = st.text_input(
                        "Tipo Lugar",
                        value=str(fila.get("Tipo Lugar", ""))
                    )

                    lugar_edit = st.text_input(
                        "Lugar",
                        value=str(fila.get("Lugar", ""))
                    )

                    centro_educativo_edit = st.text_input(
                        "Centro Educativo",
                        value=str(fila.get("Centro Educativo", ""))
                    )

                    codigo_presupuestario_edit = st.text_input(
                        "Código Presupuestario",
                        value=str(fila.get("Código Presupuestario", ""))
                    )

                    direccion_mapa_edit = st.text_input(
                        "Dirección Mapa",
                        value=str(fila.get("Dirección Mapa", ""))
                    )

                    latitud_edit = st.text_input(
                        "Latitud",
                        value=str(fila.get("Latitud", ""))
                    )

                    longitud_edit = st.text_input(
                        "Longitud",
                        value=str(fila.get("Longitud", ""))
                    )

                with col2:
                    responsable_edit = st.text_input(
                        "Responsable",
                        value=str(fila.get("Responsable", ""))
                    )

                    cantidad_edit = st.number_input(
                        "Cantidad Participantes",
                        min_value=0,
                        step=1,
                        value=int(
                            pd.to_numeric(
                                fila.get("Cantidad Participantes", 0),
                                errors="coerce"
                            ) or 0
                        )
                    )

                    hombres_edit = st.number_input(
                        "Cantidad Hombres",
                        min_value=0,
                        step=1,
                        value=int(
                            pd.to_numeric(
                                fila.get("Cantidad Hombres", 0),
                                errors="coerce"
                            ) or 0
                        )
                    )

                    mujeres_edit = st.number_input(
                        "Cantidad Mujeres",
                        min_value=0,
                        step=1,
                        value=int(
                            pd.to_numeric(
                                fila.get("Cantidad Mujeres", 0),
                                errors="coerce"
                            ) or 0
                        )
                    )

                    edad_10_18_edit = st.number_input(
                        "Edad 10 a 18",
                        min_value=0,
                        step=1,
                        value=int(
                            pd.to_numeric(
                                fila.get("Edad 10 a 18", 0),
                                errors="coerce"
                            ) or 0
                        )
                    )

                    edad_19_30_edit = st.number_input(
                        "Edad 19 a 30",
                        min_value=0,
                        step=1,
                        value=int(
                            pd.to_numeric(
                                fila.get("Edad 19 a 30", 0),
                                errors="coerce"
                            ) or 0
                        )
                    )

                    edad_31_45_edit = st.number_input(
                        "Edad 31 a 45",
                        min_value=0,
                        step=1,
                        value=int(
                            pd.to_numeric(
                                fila.get("Edad 31 a 45", 0),
                                errors="coerce"
                            ) or 0
                        )
                    )

                    edad_46_mas_edit = st.number_input(
                        "Edad 46 en adelante",
                        min_value=0,
                        step=1,
                        value=int(
                            pd.to_numeric(
                                fila.get("Edad 46 en adelante", 0),
                                errors="coerce"
                            ) or 0
                        )
                    )

                    instituciones_edit = st.text_area(
                        "Instituciones Participantes",
                        value=str(fila.get("Instituciones Participantes", ""))
                    )

                    referencia_edit = st.text_input(
                        "Número de Referencia",
                        value=str(fila.get("Número de Referencia", ""))
                    )

                    expediente_edit = st.text_input(
                        "Número de Expediente Referencia",
                        value=str(fila.get("Número de Expediente Referencia", ""))
                    )

                    observaciones_edit = st.text_area(
                        "Observaciones",
                        value=str(fila.get("Observaciones", ""))
                    )

                    usuario_edit = st.text_input(
                        "Usuario Registra",
                        value=str(fila.get("Usuario Registra", ""))
                    )

                suma_sexo_edit = hombres_edit + mujeres_edit

                suma_edades_edit = (
                    edad_10_18_edit +
                    edad_19_30_edit +
                    edad_31_45_edit +
                    edad_46_mas_edit
                )

                if cantidad_edit > 0:
                    if suma_sexo_edit != cantidad_edit:
                        st.warning(
                            f"La suma de hombres y mujeres ({suma_sexo_edit}) "
                            f"no coincide con el total ({cantidad_edit})."
                        )

                    if suma_edades_edit != cantidad_edit:
                        st.warning(
                            f"La suma de rangos de edad ({suma_edades_edit}) "
                            f"no coincide con el total ({cantidad_edit})."
                        )

                if st.button("💾 Guardar cambios del registro"):

                    if cantidad_edit > 0 and suma_sexo_edit != cantidad_edit:
                        st.error(
                            "No se puede guardar porque la suma de hombres y mujeres "
                            "no coincide con el total de participantes."
                        )
                        st.stop()

                    if cantidad_edit > 0 and suma_edades_edit != cantidad_edit:
                        st.error(
                            "No se puede guardar porque la suma de rangos de edad "
                            "no coincide con el total de participantes."
                        )
                        st.stop()

                    nuevos_datos = {
                        "ID": id_editar,
                        "Fecha Actividad": fecha_actividad_edit,
                        "Hora Actividad": hora_actividad_edit,
                        "Dirección Regional": direccion_regional_edit,
                        "Delegación": delegacion_edit,
                        "Responde a": responde_a_edit,
                        "Programa": programa_edit,
                        "Actividad": actividad_edit,
                        "Provincia": provincia_edit,
                        "Cantón": canton_edit,
                        "Distrito": distrito_edit,
                        "Tipo Lugar": tipo_lugar_edit,
                        "Lugar": lugar_edit,
                        "Centro Educativo": centro_educativo_edit,
                        "Código Presupuestario": codigo_presupuestario_edit,
                        "Dirección Mapa": direccion_mapa_edit,
                        "Latitud": latitud_edit,
                        "Longitud": longitud_edit,
                        "Responsable": responsable_edit,
                        "Cantidad Participantes": cantidad_edit,
                        "Cantidad Hombres": hombres_edit,
                        "Cantidad Mujeres": mujeres_edit,
                        "Edad 10 a 18": edad_10_18_edit,
                        "Edad 19 a 30": edad_19_30_edit,
                        "Edad 31 a 45": edad_31_45_edit,
                        "Edad 46 en adelante": edad_46_mas_edit,
                        "Instituciones Participantes": instituciones_edit,
                        "Número de Referencia": referencia_edit,
                        "Número de Expediente Referencia": expediente_edit,
                        "Observaciones": observaciones_edit,
                        "Usuario Registra": usuario_edit
                    }

                    actualizado = actualizar_registro_session(
                        id_editar,
                        nuevos_datos
                    )

                    if actualizado:
                        st.success("Registro actualizado correctamente.")
                        st.rerun()
                    else:
                        st.error("No se pudo actualizar el registro.")


elif menu == "Dashboard":

    st.markdown("## Dashboard general")

    df_actual = st.session_state.registros_pumi.copy()

    if df_actual.empty:
        st.info("Aún no hay datos para mostrar en el dashboard.")
    else:
        df_metricas = limpiar_dataframe_para_metricas(df_actual)

        total_registros = len(df_metricas)
        total_participantes = int(df_metricas["Cantidad Participantes"].sum())
        total_hombres = int(df_metricas["Cantidad Hombres"].sum())
        total_mujeres = int(df_metricas["Cantidad Mujeres"].sum())

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("Total registros", total_registros)

        with col2:
            st.metric("Participantes", total_participantes)

        with col3:
            st.metric("Hombres", total_hombres)

        with col4:
            st.metric("Mujeres", total_mujeres)

        st.markdown("---")

        colf1, colf2, colf3 = st.columns(3)

        with colf1:
            filtro_programa = st.multiselect(
                "Filtrar por programa",
                sorted(
                    df_actual["Programa"]
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                )
            )

        with colf2:
            filtro_region = st.multiselect(
                "Filtrar por Dirección Regional",
                sorted(
                    df_actual["Dirección Regional"]
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                )
            )

        with colf3:
            filtro_provincia = st.multiselect(
                "Filtrar por provincia",
                sorted(
                    df_actual["Provincia"]
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                )
            )

        df_filtrado = df_actual.copy()

        if filtro_programa:
            df_filtrado = df_filtrado[
                df_filtrado["Programa"].isin(filtro_programa)
            ]

        if filtro_region:
            df_filtrado = df_filtrado[
                df_filtrado["Dirección Regional"].isin(filtro_region)
            ]

        if filtro_provincia:
            df_filtrado = df_filtrado[
                df_filtrado["Provincia"].isin(filtro_provincia)
            ]

        df_filtrado_metricas = limpiar_dataframe_para_metricas(df_filtrado)

        st.markdown("### Registros filtrados")

        st.dataframe(
            df_filtrado,
            use_container_width=True,
            hide_index=True
        )

        boton_descargar_excel(
            df_filtrado,
            texto_boton="⬇️ Descargar Excel filtrado",
            filtrado=True,
            key="descarga_dashboard_filtrado"
        )

        st.markdown("---")
        st.markdown("### Visualizaciones")

        if not df_filtrado_metricas.empty:

            if "Programa" in df_filtrado_metricas.columns:
                conteo_programa = (
                    df_filtrado_metricas
                    .groupby("Programa")
                    .size()
                    .reset_index(name="Cantidad")
                    .sort_values("Cantidad", ascending=False)
                )

                if not conteo_programa.empty:
                    fig_programa = px.bar(
                        conteo_programa,
                        x="Programa",
                        y="Cantidad",
                        title="Cantidad de registros por programa",
                        text="Cantidad"
                    )

                    fig_programa.update_layout(
                        title_x=0.5,
                        plot_bgcolor="white",
                        paper_bgcolor="white"
                    )

                    st.plotly_chart(
                        fig_programa,
                        use_container_width=True
                    )

            if "Dirección Regional" in df_filtrado_metricas.columns:
                conteo_region = (
                    df_filtrado_metricas
                    .groupby("Dirección Regional")
                    .size()
                    .reset_index(name="Cantidad")
                    .sort_values("Cantidad", ascending=False)
                )

                if not conteo_region.empty:
                    fig_region = px.bar(
                        conteo_region,
                        x="Dirección Regional",
                        y="Cantidad",
                        title="Cantidad de registros por Dirección Regional",
                        text="Cantidad"
                    )

                    fig_region.update_layout(
                        title_x=0.5,
                        plot_bgcolor="white",
                        paper_bgcolor="white"
                    )

                    st.plotly_chart(
                        fig_region,
                        use_container_width=True
                    )

            if "Provincia" in df_filtrado_metricas.columns:
                conteo_provincia = (
                    df_filtrado_metricas
                    .groupby("Provincia")
                    .size()
                    .reset_index(name="Cantidad")
                    .sort_values("Cantidad", ascending=False)
                )

                if not conteo_provincia.empty:
                    fig_provincia = px.pie(
                        conteo_provincia,
                        names="Provincia",
                        values="Cantidad",
                        title="Distribución de registros por provincia",
                        hole=0.35
                    )

                    fig_provincia.update_layout(
                        title_x=0.5,
                        paper_bgcolor="white"
                    )

                    st.plotly_chart(
                        fig_provincia,
                        use_container_width=True
                    )

            st.markdown("### Mapa de registros")

            mostrar_mapa_registros(
                df_filtrado,
                height=560,
                key="mapa_dashboard"
            )


# ======================================================
# DESCARGA GLOBAL EN SIDEBAR
# ======================================================

st.sidebar.markdown("---")
st.sidebar.markdown("### Descargar información")

df_sidebar = st.session_state.registros_pumi

if df_sidebar.empty:
    st.sidebar.info("No hay registros para descargar.")
else:
    nombre_sidebar = generar_nombre_excel_desde_df(
        df_sidebar,
        filtrado=False
    )

    st.sidebar.download_button(
        "⬇️ Descargar Excel actual",
        data=convertir_excel(df_sidebar),
        file_name=nombre_sidebar,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="descarga_sidebar_global"
    )

    if st.sidebar.button("🧹 Limpiar registros de la sesión"):
        st.session_state.registros_pumi = pd.DataFrame(columns=ENCABEZADOS)
        st.session_state.archivo_cargado_nombre = ""
        st.session_state.latitud_registro = ""
        st.session_state.longitud_registro = ""
        st.session_state.direccion_encontrada_mapa = ""
        st.sidebar.success("Registros limpiados.")
        st.rerun()
