import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from io import BytesIO
from PIL import Image
import os

# ======================================================
# CONFIGURACIÓN GENERAL
# ======================================================

st.set_page_config(
    page_title="PUMI 2026",
    page_icon="🛡️",
    layout="wide"
)

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1O-HNa1c4ppF0-ND7BUqKA2OzZDc1O0kTxZOMVDeA-GY/edit?gid=0#gid=0"
HOJA_REGISTRO = "REGISTRO_PUMI_2026"

COLOR_AZUL = "#003366"
COLOR_VERDE = "#1B7F3A"
COLOR_GRIS = "#F4F6F8"
COLOR_DORADO = "#C9A227"


# ======================================================
# ESTILOS
# ======================================================

st.markdown(
    f"""
    <style>
    .stApp {{
        background-color: {COLOR_GRIS};
    }}

    .titulo-principal {{
        background: linear-gradient(90deg, {COLOR_AZUL}, {COLOR_VERDE});
        padding: 25px;
        border-radius: 15px;
        text-align: center;
        color: white;
        margin-bottom: 25px;
    }}

    .card {{
        background-color: white;
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0px 2px 8px rgba(0,0,0,0.12);
        margin-bottom: 15px;
    }}

    .subtitulo {{
        color: {COLOR_AZUL};
        font-weight: bold;
        font-size: 22px;
    }}
    </style>
    """,
    unsafe_allow_html=True
)


# ======================================================
# CONEXIÓN GOOGLE SHEETS
# ======================================================

@st.cache_resource
def conectar_google_sheets():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )

    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url(SPREADSHEET_URL)
    return spreadsheet


def obtener_hoja(nombre_hoja):
    spreadsheet = conectar_google_sheets()

    try:
        return spreadsheet.worksheet(nombre_hoja)
    except gspread.WorksheetNotFound:
        hoja = spreadsheet.add_worksheet(title=nombre_hoja, rows=2000, cols=30)
        return hoja


def inicializar_hoja():
    hoja = obtener_hoja(HOJA_REGISTRO)
    datos = hoja.get_all_values()

    encabezados = [
        "ID",
        "Fecha Registro",
        "Dirección Regional",
        "Delegación",
        "Programa",
        "Actividad",
        "Provincia",
        "Cantón",
        "Distrito",
        "Lugar / Centro Educativo",
        "Responsable",
        "Cantidad Participantes",
        "Instituciones Participantes",
        "Plan Estratégico Relacionado",
        "Evidencia",
        "Observaciones",
        "Estado Revisión",
        "Usuario Registra"
    ]

    if len(datos) == 0:
        hoja.append_row(encabezados)

    return hoja


def cargar_datos():
    hoja = inicializar_hoja()
    registros = hoja.get_all_records()
    return pd.DataFrame(registros)


def guardar_registro(registro):
    hoja = inicializar_hoja()
    hoja.append_row(list(registro.values()))


# ======================================================
# FUNCIONES AUXILIARES
# ======================================================

def generar_id():
    return datetime.now().strftime("%Y%m%d%H%M%S")


def convertir_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="REGISTRO_PUMI_2026")
    return output.getvalue()


def mostrar_logo():
    if os.path.exists("logo_pumi.jpeg"):
        logo = Image.open("logo_pumi.jpeg")
        st.sidebar.image(logo, use_container_width=True)
    else:
        st.sidebar.warning("Logo PUMI no encontrado. Guardá el archivo como logo_pumi.jpeg")


# ======================================================
# LISTAS BASE
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

ESTADOS_REVISION = [
    "Pendiente de revisión",
    "Aprobado",
    "Con observaciones",
    "Rechazado"
]


# ======================================================
# INTERFAZ PRINCIPAL
# ======================================================

mostrar_logo()

st.sidebar.markdown("## Sistema PUMI 2026")
menu = st.sidebar.radio(
    "Menú principal",
    [
        "Inicio",
        "Registrar actividad",
        "Consultar registros",
        "Dashboard básico",
        "Configuración"
    ]
)

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
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Bienvenido al sistema PUMI 2026")
    st.write(
        "Esta aplicación permite registrar, consultar y consolidar las actividades "
        "de los Programas Policiales Preventivos, utilizando Google Sheets como base de datos."
    )
    st.markdown("</div>", unsafe_allow_html=True)

    try:
        df = cargar_datos()

        col1, col2, col3, col4 = st.columns(4)

        col1.metric("Registros", len(df))
        col2.metric("Programas", df["Programa"].nunique() if not df.empty else 0)
        col3.metric("Delegaciones", df["Delegación"].nunique() if not df.empty else 0)

        if not df.empty and "Cantidad Participantes" in df.columns:
            participantes = pd.to_numeric(df["Cantidad Participantes"], errors="coerce").fillna(0).sum()
            col4.metric("Participantes", int(participantes))
        else:
            col4.metric("Participantes", 0)

        st.success("Conexión con Google Sheets activa.")

    except Exception as e:
        st.error("No se pudo conectar con Google Sheets.")
        st.exception(e)


# ======================================================
# REGISTRAR ACTIVIDAD
# ======================================================

elif menu == "Registrar actividad":
    st.markdown("### Registrar nueva actividad preventiva")

    with st.form("form_registro_pumi"):
        col1, col2 = st.columns(2)

        with col1:
            direccion_regional = st.selectbox("Dirección Regional", REGIONES)
            delegacion = st.text_input("Delegación")
            programa = st.selectbox("Programa", PROGRAMAS)
            actividad = st.text_input("Actividad realizada")
            provincia = st.text_input("Provincia")
            canton = st.text_input("Cantón")
            distrito = st.text_input("Distrito")

        with col2:
            lugar = st.text_input("Lugar / Centro Educativo")
            responsable = st.text_input("Funcionario responsable")
            cantidad = st.number_input("Cantidad de participantes", min_value=0, step=1)
            instituciones = st.text_area("Instituciones participantes")
            plan = st.text_input("Plan estratégico relacionado")
            evidencia = st.text_input("Enlace de evidencia")
            usuario = st.text_input("Usuario que registra")

        observaciones = st.text_area("Observaciones")

        guardar = st.form_submit_button("Guardar registro")

        if guardar:
            if not delegacion or not actividad or not responsable:
                st.warning("Debe completar al menos Delegación, Actividad y Responsable.")
            else:
                registro = {
                    "ID": generar_id(),
                    "Fecha Registro": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                    "Dirección Regional": direccion_regional,
                    "Delegación": delegacion,
                    "Programa": programa,
                    "Actividad": actividad,
                    "Provincia": provincia,
                    "Cantón": canton,
                    "Distrito": distrito,
                    "Lugar / Centro Educativo": lugar,
                    "Responsable": responsable,
                    "Cantidad Participantes": cantidad,
                    "Instituciones Participantes": instituciones,
                    "Plan Estratégico Relacionado": plan,
                    "Evidencia": evidencia,
                    "Observaciones": observaciones,
                    "Estado Revisión": "Pendiente de revisión",
                    "Usuario Registra": usuario
                }

                guardar_registro(registro)
                st.success("Registro guardado correctamente en Google Sheets.")


# ======================================================
# CONSULTAR REGISTROS
# ======================================================

elif menu == "Consultar registros":
    st.markdown("### Consulta de registros PUMI")

    df = cargar_datos()

    if df.empty:
        st.info("Aún no hay registros guardados.")
    else:
        col1, col2, col3 = st.columns(3)

        with col1:
            filtro_programa = st.selectbox(
                "Filtrar por programa",
                ["Todos"] + sorted(df["Programa"].dropna().unique().tolist())
            )

        with col2:
            filtro_region = st.selectbox(
                "Filtrar por región",
                ["Todas"] + sorted(df["Dirección Regional"].dropna().unique().tolist())
            )

        with col3:
            filtro_estado = st.selectbox(
                "Filtrar por estado",
                ["Todos"] + sorted(df["Estado Revisión"].dropna().unique().tolist())
            )

        df_filtrado = df.copy()

        if filtro_programa != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Programa"] == filtro_programa]

        if filtro_region != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Dirección Regional"] == filtro_region]

        if filtro_estado != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Estado Revisión"] == filtro_estado]

        st.dataframe(df_filtrado, use_container_width=True)

        excel = convertir_excel(df_filtrado)

        st.download_button(
            label="Descargar registros en Excel",
            data=excel,
            file_name="registro_pumi_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# ======================================================
# DASHBOARD BÁSICO
# ======================================================

elif menu == "Dashboard básico":
    st.markdown("### Dashboard básico PUMI")

    df = cargar_datos()

    if df.empty:
        st.info("Aún no hay datos para graficar.")
    else:
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("#### Actividades por programa")
            conteo_programas = df["Programa"].value_counts()
            st.bar_chart(conteo_programas)

        with col2:
            st.markdown("#### Actividades por región")
            conteo_region = df["Dirección Regional"].value_counts()
            st.bar_chart(conteo_region)

        st.markdown("#### Participantes por programa")

        df["Cantidad Participantes"] = pd.to_numeric(
            df["Cantidad Participantes"],
            errors="coerce"
        ).fillna(0)

        participantes_programa = df.groupby("Programa")["Cantidad Participantes"].sum()
        st.bar_chart(participantes_programa)


# ======================================================
# CONFIGURACIÓN
# ======================================================

elif menu == "Configuración":
    st.markdown("### Configuración inicial")

    st.info("Verificación de conexión con Google Sheets.")

    try:
        spreadsheet = conectar_google_sheets()
        st.success("Conexión exitosa.")
        st.write("Base conectada:", spreadsheet.title)
        st.write("Hoja de registro:", HOJA_REGISTRO)

        hoja = inicializar_hoja()
        st.success("Hoja REGISTRO_PUMI_2026 lista para trabajar.")

    except Exception as e:
        st.error("Error de conexión.")
        st.exception(e)









