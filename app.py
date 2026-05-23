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

ENCABEZADOS = [
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
    "R11 Chorotega Norte",
    "R6 Heredia"
]

ESTADOS_REVISION = [
    "Pendiente de revisión",
    "Aprobado",
    "Con observaciones",
    "Rechazado"
]

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
        return spreadsheet.add_worksheet(title=nombre_hoja, rows=3000, cols=30)


def inicializar_hoja():
    hoja = obtener_hoja(HOJA_REGISTRO)
    datos = hoja.get_all_values()

    if len(datos) == 0:
        hoja.append_row(ENCABEZADOS)
    else:
        primera_fila = datos[0]

        if primera_fila != ENCABEZADOS:
            hoja.insert_row(ENCABEZADOS, index=1)

    return hoja


def cargar_datos():
    hoja = inicializar_hoja()
    datos = hoja.get_all_values()

    if len(datos) <= 1:
        return pd.DataFrame(columns=ENCABEZADOS)

    df = pd.DataFrame(datos[1:], columns=ENCABEZADOS)
    return df


def guardar_registro(registro):
    hoja = inicializar_hoja()
    hoja.append_row([registro.get(col, "") for col in ENCABEZADOS])


def actualizar_registro_por_id(id_registro, nuevos_datos):
    hoja = inicializar_hoja()
    datos = hoja.get_all_values()

    for i, fila in enumerate(datos[1:], start=2):
        if str(fila[0]) == str(id_registro):
            nueva_fila = [nuevos_datos.get(col, "") for col in ENCABEZADOS]
            hoja.update(f"A{i}:R{i}", [nueva_fila])
            return True

    return False


def eliminar_registro_por_id(id_registro):
    hoja = inicializar_hoja()
    datos = hoja.get_all_values()

    for i, fila in enumerate(datos[1:], start=2):
        if str(fila[0]) == str(id_registro):
            hoja.delete_rows(i)
            return True

    return False

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
    elif os.path.exists("logo_pumi.jpg"):
        logo = Image.open("logo_pumi.jpg")
        st.sidebar.image(logo, use_container_width=True)
    elif os.path.exists("logo_pumi.png"):
        logo = Image.open("logo_pumi.png")
        st.sidebar.image(logo, use_container_width=True)
    else:
        st.sidebar.warning("Logo PUMI no encontrado.")

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
        "Consultar / editar registros",
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
    st.markdown("### Bienvenido al sistema PUMI 2026")

    try:
        df = cargar_datos()

        col1, col2, col3, col4 = st.columns(4)

        col1.metric("Registros", len(df))
        col2.metric("Programas", df["Programa"].nunique() if not df.empty else 0)
        col3.metric("Delegaciones", df["Delegación"].nunique() if not df.empty else 0)

        participantes = pd.to_numeric(
            df["Cantidad Participantes"],
            errors="coerce"
        ).fillna(0).sum() if not df.empty else 0

        col4.metric("Participantes", int(participantes))

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
                st.success("Registro guardado correctamente.")

# ======================================================
# CONSULTAR / EDITAR / ELIMINAR
# ======================================================

elif menu == "Consultar / editar registros":
    st.markdown("### Consulta, edición y eliminación de registros PUMI")

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

        st.markdown("---")
        st.markdown("### Editar o eliminar registro")

        ids = df_filtrado["ID"].astype(str).tolist()

        id_seleccionado = st.selectbox(
            "Seleccione el ID del registro",
            ids
        )

        registro_actual = df[df["ID"].astype(str) == str(id_seleccionado)].iloc[0].to_dict()

        accion = st.radio(
            "Acción a realizar",
            ["Editar registro", "Eliminar registro"],
            horizontal=True
        )

        if accion == "Editar registro":
            with st.form("form_editar_registro"):
                col1, col2 = st.columns(2)

                with col1:
                    nueva_fecha = st.text_input("Fecha Registro", registro_actual["Fecha Registro"])
                    nueva_region = st.selectbox(
                        "Dirección Regional",
                        REGIONES,
                        index=REGIONES.index(registro_actual["Dirección Regional"]) if registro_actual["Dirección Regional"] in REGIONES else 0
                    )
                    nueva_delegacion = st.text_input("Delegación", registro_actual["Delegación"])
                    nuevo_programa = st.selectbox(
                        "Programa",
                        PROGRAMAS,
                        index=PROGRAMAS.index(registro_actual["Programa"]) if registro_actual["Programa"] in PROGRAMAS else 0
                    )
                    nueva_actividad = st.text_input("Actividad realizada", registro_actual["Actividad"])
                    nueva_provincia = st.text_input("Provincia", registro_actual["Provincia"])
                    nuevo_canton = st.text_input("Cantón", registro_actual["Cantón"])
                    nuevo_distrito = st.text_input("Distrito", registro_actual["Distrito"])

                with col2:
                    nuevo_lugar = st.text_input("Lugar / Centro Educativo", registro_actual["Lugar / Centro Educativo"])
                    nuevo_responsable = st.text_input("Responsable", registro_actual["Responsable"])
                    nueva_cantidad = st.number_input(
                        "Cantidad Participantes",
                        min_value=0,
                        step=1,
                        value=int(float(registro_actual["Cantidad Participantes"])) if str(registro_actual["Cantidad Participantes"]).strip() else 0
                    )
                    nuevas_instituciones = st.text_area("Instituciones Participantes", registro_actual["Instituciones Participantes"])
                    nuevo_plan = st.text_input("Plan Estratégico Relacionado", registro_actual["Plan Estratégico Relacionado"])
                    nueva_evidencia = st.text_input("Evidencia", registro_actual["Evidencia"])
                    nuevo_estado = st.selectbox(
                        "Estado Revisión",
                        ESTADOS_REVISION,
                        index=ESTADOS_REVISION.index(registro_actual["Estado Revisión"]) if registro_actual["Estado Revisión"] in ESTADOS_REVISION else 0
                    )
                    nuevo_usuario = st.text_input("Usuario Registra", registro_actual["Usuario Registra"])

                nuevas_observaciones = st.text_area("Observaciones", registro_actual["Observaciones"])

                actualizar = st.form_submit_button("Guardar cambios")

                if actualizar:
                    nuevos_datos = {
                        "ID": registro_actual["ID"],
                        "Fecha Registro": nueva_fecha,
                        "Dirección Regional": nueva_region,
                        "Delegación": nueva_delegacion,
                        "Programa": nuevo_programa,
                        "Actividad": nueva_actividad,
                        "Provincia": nueva_provincia,
                        "Cantón": nuevo_canton,
                        "Distrito": nuevo_distrito,
                        "Lugar / Centro Educativo": nuevo_lugar,
                        "Responsable": nuevo_responsable,
                        "Cantidad Participantes": nueva_cantidad,
                        "Instituciones Participantes": nuevas_instituciones,
                        "Plan Estratégico Relacionado": nuevo_plan,
                        "Evidencia": nueva_evidencia,
                        "Observaciones": nuevas_observaciones,
                        "Estado Revisión": nuevo_estado,
                        "Usuario Registra": nuevo_usuario
                    }

                    actualizado = actualizar_registro_por_id(id_seleccionado, nuevos_datos)

                    if actualizado:
                        st.success("Registro actualizado correctamente.")
                        st.rerun()
                    else:
                        st.error("No se encontró el registro para actualizar.")

        else:
            st.warning("Esta acción eliminará el registro seleccionado de forma permanente.")

            confirmar = st.checkbox("Confirmo que deseo eliminar este registro")

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

        excel = convertir_excel(df_filtrado)

        st.download_button(
            label="Descargar registros filtrados en Excel",
            data=excel,
            file_name="registro_pumi_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ======================================================
# DASHBOARD
# ======================================================

elif menu == "Dashboard básico":
    st.markdown("### Dashboard básico PUMI")

    df = cargar_datos()

    if df.empty:
        st.info("Aún no hay datos para graficar.")
    else:
        df["Cantidad Participantes"] = pd.to_numeric(
            df["Cantidad Participantes"],
            errors="coerce"
        ).fillna(0)

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("#### Actividades por programa")
            st.bar_chart(df["Programa"].value_counts())

        with col2:
            st.markdown("#### Actividades por región")
            st.bar_chart(df["Dirección Regional"].value_counts())

        st.markdown("#### Participantes por programa")
        st.bar_chart(df.groupby("Programa")["Cantidad Participantes"].sum())

# ======================================================
# CONFIGURACIÓN
# ======================================================

elif menu == "Configuración":
    st.markdown("### Configuración inicial")

    try:
        spreadsheet = conectar_google_sheets()
        hoja = inicializar_hoja()

        st.success("Conexión exitosa.")
        st.write("Base conectada:", spreadsheet.title)
        st.write("Hoja activa:", HOJA_REGISTRO)
        st.write("Encabezados configurados correctamente.")

    except Exception as e:
        st.error("Error de conexión.")
        st.exception(e)



