import streamlit as st
import pandas as pd
from datetime import datetime
from arcgis.gis import GIS

st.set_page_config(
    page_title="Prueba ArcGIS",
    page_icon="🗺️",
    layout="centered"
)

# =========================
# CONEXIÓN ARCGIS
# =========================

@st.cache_resource
def conectar_arcgis():
    return GIS(
        st.secrets["ARCGIS_URL"],
        st.secrets["ARCGIS_USER"],
        st.secrets["ARCGIS_PASSWORD"]
    )


@st.cache_resource
def obtener_capa():
    gis = conectar_arcgis()
    item = gis.content.get(st.secrets["ARCGIS_ITEM_ID"])

    if item is None:
        st.error("No se encontró el item de ArcGIS. Revise el ARCGIS_ITEM_ID.")
        st.stop()

    # Si es Feature Layer con capas:
    if item.layers:
        return item.layers[0]

    # Si es tabla:
    if item.tables:
        return item.tables[0]

    st.error("El item no tiene capas ni tablas disponibles.")
    st.stop()


# =========================
# LOGIN
# =========================

if "logueado" not in st.session_state:
    st.session_state.logueado = False

if "usuario" not in st.session_state:
    st.session_state.usuario = ""

if not st.session_state.logueado:
    st.title("Ingreso de prueba")

    usuario = st.text_input("Usuario")
    clave = st.text_input("Clave", type="password")

    if st.button("Ingresar"):
        if usuario == st.secrets["APP_USER"] and clave == st.secrets["APP_PASSWORD"]:
            st.session_state.logueado = True
            st.session_state.usuario = usuario
            st.rerun()
        else:
            st.error("Usuario o clave incorrectos.")

    st.stop()


# =========================
# APP PRINCIPAL
# =========================

st.title("Mini app conectada a ArcGIS")

st.success(f"Sesión iniciada como: {st.session_state.usuario}")

if st.button("Cerrar sesión"):
    st.session_state.logueado = False
    st.session_state.usuario = ""
    st.rerun()

st.divider()

st.subheader("Registrar persona")

nombre = st.text_input("Nombre completo")
telefono = st.text_input("Número de teléfono")

if st.button("Guardar en ArcGIS"):
    if not nombre.strip():
        st.warning("Digite el nombre.")
        st.stop()

    if not telefono.strip():
        st.warning("Digite el teléfono.")
        st.stop()

    capa = obtener_capa()

    nuevo_registro = {
        "attributes": {
            "nombre": nombre.strip(),
            "telefono": telefono.strip(),
            "usuario_registra": st.session_state.usuario,
            "fecha_registro": int(datetime.now().timestamp() * 1000)
        }
    }

    resultado = capa.edit_features(adds=[nuevo_registro])

    exito = resultado.get("addResults", [{}])[0].get("success", False)

    if exito:
        st.success("Registro guardado correctamente en ArcGIS.")
        st.cache_data.clear()
        st.rerun()
    else:
        st.error("No se pudo guardar en ArcGIS.")
        st.write(resultado)

st.divider()

st.subheader("Registros guardados")

@st.cache_data(ttl=30)
def cargar_registros():
    capa = obtener_capa()
    resultado = capa.query(where="1=1", out_fields="*", return_geometry=False)
    datos = [feature.attributes for feature in resultado.features]
    return pd.DataFrame(datos)

df = cargar_registros()

if df.empty:
    st.info("Aún no hay registros.")
else:
    columnas_mostrar = [
        col for col in [
            "OBJECTID",
            "nombre",
            "telefono",
            "usuario_registra",
            "fecha_registro"
        ]
        if col in df.columns
    ]

    st.dataframe(
        df[columnas_mostrar],
        use_container_width=True,
        hide_index=True
    )
