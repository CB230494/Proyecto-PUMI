export const CONFIG = {
  portalUrl: "https://msp-cr.maps.arcgis.com",
  oauthClientId: "COLOQUE_AQUI_SU_CLIENT_ID",
  oauthRedirectUri: window.location.origin + window.location.pathname,
  appTitle: "PUMI 2026",

  layers: {
    actividades: "https://services5.arcgis.com/O0HjrwhVoKXT72lM/arcgis/rest/services/PUMI_ACTIVIDADES/FeatureServer/0",
    delegaciones: "https://services5.arcgis.com/O0HjrwhVoKXT72lM/arcgis/rest/services/Capa_1/FeatureServer/0",
    usuarios: "https://services5.arcgis.com/O0HjrwhVoKXT72lM/arcgis/rest/services/PUMI_USUARIOS/FeatureServer/0",
    catalogos: "https://services5.arcgis.com/O0HjrwhVoKXT72lM/arcgis/rest/services/PUMI_CATALOGOS/FeatureServer/0",
    resumen: "https://services5.arcgis.com/O0HjrwhVoKXT72lM/arcgis/rest/services/PUMI_RESUMEN/FeatureServer/0",
    validaciones: "https://services5.arcgis.com/O0HjrwhVoKXT72lM/arcgis/rest/services/PUMI_VALIDACIONES/FeatureServer/0",
    bitacora: "https://services5.arcgis.com/O0HjrwhVoKXT72lM/arcgis/rest/services/PUMI_BITACORA/FeatureServer/0",
    notificaciones: ""
  },

  fieldMap: {
    objectId: "OBJECTID",
    usuario: "usuario",
    nombre: "nombre",
    rol: "rol",
    direccionRegional: "direccion_regional",
    delegacion: "delegacion",
    programa: "programa",
    activo: "activo",

    fechaActividad: "fecha_actividad",
    horaActividad: "hora_actividad",
    actividad: "actividad",
    provincia: "provincia",
    canton: "canton",
    distrito: "distrito",
    tipoLugar: "tipo_lugar",
    lugar: "lugar",
    responsable: "responsable",
    avance: "avance",
    meta: "meta",
    pendiente: "pendiente",
    participantes: "cantidad_participantes",
    hombres: "cantidad_hombres",
    mujeres: "cantidad_mujeres",
    edad1018: "edad_10_18",
    edad1930: "edad_19_30",
    edad3145: "edad_31_45",
    edad46Mas: "edad_46_mas",
    instituciones: "instituciones_participantes",
    numeroReferencia: "numero_referencia",
    expedienteReferencia: "numero_expediente_referencia",
    observaciones: "observaciones",
    estadoRegional: "estado_verificacion_regional",
    observacionRegional: "observaciones_verificacion_regional",
    estadoNacional: "estado_validacion",
    observacionNacional: "observaciones_validacion",
    usuarioRegistra: "usuario_registra",
    fechaRegistro: "fecha_migracion"
  },

  statuses: {
    draft: "Borrador",
    pendingRegional: "Pendiente de verificación",
    approvedRegional: "Verificada para envío",
    returnedDelegation: "Devuelta a delegación",
    observedRegional: "Con observaciones regionales",
    pendingNational: "Pendiente",
    approvedNational: "Aprobada",
    rejectedNational: "Rechazada",
    observedNational: "Con observaciones"
  }
};

