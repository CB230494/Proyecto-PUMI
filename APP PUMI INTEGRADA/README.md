# PUMI 2026 — Aplicación integrada para GitHub Pages

Esta carpeta contiene una aplicación web estática en HTML, CSS y JavaScript conectada con ArcGIS Online.

## Funciones incluidas

- Inicio de sesión mediante OAuth 2.0 de ArcGIS.
- Acceso por rol: Delegación, Regional, Coordinador Nacional, Nacional y Administrador.
- Registro de actividades.
- Flujo de revisión regional.
- Flujo de validación nacional.
- Dashboard, indicadores y mapas.
- Notificaciones internas derivadas de registros pendientes.
- Administración de usuarios autorizados.
- Escritura en actividades, validaciones y bitácora mediante `applyEdits`.
- Modo demostración para revisar el diseño sin ArcGIS.

## Antes de subir a GitHub

Abra `js/config.js` y cambie:

```js
oauthClientId: "COLOQUE_AQUI_SU_CLIENT_ID"
```

por el Client ID generado en ArcGIS Developer.

En ArcGIS, agregue como URI de redirección la dirección exacta de GitHub Pages:

```text
https://SU_USUARIO.github.io/NOMBRE_REPOSITORIO/
```

## Estructura

```text
PUMI_APP_INTEGRADA/
├── index.html
├── styles.css
├── README.md
├── js/
│   ├── app.js
│   ├── arcgis.js
│   └── config.js
└── assets/
```

## Publicar en GitHub Pages

1. Cree un repositorio nuevo.
2. Suba el contenido de esta carpeta a la raíz.
3. Entre en **Settings → Pages**.
4. Seleccione **Deploy from a branch**.
5. Seleccione `main` y `/root`.
6. Guarde.
7. Copie la URL pública y agréguela como URI de redirección en ArcGIS.

## Importante sobre los campos

La app usa los nombres configurados en `CONFIG.fieldMap`. Si algún campo de ArcGIS tiene un nombre distinto, cambie solamente su equivalencia en `js/config.js`.

Ejemplo:

```js
participantes: "cantidad_participantes"
```

Si en su capa el campo se llama `participantes_total`, cambie a:

```js
participantes: "participantes_total"
```

No debe cambiar el resto de la aplicación.

## Permisos necesarios en ArcGIS

Las capas deben permitir, según corresponda:

- Consultar.
- Agregar registros.
- Actualizar registros.

Los usuarios deben existir en `PUMI_USUARIOS` y tener al menos:

- `usuario`
- `nombre`
- `rol`
- `direccion_regional`
- `delegacion`
- `programa`
- `activo`

## Estados usados

### Regional

- Pendiente de verificación
- Verificada para envío
- Devuelta a delegación
- Con observaciones regionales

### Nacional

- Pendiente
- Aprobada
- Rechazada
- Con observaciones

Solo las actividades con estado nacional `Aprobada` se muestran como validadas en la vista nacional.

