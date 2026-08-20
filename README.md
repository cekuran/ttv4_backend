# family finances backend

Google Apps Script WebApp (`code.js`). Almacenamiento en Google Sheets: una hoja **maestra** con autenticación (Usuarios, Spreadsheets, HojasUsuarios, Tokens, Config) y N hojas de datos vinculadas por usuario/rol.

Es el **único repo con lógica de negocio y datos**. Los otros 3 existen para servirlo, desplegarlo o exponerlo.

## Repos relacionados

| Repo | Función | URL |
|------|---------|-----|
| family finances frontend | HTML único que consume este backend vía `fetch` | https://github.com/cekuran/ffv3_frontend |
| family finances proxy | Cloudflare Worker que envuelve `doGet`/`doPost` (evita el límite de CORS de Apps Script) | https://github.com/cekuran/ffv3_proxy |
| family finances tools | Script PowerShell que despliega backend + proxy + frontend en una sola corrida | https://github.com/cekuran/ffv3_tools |

Flujo: `frontend → proxy (Cloudflare) → backend (Apps Script) → Sheets`. Ver **tools** para el deploy orquestado.

## Setup inicial

Requisitos: Node.js, [clasp](https://github.com/google/clasp) (`npm i -g @google/clasp`), una cuenta de Google con Apps Script habilitado.

```bash
git clone git@github.com:cekuran/ffv3_backend.git
cd ffv3_backend
clasp login
clasp clone <SCRIPT_ID>          # crear/importar el Apps Script project
# o, si .clasp.json ya tiene scriptId, basta con:
clasp push
```

`.clasp.json` lleva el `scriptId` y `rootDir`. No commitear credenciales.

## Configuración básica

1. **Hoja maestra**: crear un Google Sheet con las pestañas `Usuarios`, `Spreadsheets`, `HojasUsuarios`, `Tokens`, `Config` (encabezados en la primera fila). Es la fuente de verdad de auth.
2. **Vincular la hoja al script**: en Apps Script editor → *Project Settings* → *Script Properties*, o desde código con `PropertiesService`.
3. **Hojas de datos**: cada usuario/rol se asocia a una hoja de datos vía `HojasUsuarios` (ver `_authadmin` en `code.js:116`). Los endpoints de datos usan `ssActiva_()`, nunca `ss_()` directo.
5. **Tokens**: la hoja `Tokens` valida cada request. El helper `_currentToken` cachea por-request.
6. **Sin hojas**: usuarios sin `HojasUsuarios` devuelven `{sin_hojas: true, sin_datos_financieros: true}` (ver `ALWAYS_ALLOWED_FOR_NO_HOJAS` en `code.js:103`). No se filtran datos de producción.

## Despliegue

Despliegue manual:

```bash
clasp push
clasp deploy -d "mensaje del cambio"
```

Para el flujo orquestado (backend + proxy + frontend), usar **family finances tools**.

## Añadir un endpoint nuevo

1. Agregar el nombre de acción a `API_ACTIONS` en `code.js:1053`. Si debe funcionar sin token, también a `API_PUBLIC_ACTIONS` en `code.js:1082`.
2. Implementar la función en `code.js` usando `currentUser_()` / `requireUsuario_()` / `requireAdmin_()` para auth, y `ssActiva_()` para datos.
3. Frontend lo invoca con `call('miAccion', arg1, arg2)`. Errores llegan como `payload.error`.

Convención: helpers con sufijo `_` son internos del módulo. No leer `_currentToken` directamente.

## Estructura

```
backend/
├── code.js              # todo el backend (Apps Script V8)
├── code.js.bak          # respaldo rotativo — ignorar para deploy
├── appsscript.json      # manifest del Apps Script project
└── .clasp.json          # scriptId + rootDir para clasp
```