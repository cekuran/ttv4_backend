# TaskTime backend

Google Apps Script WebApp (`Code.js`, runtime V8). Es el **único componente con lógica de negocio y datos** de TaskTime: autenticación multiusuario, roles y CRUD de tareas, etiquetas, registros de tiempo, temporizador activo y rutinas recurrentes. Almacena todo en Google Sheets.

```
frontend (index.html) → proxy (Cloudflare Worker) → backend (Apps Script /exec) → Google Sheets
```

Documentación general y deploy orquestado: ver [../CLAUDE.md](../CLAUDE.md) y [../tools/Set-ApiDeployment.ps1](../tools/Set-ApiDeployment.ps1).

## Modelo de almacenamiento

Dos niveles de spreadsheets:

| Nivel | Cómo se identifica | Contenido |
|-------|--------------------|-----------|
| **Maestra** | Script Property `SPREADSHEET_ID` (se fija con la acción pública `configurarSpreadsheetMaestro`) | Tablas de auth: `Usuarios`, `Spreadsheets`, `HojasUsuarios`, `Tokens`, `Config` |
| **Por usuario** | Vinculada en `HojasUsuarios` (`username`, `spreadsheet_id`, `por_defecto`) | Datos de tiempo: `Tasks`, `Labels`, `TimeEntries`, `Routines`, `DailyCompletions`, `WeeklyCompletions`, `MonthlyCompletions` |

Los esquemas (columnas) están en `AUTH_SCHEMA` y `SHEETS` al inicio de [Code.js](Code.js). Las pestañas de la hoja de usuario se crean/reparan solas con `ensureUserSchema_` en el primer `bootstrap`.

El temporizador activo se guarda en `PropertiesService.getUserProperties()` y no depende de la hoja abierta.

## Autenticación y roles

- `doGet` / `doPost` → `dispatchApi_()` valida la acción contra `API_ACTIONS`.
- Acciones públicas (`API_PUBLIC_ACTIONS`): `ping`, `authStatus`, `loginUsuario`, `logoutUsuario`, `configurarSpreadsheetMaestro`.
- El resto pasa por `_authadmin(token, action, ...args)`: valida el token (`Tokens`), resuelve la hoja activa del usuario y ejecuta la acción.
- Roles: `admin` y `basico`. Las acciones `*Admin` y `adminOverview` requieren `admin` (`requireAdmin_()`).
- Contraseñas: SHA-256 iterado (`PASSWORD_ITERATIONS = 1500`) con `salt` por usuario.
- Usuarios sin hojas vinculadas solo pueden usar `ALWAYS_ALLOWED_FOR_NO_HOJAS` (`bootstrap`, `bootstrapBase`, `listarMisHojas`, `cambiarHojaActiva`, `cambiarMiContrasena`, y las públicas).
- Primeros pasos: `configurarSpreadsheetMaestro(id, adminUser, adminPass)` siembra `Usuarios` con ese admin. Si no se indica, se usa el fallback `admin` / `admin1234` (cambiarlo desde **Mi cuenta**).

## Caché y concurrencia

- `CacheService` (script cache) para token (30 min), hojas de la maestra (5 min), datos de usuario (10 min) y verificación de esquema (30 min). Las entradas grandes se comprimen con gzip+base64.
- Las acciones que modifican datos corren dentro de `LockService` (`withLock_`); las de `READ_ONLY_ACTIONS` no.
- Las escrituras invalidan la caché correspondiente (`markUserDirty_`).

## Setup

Requisitos: Node.js, [clasp](https://github.com/google/clasp) (`npm i -g @google/clasp`) y una cuenta de Google con Apps Script habilitado.

```bash
cd backend
clasp login
clasp clone <SCRIPT_ID>      # o crear .clasp.json con scriptId y rootDir
```

`.clasp.json` está en `.gitignore` (no se versiona). `appsscript.json` define:

- `executeAs: USER_DEPLOYING`, `access: ANYONE_ANONYMOUS` (la seguridad la da el token de sesión propio, no Google).
- Scope `https://www.googleapis.com/auth/spreadsheets`.
- `timeZone: Europe/Madrid`.

Para configurar la maestra: crear un Google Sheet vacío, compartirlo con la cuenta que despliega y pegar su ID en la pantalla de **setup** del frontend (o definir la Script Property `SPREADSHEET_ID` a mano). Las pestañas de auth se crean automáticamente.

## Despliegue

Recomendado, desde la raíz del repo (publica una nueva versión y actualiza `release_info.txt`):

```powershell
pwsh ./tools/Set-ApiDeployment.ps1 -Description 'mensaje'
```

Manual:

```bash
cd backend
clasp push
clasp deploy -d "mensaje"
```

Si cambia el `DeploymentId`, actualizar el `TARGET` del proxy (ver [../proxy/README.md](../proxy/README.md)).

## Añadir un endpoint

1. Añadir la acción a `API_ACTIONS` en [Code.js](Code.js). Si no requiere token, añadirla también a `API_PUBLIC_ACTIONS`. Si solo lee, añadirla a `READ_ONLY_ACTIONS`.
2. Implementar la función en `Code.js`. Leer/escribir con `readRows_` / `appendRow_` / `updateRow_` / `deleteRow_` (usan `ssActiva_()`); autorizar con `currentUser_()` / `requireUsuario_()` / `requireAdmin_()`.
3. Frontend: `await call('miAccion', arg1, arg2)`. Los errores llegan como `Error` desde `rawCall`.

Convención: los helpers con sufijo `_` son internos del módulo.

## Respuesta de la API

Todas las respuestas son JSON: `{ ok: true, data }` o `{ ok: false, error }`. `GET` sin `action` devuelve `{ service: 'tasktime', version }`. `POST` espera `{ action, args, token }` con `Content-Type: text/plain` (Apps Script rechaza `application/json`).

## Estructura

```
backend/
├── Code.js            # todo el backend
├── Code.js.bak        # respaldo local (ignorado por git)
├── appsscript.json    # manifest (V8, scopes, webapp)
├── .clasp.json        # scriptId + rootDir (ignorado por git)
└── .gitignore
```
