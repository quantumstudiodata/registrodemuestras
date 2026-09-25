# Lab — Registro de Muestras

Migrado de Google Apps Script a un backend propio (funciones serverless de Vercel +
API de Google Sheets/Drive con una cuenta de servicio). Ya no pasa por ningún
consentimiento de Google, así que no vuelve a aparecer el banner de "Esta aplicación
la ha creado un usuario de Google Apps Script".

Los datos siguen viviendo en las mismas hojas de Google Sheets de siempre — no hay
migración de datos, solo de dónde corre el código.

`apps-script-archive/` guarda el `Code.gs`/`Index.html` originales como respaldo
histórico; ya no se usan (puedes dejar esa implementación de Apps Script como está,
o borrarla más adelante, sin prisa).

## Qué reemplaza a qué

| Antes (Apps Script)              | Ahora                                              |
|-----------------------------------|-----------------------------------------------------|
| `Code.gs` (funciones del servidor)| `lib/muestras.js` + `api/rpc.js`                   |
| `SpreadsheetApp`                  | API de Google Sheets vía cuenta de servicio (`lib/sheetsHelper.js`) |
| `DriveApp` (fotos de perfil)      | API de Google Drive vía cuenta de servicio (`lib/drive.js`) |
| `InventarioLib.sincronizarTotalesDelDia` | `lib/inventarioSync.js` — escribe directo en la misma hoja de Inventario, sin tocar ese proyecto de Apps Script |
| `google.script.run`               | `fetch('/api/rpc')` (un solo cambio, en la función `servidor()` de `index.html`) |
| Hosting Apps Script (`/exec`)     | Vercel                                              |

Todo el resto de la lógica de negocio y del front-end quedó exactamente igual.

## Pasos que solo tú puedes hacer (una sola vez)

### 1. Cuenta de servicio de Google Cloud

1. Ve a [console.cloud.google.com](https://console.cloud.google.com), crea un proyecto
   (o usa uno existente).
2. Habilita **Google Sheets API** y **Google Drive API** (menú "APIs y servicios" →
   "Habilitar APIs y servicios").
3. Crea una cuenta de servicio ("APIs y servicios" → "Credenciales" → "Crear
   credenciales" → "Cuenta de servicio"). No necesita ningún rol de proyecto.
4. Entra a la cuenta de servicio recién creada → pestaña "Claves" → "Agregar clave" →
   "Crear clave nueva" → tipo **JSON**. Se descarga un archivo `.json`; guárdalo, es
   la única copia.
5. Copia el correo de la cuenta de servicio (algo como
   `nombre@proyecto.iam.gserviceaccount.com`) — lo necesitas en los siguientes pasos.

### 2. Compartir tus hojas de cálculo con la cuenta de servicio

1. Abre tu hoja de **Registro de Muestras**
   (`1lbItZAdR7NmAAAB4iykvryPWD4nX13lWwzKEG_ilLLY`) → botón "Compartir" → agrega el
   correo de la cuenta de servicio con permiso de **Editor**.
2. Si quieres que siga funcionando la sincronización automática de inventario, haz lo
   mismo con la hoja de **Inventario** (`14lTCBum167F1Muw8AtCBjTosIJ5yGgDNY9uyWjOh55s`).
   Si la omites, la app de Registro de Muestras funciona igual, solo sin sincronizar
   el inventario.

### 3. Carpeta de Drive para fotos de perfil

1. Crea una carpeta nueva en tu Google Drive (ej. "Fotos de Perfil - Lab").
2. Compártela con el correo de la cuenta de servicio, permiso **Editor**.
3. Copia el ID de la carpeta (está en la URL: `drive.google.com/drive/folders/ESTE_ES_EL_ID`).

### 4. Cuenta de Vercel y despliegue

1. Crea una cuenta en [vercel.com](https://vercel.com) (puedes entrar con tu cuenta de
   GitHub directamente).
2. "Add New..." → "Project" → importa este repositorio
   (`quantumstudiodata/registrodemuestras`).
3. En "Environment Variables", agrega (usa `.env.example` como referencia):
   - `GOOGLE_SERVICE_ACCOUNT_KEY` — pega el contenido completo del archivo `.json`
     descargado en el paso 1 (como una sola línea de texto).
   - `GOOGLE_SHEET_ID_MUESTRAS` — `1lbItZAdR7NmAAAB4iykvryPWD4nX13lWwzKEG_ilLLY`
   - `GOOGLE_SHEET_ID_INVENTARIO` — `14lTCBum167F1Muw8AtCBjTosIJ5yGgDNY9uyWjOh55s`
     (opcional, ver paso 2)
   - `DRIVE_FOLDER_FOTOS_ID` — el ID de la carpeta del paso 3.
4. Dale "Deploy". Cuando termine, te da una URL (ej. `registro-de-muestras.vercel.app`)
   — esa es la nueva dirección de la app, sin ningún banner de Google.
5. Cada vez que se haga `git push` a `master`, Vercel vuelve a desplegar solo, sin
   pasos manuales.

### 5. Dominio propio (opcional)

Si quieres una URL con tu propio dominio en vez de `*.vercel.app`, en el proyecto de
Vercel ve a "Settings" → "Domains" y sigue las instrucciones para conectar un dominio
que ya tengas.

## Diferencias de comportamiento a tener en cuenta

- **Fotos de perfil grandes**: el límite de tamaño de request en el plan gratuito de
  Vercel es de ~4.5MB. Una foto sin comprimir muy grande podría fallar al subirse; si
  pasa, lo más simple es comprimir la imagen antes de subirla.
- **Sincronización de inventario concurrente**: el `Code.gs` original usaba
  `LockService` para evitar que dos sincronizaciones casi simultáneas pisaran datos.
  Aquí no hay un candado equivalente entre invocaciones serverless. Para el volumen de
  este laboratorio el riesgo es mínimo, pero si algún día se vuelve un problema real,
  se puede agregar un candado externo (ej. con Vercel KV).
