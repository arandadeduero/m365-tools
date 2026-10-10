# m365-users

Herramienta de línea de comandos (CLI) para gestionar usuarios de Microsoft 365 a través de la API de Microsoft Graph. Permite importar usuarios en masa desde CSV, buscarlos, editarlos de forma interactiva, asignar managers, y corregir problemas de calidad de datos en el directorio.

---

## Requisitos previos

- **Node.js 24.0.0 o superior**
- Una suscripción de Microsoft 365 con permisos de administrador
- Un registro de aplicación en Azure Active Directory con los permisos delegados:
  - `User.ReadWrite.All`
  - `Directory.ReadWrite.All`
  - `MailboxSettings.ReadWrite` (necesario para `fix timezone`)

---

## Instalación

```bash
# Clonar el repositorio
git clone <url-del-repositorio>
cd m365-users

# Instalar dependencias
npm install

# (Opcional) Instalar globalmente para usar el comando m365-users desde cualquier directorio
npm install -g .
```

---

## Configuración

Copia el archivo de ejemplo y rellena los valores de tu tenant:

```bash
cp config.example.json config.json
```

Edita `config.json`:

```json
{
  "tenantId": "TU_TENANT_ID",
  "clientId": "TU_CLIENT_ID",
  "domain": "tudominio.com",
  "scopes": [
    "User.ReadWrite.All",
    "Directory.ReadWrite.All",
    "MailboxSettings.ReadWrite",
    "User-PasswordProfile.ReadWrite.All"
  ]
}
```

> El scope `User-PasswordProfile.ReadWrite.All` es necesario para `reset-password`. `MailboxSettings.ReadWrite` es necesario para `fix timezone`.

> **Importante:** `config.json` está incluido en `.gitignore` para evitar exponer credenciales. Nunca lo subas a un repositorio.

### Cómo obtener los valores

| Campo | Dónde encontrarlo |
|---|---|
| `tenantId` | Azure Portal → Azure Active Directory → Información general |
| `clientId` | Azure Portal → Registros de aplicaciones → tu app → Información general |
| `domain` | Dominio principal de tu tenant (ej. `contoso.com`) |

El campo `domain` es opcional pero recomendado: la herramienta lo usa para completar automáticamente los UPN (p. ej., escribe `jdoe` y se convierte en `jdoe@contoso.com`).

---

## Autenticación

La herramienta usa el flujo **Device Code** de OAuth 2.0 (autenticación delegada). No requiere contraseña ni secreto de aplicación almacenado; el usuario se autentica en el navegador.

```bash
m365-users login
```

Se mostrará un código y una URL. Abre la URL en tu navegador, introduce el código y autentica con tu cuenta de administrador.

La sesión completa de MSAL (incluyendo el *refresh token*) se guarda cifrada en `~/.config/m365-users/msal-cache.json` con AES-256-GCM y una clave derivada de la máquina. Esto permite **renovar el acceso de forma silenciosa** en arranques posteriores sin volver a mostrar el código de dispositivo. El *refresh token* de Azure AD tiene una vigencia de **90 días renovables**: mientras ejecutes la herramienta al menos una vez cada 90 días, nunca tendrás que volver a autenticarte.

```bash
# Cerrar sesión (elimina la caché de MSAL del disco)
m365-users logout
```

---

## Comandos disponibles

### `login` / `logout`

```bash
m365-users login    # Iniciar sesión (Device Code Flow)
m365-users logout   # Cerrar sesión y borrar el token guardado
```

---

### `import <archivo.xlsx>`

Crea o actualiza usuarios en masa a partir del Excel de trabajadores (hoja `ACTIVOS`).

```bash
m365-users import usuarios.xlsx

# Limitar el número de usuarios a procesar (útil para pruebas)
m365-users import usuarios.xlsx --limit 5
```

- `--limit <number>`: Limita la cantidad de filas válidas que se procesarán en esta ejecución.

> `import` no pide confirmación: escribe en el tenant según va leyendo las filas. Usa `--limit` para una primera pasada.

#### Formato del archivo Excel

El archivo debe ser un `.xlsx` con una hoja llamada `ACTIVOS`; la primera fila son las cabeceras. Se reconocen las cabeceras del listado de trabajadores (sin distinguir mayúsculas de minúsculas):

| Cabecera en el Excel | Campo en Graph API | Notas |
|---|---|---|
| `e_mail` | `userPrincipalName` | **Obligatorio.** Si no lleva `@`, se añade el `domain` de `config.json`. Si `domain` está configurado, debe terminar en ese dominio. |
| `Trabajador` | `displayName`, `givenName`, `surname` | **Obligatorio.** `APELLIDOS, NOMBRE` se convierte en `Nombre Apellidos`. |
| `id_empleado` | `employeeId` | |
| `Departamento` | `department` | |
| `Descripción` | `jobTitle` | |
| `Fecha de Alta` | `employeeHireDate` | `dd/mm/yyyy` (o celda de fecha). También se acepta ISO 8601. Una fecha imposible se ignora. |
| `ID responsable` | manager | Si es numérico, es el `id_empleado` del responsable y se busca entre los usuarios activos de M365. Si no, se trata como nombre de usuario y se le añade el dominio. |

`mailNickname` se deduce de la parte local del correo si no hay una columna con ese nombre.

También se aceptan directamente como cabecera los nombres de campo de Graph API. [`sample.csv`](./sample.csv) muestra esas columnas (es solo una referencia: `import` necesita un `.xlsx` real):

| Columna | Campo en Graph API | Notas |
|---|---|---|
| `givenName` / `surname` | `givenName` / `surname` | |
| `accountEnabled` | `accountEnabled` | `true` o `false`. Por defecto `true` |
| `companyName`, `employeeType`, `officeLocation` | ídem | |
| `mobilePhone` / `businessPhone` | `mobilePhone` / `businessPhones[0]` | |
| `streetAddress`, `city`, `state`, `postalCode`, `country` | Dirección | |
| `usageLocation` | `usageLocation` | Código ISO 3166-1 (ej. `ES`) |
| `preferredLanguage` | `preferredLanguage` | ej. `es-ES` |
| `password` | `passwordProfile.password` | Opcional, solo al crear |
| `forceChangePasswordNextSignIn` | `passwordProfile` | Por defecto `true` |

#### Qué hace con cada fila

| Salida | Significado |
|---|---|
| `[CREATE]` | El UPN no existe: se crea la cuenta. Si no hay columna `password`, se genera una contraseña temporal (`Palabra-NNNN-Palabra`) con cambio obligatorio en el primer inicio de sesión. |
| `[UPDATE]` | La cuenta existe y algún campo difiere: se actualiza y se listan los cambios (`Change campo: antes -> después`). Nunca se toca la contraseña ni el `mailNickname`. |
| `[IS-READY]` | La cuenta existe y ya coincide con el Excel: no se escribe nada. |
| `[FAILED]` | Graph rechazó la operación; se continúa con la siguiente fila. |

Después, para cada fila: se asigna el responsable solo si es distinto del actual, y se avisa si la cuenta no tiene ninguna licencia.

Las filas sin `e_mail` o `Trabajador`, o con un dominio distinto, se listan al principio como errores de validación y se omiten. Al final se imprime un resumen (creados / actualizados / fallidos / errores de responsable).

#### Limitaciones conocidas

- La contraseña generada para una cuenta nueva **no se muestra**. Ejecuta `reset-password <usuario>` para obtener unas credenciales que entregar.
- El responsable por `id_empleado` solo se resuelve si su cuenta está activa en M365.

---

### `search <consulta>`

Busca usuarios por nombre o correo electrónico.

```bash
m365-users search "Juan García"
m365-users search jgarcia@contoso.com
m365-users search "Juan García" --edit   # Abre el editor directamente tras seleccionar
```

Si hay varios resultados, muestra un selector interactivo. Con `--edit` se abre el editor de usuario inmediatamente después de seleccionar.

---

### `list`

Lista todos los usuarios activos del directorio en una tabla con las siguientes columnas: **ID empleado**, **UPN**, **Nombre**, **Puesto**, **Departamento**, **Manager**, **Grupos**, **Licenses** y **Fechas**.

Los grupos automáticos (`Todo Ayuntamiento`, `Todos los usuarios`, `Expertos 365`) se filtran siempre de la columna Grupos.

```bash
m365-users list

# Solo usuarios deshabilitados
m365-users list --disabled

# Filtrar usuarios con datos incompletos
m365-users list --missing-manager
m365-users list --missing-department
m365-users list --missing-job-title

# Filtrar usuarios que nunca han iniciado sesión
m365-users list --never-signed-in

# Generar lista de correos de usuarios con licencia (separados por ;)
m365-users list --for-mailing

# Combinar filtro con edición interactiva
m365-users list --missing-manager --edit
```

- `--for-mailing`: En lugar de mostrar la tabla, lista los UPNs de todos los usuarios que tienen al menos una licencia asignada, separados por `;`.
- `--never-signed-in`: Muestra únicamente los usuarios que nunca han iniciado sesión en el tenant.
- `--edit`: Abre el editor interactivo para los usuarios filtrados.

---

### `edit <upn-o-id>`

Editor interactivo para un usuario concreto. Muestra los datos actuales, agrupa los campos por sección y aplica todos los cambios en una sola llamada a la API tras confirmación.

```bash
m365-users edit jdoe@arandadeduero.es
m365-users edit <objectId>
```

**Secciones disponibles:**

| Sección | Qué permite editar |
|---|---|
| Identity | displayName, givenName, surname, mailNickname |
| Job Information | jobTitle, department, companyName, employeeId, employeeType, hireDate |
| Contact | mobilePhone, businessPhone, officeLocation |
| Address | streetAddress, city, state, postalCode, country |
| Settings | usageLocation, preferredLanguage, accountEnabled |
| Manager (Org Chart) | Manager del usuario (por UPN) |
| Password | Nueva contraseña |
| **Groups** | Membresías de grupos — muestra todos los grupos del tenant con los actuales pre-seleccionados; espacio para activar/desactivar, Enter para confirmar. Aplica los cambios necesarios (añade y elimina) al guardar. |

Los cambios pendientes se muestran en pantalla antes de confirmar, incluyendo los grupos que se van a añadir (en verde `+`) y los que se van a eliminar (en rojo `-`).

---

### `whoami <upn-o-id>`

Muestra la ficha completa de un usuario (nombre, título, departamento, teléfonos, dirección, estado, manager).

```bash
m365-users whoami jdoe@contoso.com
```

---

### `assign-manager`

Asigna un manager en masa a todos los usuarios de un departamento o con un título de trabajo concreto.

```bash
m365-users assign-manager
```

El comando guía de forma interactiva: selecciona el filtro (departamento o título), elige el valor de una lista obtenida en tiempo real, introduce el UPN del nuevo manager y confirma antes de aplicar.

---

### `fix job-titles`

Normaliza los títulos de trabajo a formato Oración (primera letra en mayúscula, resto en minúscula). Muestra una previsualización de los cambios, permite seleccionar cuáles aplicar y requiere doble confirmación antes de modificar nada.

```bash
m365-users fix job-titles
```

---

### `fix org-chart`

Analiza la integridad del organigrama completo: detecta usuarios sin manager, cadenas que no llegan al usuario raíz y ciclos. Ofrece un bucle interactivo para corregir los problemas uno a uno.

```bash
m365-users fix org-chart
```

---

### `fix timezone`

Establece la zona horaria del buzón al valor **Romance Standard Time** (equivalente a `Europe/Madrid`, UTC+1 / UTC+2 en verano) para todos los usuarios activos que tengan un valor distinto.

Corrige **dos campos independientes** en un único PATCH por usuario:

| Campo en Graph API | Qué afecta |
|---|---|
| `mailboxSettings.timeZone` | Vista del calendario en Outlook y OWA |
| `mailboxSettings.workingHours.timeZone` | Disponibilidad en Teams, franjas de horario laboral, vista libre/ocupado |

Ambos campos deben estar sincronizados. Un tenant configurado con valores por defecto de Microsoft puede tener `workingHours.timeZone` apuntando a `UTC-7` (Mountain Standard Time, zona horaria de los centros de datos de EE. UU.) incluso aunque `timeZone` ya esté correcto, lo que hace que Teams y el calendario de Outlook muestren horas incorrectas.

El horario laboral de cada usuario (días de la semana, hora de inicio y fin) se **preserva exactamente** tal y como está; solo se cambia la zona horaria dentro de `workingHours`.

```bash
m365-users fix timezone

# Previsualizar cambios sin aplicar nada
m365-users fix timezone --dry-run
```

**Requisito previo:** el registro de aplicación en Azure AD debe tener el permiso delegado `MailboxSettings.ReadWrite` concedido por un administrador del tenant. Además, la **cuenta con la que se ejecuta `m365-users login`** debe tener el rol **Exchange Administrator** (o Global Administrator) asignado en Azure AD. Sin ese rol, Exchange devuelve `403 ErrorAccessDenied` al intentar acceder al buzón de otros usuarios.

> Para asignar el rol: **Azure Portal → Azure Active Directory → Roles y administradores → Exchange Administrator → + Agregar asignaciones → selecciona tu cuenta de administrador**.
> Tras asignar el rol ejecuta `m365-users logout && m365-users login` para obtener un nuevo token.

**Flujo:**

1. Obtiene todos los usuarios activos del directorio.
2. Lee `mailboxSettings.timeZone` y `mailboxSettings.workingHours.timeZone` de cada usuario (llamadas secuenciales para respetar los límites de Exchange Online).
3. Muestra un resumen: cuántos ya tienen ambos valores correctos, cuántos necesitan actualización y cuántos no tienen buzón de Exchange.
4. Muestra una tabla de previsualización con los usuarios que van a cambiar, indicando para cada uno el valor actual de ambas zonas horarias (resaltado en rojo si está mal) frente al valor objetivo.
5. Solicita **doble confirmación** antes de aplicar ningún cambio.
6. Aplica las actualizaciones de forma secuencial e imprime el resultado línea a línea.
7. Muestra un resumen final: actualizados / ya correctos / sin buzón / fallidos.

Los usuarios sin buzón de Exchange (sin licencia de Outlook/M365) se detectan automáticamente y se omiten sin error.

> **Nota:** los marcadores de hora en los mensajes de chat de Teams usan la **zona horaria del sistema operativo** del dispositivo del usuario, no la configuración del buzón. Si los timestamps del chat siguen siendo incorrectos tras ejecutar este comando, el usuario debe corregir la zona horaria de su PC o móvil.

**Opción `--dry-run`:** muestra la previsualización completa pero no realiza ninguna llamada de escritura a la API.

**Opción `--debug`:** imprime por consola la respuesta cruda de la API `mailboxSettings` para cada usuario (útil para diagnosticar problemas de permisos o respuestas inesperadas).

---

### `validate <archivo.xlsx>`

Comprueba la calidad de los datos según un conjunto de reglas. No requiere conexión a Microsoft 365. Sale con código `1` si se encuentran problemas, `0` si todo es correcto.

```bash
m365-users validate usuarios.xlsx
```

**Reglas aplicadas:**

| ID de regla | Columna | Qué comprueba |
|---|---|---|
| `id_empleado-required-unique` | `id_empleado` | Obligatorio y único en todas las filas |
| `email-required-domain` | `e_mail` | Obligatorio y debe terminar en `@arandadeduero.es` |
| `email-unique` | `e_mail` | Debe ser único — no puede repetirse entre trabajadores |
| `id_responsable-exists` | `ID responsable` | Si está relleno, debe coincidir con un `id_empleado` existente (puede estar vacío para el director/a) |
| `fecha-de-baja-format` | `Fecha de Baja` | Si está relleno, debe tener formato `dd/mm/yyyy` |
| `fecha-de-baja-not-future` | `Fecha de Baja` | Si está relleno y es una fecha válida, no puede ser futura |

El informe muestra por cada problema: número de fila, nombre del trabajador, columna afectada, valor incorrecto y descripción del error.

---

### `validate-users`

Lista los usuarios **activos** de Microsoft 365 junto con su manager y los grupos a los que pertenecen (las cuentas deshabilitadas no se consultan, así que el contador de desactivados es siempre 0). Útil para auditar membresías de grupos antes de aplicar correcciones.

Los grupos automáticos (`Todo Ayuntamiento`, `Todos los usuarios`, `Expertos 365`) se filtran siempre del listado.

```bash
m365-users validate-users
```

Por cada usuario muestra:
- **Nombre** (`displayName`)
- **UPN** (correo de inicio de sesión)
- **Estado** — `activo` (verde) o `desactivado` (rojo)
- **Manager** — nombre del manager directo
- **Grupos** — lista de grupos, ordenados alfabéticamente. `(sin grupos)` si no tiene ninguno

Al final imprime un resumen: total / activos / desactivados / sin grupos.

Las membresías se obtienen usando peticiones `$batch` (20 usuarios por lote).

#### Opción `--empty-groups`

Muestra solo los usuarios que no pertenecen a ningún grupo (excluidos los automáticos).

```bash
m365-users validate-users --empty-groups
```

#### Opción `--fix-groups-by-manager`

Flujo interactivo para añadir en masa a un grupo a todos los subordinados de un manager concreto.

```bash
m365-users validate-users --fix-groups-by-manager
```

1. Busca y selecciona el manager
2. Muestra la tabla de todos sus subordinados con sus grupos actuales
3. Busca y selecciona el grupo destino
4. Muestra la previsualización (omite los que ya son miembros)
5. Doble confirmación
6. Añade cada usuario al grupo, mostrando el resultado por línea

---

### `sync-check <archivo.xlsx>`

Cruza el archivo Excel con los usuarios reales de Microsoft 365 y detecta desincronizaciones. Solo lectura por defecto.

```bash
m365-users sync-check usuarios.xlsx
```

**Check 1a — Cuentas que deben existir en la nube:**
Trabajadores con email `@arandadeduero.es` y sin `Fecha de Baja` deben tener una cuenta activa en M365. Se marcan los que no se encuentran.

**Check 1b — Sin email de dominio:**
Trabajadores activos (sin `Fecha de Baja`) cuyo email no es `@arandadeduero.es` o está vacío. No se pueden verificar por UPN: es un problema de datos a corregir.

**Check 2 — Cuentas que deben estar deshabilitadas:**
Trabajadores con una `Fecha de Baja` pasada o de hoy, y con email `@arandadeduero.es`, deben tener su cuenta deshabilitada en M365. Se marcan las cuentas que siguen activas.

> Las fechas de baja futuras no cuentan como salida efectiva y no se incluyen en el Check 2.

**Check 3 — Cuentas en la nube que no aparecen en el Excel:**
Compara todas las cuentas **activas** `@arandadeduero.es` de M365 contra los emails encontrados en **ambas** hojas del Excel (`ACTIVOS` y, si existe, `CONCEJALES`). Las cuentas de sala (UPN que empieza por `sala-` o `salapz`) se excluyen siempre de este check.

#### Opción `--disable-left-workers`

Deshabilita y retira las licencias de las cuentas identificadas en el **Check 2** que siguen activas. Muestra una previsualización completa (incluyendo las licencias asignadas y un resumen del total que se va a retirar) y solicita **doble confirmación** antes de realizar ningún cambio.

```bash
m365-users sync-check usuarios.xlsx --disable-left-workers
```

El proceso por cada cuenta:
1. Retira todas las licencias asignadas (la cuenta queda sin licencia)
2. Deshabilita la cuenta (`accountEnabled = false`)

Las licencias se retiran antes de deshabilitar; si la retirada falla, esa cuenta **no** se deshabilita y se informa del error. Los trabajadores de baja que ya no tienen cuenta activa en la nube no aparecen (no hay nada que hacer).

Si se rechaza cualquiera de las dos confirmaciones no se modifica nada.

> `sync-check` solo consulta las cuentas **activas** de M365. Por eso, en el Check 1a un trabajador activo cuya cuenta existe pero está deshabilitada aparece como "no encontrado".

**Códigos de salida:** `0` si todo está sincronizado (o si todas las correcciones se aplicaron); `1` si hay incidencias, se cancela la operación o falla alguna corrección.

---

### `reset-password <username>`

Restablece la contraseña de un usuario y genera un enlace `mailto:` listo para abrir en el cliente de correo con las credenciales incluidas en el cuerpo.

```bash
m365-users reset-password jgarcia
m365-users reset-password jgarcia@arandadeduero.es
```

El nombre de usuario acepta el formato corto (sin dominio): si se omite `@arandadeduero.es` se completa automáticamente.

**Flujo:**

1. Busca el usuario en Azure AD y muestra su ficha (nombre, UPN, puesto, departamento).
2. Genera una contraseña legible con el formato `Palabra-NNNN-Palabra` (ej. `Monte-3847-Roble`).
3. Muestra la contraseña generada **antes** de pedir confirmación.
4. Solicita confirmación única (por defecto: No).
5. Aplica el cambio vía Graph API con `forceChangePasswordNextSignIn: true` (el usuario deberá cambiarla en el primer inicio de sesión).
6. Imprime la nueva contraseña y un enlace `mailto:` con los datos de acceso.

**Enlace `mailto:` generado:**

```
Para:    usuario@arandadeduero.es
Asunto:  Microsoft365
Cuerpo:
  username: usuario@arandadeduero.es
  password: Monte-3847-Roble
  url: https://aytoarandaduero.sharepoint.com/
```

---

## Opción global

Todos los comandos aceptan `-c` para usar un archivo de configuración alternativo:

```bash
m365-users -c /ruta/a/otro-config.json list
```

---

## Estructura del proyecto

```
m365-users/
├── CLAUDE.md               # Guía para trabajar en el código con Claude Code
├── sample.csv              # Referencia de columnas para importación (no importable directamente)
├── config.example.json     # Plantilla de configuración (sin credenciales reales)
├── config.json             # Tu configuración real (gitignoreado)
├── package.json
├── src/
│   ├── index.js            # Punto de entrada CLI (definición de comandos)
│   ├── auth.js             # Autenticación (Device Code Flow + caché MSAL cifrada)
│   ├── constants.js        # Constantes compartidas (dominio requerido)
│   ├── graph.js            # Único módulo que llama a Microsoft Graph (con caché en memoria)
│   ├── commands/
│   │   ├── import.js           # Importación masiva desde Excel
│   │   ├── search.js           # Búsqueda de usuarios y ficha (whoami)
│   │   ├── edit.js             # Editor interactivo
│   │   ├── list.js             # Listado con filtros (incluye grupos y licencias)
│   │   ├── assign-manager.js   # Asignación masiva de manager
│   │   ├── fix.js              # Correcciones de calidad de datos (job-titles, org-chart)
│   │   ├── fix-timezone.js     # Corrección masiva de zona horaria de buzón
│   │   ├── validate.js         # Validación del Excel de trabajadores
│   │   ├── validate-users.js   # Listado de usuarios con sus grupos
│   │   ├── sync-check.js       # Comprobación de sincronización con M365
│   │   └── reset-password.js   # Restablecimiento de contraseña con enlace mailto:
│   ├── validators/
│   │   └── excel-rules.js      # Reglas de validación del Excel
│   └── utils/
│       ├── ansi.js             # Eliminación de códigos de escape ANSI
│       ├── cache.js            # Caché en memoria para datos de Graph API
│       ├── excel-importer.js   # Mapeo de filas del Excel a usuarios de Graph
│       ├── excel.js            # Lectura de archivos .xlsx (hojas ACTIVOS y CONCEJALES)
│       ├── token-store.js      # Almacenamiento cifrado genérico (no usado por la CLI)
│       └── upn.js              # Utilidades para UPN
└── test/
    ├── *.test.js           # Tests (ava)
    └── helpers/            # Tenant simulado, prompts simulados, HOME temporal, generador de Excel
```

---

## Desarrollo y tests

```bash
npm test                # Ejecuta todos los tests
npm run test:coverage   # Tests + informe de cobertura (texto y lcov en coverage/)
npx ava test/import.test.js   # Un solo archivo
```

Los tests **no necesitan `config.json`, ni sesión iniciada, ni conexión**: ejecutan los comandos completos contra un tenant simulado (se sustituye `fetch`) y contra archivos Excel generados al vuelo. Tampoco tocan `~/.config/m365-users`: cada archivo de test se ejecuta con un `HOME` temporal.

Para los comandos que modifican el tenant, los tests comprueban exactamente qué peticiones de escritura se envían, y que no se envía ninguna cuando se rechaza una confirmación o se usa `--dry-run`.

La integración continua (GitHub Actions, `.github/workflows/ci.yml`) ejecuta `npm test` con Node 24 en cada push a `main` y en cada pull request.

---

## Seguridad

- `config.json` está en `.gitignore` y **nunca debe subirse al repositorio**.
- La caché de sesión MSAL (incluyendo el *refresh token*) se almacena cifrada (AES-256-GCM) en `~/.config/m365-users/msal-cache.json` con permisos `600` (solo lectura del propietario) y una clave derivada de la máquina, por lo que no es portable entre equipos.
- La autenticación es **delegada** (el usuario que hace login debe tener los permisos necesarios en el tenant). No se usan secretos de aplicación en el flujo de autenticación.
- Los permisos `User.ReadWrite.All` y `Directory.ReadWrite.All` otorgan acceso de lectura y escritura completo sobre todos los usuarios del directorio. Úsalos con precaución y solo en cuentas de administrador dedicadas.

---

## Licencia

MIT
