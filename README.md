# m365-users

Herramienta de línea de comandos (CLI) para gestionar los usuarios de Microsoft 365 del Ayuntamiento de Aranda de Duero a través de la API de Microsoft Graph. Permite dar de alta y actualizar usuarios en masa desde el Excel de trabajadores, comprobar que el Excel y la nube están sincronizados, buscar y editar usuarios de forma interactiva, restablecer contraseñas y corregir problemas de calidad de datos en el directorio.

> **Atención:** la herramienta trabaja contra el tenant **de producción**. No hay entorno de pruebas. Los comandos que escriben en masa muestran una previsualización y piden confirmación; `import` es la excepción (ver su sección).

## Resumen de comandos

| Comando | Para qué sirve | ¿Modifica M365? |
|---|---|---|
| [`login` / `logout`](#login--logout) | Iniciar o cerrar sesión | No |
| [`validate <excel>`](#validate-archivoxlsx) | Revisar la calidad de datos del Excel (sin conexión) | No |
| [`sync-check <excel>`](#sync-check-archivoxlsx) | Comparar el Excel con las cuentas de M365 | Solo con `--disable-left-workers` |
| [`import <excel>`](#import-archivoxlsx) | Crear y actualizar usuarios desde el Excel | **Sí, sin confirmación** |
| [`reset-password <usuario>`](#reset-password-username) | Nueva contraseña temporal + enlace `mailto:` | Sí, con confirmación |
| [`list`](#list) | Tabla de usuarios, con filtros | Solo con `--edit` |
| [`search <consulta>`](#search-consulta) | Buscar por nombre o correo | Solo si se elige editar |
| [`whoami <usuario>`](#whoami-upn-o-id) | Ficha de un usuario | No |
| [`edit <usuario>`](#edit-upn-o-id) | Editor interactivo de un usuario | Sí, con confirmación |
| [`assign-manager`](#assign-manager) | Asignar responsable a un departamento o puesto | Sí, con confirmación |
| [`validate-users`](#validate-users) | Usuarios con su responsable y sus grupos | Solo con `--fix-groups-by-manager` |
| [`fix job-titles`](#fix-job-titles) | Normalizar mayúsculas de los puestos | Sí, con doble confirmación |
| [`fix org-chart`](#fix-org-chart) | Revisar y reparar el organigrama | Sí, usuario a usuario |
| [`fix timezone`](#fix-timezone) | Corregir la zona horaria de los buzones | Sí, con doble confirmación (`--dry-run` no escribe) |

### Flujo habitual con el Excel de trabajadores

```bash
m365-users validate trabajadores.xlsx      # 1. ¿El Excel está bien? (sin conexión)
m365-users sync-check trabajadores.xlsx    # 2. ¿Qué diferencias hay con la nube?
m365-users import trabajadores.xlsx        # 3. Crear las cuentas nuevas y actualizar datos
m365-users reset-password cromero          # 4. Credenciales para cada cuenta nueva
m365-users sync-check trabajadores.xlsx --disable-left-workers   # 5. Deshabilitar las bajas
```

> Los ejemplos de salida de este documento usan **personas y datos ficticios**. Las tablas anchas están recortadas (`…`) para que quepan.

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

Ejemplo de `login`:

```text
No valid session found. Starting authentication...

========================================
  Microsoft 365 Authentication
========================================

  1. Open your browser and go to (code pre-filled):
     https://microsoft.com/devicelogin?usercode=ABCD1234

  2. Enter the code (if not already filled):
     ABCD1234

  3. Sign in with your Microsoft 365 account

========================================

Authentication successful. Session saved.
```

`logout` responde `Logged out. Session token cleared.`

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

#### Ejemplo

Con un Excel de siete filas en el que dos puestos están escritos distinto que en la nube, una trabajadora no tiene responsable asignado, hay un trabajador nuevo y una fila con correo de otro dominio:

```text
Parsing Excel: /ruta/trabajadores.xlsx

✔ Loaded 8 users.
Found 1 row(s) with validation errors:

  Row 8 (elena.vidal@gmail.com): UPN must end with @arandadeduero.es

Processing 6 user(s)...

  [IS-READY] mgarcia@arandadeduero.es — María García López
  [UPDATE] jperez@arandadeduero.es — Juan Pérez Ruiz
           Change jobTitle: ADMINISTRATIVO -> Administrativo
  [IS-READY] lmartin@arandadeduero.es — Laura Martín Gil
  [UPDATE] asanz@arandadeduero.es — Ana Sanz Vega
           Change jobTitle: auxiliar ADMINISTRATIVO -> Auxiliar administrativo
           Manager: Cloud=(none) | Excel=102
           Manager set: 102 (resolved to 5f0c…)
  [IS-READY] pdiaz@arandadeduero.es — Pedro Díaz Mora
  [CREATE] cromero@arandadeduero.es — Carlos Romero Cano
           Manager: Cloud=(none) | Excel=102
           Manager set: 102 (resolved to 5f0c…)
           Warning: User cromero@arandadeduero.es has no active licenses.

----------------------------------------
Import Summary:
  Created: 1
  Updated: 2
  Failed:  0
----------------------------------------
```

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

Ejemplo con varios resultados:

```text
$ m365-users search martín
✔ Found 2 user(s).
Found 2 user(s):

? Select a user:
❯ Laura Martín Gil <lmartin@arandadeduero.es> — Técnica de cultura (Cultura)
  Martín Soto Rey <msoto@arandadeduero.es> — Policía local (Seguridad) [DISABLED]
  (cancel)
```

Al elegir uno se muestra su ficha (la misma que [`whoami`](#whoami-upn-o-id)) y se pregunta qué hacer:

```text
? What would you like to do?
❯ Edit this user
  Nothing (exit)
```

Si solo hay un resultado se muestra su ficha directamente. Si no hay ninguno: `No users found matching that query.`

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

Ejemplo de `m365-users list --missing-manager` (faltan a la derecha las columnas Groups, Licenses, Last Sign-In y Last Pass Change):

```text
✔ Fetched 8 users.
✔ Fetched group memberships.
✔ Fetched licenses.
Found 2 user(s) out of 8 active total [filters: no manager]

──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
Emp ID     UPN / Email                 Display Name          Job Title                Department    Manager    …
──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
100        mgarcia@arandadeduero.es    María García López    Secretaria general       Secretaría    (none)     …
103        asanz@arandadeduero.es      Ana Sanz Vega         auxiliar ADMINISTRATIVO  Cultura       (none)     …
──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
```

Ejemplo de `m365-users list --for-mailing`:

```text
✔ Fetched 8 users.
✔ Licensed users:
mgarcia@arandadeduero.es;jperez@arandadeduero.es;lmartin@arandadeduero.es;asanz@arandadeduero.es;pdiaz@arandadeduero.es
```

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

Ejemplo de sesión (tras la ficha del usuario aparece el menú de secciones; se vuelve a él después de cada sección hasta elegir guardar o cancelar):

```text
Editing user. Leave fields blank to keep current values.

? Select section to edit:
❯ Identity
  Job Information
  Contact
  Address
  Settings
  Manager (Org Chart)
  Password
  Groups
  Save all changes
  Cancel (discard)

  … (se edita "Job Information" y se elige "Save all changes") …

? Apply 1 change(s) to Juan Pérez Ruiz? Yes

User properties updated successfully.

Done.
```

Con `Cancel (discard)` no se envía nada: `Edit cancelled. No changes made.`

---

### `whoami <upn-o-id>`

Muestra la ficha completa de un usuario (nombre, título, departamento, teléfonos, dirección, estado, manager).

```bash
m365-users whoami jdoe@contoso.com
```

El nombre corto también vale (`m365-users whoami jperez`): se completa con el `domain` de `config.json`.

```text
────────────────────────────────────────────────────
  Juan Pérez Ruiz
  Title:     Administrativo
  Dept:      Secretaría
  UPN:       jperez@arandadeduero.es
  Usage loc: ES
  Enabled:   Yes
  Manager:   María García López <mgarcia@arandadeduero.es>
  ID:        2b7e1c44-…
────────────────────────────────────────────────────
```

Solo se imprimen los campos que tienen valor (móvil, teléfono, oficina y localidad aparecen si están rellenos).

---

### `assign-manager`

Asigna un manager en masa a todos los usuarios de un departamento o con un título de trabajo concreto.

```bash
m365-users assign-manager
```

El comando guía de forma interactiva: selecciona el filtro (departamento o título), elige el valor de una lista obtenida en tiempo real, introduce el UPN del nuevo manager y confirma antes de aplicar.

Ejemplo (eligiendo el departamento `Cultura`):

```text
? Assign manager by: Department
✔ Found 3 distinct department(s).
? Select department: Cultura
✔ Fetched 2 users.

Found 2 user(s) for department "Cultura":

  Laura Martín Gil — lmartin@arandadeduero.es  [Técnica de cultura — Cultura]
  Ana Sanz Vega — asanz@arandadeduero.es  [Auxiliar administrativo — Cultura]

? Enter the manager UPN (email) [@arandadeduero.es added automatically if missing]: mgarcia

Looking up manager "mgarcia"...
  Manager found: María García López <mgarcia@arandadeduero.es>

? Assign "María García López" as manager to all 2 user(s) with department "Cultura"? Yes
✔ Manager assignment complete.

─────────────────────────────────────────
Summary:
  Updated: 2
  Failed:  0
─────────────────────────────────────────
```

---

### `fix job-titles`

Normaliza los títulos de trabajo a formato Oración (primera letra en mayúscula, resto en minúscula). Muestra una previsualización de los cambios, permite seleccionar cuáles aplicar y requiere doble confirmación antes de modificar nada.

```bash
m365-users fix job-titles
```

```text
✔ Found 2 candidates.
? Select users to apply the fix to (Space to toggle, A to select all):
 ◉ Juan Pérez Ruiz <jperez@arandadeduero.es>  ADMINISTRATIVO → Administrativo
 ◉ Ana Sanz Vega <asanz@arandadeduero.es>  auxiliar ADMINISTRATIVO → Auxiliar administrativo

You are about to update 2 user(s):
  Juan Pérez Ruiz — ADMINISTRATIVO → Administrativo
  Ana Sanz Vega — auxiliar ADMINISTRATIVO → Auxiliar administrativo
? Apply Sentence case fix to 2 user(s)? Yes
? Are you sure? This will overwrite job titles in Microsoft 365. Yes
✔ Changes applied.

─────────────────────────────────────────
Summary:
  Updated: 2
  Failed:  0
─────────────────────────────────────────
```

Si no hay nada que corregir: `All job titles are already in Sentence case. Nothing to do.`

---

### `fix org-chart`

Analiza la integridad del organigrama completo: detecta usuarios sin manager, cadenas que no llegan al usuario raíz y ciclos. Ofrece un bucle interactivo para corregir los problemas uno a uno.

```bash
m365-users fix org-chart
```

Primero pide el UPN de la persona que debe ser la raíz del organigrama y después lista los problemas:

```text
✔ Loaded 8 active users.
✔ Manager relationships loaded.
? Enter the UPN of the expected org root (e.g. alinaje@arandadeduero.es) [@arandadeduero.es added if missing]: mgarcia

  Root: María García López <mgarcia@arandadeduero.es>

──────────────────────────────────────────────────────────────────────────────────────
Display Name                   UPN                                          Issue
──────────────────────────────────────────────────────────────────────────────────────
Ana Sanz Vega                  asanz@arandadeduero.es                       No manager
Becario Uno                    becario1@arandadeduero.es                    No manager
──────────────────────────────────────────────────────────────────────────────────────

  No manager: 2  Wrong root: 0  Cycles: 0  Total affected: 2

? What do you want to do?
❯ Fix interactively — assign a manager to each affected user
  Exit without changes
```

En el modo interactivo se elige un usuario de la lista, se escribe el UPN de su responsable (se comprueba que existe) y se aplica en el momento: `Manager set: lmartin@arandadeduero.es`. Se repite hasta elegir `Done — exit`.

Si todo está bien: `✓ Org chart is valid. All users connect to the root with no cycles.`

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

Ejemplo con `--dry-run`:

```text
✔ Loaded 8 active users.
✔ Read complete.
─────────────────────────────────────────
Mailbox timezone audit:
  Already correct (Romance Standard Time): 1
  Need update:                    2
  No mailbox / no license:        5
─────────────────────────────────────────

─────────────────────────────────────────────────────────────────────────────────────────────────────────────
UPN / Email                    Display Name          Calendar TZ                Working Hours TZ
─────────────────────────────────────────────────────────────────────────────────────────────────────────────
jperez@arandadeduero.es        Juan Pérez Ruiz       Pacific Standard Time      Pacific Standard Time
lmartin@arandadeduero.es       Laura Martín Gil      Romance Standard Time      Mountain Standard Time
─────────────────────────────────────────────────────────────────────────────────────────────────────────────

[dry-run] 2 user(s) would be updated. No changes were made.
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

El informe muestra por cada problema: número de fila, nombre del trabajador, columna afectada, valor incorrecto y descripción del error.

Ejemplo:

```text
Reading: /ruta/trabajadores.xlsx
Found 4 data row(s) with 8 column(s).

● Rule: id_empleado-required-unique
  id_empleado must be filled in every row and must be unique

  ✖ Row    3  MARTÍN GIL, LAURA     id_empleado      "101"                      Duplicate id_empleado — first seen at row 2
  ✖ Row    4  SANZ VEGA, ANA        id_empleado      (empty)                    Field is empty — id_empleado is required

● Rule: email-required-domain
  e_mail must be filled and must end with @arandadeduero.es

  ✖ Row    3  MARTÍN GIL, LAURA     e_mail           "lmartin@gmail.com"        Email does not end with @arandadeduero.es

● Rule: email-unique
  e_mail must be unique — no two employees can share the same address

  ✖ Row    4  SANZ VEGA, ANA        e_mail           "jperez@arandadeduero.es"  Duplicate email — first seen at row 2

● Rule: id_responsable-exists
  ID Responsable must reference an existing id_empleado (can be empty for the top-level director)

  ✖ Row    3  MARTÍN GIL, LAURA     ID responsable   "999"                      ID responsable "999" does not match any id_empleado in the file

● Rule: fecha-de-baja-format
  Fecha de Baja, if filled, must be in dd/mm/yyyy format (2-digit day and month)

  ✖ Row    4  SANZ VEGA, ANA        Fecha de Baja    "2026-10-01"               Expected dd/mm/yyyy format (e.g. 31/12/2025), got "2026-10-01"

✖  6 issue(s) found across 2 row(s).
```

Solo se imprimen las reglas que tienen algún problema.

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

```text
✔ Found 8 user(s).
✔ Fetched group memberships.
Nombre                UPN                           Estado    Manager                     Grupos
──────────────────────────────────────────────────────────────────────────────────────────────────────────────
Ana Sanz Vega         asanz@arandadeduero.es        activo                                (sin grupos)
Juan Pérez Ruiz       jperez@arandadeduero.es       activo    mgarcia@arandadeduero.es    Secretaría
Laura Martín Gil      lmartin@arandadeduero.es      activo    mgarcia@arandadeduero.es    Biblioteca, Cultura
María García López    mgarcia@arandadeduero.es      activo                                Secretaría
Pedro Díaz Mora       pdiaz@arandadeduero.es        activo    lmartin@arandadeduero.es    Deportes
…
──────────────────────────────────────────────────────────────────────────────────────────────────────────────
Total: 8 usuario(s)  |  8 activo(s)  |  0 desactivado(s)  |  4 sin grupos
```

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
Trabajadores cuya `Fecha de Baja` es de ayer o anterior, y con email `@arandadeduero.es`, deben tener su cuenta deshabilitada en M365. Se marcan las cuentas que siguen activas.

> La baja es efectiva a partir del día siguiente a la `Fecha de Baja`. Si la fecha es de hoy o futura, el trabajador sigue contando como activo: no se incluye en el Check 2 ni se deshabilita, y `validate` no lo marca como error.

**Check 3 — Cuentas en la nube que no aparecen en el Excel:**
Compara todas las cuentas **activas** `@arandadeduero.es` de M365 contra los emails encontrados en **ambas** hojas del Excel (`ACTIVOS` y, si existe, `CONCEJALES`). Las cuentas de sala (UPN que empieza por `sala-` o `salapz`) se excluyen siempre de este check.

**Cuentas especiales:**
Algunas cuentas existen en la nube sin fila en el Excel a propósito (buzones compartidos, consultoras, lectores…). Están en la lista `SPECIAL_ACCOUNTS` de [`src/constants.js`](./src/constants.js): `asac`, `info_practicas1`, `lector01agua`, `consultorahyk` y `lector02agua`. No cuentan como incidencia en el Check 3 ni afectan al código de salida, pero si se encuentran activas en M365 se muestran siempre en un aviso (`⚠ Special accounts present in M365`) para recordar que hay que revisarlas. Si una de ellas no existe en la nube no se avisa. Para añadir o quitar cuentas, edita esa lista.

#### Ejemplo

```text
Reading: /ruta/trabajadores.xlsx
Excel: 5 active employee(s) to check, 1 left employee(s) to check, 1 active employee(s) with no @arandadeduero.es email, 0 concejales found.

✔ Cloud: 8 user(s) found.
⚠ Special accounts present in M365 (not in the Excel file by design)
  These accounts are allowed to exist without an Excel row. Check that they are still needed.

  ⚠ ASAC                    asac@arandadeduero.es

● Check 1a: Active employees missing from M365
  These employees have no Fecha de Baja but no matching account was found in the cloud.

  ✖ ROMERO CANO, CARLOS     cromero@arandadeduero.es

● Check 1b: Active employees with no @arandadeduero.es email
  These employees have no Fecha de Baja but their email is missing or uses a
  different domain, so their cloud account cannot be verified by UPN.

  ✖ VIDAL POLO, ELENA       elena.vidal@gmail.com

● Check 2: Left employees whose cloud account is not disabled
  These employees have a Fecha de Baja in the file,
  but their cloud account is still active (or not found).

  ✖ DÍAZ MORA, PEDRO        pdiaz@arandadeduero.es       Fecha de Baja: 01/10/2026       [account is ENABLED]

  Run with --disable-left-workers to disable and unlicense these accounts.

● Check 3: Cloud accounts not present in the Excel file
  These users exist in M365 with an @arandadeduero.es email but do not appear
  in the employee Excel file (neither as active nor as left).

  ✖ Becario Uno             becario1@arandadeduero.es    [enabled]

✖  4 sync issue(s) found.
```

En este ejemplo, una trabajadora con `Fecha de Baja` dentro de 20 días no aparece en ningún check (sigue activa), y la cuenta `sala-plenos` no aparece en el Check 3 (es una sala).

#### Opción `--disable-left-workers`

Deshabilita y retira las licencias de las cuentas identificadas en el **Check 2** que siguen activas. Muestra una previsualización completa (incluyendo las licencias asignadas y un resumen del total que se va a retirar) y solicita **doble confirmación** antes de realizar ningún cambio.

```bash
m365-users sync-check usuarios.xlsx --disable-left-workers
```

El proceso por cada cuenta:
1. Retira todas las licencias asignadas (la cuenta queda sin licencia)
2. Deshabilita la cuenta (`accountEnabled = false`)

Las licencias se retiran antes de deshabilitar; si la retirada falla, esa cuenta **no** se deshabilita y se informa del error. Los trabajadores de baja que ya no tienen cuenta activa en la nube no aparecen (no hay nada que hacer).

Ejemplo (tras el informe de los checks):

```text
  The following accounts will be DISABLED and UNLICENSED:

  Fetching current licence assignments…

  ► DÍAZ MORA, PEDRO        pdiaz@arandadeduero.es       Fecha de Baja: 01/10/2026    Licences: O365_BUSINESS_ESSENTIALS

  Licence removal summary:
    O365_BUSINESS_ESSENTIALS                  1 assignment(s) will be removed

    Total: 1 licence assignment(s) across 1 account(s).

  1 account(s) will be disabled and unlicensed.
? Are you sure you want to disable and unlicense these accounts? Yes
? FINAL confirmation: disable and remove licences for 1 account(s)? Yes

  pdiaz@arandadeduero.es                         removed 1 licence(s)  disabled  ✔

✔  1 account(s) disabled and unlicensed successfully.
```

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

**Ejemplo:**

```text
────────────────────────────────────────────────────
  Juan Pérez Ruiz
  UPN:        jperez@arandadeduero.es
  Title:      Administrativo
  Department: Secretaría
────────────────────────────────────────────────────

  New password will be: Solar-1421-Ruido
  (user will be required to change it on first sign-in)

? Reset password for Juan Pérez Ruiz (jperez@arandadeduero.es)? Yes
✔ Password reset applied.

✓ Password reset for Juan Pérez Ruiz
────────────────────────────────────────────────────
  UPN:      jperez@arandadeduero.es
  Password: Solar-1421-Ruido
────────────────────────────────────────────────────

  mailto link (open in your email client):

  mailto:jperez%40arandadeduero.es?subject=Microsoft365&body=username%3A%20jperez%40arandadeduero.es%0Apassword%3A%20Solar-1421-Ruido%0A…
```

**Correo que abre el enlace `mailto:`:**

```text
Para:    jperez@arandadeduero.es
Asunto:  Microsoft365
Cuerpo:
  username: jperez@arandadeduero.es
  password: Solar-1421-Ruido
  Recuerda que la contraseña de Microsoft365 es diferente a la de tu actual correo
  corporativo y a la de tu ordenador. Esta contraseña es temporal y deberás cambiarla
  en tu primer inicio de sesión.

  url: https://aytoarandaduero.sharepoint.com/

  Sigue la guía de bienvenida en este enlace https://…
```

Si se responde `No` a la confirmación, la contraseña no se cambia.

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
├── LICENSE                 # Licencia MIT
├── CLAUDE.md               # Guía para trabajar en el código con Claude Code
├── .nvmrc                  # Versión de Node (la usa también el CI)
├── .github/                # CI (workflows/ci.yml) y Dependabot
├── sample.csv              # Referencia de columnas para importación (no importable directamente)
├── config.example.json     # Plantilla de configuración (sin credenciales reales)
├── config.json             # Tu configuración real (gitignoreado)
├── package.json
├── src/
│   ├── index.js            # Punto de entrada CLI (definición de comandos)
│   ├── auth.js             # Autenticación (Device Code Flow + caché MSAL cifrada)
│   ├── constants.js        # Constantes compartidas (dominio, cuentas especiales)
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

La integración continua (GitHub Actions, `.github/workflows/ci.yml`) ejecuta `npm test` en cada push a `main` y en cada pull request, con la versión de Node indicada en `.nvmrc`.

Dependabot (`.github/dependabot.yml`) revisa cada lunes las dependencias de npm y las acciones de GitHub: las actualizaciones *minor* y *patch* llegan agrupadas en un único PR por ecosistema, y cada *major* en un PR propio.

---

## Seguridad

- `config.json` está en `.gitignore` y **nunca debe subirse al repositorio**.
- La caché de sesión MSAL (incluyendo el *refresh token*) se almacena cifrada (AES-256-GCM) en `~/.config/m365-users/msal-cache.json` con permisos `600` (solo lectura del propietario) y una clave derivada de la máquina, por lo que no es portable entre equipos.
- La autenticación es **delegada** (el usuario que hace login debe tener los permisos necesarios en el tenant). No se usan secretos de aplicación en el flujo de autenticación.
- Los permisos `User.ReadWrite.All` y `Directory.ReadWrite.All` otorgan acceso de lectura y escritura completo sobre todos los usuarios del directorio. Úsalos con precaución y solo en cuentas de administrador dedicadas.

---

## Licencia

[MIT](./LICENSE) © Ayuntamiento de Aranda de Duero
