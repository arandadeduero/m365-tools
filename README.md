# m365-users

Herramienta de línea de comandos (CLI) para gestionar usuarios de Microsoft 365 a través de la API de Microsoft Graph. Permite importar usuarios en masa desde CSV, buscarlos, editarlos de forma interactiva, asignar managers, y corregir problemas de calidad de datos en el directorio.

---

## Requisitos previos

- **Node.js 24.0.0 o superior**
- Una suscripción de Microsoft 365 con permisos de administrador
- Un registro de aplicación en Azure Active Directory con los permisos delegados:
  - `User.ReadWrite.All`
  - `Directory.ReadWrite.All`

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
  "scopes": ["User.ReadWrite.All", "Directory.ReadWrite.All"]
}
```

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

### `import <archivo.csv>`

Importa o actualiza usuarios en masa desde un CSV.

```bash
m365-users import usuarios.csv
```

- Si el usuario **no existe**: lo crea.
- Si el usuario **ya existe**: lo actualiza (sin tocar la contraseña).
- Si la columna `manager` está presente: asigna el manager al final.
- La operación es **idempotente**: puedes ejecutarla varias veces de forma segura.

#### Formato del CSV

Consulta [`sample.csv`](./sample.csv) para ver un ejemplo completo.

**Columnas obligatorias:** `userPrincipalName`, `displayName`, `mailNickname`, `password`

**Columnas opcionales:**

| Columna CSV | Campo en Graph API | Notas |
|---|---|---|
| `givenName` / `surname` | `givenName` / `surname` | Nombre y apellidos |
| `accountEnabled` | `accountEnabled` | `true` o `false` |
| `jobTitle` | `jobTitle` | |
| `department` | `department` | |
| `companyName` | `companyName` | |
| `employeeId` | `employeeId` | |
| `employeeType` | `employeeType` | Employee, Contractor… |
| `employeeHireDate` | `employeeHireDate` | Formato ISO 8601 |
| `mobilePhone` / `businessPhone` | `mobilePhone` / `businessPhones[0]` | |
| `officeLocation` | `officeLocation` | |
| `streetAddress`, `city`, `state`, `postalCode`, `country` | Dirección | |
| `usageLocation` | `usageLocation` | Código ISO 3166-1 (ej. `ES`) |
| `preferredLanguage` | `preferredLanguage` | ej. `es-ES` |
| `forceChangePasswordNextSignIn` | `passwordProfile` | Por defecto `true` |
| `manager` | Llamada separada a `setManager` | UPN del manager |

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

Lista todos los usuarios activos del directorio en una tabla con las siguientes columnas: **ID empleado**, **UPN**, **Nombre**, **Puesto**, **Departamento**, **Manager** y **Grupos**.

Los grupos automáticos (`Todo Ayuntamiento`, `Todos los usuarios`, `Expertos 365`) se filtran siempre de la columna Grupos.

```bash
m365-users list

# Solo usuarios deshabilitados
m365-users list --disabled

# Filtrar usuarios con datos incompletos
m365-users list --missing-manager
m365-users list --missing-department
m365-users list --missing-job-title

# Combinar filtro con edición interactiva
m365-users list --missing-manager --edit
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

### `validate`

Lee el archivo Excel de la carpeta `real-csv/` y comprueba la calidad de los datos según un conjunto de reglas. No requiere conexión a Microsoft 365. Sale con código `1` si se encuentran problemas, `0` si todo es correcto.

```bash
m365-users validate
```

**Reglas aplicadas:**

| ID de regla | Columna | Qué comprueba |
|---|---|---|
| `id_empleado-required-unique` | `id_empleado` | Obligatorio y único en todas las filas |
| `email-required-domain` | `e_mail` | Obligatorio y debe terminar en `@arandadeduero.es` |
| `email-unique` | `e_mail` | Debe ser único — no puede repetirse entre trabajadores |
| `id_responsable-exists` | `ID Responsable` | Si está relleno, debe coincidir con un `id_empleado` existente (puede estar vacío para el director/a) |
| `fecha-de-baja-format` | `Fecha de Baja` | Si está relleno, debe tener formato `dd/mm/yyyy` |
| `fecha-de-baja-not-future` | `Fecha de Baja` | Si está relleno y es una fecha válida, no puede ser futura |

El informe muestra por cada problema: número de fila, nombre del trabajador, columna afectada, valor incorrecto y descripción del error.

---

### `validate-users`

Lista todos los usuarios de Microsoft 365 junto con su manager y los grupos a los que pertenecen. Útil para auditar membresías de grupos antes de aplicar correcciones.

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

### `sync-check`

Cruza el archivo Excel de `real-csv/` con los usuarios reales de Microsoft 365 y detecta desincronizaciones. Solo lectura por defecto.

```bash
m365-users sync-check
```

**Check 1a — Cuentas que deben existir en la nube:**
Trabajadores con email `@arandadeduero.es` y sin `Fecha de Baja` deben tener una cuenta activa en M365. Se marcan los que no se encuentran.

**Check 1b — Sin email de dominio:**
Trabajadores activos (sin `Fecha de Baja`) cuyo email no es `@arandadeduero.es` o está vacío. No se pueden verificar por UPN: es un problema de datos a corregir.

**Check 2 — Cuentas que deben estar deshabilitadas:**
Trabajadores con una `Fecha de Baja` pasada o de hoy, y con email `@arandadeduero.es`, deben tener su cuenta deshabilitada en M365. Se marcan las cuentas que siguen activas.

> Las fechas de baja futuras no cuentan como salida efectiva y no se incluyen en el Check 2.

#### Opción `--disable-left-workers`

Deshabilita y retira las licencias de las cuentas identificadas en el **Check 2** que siguen activas. Muestra una previsualización completa (incluyendo las licencias asignadas y un resumen del total que se va a retirar) y solicita **doble confirmación** antes de realizar ningún cambio.

```bash
m365-users sync-check --disable-left-workers
```

El proceso por cada cuenta:
1. Retira todas las licencias asignadas (la cuenta queda sin licencia)
2. Deshabilita la cuenta (`accountEnabled = false`)

Las cuentas que no se encuentran en la nube se muestran en la previsualización pero se omiten de la operación (no hay nada que deshabilitar).

---

### `sync-employee-ids`

Lee las columnas `id_empleado` y `Tipo de empleado` del archivo Excel en `real-csv/` y las escribe en los campos `employeeId` y `employeeType` de cada usuario en Microsoft 365 en una sola pasada. El Excel es la fuente de verdad: los valores siempre se sobreescriben.

La coincidencia entre el Excel y la nube se hace por email (`e_mail` → `userPrincipalName`).

```bash
m365-users sync-employee-ids
```

**Valores permitidos para `Tipo de empleado`:** `Funcionario`, `Laboral`.  
Las filas con un valor distinto se muestran como advertencia y no se sincronizan (las demás filas sí).

**Se omiten las filas donde:**
- `e_mail` no tiene el dominio `@arandadeduero.es` o está vacío
- No se encuentra la cuenta en la nube
- Ambos campos ya están sincronizados (sin cambios necesarios)

**Flujo:**
1. Muestra advertencias para filas con `Tipo de empleado` inválido
2. Muestra tabla de preview con los cambios pendientes (campo, valor Excel, valor actual en la nube)
3. Solicita confirmación única antes de aplicar cambios
4. Aplica un único PATCH por usuario con todos los campos a actualizar
5. Muestra un resumen final (correctos / fallidos / advertencias)

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
├── config.example.json     # Plantilla de configuración (sin credenciales reales)
├── config.json             # Tu configuración real (gitignoreado)
├── sample.csv              # CSV de ejemplo para importación
├── real-csv/               # Archivos Excel reales (gitignoreado — contiene datos privados)
├── package.json
└── src/
    ├── index.js            # Punto de entrada CLI
    ├── auth.js             # Autenticación (Device Code Flow + caché MSAL cifrada)
    ├── graph.js            # Cliente de Microsoft Graph API (con caché en memoria)
    ├── commands/
    │   ├── import.js       # Importación masiva desde CSV
    │   ├── search.js       # Búsqueda de usuarios
    │   ├── edit.js         # Editor interactivo
    │   ├── list.js         # Listado con filtros (incluye grupos y ID empleado)
    │   ├── assign-manager.js  # Asignación masiva de manager
    │   ├── fix.js          # Correcciones de calidad de datos
    │   ├── validate.js     # Validación del Excel de trabajadores
    │   ├── validate-users.js  # Listado de usuarios con sus grupos
    │   ├── sync-check.js   # Comprobación de sincronización con M365
    │   ├── sync-employee-ids.js  # Sincronización de id_empleado/Tipo → employeeId/employeeType
    │   └── reset-password.js    # Restablecimiento de contraseña con enlace mailto:
    ├── validators/
    │   └── csv-rules.js    # Reglas de validación del Excel
    └── utils/
        ├── ansi.js         # Eliminación de códigos de escape ANSI
        ├── cache.js        # Caché en memoria para datos de Graph API (usuarios, grupos, etc.)
        ├── csv.js          # Parseo y validación de CSV
        ├── excel.js        # Lectura de archivos .xlsx
        ├── token-store.js  # Almacenamiento cifrado genérico (AES-256-GCM)
        └── upn.js          # Utilidades para UPN
```

---

## Seguridad

- `config.json` está en `.gitignore` y **nunca debe subirse al repositorio**.
- La caché de sesión MSAL (incluyendo el *refresh token*) se almacena cifrada (AES-256-GCM) en `~/.config/m365-users/msal-cache.json` con permisos `600` (solo lectura del propietario) y una clave derivada de la máquina, por lo que no es portable entre equipos.
- La autenticación es **delegada** (el usuario que hace login debe tener los permisos necesarios en el tenant). No se usan secretos de aplicación en el flujo de autenticación.
- Los permisos `User.ReadWrite.All` y `Directory.ReadWrite.All` otorgan acceso de lectura y escritura completo sobre todos los usuarios del directorio. Úsalos con precaución y solo en cuentas de administrador dedicadas.

---

## Licencia

MIT
