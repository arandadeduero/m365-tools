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

Se mostrará un código y una URL. Abre la URL en tu navegador, introduce el código y autentica con tu cuenta de administrador. El token se guarda cifrado en `~/.config/m365-users/token.json` con AES-256-GCM y solo es válido en la misma máquina.

```bash
# Cerrar sesión
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
```

Si hay varios resultados, muestra un selector interactivo. Desde el resultado puedes pasar directamente a editar el usuario.

---

### `list`

Lista todos los usuarios activos del directorio en una tabla.

```bash
m365-users list

# Filtrar usuarios con datos incompletos
m365-users list --missing-manager
m365-users list --missing-department
m365-users list --missing-job-title

# Combinar filtro con edición interactiva
m365-users list --missing-manager --edit
```

---

### `edit <upn-o-id>`

Editor interactivo para un usuario concreto. Muestra los datos actuales, agrupa los campos por sección (Identidad, Información laboral, Contacto, Dirección, Configuración, Manager, Contraseña) y aplica todos los cambios en una sola llamada a la API tras confirmación.

```bash
m365-users edit jdoe@contoso.com
m365-users edit <objectId>
```

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

Normaliza los títulos de trabajo a formato Oración (primera letra en mayúscula). Muestra una previsualización de los cambios, permite seleccionar cuáles aplicar y requiere doble confirmación antes de modificar nada.

```bash
m365-users fix job-titles
```

---

### `fix org-chart`

Analiza la integridad del organigrama completo: detecta usuarios sin manager, cadenas que no llegan al usuario raíz (CEO) y ciclos. Ofrece un bucle interactivo para corregir los problemas uno a uno.

```bash
m365-users fix org-chart
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
├── package.json
└── src/
    ├── index.js            # Punto de entrada CLI
    ├── auth.js             # Autenticación (Device Code Flow)
    ├── graph.js            # Cliente de Microsoft Graph API
    ├── commands/
    │   ├── import.js       # Importación masiva desde CSV
    │   ├── search.js       # Búsqueda de usuarios
    │   ├── edit.js         # Editor interactivo
    │   ├── list.js         # Listado con filtros
    │   ├── assign-manager.js  # Asignación masiva de manager
    │   └── fix.js          # Correcciones de calidad de datos
    └── utils/
        ├── csv.js          # Parseo y validación de CSV
        ├── token-store.js  # Almacenamiento cifrado del token
        └── upn.js          # Utilidades para UPN
```

---

## Seguridad

- `config.json` está en `.gitignore` y **nunca debe subirse al repositorio**.
- El token de sesión se almacena cifrado (AES-256-GCM) en `~/.config/m365-users/token.json` con permisos `600` (solo lectura del propietario) y una clave derivada de la máquina, por lo que no es portable.
- La autenticación es **delegada** (el usuario que hace login debe tener los permisos necesarios en el tenant). No se usan secretos de aplicación en el flujo de autenticación.
- Los permisos `User.ReadWrite.All` y `Directory.ReadWrite.All` otorgan acceso de lectura y escritura completo sobre todos los usuarios del directorio. Úsalos con precaución y solo en cuentas de administrador dedicadas.

---

## Licencia

MIT
