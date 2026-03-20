# XGD - Utilidad General de Active Directory

Herramienta gráfica (GUI) en PowerShell para la gestión de **grupos** y **equipos** en Active Directory. Proporciona una interfaz WinForms completa para operaciones comunes de administración AD sin necesidad de trabajar directamente con consolas RSAT.

---

## Requisitos

| Requisito | Detalle |
|-----------|---------|
| **Sistema operativo** | Windows 10 / Windows Server 2016 o superior |
| **PowerShell** | 5.1 (incluido en Windows) |
| **Módulo AD** | `ActiveDirectory` (RSAT: Active Directory Domain Services) |
| **Permisos** | Cuenta de dominio con permisos de lectura en AD; permisos de escritura para operaciones de creación/modificación |

### Instalación de RSAT (si no está disponible)

```powershell
# Windows 10/11
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0

# Windows Server
Install-WindowsFeature RSAT-AD-PowerShell
```

---

## Estructura de archivos

```
XGD/
├── XGD.ps1              # Punto de entrada (lanzador)
├── XGD.General.ps1      # Aplicación completa (funciones, UI, lógica)
├── XGD.config.json      # Configuración persistente (JSON)
└── README.md            # Este archivo
```

| Archivo | Función |
|---------|---------|
| `XGD.ps1` | Script lanzador. Establece el directorio de trabajo y carga `XGD.General.ps1` mediante dot-sourcing. |
| `XGD.General.ps1` | Contiene toda la lógica: definiciones de funciones, construcción de la interfaz, event handlers y flujo principal. |
| `XGD.config.json` | Almacena la configuración operativa y las OUs de trabajo guardadas. Se genera automáticamente con valores por defecto si no existe. |

---

## Ejecución

```powershell
# Opción 1: Desde PowerShell
cd "C:\ruta\a\XGD"
.\XGD.ps1

# Opción 2: Click derecho → "Ejecutar con PowerShell" sobre XGD.ps1
```

> **Nota:** Si la política de ejecución impide ejecutar scripts, ejecútelo con:
> ```powershell
> powershell -ExecutionPolicy Bypass -File "C:\ruta\a\XGD\XGD.ps1"
> ```

---

## Interfaz de usuario

La ventana principal (1500×820 px) se divide en dos zonas:

### Zona izquierda (controles y registro)

| Elemento | Descripción |
|----------|-------------|
| **Título** | Nombre de la aplicación (configurable en `UiTitle`) |
| **OUs de trabajo** | Lista de OUs seleccionadas como ámbito de búsqueda de grupos |
| **+ OU / - OU** | Botones para agregar o quitar OUs de trabajo |
| **Guardar OUs** | Persiste las OUs seleccionadas en `XGD.config.json` para futuras sesiones |
| **Grupo AD** | ComboBox con autocompletado para seleccionar un grupo cargado |
| **Cargar Grupos** | Busca grupos dentro de las OUs de trabajo seleccionadas |
| **Panel de acciones** | 9 botones con las operaciones principales |
| **Registro de actividad** | Log en tiempo real de todas las operaciones realizadas |

### Zona derecha (resultados)

| Elemento | Descripción |
|----------|-------------|
| **Resultados** | Área de texto o árbol donde se muestran los resultados |
| **Contador de equipos** | Muestra el total de equipos encontrados |
| **Buscar / Siguiente** | Búsqueda incremental dentro de los resultados |
| **Copiar** | Copia los resultados al portapapeles |
| **Limpiar** | Limpia el área de resultados |
| **Exportar CSV** | Exporta los resultados de equipos a archivo CSV |
| **Guardar TXT** | Guarda el texto de resultados en archivo TXT |

---

## Flujo de trabajo

```
1. Seleccionar OUs de trabajo  →  + OU (navegar árbol AD)
2. Cargar Grupos               →  Se buscan grupos en las OUs seleccionadas
3. Seleccionar grupo           →  ComboBox con autocompletado
4. Ejecutar operación          →  Botón de acción correspondiente
5. Ver resultados              →  Panel derecho (texto o árbol)
6. Exportar/copiar             →  Botones inferiores del panel de resultados
```

---

## Operaciones disponibles

### 1. Extraer Equipos
Obtiene todos los equipos miembros del grupo AD seleccionado. Muestra nombre, sistema operativo, última conexión y ruta OU en formato de árbol jerárquico.

### 2. Explorar Equipos
Explora los equipos en las bases LDAP configuradas (`ExploreSearchBases`), mostrando su distribución por OUs en formato de árbol. Útil para auditar la estructura de equipos.

### 3. Comparar Equipos
Compara los equipos de dos grupos AD. Abre un diálogo para seleccionar un segundo grupo y muestra:
- Equipos solo en el primer grupo
- Equipos solo en el segundo grupo
- Equipos en ambos grupos

### 4. Incluir Equipos
Permite añadir equipos al grupo seleccionado. Abre un diálogo con navegador de OUs en árbol donde se pueden seleccionar equipos individualmente o por OU completa mediante checkboxes.

### 5. Añadir Equipos Filtrados
Igual que "Incluir Equipos" pero aplica un filtro regex adicional (`FilteredComputerIncludeRegex`). Útil para seleccionar solo equipos que cumplan un patrón (p.ej., nombres que terminen en "TV").

### 6. Modificar Grupo
Abre un gestor completo del grupo seleccionado que permite:
- Ver todos los miembros (equipos) del grupo
- Eliminar equipos individualmente o en bloque
- Añadir equipos desde el navegador OU
- Mover equipos entre OUs
- Ver resumen de cambios

### 7. Extraer Grupos
Para un equipo dado, muestra todos los grupos AD a los que pertenece. Permite buscar por nombre de equipo.

### 8. Crear Equipos
Diálogo para crear nuevas cuentas de equipo en AD:
- Seleccionar OU destino mediante navegador de árbol
- Introducir nombres de equipos (uno por línea)
- Opcionalmente aplicar delegación ACL (configurable)
- Asignar descripción automática con plantilla

### 9. CSV a Grupos
Permite cargar un archivo CSV con nombres de equipos y añadirlos a uno o varios grupos AD seleccionados. Los grupos destino pueden preconfigurarse en `FixedTargetGroups`.

---

## Configuración (`XGD.config.json`)

El archivo de configuración se crea automáticamente con valores por defecto en la primera ejecución. Para modificarlo, edítelo directamente con un editor de texto.

### Propiedades

| Propiedad | Tipo | Descripción | Valor por defecto |
|-----------|------|-------------|-------------------|
| `UiTitle` | string | Título de la ventana principal | `"XGD - Utilidad general AD"` |
| `Server` | string | DC preferido para conexiones LDAP (vacío = auto-detect) | `""` |
| `GroupSearchBases` | string[] | Bases LDAP para búsqueda de grupos (legacy) | Auto-detect |
| `GroupNameFilter` | string | Filtro wildcard para carga de grupos | `"*"` |
| `ExploreSearchBases` | string[] | Bases LDAP para la operación "Explorar Equipos" | Auto-detect |
| `BrowseRoots` | string[] | Raíces del navegador de OUs (diálogos de selección) | Auto-detect |
| `ComputerContainerName` | string | Nombre de la sub-OU que contiene equipos dentro de cada OU | `"Equipos"` |
| `HiddenOUSegments` | string[] | Segmentos de OU a ocultar en rutas de visualización | `["Equipos"]` |
| `ExcludedOUPatterns` | string[] | Patrones de OUs a excluir de resultados | `["_Cuentas deshabilitadas", "Transito", "Pre-Windows 10"]` |
| `ExcludedComputerNameRegex` | string | Regex para excluir equipos por nombre | `""` |
| `FilteredComputerLabel` | string | Etiqueta del botón de filtro especial | `"Anadir Equipos Filtrados"` |
| `FilteredComputerIncludeRegex` | string | Regex del filtro especial | `""` |
| `ApplyDelegationOnCreate` | bool | Aplicar delegación ACL al crear equipos | `false` |
| `DelegationGroupDN` | string | DN del grupo para delegación ACL | `""` |
| `CreatedComputerDescriptionTemplate` | string | Plantilla de descripción al crear equipos (`{date}` se reemplaza) | `"Equipo creado {date}"` |
| `FixedTargetGroups` | string[] | Grupos sugeridos para "CSV a Grupos" | `[]` |
| `SavedWorkOUs` | string[] | OUs de trabajo guardadas (persiste la selección entre sesiones) | `[]` |

### Ejemplo de configuración

```json
{
    "UiTitle": "XGD - Mi Organización",
    "Server": "dc01.midominio.local",
    "GroupSearchBases": [],
    "GroupNameFilter": "GRP-*",
    "ExploreSearchBases": ["OU=Empresa,DC=midominio,DC=local"],
    "BrowseRoots": ["OU=Empresa,DC=midominio,DC=local"],
    "ComputerContainerName": "Equipos",
    "HiddenOUSegments": ["Equipos"],
    "ExcludedOUPatterns": ["_Cuentas deshabilitadas", "Transito"],
    "ExcludedComputerNameRegex": "^TEST-",
    "FilteredComputerLabel": "Añadir TVs",
    "FilteredComputerIncludeRegex": "[TV]$",
    "ApplyDelegationOnCreate": true,
    "DelegationGroupDN": "CN=GRP-Delegacion,OU=Grupos,DC=midominio,DC=local",
    "CreatedComputerDescriptionTemplate": "Equipo creado {date} por XGD",
    "FixedTargetGroups": ["GRP-Todos-Equipos", "GRP-Actualizaciones"],
    "SavedWorkOUs": [
        "OU=Sede-Madrid,OU=Empresa,DC=midominio,DC=local",
        "OU=Sede-Barcelona,OU=Empresa,DC=midominio,DC=local"
    ]
}
```

---

## Persistencia de OUs de trabajo

Las OUs de trabajo seleccionadas se pueden guardar pulsando el botón **"Guardar OUs"**. Al iniciar la aplicación, se restauran automáticamente las OUs guardadas previamente desde `XGD.config.json`.

Esto permite mantener el ámbito de trabajo entre sesiones sin necesidad de volver a seleccionar las OUs cada vez.

---

## Arquitectura interna

### Módulos funcionales

| Módulo | Funciones principales | Líneas aprox. |
|--------|----------------------|---------------|
| **Utilidades** | `Convert-ToStringArray`, `Join-Lines`, `Escape-ADFilterValue` | 30–68 |
| **AD Core** | `Get-ADCommonParameters`, `Get-DefaultNamingContext`, `Resolve-ConfiguredBases` | 71–103 |
| **Configuración** | `Get-DefaultConfig`, `Normalize-Config`, `Save-Config`, `Load-Config` | 105–185 |
| **UI Helpers** | `Set-Status`, `Mostrar-Mensaje`, `Show-ErrorDialog`, `Clear-Results` | 187–237 |
| **DN/OU Processing** | `Split-DistinguishedName`, `Get-DnFriendlyName`, `Get-OUPathSegments`, `Get-DisplayPathSegments`, `Test-ExcludedOU`, `Test-ComputerNameAllowed` | 239–350 |
| **AD Queries** | `Get-ChildOUs`, `Get-ImmediateComputersForTree`, `Get-ComputersFromOU`, `Get-ComputersFromSearchBases`, `Get-ADGroupDirectMembersRanged`, `Get-ADGroupMemberSafe` | 353–592 |
| **Record Resolution** | `Resolve-ComputerRecord`, `Resolve-GroupRecord`, `Get-SelectedGroupRecord` | 555–657 |
| **Results Display** | `Get-ComputerSummaryText`, `Show-ResultsText`, `Show-ResultsComputerTree`, `Focus-TreeNode`, `Search-LoadedTreeNodes` | 659–804 |
| **OU Management** | `Refresh-OUListBox`, `Add-WorkOU`, `Remove-WorkOU`, `Load-GroupList`, `Save-WorkOUs`, `Load-SavedWorkOUs` | 806–351 |
| **Tree Builder** | `Add-PlaceholderNode`, `New-OUNode`, `Load-OUNodeChildren`, `New-OUTree` | 923–1054 |
| **CSV Processing** | `Resolve-ComputerEntriesFromCsv` | 1011–1054 |
| **Dialogs** | `Show-ComputerSelectionDialog`, `Show-OUSelectionDialog`, `Show-ModifyGroupDialog`, `Show-CreateComputersDialog`, `Show-CompareComputersDialog`, `Show-ExtractComputerGroupsDialog`, `Show-CsvToGroupsDialog` | 1056–2313 |
| **Operations** | `Invoke-ExtractSelectedGroupComputers`, `Invoke-ExploreEnvironmentComputers`, `Invoke-AddComputersToGroup` | 2352–2420 |
| **Main Form** | Construcción UI, event handlers, inicialización | 2422–2834 |

### Flujo de datos

```
XGD.ps1
  └─→ dot-source XGD.General.ps1
        ├─→ Import-Module ActiveDirectory
        ├─→ Declaración de variables Script-scope
        ├─→ Definición de funciones
        ├─→ Load-Config (lee o crea XGD.config.json)
        ├─→ Construcción del formulario WinForms
        ├─→ Registro de event handlers
        ├─→ Load-SavedWorkOUs (restaura OUs guardadas)
        └─→ MainForm.ShowDialog() (bucle principal)
```

### Patrones de diseño

- **Scope-safe state en diálogos**: Las variables que necesitan sobrevivir entre event handlers de WinForms se almacenan en la propiedad `.Tag` de controles, evitando problemas de scope en closures de PowerShell.
- **Ranged member retrieval**: Para grupos con más de 1500 miembros, se usa paginación manual con `msds-membersCid` y rangos (`member;range=N-M`) para superar las limitaciones de LDAP.
- **Lazy tree loading**: Los nodos del árbol OU se cargan bajo demanda al expandirse, usando nodos placeholder para indicar contenido pendiente.
- **Normalización de configuración**: Al cargar la configuración, se normalizan todos los valores contra los defaults, asegurando que propiedades nuevas se incorporen sin romper configuraciones existentes.

---

## Solución de problemas

| Problema | Solución |
|----------|----------|
| No se carga el módulo ActiveDirectory | Instalar RSAT (ver sección Requisitos) |
| No aparecen grupos al cargar | Verificar que las OUs de trabajo seleccionadas contienen grupos. Revisar `GroupNameFilter` en config |
| No se ve el servidor correcto | Editar `Server` en `XGD.config.json` con el FQDN del DC deseado |
| Los equipos no aparecen en "Explorar" | Verificar `ExploreSearchBases` en la configuración. Si está vacío, se usa el naming context por defecto |
| Error de permisos al crear equipos | La cuenta debe tener permisos de creación de objetos Computer en la OU destino |
| El archivo config se reinicia | Si el JSON tiene errores de sintaxis, se regenera con defaults. Hacer backup antes de editar manualmente |

---

## Licencia

Herramienta interna de administración. Uso restringido al ámbito organizacional autorizado.
