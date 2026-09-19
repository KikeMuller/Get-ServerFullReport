# Changelog

Todos los cambios notables de este proyecto se documentan acá.

El formato sigue [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/) y el versionado sigue [Versionado Semántico](https://semver.org/lang/es/).

## [Sin publicar]

Correcciones surgidas de ejecutar el script completo contra un controlador de dominio real (Windows Server 2025, PowerShell 5.1) por WinRM.

### Corregido

- **En toda ejecución remota, la detección de capacidades quedaba con todo en `False`.** `Remove-PropiedadesDeRemoting` copiaba las propiedades de un hashtable (`Count`, `Keys`, `Values`) en vez de sus entradas. Efecto: un DC con DNS instalado salía con la sección 1.9 diciendo "El rol de servidor DNS no esta instalado", y se omitían las secciones dependientes de rol. Ahora un hashtable pasa intacto.
- **Los enums remotos llegaban como `1`/`0`.** Un enum que vuelve de una PSSession es un PSObject cuyo valor base es el entero. Las reglas de firewall que comparan contra `'False'` o `'Allow'` nunca coincidían, así que un perfil de firewall deshabilitado no generaba hallazgo en equipos remotos. Ahora se normalizan a su nombre (`True`, `Allow`), igual que en ejecución local.
- **El JSON de una ejecución remota no se podía volver a leer, y `-BaselinePath` fallaba.** Los mismos enums se serializaban como `{"value":1,"Value":"True"}`, dos claves que solo difieren en mayúsculas. La 3.1.0 lo cubría solo en parte: el lector reintentaba únicamente si el mensaje de error contenía el texto en inglés `different casing`, así que en un SO en español nunca reintentaba. El exportador ahora escribe el nombre del enum, y el lector reintenta ante cualquier fallo, lo que además recupera los JSON ya guardados con la 3.2.0 y anteriores.
- **Falso CRIT "Volumen D con espacio crítico" sobre una ISO/DVD montada.** La exclusión de unidades ópticas leía una columna `DriveType` que la sección 1.5.2 nunca recolectaba, y una ISO de ~7 GB no cae en el filtro por tamaño. La sección ahora incluye `DriveType`, la exclusión también reconoce `CDFS`/`UDF`, y el mismo criterio se aplica al espacio libre mínimo del resumen de flota.

## [3.2.0] — 2026-09-04

Correcciones surgidas de ejecutar el script completo de punta a punta, en vez de sus funciones por separado.

### Corregido

- **El script abortaba al terminar, después de recolectar todo.** `Write-ReportLog -Message ''` (la línea en blanco del resumen final) fallaba porque el parámetro era `[Parameter(Mandatory)][string]`, que en PowerShell rechaza la cadena vacía. Falta `[AllowEmptyString()]`.
- **Se afirmaban incumplimientos sobre datos que no se pudieron leer.** Cuando un colector falla escribe `ValorActual='No se pudo determinar'` con `Cumple=$false` por defecto, y las reglas leían ese `$false` como "no cumple". Producía un CRIT de "SMBv1 habilitado" en equipos donde el dato simplemente no era consultable.
- **Se emitían mensajes OK sobre verificaciones que nunca ocurrieron.** Si el nombre de la fila no coincidía con el patrón de la regla (`"UAC"` contra `"EnableLUA"`), la regla no emitía nada y el bloque cerraba con un OK. Ahora un tópico que no pudo evaluar sus datos queda en silencio.
- UAC, RDP/NLA y TPM/Secure Boot exigen evidencia afirmativa antes de concluir en cualquier dirección.
- `-Format HTML,JSON` fallaba al invocar con `powershell.exe -File` (la forma típica de una tarea programada): PowerShell pasa la lista como un solo string y el `ValidateSet` la rechazaba con un mensaje contradictorio. Ahora se separa y valida manualmente.
- El nombre del equipo vacío (`$env:COMPUTERNAME` ausente en sesiones no interactivas) abortaba el arranque. Respaldo a `[Environment]::MachineName`.

## [3.1.0] — 2026-09-04

Correcciones surgidas de la primera corrida en un equipo real.

### Corregido

- **Toda sección de exactamente una fila se renderizaba vacía en PowerShell 5.1** (17 secciones afectadas, entre ellas hardware, configuración del SO e historial de logon). `return @($x)` desenrolla un arreglo de un solo elemento, y en PS 5.1 indexar el objeto suelto resultante devuelve `$null`, así que la tabla se quedaba sin columnas. Resuelto con el operador coma unario en todos los normalizadores. En PS 7 el mismo código funcionaba, que es por qué no apareció antes.
- **`Invoke-Remote` levantaba un Job incluso en ejecución local**, lo que agregaba las columnas `PSComputerName`, `RunspaceId` y `PSShowComputerName` a todas las tablas, serializaba las fechas y levantaba un proceso PowerShell por cada una de las ~100 secciones. Ahora en local usa un runspace propio: mantiene el timeout duro, sin serializar.
- **Parseo de fechas con cultura invariante**: `03-09-2026` (3 de septiembre) se leía como 9 de marzo, y reportaba "firmas de antivirus desactualizadas hace 179 días" sobre un equipo actualizado el día anterior. Afectaba también a certificados, hotfixes y eventos.
- **Seis falsos "CA raíz sospechosa"** sobre autoridades del Microsoft Trusted Root Program (SSL.com, SecureTrust, SECOM, HARICA). La lista blanca pasó a ~90 emisores públicos; una raíz desconocida es INFO agrupado, y el WARN queda para raíces autofirmadas con el nombre del equipo o patrones de interceptación TLS.
- CRIT falso sobre unidades ópticas o removibles vacías (0 GB de 0 GB).
- Reglas que disparaban sobre valores `N/D`.
- "Windows sin activar" evaluaba un SKU de Office en vez del sistema operativo.
- **El JSON no lo podía leer ni el propio PowerShell**: salía con BOM, y un enum serializado como `{"value":0,"Value":"..."}` (dos claves que solo difieren en mayúsculas) hacía fallar a `ConvertFrom-Json`. Eso dejaba inutilizable la comparación contra baseline.
- Fechas serializadas como `/Date(1743109017000)/` en el JSON (formato de `JavaScriptSerializer` en PS 5.1). Ahora ISO 8601.
- `TermService` corre en todo Windows, así que cualquier equipo cliente quedaba marcado como RDS Session Host.
- Doble conteo en los tiempos por sección.
- Encoding de la salida de `netsh`, bytes nulos en la versión del TPM, y la etiqueta de BitLocker que llamaba "volumen de datos" a `C:`.

### Rendimiento

- Top de procesos: una consulta WMI en vez de una por proceso.
- Tareas programadas: una tubería en vez de una llamada por tarea.
- Búsqueda de venv: poda la recursión antes de descender a `node_modules` y las cachés de AppData.
- Configuración de WinRM: lee el proveedor `WSMan:` en vez de invocar `winrm.cmd`, que podía quedar esperando entrada del usuario.

## [3.0.0] — 2026-09-04

Reescritura completa. De 1.126 a más de 8.000 líneas.

### Agregado

- **Motor de hallazgos**: 60 reglas en 34 tópicos, con severidad, evidencia, recomendación y mapeo a controles CIS / ISO 27001.
- **Resumen ejecutivo** al inicio del reporte, con puntaje de salud, tarjetas KPI y tabla de hallazgos filtrable.
- **Salida en JSON, CSV y Markdown** además del HTML, con metadatos de auditoría y hash SHA256.
- **Comparación contra baseline** (`-BaselinePath`) para detectar cambios entre corridas.
- **Modo flota**: varios equipos en paralelo, con índice consolidado y CSV de hallazgos.
- **64 secciones nuevas** (de ~40 a 104): administradores locales, usuarios locales, política de contraseñas y auditoría, derechos de usuario, puertos en escucha con proceso dueño, hardening del SO, TPM y Secure Boot, Device Guard, LAPS, activación, CAs raíz, dominio y replicación AD, virtualización, SQL Server, DFS, impresoras, clúster, NIC teaming, iSCSI/MPIO, WinRM, uptime, apagados inesperados, event log, rendimiento, salud de discos, ACLs NTFS de shares, rutas, hosts y proxy.
- `-Sections` / `-SkipSections` para acotar la corrida.
- `-Redact` para enmascarar datos sensibles.
- Códigos de salida (`0`/`1`/`2`) para encadenar en un scheduler.

### Cambiado

- **Una sola sesión WinRM** reutilizada para todas las secciones, en vez de una por bloque (~60 conexiones por servidor).
- **TOC generado automáticamente** desde las secciones registradas, en vez del listado escrito a mano.
- **Reporte HTML rediseñado**: tema claro/oscuro, buscador global, tablas ordenables y filtrables, secciones colapsables, CSS de impresión.
- El HTML se escribe con escapado explícito en vez de `ConvertTo-Html`: un nombre de servicio con `<` ya no rompe ni inyecta en el reporte.
- Las reglas de firewall se enriquecen con puertos y direcciones mediante un join indexado por `InstanceID`.
- La consulta de actualizaciones faltantes pasó a ser opt-in (`-IncludeMissingUpdates`) porque puede colgarse minutos.
- Timeout explícito por sección: un cmdlet colgado devuelve un mensaje en su sección en vez de detener el script.

### Corregido

- **La IP del broker PAM estaba hardcodeada en el script.** Ahora es el parámetro `-PamBrokerEndpoint`; sin él no se prueba conectividad.
- Servicios: se agregó `PathName` y la detección de *unquoted service path*.
- Certificados: CN separado del Subject, estado de expiración, uso mejorado y presencia de clave privada.

## [2.0.0] — anterior

Versión inicial: reporte HTML con TOC escrito a mano, ~40 secciones, detección de roles IIS/DHCP/DNS/File Server/RDS/WSUS, y secciones de seguridad y entornos Python.

[3.2.0]: https://github.com/KikeMuller/Get-ServerFullReport/releases/tag/v3.2.0
[3.1.0]: https://github.com/KikeMuller/Get-ServerFullReport/releases/tag/v3.1.0
[3.0.0]: https://github.com/KikeMuller/Get-ServerFullReport/releases/tag/v3.0.0
