<#
.SYNOPSIS
    Genera un reporte "as-built" completo de configuracion de uno o varios servidores
    Windows (local o remoto via WinRM), siguiendo la estructura estandar de
    "Server Health Check / Documentacion de Configuracion": Hardware, Sistema Operativo,
    Firewall, Red, Storage, IIS, File Server, DHCP, DNS, Terminal Services/RD Licensing,
    Seguridad y Cumplimiento, WSUS, Group Policy, Cuentas y Accesos, Entornos de
    Desarrollo (Python/Git), Plataforma y Aplicaciones, y Salud y Eventos.

.DESCRIPTION
    Reemplaza / complementa la exportacion de "System Information" (msinfo32), que solo
    cubre hardware/drivers y omite software instalado, servicios, licencias, roles,
    configuracion de IIS/File Server/DHCP/DNS/Terminal Services, postura de seguridad,
    cumplimiento normativo y salud operativa.

    El script detecta automaticamente que roles/capacidades estan presentes en el equipo
    objetivo (IIS, File Server, DHCP, DNS, RDS Session Host / RD Licensing, WSUS, Hyper-V,
    SQL Server, Cluster de failover, Docker, si es maquina virtual, si es Controlador de
    Dominio, etc.) y omite u obvia (indicando "No instalado" o "No aplica") las secciones
    que no correspondan. Es compatible con Windows Server y Windows 10/11.

    Soporta ejecucion contra multiples equipos en una sola corrida (-ComputerName acepta
    un arreglo, por ejemplo el resultado de una consulta a Active Directory), generacion
    de reportes en HTML / JSON / CSV / Markdown, filtrado de secciones por prefijo
    numerico (-Sections / -SkipSections), comparacion contra una linea base previa
    (-BaselinePath) para detectar "drift" de configuracion, y enmascarado de datos
    sensibles (-Redact) para poder compartir el reporte sin exponer IPs completas,
    cuentas de dominio, rutas internas, numeros de serie ni huellas digitales completas.

    Cada bloque de recoleccion de datos esta aislado: un rol ausente, un modulo no
    disponible, o un error de permisos en un equipo NUNCA aborta el reporte completo;
    se registra como hallazgo o nota y el resto de las secciones continua normalmente.

.PARAMETER ComputerName
    Uno o mas nombres de equipo o direcciones IP a inventariar. Acepta el resultado de
    una consulta a Active Directory (por ejemplo Get-ADComputer). Por defecto, el
    equipo local.

.PARAMETER OutputPath
    Carpeta donde se guardaran los archivos de reporte generados (uno por -Format
    solicitado, por equipo). Por defecto: el mismo directorio desde donde se ejecuta
    el script. Si el script se ejecuta pegando el codigo directamente en una consola
    (sin archivo .ps1), se usa el directorio de trabajo actual como respaldo.

.PARAMETER Format
    Uno o mas formatos de salida: HTML, JSON, CSV, Markdown. Por defecto HTML y JSON.

.PARAMETER Sections
    Lista de prefijos de seccion a incluir (ej '1.5','1.11'). Si esta vacio (por
    defecto), se incluyen todas las secciones salvo las indicadas en -SkipSections.
    El filtro es por prefijo de punto: '1.11' habilita '1.11' y toda su descendencia
    ('1.11.1', '1.11.3', etc), pero no habilita '1.1' ni '1.112' (si existiera).

.PARAMETER SkipSections
    Lista de prefijos de seccion a excluir, con la misma logica de coincidencia por
    segmentos que -Sections. Por ejemplo, '1.2' excluye '1.2' y '1.2.3' pero NO
    excluye '1.12' (no es un match de prefijo por punto, solo de texto).

.PARAMETER BaselinePath
    Ruta a un JSON generado en una corrida anterior de este mismo script (formato
    -Format JSON). Si se especifica, se agrega una seccion de comparacion ("drift")
    entre el estado actual y esa linea base.

.PARAMETER Redact
    Si esta presente, enmascara datos sensibles en el reporte (IPs, cuentas de
    dominio, rutas, numeros de serie y huellas digitales) via Protect-Sensitive.
    Util para compartir reportes con terceros o adjuntarlos a tickets.

.PARAMETER ThrottleLimit
    Cantidad maxima de equipos a procesar en paralelo cuando -ComputerName tiene mas
    de un elemento (modo flota). Por defecto 8.

.PARAMETER Credential
    Credencial a usar para las sesiones remotas (WinRM). Si no se especifica, se usa
    el contexto de seguridad del usuario que ejecuta el script.

.PARAMETER PamBrokerEndpoint
    Direccion "IP:Puerto" (o solo IP) del Resource Broker de la solucion PAM
    (BeyondTrust/Bomgar u otra) contra el cual probar conectividad en la seccion 1.11.
    Vacio por defecto, en cuyo caso el test de conectividad se omite.

.PARAMETER IncludeMissingUpdates
    Switch opcional. Si esta presente, consulta el catalogo de Windows Update via COM
    (Microsoft.Update.Session) para listar actualizaciones faltantes. Desactivado por
    defecto porque esa consulta puede ser lenta (varios minutos) en algunos equipos.

.PARAMETER EventLogDays
    Cantidad de dias hacia atras a considerar al revisar el visor de eventos para la
    seccion de Salud y Eventos. Por defecto 7.

.PARAMETER PerfSampleSeconds
    Segundos de muestreo de contadores de rendimiento (CPU/Memoria/Disco) para incluir
    una linea base de performance en el reporte. Por defecto 0 (se omite el muestreo).

.PARAMETER ScanGitRepos
    Switch opcional. Si se especifica, busca repositorios Git dentro de -GitScanPaths.
    Desactivado por defecto porque puede ser lento en discos grandes. Cualquier
    credencial embebida en la URL del remoto se enmascara automaticamente.

.PARAMETER GitScanPaths
    Rutas donde buscar entornos virtuales (venv) y, si -ScanGitRepos esta activo,
    repositorios Git. Por defecto cubre carpetas de proyecto tipicas y el perfil de
    usuario.

.PARAMETER Quiet
    Switch opcional. Suprime la salida de progreso por consola (Write-ReportLog sigue
    acumulando las lineas para incluirlas en el reporte, solo se omite Write-Host).

.EXAMPLE
    .\Get-ServerFullReport.ps1
    Genera el reporte del equipo local (HTML + JSON) en la carpeta del script.

.EXAMPLE
    .\Get-ServerFullReport.ps1 -ComputerName SRVWEB01 -Credential (Get-Credential) -OutputPath D:\Reportes
    Genera el reporte de un servidor remoto via WinRM usando credenciales especificas.

.EXAMPLE
    .\Get-ServerFullReport.ps1 -ComputerName (Get-ADComputer -Filter { OperatingSystem -like '*Server*' } | Select-Object -ExpandProperty Name) -ThrottleLimit 10 -Format HTML,CSV
    Releva toda la flota de servidores del dominio (consulta a Active Directory) con
    hasta 10 equipos procesados en paralelo.

.EXAMPLE
    .\Get-ServerFullReport.ps1 -ComputerName SRVAPP02 -Sections '1.11' -SkipSections '1.11.5'
    Genera unicamente la seccion 1.11 (Seguridad y Cumplimiento) de un servidor,
    excluyendo el detalle de Tareas Programadas (1.11.5).

.EXAMPLE
    .\Get-ServerFullReport.ps1 -ComputerName SRVDB01 -BaselinePath C:\Baselines\SRVDB01_20260601.json
    Compara la configuracion actual del servidor contra una linea base previa y agrega
    una seccion de "drift" con los cambios detectados.

.EXAMPLE
    .\Get-ServerFullReport.ps1 -ComputerName SRVFILE01 -Format JSON,CSV -Redact -OutputPath \\nas\Auditorias
    Exporta el inventario en JSON y CSV con datos sensibles enmascarados (IPs, cuentas
    de dominio, rutas y numeros de serie/huellas digitales parciales), apto para
    compartir con auditoria externa.

.NOTES
    Requisitos:
      - PowerShell 5.1 o superior en el equipo desde donde se ejecuta el script.
      - Para equipos remotos: WinRM habilitado en el equipo objetivo y accesible por
        red (puertos TCP 5985 HTTP / 5986 HTTPS), y una cuenta con permisos de
        administrador local en dicho equipo (via -Credential o el contexto actual).
      - Para inventario local completo: se recomienda ejecutar la consola como
        Administrador. El script NO exige privilegios elevados para poder correr
        (se detecta en runtime y se advierte que secciones quedaran incompletas),
        pero varias secciones (Seguridad 1.11, Cuentas 1.14, ciertas consultas CIM y
        de Storage) requieren privilegios elevados para devolver datos completos.
      - Es de SOLO LECTURA: no crea, modifica ni elimina nada en el equipo auditado
        (salvo archivos temporales propios que el propio script limpia).

    Version 3.0.0:
      - Reescritura completa sobre una arquitectura modular: cada colector llama a
        Add-ReportSection / Add-Finding en lugar de construir el HTML directamente.
      - Deteccion de capacidades unificada en una sola llamada remota (Get-TargetCapabilities).
      - Soporte de flota (-ComputerName como arreglo), filtrado de secciones,
        comparacion contra baseline, exportacion multi-formato y enmascarado de datos
        sensibles.
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Position = 0, HelpMessage = 'Uno o mas equipos a inventariar. Por defecto, el equipo local.')]
    # Respaldo a [Environment]::MachineName: la variable de entorno COMPUTERNAME
    # puede no existir (sesiones no interactivas, contenedores, PowerShell sobre
    # Linux), y un nombre vacio aborta el script al enlazar los parametros.
    [string[]]$ComputerName = @($(if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName })),

    [Parameter(HelpMessage = 'Carpeta donde se guardaran los reportes generados.')]
    [string]$OutputPath = $(if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }),

    # Sin ValidateSet a proposito: al invocar el script con
    # 'powershell.exe -File script.ps1 -Format HTML,JSON' (la forma tipica de una
    # tarea programada), PowerShell pasa "HTML,JSON" como UN solo string y el
    # ValidateSet lo rechaza con un mensaje absurdo ('el argumento "HTML,JSON" no
    # pertenece al conjunto "HTML,JSON,CSV,Markdown"'). La validacion se hace mas
    # abajo, despues de separar por coma.
    [Parameter(HelpMessage = 'Formatos de salida a generar: HTML, JSON, CSV, Markdown.')]
    [string[]]$Format = @('HTML', 'JSON'),

    [Parameter(HelpMessage = 'Prefijos de seccion a incluir (ej 1.5, 1.11). Vacio = todas.')]
    [string[]]$Sections = @(),

    [Parameter(HelpMessage = 'Prefijos de seccion a excluir (ej 1.2, 1.15.5).')]
    [string[]]$SkipSections = @(),

    [Parameter(HelpMessage = 'Ruta a un JSON de una corrida anterior para generar seccion de drift.')]
    [string]$BaselinePath = '',

    [Parameter(HelpMessage = 'Enmascara IPs, cuentas de dominio, rutas, seriales y huellas digitales.')]
    [switch]$Redact,

    [Parameter(HelpMessage = 'Maximo de equipos procesados en paralelo en modo flota.')]
    [int]$ThrottleLimit = 8,

    [Parameter(HelpMessage = 'Credencial para las sesiones remotas WinRM.')]
    [pscredential]$Credential,

    [Parameter(HelpMessage = 'IP:Puerto del Resource Broker PAM a probar conectividad. Vacio = se omite.')]
    [string]$PamBrokerEndpoint = '',

    [Parameter(HelpMessage = 'Consulta Windows Update via COM para listar actualizaciones faltantes (lento).')]
    [switch]$IncludeMissingUpdates,

    [Parameter(HelpMessage = 'Dias hacia atras a revisar en el visor de eventos.')]
    [int]$EventLogDays = 7,

    [Parameter(HelpMessage = 'Segundos de muestreo de contadores de rendimiento. 0 = omitir.')]
    [int]$PerfSampleSeconds = 0,

    [Parameter(HelpMessage = 'Busca repositorios Git dentro de -GitScanPaths (desactivado por defecto).')]
    [switch]$ScanGitRepos,

    [Parameter(HelpMessage = 'Rutas donde buscar entornos virtuales de Python y repositorios Git.')]
    [string[]]$GitScanPaths = @(
        "$env:SystemDrive\inetpub", "$env:SystemDrive\Scripts", "$env:SystemDrive\Source",
        "$env:SystemDrive\Repos", "$env:SystemDrive\Proyectos", "$env:USERPROFILE"
    ),

    [Parameter(HelpMessage = 'Suprime la salida de progreso por consola.')]
    [switch]$Quiet
)

# =============================================================================
# REGION: CORE - Variables de script y utilidades base
# =============================================================================
#region CORE-Utilidades

$script:ScriptVersion = '3.2.1'

# --- Normalizacion de -Format -------------------------------------------------
# Acepta 'HTML,JSON', 'html json', @('HTML','JSON') y cualquier combinacion.
$script:FormatosValidos = @('HTML', 'JSON', 'CSV', 'Markdown')
$formatosPedidos = New-Object System.Collections.ArrayList
foreach ($entrada in @($Format)) {
    foreach ($pieza in ("$entrada" -split '[,;\s]+')) {
        $t = "$pieza".Trim()
        if ($t -eq '') { continue }
        $match = $script:FormatosValidos | Where-Object { $_ -ieq $t } | Select-Object -First 1
        if ($match) {
            if (-not $formatosPedidos.Contains($match)) { [void]$formatosPedidos.Add($match) }
        } else {
            Write-Warning "Formato de salida desconocido y omitido: '$t'. Validos: $($script:FormatosValidos -join ', ')."
        }
    }
}
if ($formatosPedidos.Count -eq 0) { [void]$formatosPedidos.Add('HTML'); [void]$formatosPedidos.Add('JSON') }
$Format = $formatosPedidos.ToArray()
$script:IsPS7         = $PSVersionTable.PSVersion.Major -ge 7
$script:StartTime     = Get-Date

# Copias en script scope de parametros consumidos por las funciones de este bloque,
# para que Test-SectionEnabled / Protect-Sensitive no dependan del scope del caller.
$script:Sections     = $Sections
$script:SkipSections = $SkipSections
$script:Redact       = [bool]$Redact

# Estado por-equipo. En modo flota, el bucle principal (fuera de este fragmento)
# actualiza ComputerName / IsRemote / Session en cada iteracion, llamando a
# New-TargetSession antes de recolectar y a Remove-TargetSession al terminar.
$script:TargetName  = $null
$script:IsRemote      = $false
$script:Session       = $null
$script:SessionFailed = $false

# Colecciones acumuladas durante la recoleccion, consumidas por el renderer.
$script:LogLines       = New-Object System.Collections.Generic.List[string]
$script:Timings        = New-Object System.Collections.Generic.List[object]
$script:Findings       = New-Object System.Collections.Generic.List[object]
$script:ReportSections = New-Object System.Collections.Generic.List[object]

# Capacidades detectadas del equipo objetivo actual (ver Get-TargetCapabilities).
$script:Caps = @{}

function Write-ReportLog {
    <#
        Registra un mensaje con timestamp y nivel. Siempre acumula la linea en
        $script:LogLines (para poder incluir el log en el reporte final); solo
        escribe por consola si -Quiet no fue especificado.
    #>
    param(
        # AllowEmptyString es necesario: [Parameter(Mandatory)][string] rechaza la
        # cadena vacia, y el resumen final usa -Message '' para imprimir lineas en
        # blanco separadoras. Sin esto el script aborta al terminar la recoleccion.
        [Parameter(Mandatory)][AllowEmptyString()][AllowNull()][string]$Message,
        [ValidateSet('Info', 'Warn', 'Error', 'Success', 'Debug')][string]$Level = 'Info'
    )

    if ($null -eq $Message) { $Message = '' }

    $ts = Get-Date -Format 'HH:mm:ss'
    # Una linea vacia se imprime vacia de verdad, sin el prefijo de timestamp.
    $line = if ($Message.Trim() -eq '') { '' } else { "[$ts] [$Level] $Message" }
    $script:LogLines.Add($line) | Out-Null

    if ($Quiet) { return }

    $color = switch ($Level) {
        'Info'    { 'Gray' }
        'Warn'    { 'Yellow' }
        'Error'   { 'Red' }
        'Success' { 'Green' }
        'Debug'   { 'DarkGray' }
        default   { 'White' }
    }
    Write-Host $line -ForegroundColor $color
}

function Measure-Section {
    <#
        Ejecuta $ScriptBlock cronometrando su duracion. Nunca propaga excepciones:
        cualquier error queda registrado en el log y en $script:Timings, y la funcion
        devuelve $null en ese caso para que el llamador pueda seguir adelante.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock
    )

    $sw       = [System.Diagnostics.Stopwatch]::StartNew()
    $estado   = 'OK'
    $resultado = $null

    try {
        $resultado = & $ScriptBlock
        $sw.Stop()
        Write-ReportLog -Message "$Name : OK ($([math]::Round($sw.Elapsed.TotalSeconds, 1))s)" -Level Success
    } catch {
        $sw.Stop()
        $estado = "Error: $($_.Exception.Message)"
        Write-ReportLog -Message "$Name : ERROR - $($_.Exception.Message)" -Level Error
    }

    $script:Timings.Add([PSCustomObject]@{
        Seccion  = $Name
        Segundos = [math]::Round($sw.Elapsed.TotalSeconds, 2)
        Estado   = $estado
    }) | Out-Null

    return $resultado
}

function Test-Administrator {
    <# Devuelve $true si el proceso actual corre con privilegios de Administrador. #>
    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function New-TargetSession {
    <#
        Crea una unica PSSession contra $ComputerName. Si el equipo es el local,
        devuelve $null (Invoke-Remote ejecuta localmente en ese caso). Si falla la
        conexion remota, loguea el error, marca $script:SessionFailed = $true y
        devuelve $null (el bucle principal decide como reportar ese equipo).
    #>
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [pscredential]$Credential
    )

    $esLocal = $ComputerName -eq (Get-EquipoActual) -or $ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq 'localhost' -or $ComputerName -eq '.' -or [string]::IsNullOrWhiteSpace($ComputerName)
    if ($esLocal) { return $null }

    try {
        $opciones = New-PSSessionOption -OpenTimeout 15000 -OperationTimeout 180000
        $parametros = @{
            ComputerName  = $ComputerName
            SessionOption = $opciones
            ErrorAction   = 'Stop'
        }
        if ($Credential) { $parametros['Credential'] = $Credential }

        return New-PSSession @parametros
    } catch {
        Write-ReportLog -Message "No se pudo abrir sesion remota hacia $ComputerName : $($_.Exception.Message)" -Level Error
        $script:SessionFailed = $true
        return $null
    }
}

function Remove-TargetSession {
    <# Cierra una PSSession de forma prolija. Tolera $null. #>
    param([System.Management.Automation.Runspaces.PSSession]$Session)

    if ($Session) {
        try { Remove-PSSession -Session $Session -ErrorAction SilentlyContinue } catch { }
    }
}

function Invoke-Remote {
    <#
        Ejecuta $ScriptBlock en $script:Session (remoto) o localmente si no hay
        sesion activa, con un timeout duro de $TimeoutSec segundos.

        Implementacion:
          - LOCAL: se ejecuta en un runspace propio via [PowerShell]::Create(),
            con espera acotada sobre el AsyncWaitHandle. Esto da timeout duro SIN
            pagar el costo de serializacion. Antes se usaba Start-Job tambien en
            local, lo que tenia tres consecuencias malas: cada objeto volvia
            deserializado con las propiedades basura PSComputerName / RunspaceId /
            PSShowComputerName (que terminaban como columnas en todas las tablas
            del reporte), se levantaba un proceso PowerShell completo por cada una
            de las ~76 secciones, y algunos comandos interactivos hacian que
            Wait-Job fallara con "uno o varios trabajos estan bloqueados esperando
            la entrada del usuario".
          - REMOTO: Invoke-Command -AsJob + Wait-Job -Timeout, que sigue siendo el
            mecanismo mas confiable en 5.1 para cortar una llamada colgada.

        NUNCA propaga una excepcion: ante timeout o error devuelve un string
        explicativo (los colectores deben tratar un resultado string como mensaje,
        segun el contrato de Add-ReportSection).

        Nota para los colectores: cuando se ejecuta contra $script:Session, los
        objetos que vuelven estan deserializados (PSObject "plano", sin metodos de
        tipo). Usar siempre Select-Object con nombres de propiedad explicitos (o
        Select-Object @{N=...;E=...}) dentro del ScriptBlock, nunca depender de
        metodos .NET sobre el resultado ya recibido en el llamador.
    #>
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @(),
        [int]$TimeoutSec = 120
    )

    if ($script:Session) {
        return (Invoke-RemoteViaJob -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList -TimeoutSec $TimeoutSec)
    }
    return (Invoke-LocalConTimeout -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList -TimeoutSec $TimeoutSec)
}

function Invoke-LocalConTimeout {
    <#
        Ejecuta el bloque en un runspace propio con timeout duro. A diferencia de
        Start-Job, no serializa: los objetos vuelven vivos, con sus DateTime
        reales y sin las propiedades PSComputerName/RunspaceId/PSShowComputerName.
        Tampoco levanta un proceso nuevo, asi que es mucho mas rapido.
    #>
    param([scriptblock]$ScriptBlock, [object[]]$ArgumentList = @(), [int]$TimeoutSec = 120)

    $ps = $null
    $runspace = $null
    try {
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.ApartmentState = 'MTA'
        $runspace.ThreadOptions   = 'ReuseThread'
        $runspace.Open()

        $ps = [powershell]::Create()
        $ps.Runspace = $runspace
        # Se re-crea el scriptblock en el runspace destino: un scriptblock creado
        # en otra sesion arrastra su propio session state y falla al invocarse.
        [void]$ps.AddScript($ScriptBlock.ToString())
        foreach ($arg in $ArgumentList) { [void]$ps.AddArgument($arg) }

        $handle = $ps.BeginInvoke()
        if (-not $handle.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds($TimeoutSec))) {
            try { $ps.Stop() } catch {}
            Write-ReportLog -Message "Tiempo de espera agotado (${TimeoutSec}s) ejecutando un bloque de recoleccion." -Level Warn
            return "Tiempo de espera agotado (${TimeoutSec}s) al recolectar este dato."
        }

        $resultado = $ps.EndInvoke($handle)

        if ($ps.Streams.Error.Count -gt 0 -and ($null -eq $resultado -or $resultado.Count -eq 0)) {
            $msg = 'Error desconocido'
            try { $msg = $ps.Streams.Error[0].Exception.Message } catch {}
            return "Error al recolectar: $msg"
        }

        if ($null -eq $resultado) { return $null }
        # EndInvoke siempre devuelve una coleccion; se normaliza para que los
        # colectores reciban lo mismo que si el bloque se hubiera ejecutado en linea.
        $arr = @($resultado)
        if ($arr.Count -eq 0) { return $null }
        if ($arr.Count -eq 1) { return $arr[0] }
        return ,$arr
    } catch {
        Write-ReportLog -Message "Error en Invoke-Remote (local): $($_.Exception.Message)" -Level Warn
        return "Error al recolectar: $($_.Exception.Message)"
    } finally {
        if ($ps)       { try { $ps.Dispose() } catch {} }
        if ($runspace) { try { $runspace.Close(); $runspace.Dispose() } catch {} }
    }
}

function Invoke-RemoteViaJob {
    <# Ejecucion contra la PSSession remota, con timeout via Wait-Job. #>
    param([scriptblock]$ScriptBlock, [object[]]$ArgumentList = @(), [int]$TimeoutSec = 120)

    $job = $null
    try {
        $job = Invoke-Command -Session $script:Session -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList -AsJob

        $terminado = Wait-Job -Job $job -Timeout $TimeoutSec

        if (-not $terminado) {
            Write-ReportLog -Message "Tiempo de espera agotado (${TimeoutSec}s) ejecutando un bloque de recoleccion." -Level Warn
            Stop-Job -Job $job -ErrorAction SilentlyContinue | Out-Null
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null
            return "Tiempo de espera agotado (${TimeoutSec}s) al recolectar este dato."
        }

        if ($job.State -eq 'Failed') {
            $mensajeError = 'Error desconocido'
            try { $mensajeError = $job.ChildJobs[0].JobStateInfo.Reason.Message } catch { }
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null
            return "Error al recolectar: $mensajeError"
        }

        $resultado = Receive-Job -Job $job -ErrorAction Stop
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null
        return (Remove-PropiedadesDeRemoting $resultado)
    } catch {
        if ($job) { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null }
        Write-ReportLog -Message "Error en Invoke-Remote: $($_.Exception.Message)" -Level Warn
        return "Error al recolectar: $($_.Exception.Message)"
    }
}

function Get-NombreDeEnum {
    <#
        Devuelve el NOMBRE (texto) de un valor enum, o $null si $Value no es un
        enum. Cubre los dos casos:
          - enum real (ejecucion local): ToString() da el nombre.
          - enum que vuelve de una PSSession: llega deserializado, como un
            PSObject cuyo valor base es el entero (ToString() da '1', no
            'True'), con TypeNames 'Deserialized.System.Enum' y una
            ScriptProperty 'Value' que trae el nombre.
        Sin esta normalizacion un enum remoto se veia como '1'/'0' en las tablas,
        las reglas que comparan contra 'True'/'False'/'Allow' nunca coincidian,
        y ConvertTo-Json lo serializaba como {"value":1,"Value":"True"}: dos
        claves que solo difieren en mayusculas, que ni ConvertFrom-Json puede
        leer (dejaba inutilizable -BaselinePath).
    #>
    param($Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Enum]) { return $Value.ToString() }
    try {
        if ($Value -isnot [string] -and ($Value.PSObject.TypeNames -contains 'Deserialized.System.Enum')) {
            $propNombre = $Value.PSObject.Properties['Value']
            if ($propNombre -and "$($propNombre.Value)" -ne '') { return "$($propNombre.Value)" }
            return "$($Value.PSObject.BaseObject)"
        }
    } catch {}
    return $null
}

function Remove-PropiedadesDeRemoting {
    <#
        Quita las propiedades que agrega PowerShell Remoting a cada objeto
        (PSComputerName, RunspaceId, PSShowComputerName). Sin esto aparecen como
        tres columnas inutiles en todas las tablas del reporte.
    #>
    param($InputObject)

    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [string] -or $InputObject.GetType().IsPrimitive) { return $InputObject }

    $basura = @('PSComputerName', 'RunspaceId', 'PSShowComputerName', 'PSSourceJobInstanceId')
    $salida = New-Object System.Collections.ArrayList

    foreach ($item in @($InputObject)) {
        # Un hashtable / diccionario se devuelve tal cual. Si pasara por la
        # copia de propiedades de mas abajo, quedaria convertido en un PSObject
        # con las propiedades del CONTENEDOR (Count, Keys, Values, SyncRoot...)
        # y sin sus entradas: Get-TargetCapabilities recibia asi todas las
        # capacidades en $false en cualquier corrida remota.
        if ($item -is [System.Collections.IDictionary]) {
            [void]$salida.Add($item); continue
        }
        if ($item -is [ValueType]) {
            $nombreEnum = Get-NombreDeEnum $item
            if ($null -ne $nombreEnum) { [void]$salida.Add($nombreEnum); continue }
        }
        if ($null -eq $item -or $item -is [string] -or $item -is [ValueType]) {
            [void]$salida.Add($item); continue
        }
        try {
            $limpio = New-Object PSObject
            foreach ($prop in $item.PSObject.Properties) {
                if ($basura -contains $prop.Name) { continue }
                $valorProp = $prop.Value
                if ($null -ne $valorProp -and $valorProp -isnot [string]) {
                    $nombreEnum = Get-NombreDeEnum $valorProp
                    if ($null -ne $nombreEnum) { $valorProp = $nombreEnum }
                }
                Add-Member -InputObject $limpio -MemberType NoteProperty -Name $prop.Name -Value $valorProp -Force
            }
            [void]$salida.Add($limpio)
        } catch {
            [void]$salida.Add($item)
        }
    }

    if ($salida.Count -eq 0) { return $null }
    if ($salida.Count -eq 1) { return $salida[0] }
    return ,$salida.ToArray()
}

function Test-SectionPrefixMatch {
    <# Helper privado: compara un prefijo contra segmentos de Id, por SEGMENTO (no substring). #>
    param(
        [Parameter(Mandatory)][string]$Prefix,
        [Parameter(Mandatory)][string[]]$TargetSegments
    )

    $segmentosPrefijo = $Prefix.Split('.')
    if ($segmentosPrefijo.Count -gt $TargetSegments.Count) { return $false }

    for ($i = 0; $i -lt $segmentosPrefijo.Count; $i++) {
        if ($segmentosPrefijo[$i] -ne $TargetSegments[$i]) { return $false }
    }
    return $true
}

function Test-SectionEnabled {
    <#
        Devuelve $true si la seccion $Id debe recolectarse/mostrarse, segun
        -Sections / -SkipSections. Si -Sections esta vacio, todo esta habilitado
        salvo lo que matchee -SkipSections. La coincidencia es por segmento de
        punto: '1.11' habilita '1.11.3' pero '1.2' NO deshabilita '1.12'.
    #>
    param([Parameter(Mandatory)][string]$Id)

    $segmentos = $Id.Split('.')

    if ($script:SkipSections -and $script:SkipSections.Count -gt 0) {
        foreach ($skip in $script:SkipSections) {
            if (Test-SectionPrefixMatch -Prefix $skip -TargetSegments $segmentos) { return $false }
        }
    }

    if (-not $script:Sections -or $script:Sections.Count -eq 0) {
        return $true
    }

    foreach ($sec in $script:Sections) {
        if (Test-SectionPrefixMatch -Prefix $sec -TargetSegments $segmentos) { return $true }
    }
    return $false
}

function Protect-Sensitive {
    <#
        Si $script:Redact es falso, devuelve $Value sin modificar. Si es verdadero,
        enmascara segun $Kind (o lo detecta automaticamente con 'Auto'):
          IP         -> conserva los dos primeros octetos: 10.20.x.x
          Account    -> DOMINIO\u****
          Path       -> conserva la raiz, enmascara el resto
          Serial / Thumbprint -> primeros 4 + '...' + ultimos 4 caracteres
    #>
    param(
        [string]$Value,
        [ValidateSet('IP', 'Account', 'Path', 'Serial', 'Thumbprint', 'Auto')][string]$Kind = 'Auto'
    )

    if (-not $script:Redact) { return $Value }
    if ([string]::IsNullOrEmpty($Value)) { return $Value }

    $tipo = $Kind
    if ($tipo -eq 'Auto') {
        if ($Value -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
            $tipo = 'IP'
        } elseif ($Value -match '^[A-Za-z0-9_.-]+\\[A-Za-z0-9_.-]+$') {
            $tipo = 'Account'
        } elseif ($Value -match '^[A-Za-z]:\\' -or $Value -match '^\\\\') {
            $tipo = 'Path'
        } elseif ($Value -match '^[0-9A-Fa-f]{16,64}$') {
            $tipo = 'Thumbprint'
        } else {
            $tipo = 'Ninguno'
        }
    }

    switch ($tipo) {
        'IP' {
            if ($Value -match '^(\d{1,3})\.(\d{1,3})\.\d{1,3}\.\d{1,3}$') {
                return "$($matches[1]).$($matches[2]).x.x"
            }
            return $Value
        }
        'Account' {
            if ($Value -match '^([A-Za-z0-9_.-]+)\\([A-Za-z0-9_.-]+)$') {
                return "$($matches[1])\u****"
            }
            return $Value
        }
        'Path' {
            if ($Value -match '^([A-Za-z]:\\[^\\]*)') {
                return "$($matches[1])\...(enmascarado)"
            } elseif ($Value -match '^(\\\\[^\\]+\\[^\\]+)') {
                return "$($matches[1])\...(enmascarado)"
            }
            return $Value
        }
        { $_ -eq 'Serial' -or $_ -eq 'Thumbprint' } {
            if ($Value.Length -gt 8) {
                return "$($Value.Substring(0, 4))...$($Value.Substring($Value.Length - 4, 4))"
            }
            return $Value
        }
        default { return $Value }
    }
}

function Get-UsuarioActual {
    <# Usuario que ejecuta el reporte, con respaldo si la variable de entorno no esta. #>
    if ($env:USERNAME) { return $env:USERNAME }
    try { return [Environment]::UserName } catch { return 'desconocido' }
}

function Get-EquipoActual {
    <# Equipo desde el que se ejecuta el reporte (no el auditado), con respaldo. #>
    if ($env:COMPUTERNAME) { return $env:COMPUTERNAME }
    try { return [Environment]::MachineName } catch { return 'desconocido' }
}

function ConvertTo-SafeArray {
    <# Normaliza $InputObject a un arreglo, sin importar si es $null, escalar o arreglo. #>
    param($InputObject)

    if ($null -eq $InputObject) { return ,@() }
    if ($InputObject -is [string]) { return ,@($InputObject) }
    # OJO: @() sobre System.Collections.Generic.List[object] falla en algunos
    # runtimes ('Argument types do not match'). Enumeramos a mano, que siempre anda.
    if ($InputObject -is [System.Collections.IEnumerable]) {
        $acc = New-Object System.Collections.ArrayList
        foreach ($item in $InputObject) { [void]$acc.Add($item) }
        return ,$acc.ToArray()
    }
    return ,@($InputObject)
}

function Add-ReportSection {
    <#
        Registra una seccion del reporte para que el renderer la consuma. Ignora el
        llamado silenciosamente si Test-SectionEnabled($Id) es falso Y el Id tiene
        al menos 2 segmentos (las secciones de nivel 1, ej '1', nunca se filtran).
        Asi los colectores no necesitan chequear -Sections/-SkipSections por su cuenta.
    #>
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][string]$Title,
        $Data,
        [string]$Note,
        [string]$Icon,
        [switch]$Wide
    )

    $segmentos = $Id.Split('.')
    if ($segmentos.Count -ge 2 -and -not (Test-SectionEnabled -Id $Id)) {
        return
    }

    $esMensaje       = $false
    $cantidadFilas   = 0
    $datosNormalizados = $Data

    if ($null -eq $Data) {
        $datosNormalizados = @()
        $cantidadFilas = 0
    } elseif ($Data -is [string]) {
        $esMensaje = $true
        $cantidadFilas = 0
    } elseif ($Data -is [array]) {
        $datosNormalizados = $Data
        $cantidadFilas = $Data.Count
    } else {
        # Un solo PSCustomObject (o cualquier objeto suelto, incluyendo objetos
        # deserializados que vuelven de una PSSession) se envuelve en arreglo.
        $datosNormalizados = @($Data)
        $cantidadFilas = 1
    }

    $nivel = [Math]::Min(($segmentos.Count + 1), 6)

    $script:ReportSections.Add([PSCustomObject]@{
        Id        = $Id
        Title     = $Title
        Data      = $datosNormalizados
        Note      = $Note
        Icon      = $Icon
        Wide       = [bool]$Wide
        Level     = $nivel
        IsMessage = $esMensaje
        RowCount  = $cantidadFilas
    }) | Out-Null
}

function Add-Finding {
    <# Registra un hallazgo para el resumen ejecutivo (semaforo CRIT/WARN/INFO/OK). #>
    param(
        [Parameter(Mandatory)][ValidateSet('CRIT', 'WARN', 'INFO', 'OK')][string]$Severity,
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$Item,
        [string]$Detail,
        [string]$Recommendation,
        [string]$Control,
        [string]$SectionRef
    )

    $script:Findings.Add([PSCustomObject]@{
        Severity       = $Severity
        Category       = $Category
        Item           = $Item
        Detail         = $Detail
        Recommendation = $Recommendation
        Control        = $Control
        SectionRef     = $SectionRef
        Timestamp      = Get-Date
    }) | Out-Null
}

#endregion CORE-Utilidades

# =============================================================================
# REGION: CORE - Deteccion de capacidades del equipo objetivo
# =============================================================================
#region CORE-Capacidades

function Get-TargetCapabilities {
    <#
        Llena $script:Caps con todas las claves del contrato compartido, en UNA sola
        llamada a Invoke-Remote (un unico round-trip contra el equipo objetivo, en
        vez de una llamada remota por cada capacidad). Si la llamada falla o expira,
        $script:Caps queda con valores por defecto seguros (todo $false) para que el
        resto del script pueda seguir funcionando sin romperse.
    #>

    $valoresPorDefecto = @{
        IIS = $false; DHCP = $false; DNS = $false; FileServer = $false; RDSH = $false
        RDSLicensing = $false; WSUS = $false; Python = $false; HyperV = $false; SQL = $false
        Cluster = $false; DomainMember = $false; IsDC = $false; IsServerOS = $false
        IsVM = $false; BitLocker = $false; Docker = $false
        OSCaption = 'Desconocido'; OSBuild = 'Desconocido'; PSVersionRemota = 'Desconocido'
    }

    $crudo = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        $resultado = @{}
        $servicios = Get-Service -ErrorAction SilentlyContinue

        $resultado.IIS          = [bool]($servicios | Where-Object { $_.Name -eq 'W3SVC' })
        $resultado.DHCP         = [bool]($servicios | Where-Object { $_.Name -eq 'DHCPServer' })
        $resultado.DNS          = [bool]($servicios | Where-Object { $_.Name -eq 'DNS' })
        $resultado.RDSH         = [bool]($servicios | Where-Object { $_.Name -eq 'TermService' -and $_.Status -eq 'Running' })
        $resultado.RDSLicensing = [bool]($servicios | Where-Object { $_.Name -eq 'TermServLicensing' })
        $resultado.WSUS         = [bool]($servicios | Where-Object { $_.Name -eq 'WsusService' })
        $resultado.HyperV       = [bool]($servicios | Where-Object { $_.Name -eq 'vmms' })
        $resultado.SQL          = [bool]($servicios | Where-Object { $_.Name -like 'MSSQL*' })
        $resultado.Cluster      = [bool]($servicios | Where-Object { $_.Name -eq 'ClusSvc' })
        $resultado.Docker       = [bool]($servicios | Where-Object { $_.Name -eq 'docker' -or $_.Name -eq 'com.docker.service' })

        try {
            $shares = Get-SmbShare -ErrorAction Stop | Where-Object { $_.Name -notmatch '^[A-Z]\$$|^ADMIN\$$|^IPC\$$|^PRINT\$$' }
            $resultado.FileServer = [bool]$shares
        } catch {
            $resultado.FileServer = $false
        }

        try {
            $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
            $resultado.DomainMember = [bool]$cs.PartOfDomain
            $fabricanteModelo = "$($cs.Manufacturer) $($cs.Model)"
            $resultado.IsVM = [bool]($fabricanteModelo -match 'VMware|Virtual|KVM|Xen|QEMU|Hyper-V|Parallels|Amazon|Google')
        } catch {
            $resultado.DomainMember = $false
            $resultado.IsVM = $false
        }

        try {
            $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
            $resultado.IsDC        = [bool]($os.ProductType -eq 2)
            $resultado.IsServerOS  = [bool]($os.ProductType -ne 1)
            $resultado.OSCaption   = $os.Caption
            $resultado.OSBuild     = $os.BuildNumber
        } catch {
            $resultado.IsDC       = $false
            $resultado.IsServerOS = $false
            $resultado.OSCaption  = 'Desconocido'
            $resultado.OSBuild    = 'Desconocido'
        }

        try {
            $resultado.BitLocker = [bool](Get-Module -ListAvailable -Name BitLocker -ErrorAction Stop)
        } catch {
            $resultado.BitLocker = $false
        }

        # Python se resuelve mas adelante (seccion 1.15), aca queda en $false.
        $resultado.Python = $false
        $resultado.PSVersionRemota = $PSVersionTable.PSVersion.ToString()

        $resultado
    }

    if ($crudo -is [string]) {
        Write-ReportLog -Message "No se pudieron detectar capacidades del equipo objetivo: $crudo" -Level Warn
        $script:Caps = $valoresPorDefecto
        return
    }

    $combinado = $valoresPorDefecto.Clone()

    if ($crudo -is [hashtable]) {
        foreach ($clave in $crudo.Keys) { $combinado[$clave] = $crudo[$clave] }
    } elseif ($crudo -and $crudo.PSObject -and $crudo.PSObject.Properties) {
        # Ruta de respaldo por si la deserializacion remota devuelve un PSCustomObject
        # en lugar de un Hashtable nativo.
        foreach ($propiedad in $crudo.PSObject.Properties) { $combinado[$propiedad.Name] = $propiedad.Value }
    } else {
        Write-ReportLog -Message 'Respuesta de deteccion de capacidades con formato inesperado, se usan valores por defecto.' -Level Warn
    }

    $script:Caps = $combinado
}

#endregion CORE-Capacidades

# =============================================================================
# BANNER DE ARRANQUE
# =============================================================================
#region Banner

$script:IsAdmin = Test-Administrator

Write-ReportLog -Message "Get-ServerFullReport v$($script:ScriptVersion) iniciando..." -Level Info
Write-ReportLog -Message "Equipo(s) objetivo: $($ComputerName -join ', ')" -Level Info
Write-ReportLog -Message "Formatos de salida solicitados: $($Format -join ', ')" -Level Info
Write-ReportLog -Message "Carpeta de salida: $OutputPath" -Level Info
Write-ReportLog -Message "Ejecutando como Administrador: $($script:IsAdmin)" -Level Info

if (-not $script:IsAdmin) {
    Write-ReportLog -Message ('El script NO se esta ejecutando con privilegios de Administrador. ' +
        'Las secciones 1.11 (Seguridad y Cumplimiento), 1.14 (Cuentas y Accesos), parte de 1.5 (Storage) ' +
        'y algunas consultas CIM/WMI pueden devolver datos incompletos o errores de acceso denegado. ' +
        'Se recomienda re-ejecutar como Administrador (local) o con -Credential de un administrador (remoto) ' +
        'para un inventario completo.') -Level Warn
}

#endregion Banner

# #############################################################################
# ##  MOTOR DE RENDERIZADO HTML
# #############################################################################

# =============================================================================
# FRAGMENTO: Motor de renderizado HTML (frag_05_render.ps1)
# Consume $script:ReportSections, $script:Findings, $script:FindingsSummary,
# $script:Caps, $script:Timings, $script:LogLines, $script:ScriptVersion y
# $script:TargetName (provistos por CORE y los colectores) y produce el
# documento HTML final. No modifica el equipo auditado: solo genera texto.
# =============================================================================

# (Requires declarado en el encabezado del script)
# =============================================================================
# REGION: CSS del reporte (hoja de estilos completa, reutilizable)
# =============================================================================
#region Render-CSS

$script:ReportCss = @'
:root{
  --bg:#f4f6f9; --surface:#ffffff; --surface-2:#eef2f7; --surface-3:#e4eaf2;
  --border:#d8e0ea; --border-strong:#b9c6d6; --text:#131b26; --text-dim:#5a6a7d;
  --accent:#0f4c81; --accent-hover:#0c3d68; --accent-soft:#e7eff8;
  --shadow:0 1px 2px rgba(16,32,52,.06),0 4px 12px rgba(16,32,52,.05);

  --sev-crit:#a4231b; --sev-crit-bg:#fbe4e2;
  --sev-warn:#7a4c00; --sev-warn-bg:#fbead0;
  --sev-info:#0d3f68; --sev-info-bg:#e3edf9;
  --sev-ok:#146c2c;   --sev-ok-bg:#e2f3e6;
  --sev-neutral:#4c5a6b; --sev-neutral-bg:#e9edf3;

  --meter-track:#e4eaf2; --meter-ok:#1e9e4a; --meter-warn:#c98a00; --meter-crit:#cf3428;

  --font-sans:-apple-system,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;
  --font-mono:"Cascadia Mono","SF Mono",Consolas,"Liberation Mono",monospace;

  --radius-card:8px; --radius-control:5px; --radius-badge:999px;
  color-scheme: light;
}

@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --bg:#0d1117; --surface:#151b23; --surface-2:#1b222c; --surface-3:#222b36;
    --border:#2a333f; --border-strong:#3a4553; --text:#e6edf3; --text-dim:#93a1b1;
    --accent:#4d9fe8; --accent-hover:#6fb3ee; --accent-soft:#12283d;
    --shadow:0 1px 2px rgba(0,0,0,.4),0 4px 14px rgba(0,0,0,.3);

    --sev-crit:#ff9d93; --sev-crit-bg:#3a1613;
    --sev-warn:#ffcf87; --sev-warn-bg:#3a2a10;
    --sev-info:#8dcbf7; --sev-info-bg:#112a40;
    --sev-ok:#8fe6ab;   --sev-ok-bg:#0f3320;
    --sev-neutral:#b7c2d0; --sev-neutral-bg:#232c38;

    --meter-track:#2a333f; --meter-ok:#3ecb6c; --meter-warn:#e0aa3e; --meter-crit:#ec5a4d;
    color-scheme: dark;
  }
}

:root[data-theme="dark"]{
  --bg:#0d1117; --surface:#151b23; --surface-2:#1b222c; --surface-3:#222b36;
  --border:#2a333f; --border-strong:#3a4553; --text:#e6edf3; --text-dim:#93a1b1;
  --accent:#4d9fe8; --accent-hover:#6fb3ee; --accent-soft:#12283d;
  --shadow:0 1px 2px rgba(0,0,0,.4),0 4px 14px rgba(0,0,0,.3);

  --sev-crit:#ff9d93; --sev-crit-bg:#3a1613;
  --sev-warn:#ffcf87; --sev-warn-bg:#3a2a10;
  --sev-info:#8dcbf7; --sev-info-bg:#112a40;
  --sev-ok:#8fe6ab;   --sev-ok-bg:#0f3320;
  --sev-neutral:#b7c2d0; --sev-neutral-bg:#232c38;

  --meter-track:#2a333f; --meter-ok:#3ecb6c; --meter-warn:#e0aa3e; --meter-crit:#ec5a4d;
  color-scheme: dark;
}

*{box-sizing:border-box;}
html,body{margin:0;padding:0;}
body{
  background:var(--bg); color:var(--text); font-family:var(--font-sans);
  font-size:13.5px; line-height:1.55; -webkit-font-smoothing:antialiased;
}
.sr-only{
  position:absolute; width:1px; height:1px; padding:0; margin:-1px; overflow:hidden;
  clip:rect(0,0,0,0); white-space:nowrap; border:0;
}
a{color:var(--accent);}
a:hover{color:var(--accent-hover);}
:focus-visible{outline:2px solid var(--accent); outline-offset:2px;}

h1{font-size:26px;font-weight:600;margin:0 0 6px;}
h2{font-size:19px;font-weight:600;margin:0 0 10px;}
h3{font-size:15.5px;font-weight:600;margin:0 0 8px;}
h4{
  font-size:13.5px;font-weight:600;margin:0 0 8px;
  text-transform:uppercase; letter-spacing:.06em; color:var(--text-dim);
}

/* ---------- layout ---------- */
body{display:flex; align-items:stretch; min-height:100vh;}
.sidebar{
  width:288px; min-width:288px; position:sticky; top:0; height:100vh;
  overflow-y:auto; background:var(--surface); border-right:1px solid var(--border);
  padding:18px 14px 40px;
}
.sidebar-host{
  font-size:17px; font-weight:700; margin-bottom:12px; word-break:break-word;
}
.toc-filter-input{
  width:100%; padding:7px 10px; margin-bottom:12px; border:1px solid var(--border);
  border-radius:var(--radius-control); background:var(--surface-2); color:var(--text);
  font-family:var(--font-sans); font-size:12.5px;
}
.toc-list{list-style:none; margin:0; padding-left:0;}
.toc-list[data-depth="0"]{padding-left:0;}
.toc-list:not([data-depth="0"]){padding-left:12px; border-left:1px solid var(--border);}
.toc-item{margin:1px 0;}
.toc-link{
  display:flex; align-items:center; gap:6px; padding:5px 8px; border-radius:var(--radius-control);
  color:var(--text); text-decoration:none; font-size:12.5px; line-height:1.35;
}
.toc-link:hover{background:var(--surface-2);}
.toc-link.active{background:var(--accent-soft); color:var(--accent); font-weight:600;}
.toc-link-summary{font-weight:600; margin-bottom:8px; border-bottom:1px dashed var(--border); padding-bottom:8px;}
.toc-id{color:var(--text-dim); font-variant-numeric:tabular-nums; flex-shrink:0;}
.toc-count{
  margin-left:auto; font-size:10.5px; background:var(--surface-3); color:var(--text-dim);
  border-radius:var(--radius-badge); padding:1px 6px;
}

.main-col{flex:1; min-width:0; display:flex; flex-direction:column;}
.content{max-width:1180px; padding:28px 36px; width:100%; margin:0 auto;}

.topbar{
  position:sticky; top:0; z-index:20; display:flex; align-items:center; justify-content:space-between;
  gap:16px; padding:10px 28px; background:color-mix(in srgb, var(--surface) 85%, transparent);
  background:var(--surface); backdrop-filter:blur(8px); -webkit-backdrop-filter:blur(8px);
  border-bottom:1px solid var(--border);
}
.topbar-id{display:flex; align-items:center; gap:10px; min-width:0; flex-wrap:wrap;}
.topbar-host{font-weight:700; font-size:14px;}
.topbar-date{color:var(--text-dim); font-size:12px;}
.topbar-score{white-space:nowrap;}
.topbar-actions{display:flex; align-items:center; gap:8px;}
.icon-btn{
  border:1px solid var(--border); background:var(--surface-2); color:var(--text);
  border-radius:var(--radius-control); width:32px; height:32px; cursor:pointer; font-size:15px;
  display:flex; align-items:center; justify-content:center;
}
.icon-btn:hover{background:var(--surface-3);}

.search-wrap{position:relative;}
.global-search-input{
  width:220px; padding:7px 10px; border:1px solid var(--border); border-radius:var(--radius-control);
  background:var(--surface-2); color:var(--text); font-family:var(--font-sans); font-size:12.5px;
}
.search-panel{
  position:absolute; top:calc(100% + 6px); right:0; width:340px; max-height:420px; overflow-y:auto;
  background:var(--surface); border:1px solid var(--border-strong); border-radius:var(--radius-card);
  box-shadow:var(--shadow); z-index:30; padding:6px;
}
.search-result{
  display:flex; justify-content:space-between; gap:10px; width:100%; text-align:left; padding:8px 10px;
  border:none; background:transparent; color:var(--text); border-radius:var(--radius-control); cursor:pointer;
  font-size:12.5px;
}
.search-result:hover{background:var(--surface-2);}
.search-result-count{
  background:var(--surface-3); color:var(--text-dim); border-radius:var(--radius-badge); padding:0 7px; font-size:11px;
}
.search-empty{padding:10px; color:var(--text-dim); font-size:12.5px;}

/* ---------- hero / portada ---------- */
.hero{padding:8px 0 26px; border-bottom:1px solid var(--border); margin-bottom:26px;}
.hero-subtitle{color:var(--text-dim); margin:0 0 14px; font-size:14px;}
.role-chips{display:flex; flex-wrap:wrap; gap:8px; margin-bottom:20px;}
.role-chip{
  background:var(--accent-soft); color:var(--accent); border-radius:var(--radius-badge);
  padding:5px 12px; font-size:12px; font-weight:600;
}
.role-chip-muted{background:var(--surface-2); color:var(--text-dim); font-weight:400;}

.kpi-grid{
  display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:14px;
}
.kpi-card{
  background:var(--surface); border:1px solid var(--border); border-radius:var(--radius-card);
  padding:14px 16px; box-shadow:var(--shadow);
}
.kpi-label{
  font-size:11px; text-transform:uppercase; letter-spacing:.06em; color:var(--text-dim); margin-bottom:6px;
}
.kpi-value{
  font-size:26px; font-weight:600; font-variant-numeric:tabular-nums; line-height:1.1;
}
.kpi-unit{font-size:13px; font-weight:400; color:var(--text-dim); margin-left:2px;}
.kpi-context{font-size:11.5px; color:var(--text-dim); margin-top:6px;}
.kpi-score .kpi-value{font-size:30px;}
.kpi-crit .kpi-value{color:var(--sev-crit);}
.kpi-warn .kpi-value{color:var(--sev-warn);}

/* ---------- meters ---------- */
.meter-track{
  width:100%; height:7px; border-radius:999px; background:var(--meter-track); overflow:hidden; margin-top:8px;
}
.meter-fill{height:100%; border-radius:999px; transition:width .25s ease;}
.meter-ok{background:var(--meter-ok);}
.meter-warn{background:var(--meter-warn);}
.meter-crit{background:var(--meter-crit);}
.score-track{margin-top:10px; height:8px;}

.cell-meter{display:inline-flex; align-items:center; gap:8px; min-width:110px;}
.cell-meter .meter-track{width:60px; margin-top:0; flex-shrink:0;}
.meter-label{font-variant-numeric:tabular-nums; font-size:12px;}

/* ---------- badges ---------- */
.badge{
  display:inline-block; padding:2px 9px; border-radius:var(--radius-badge); font-size:11.5px;
  font-weight:600; line-height:1.6; white-space:nowrap;
}
.badge-crit{color:var(--sev-crit); background:var(--sev-crit-bg);}
.badge-warn{color:var(--sev-warn); background:var(--sev-warn-bg);}
.badge-info{color:var(--sev-info); background:var(--sev-info-bg);}
.badge-ok{color:var(--sev-ok); background:var(--sev-ok-bg);}
.badge-neutral{color:var(--sev-neutral); background:var(--sev-neutral-bg);}
.count-badge{margin-left:8px; font-weight:400;}

/* ---------- cards / sections ---------- */
.card{
  background:var(--surface); border:1px solid var(--border); border-radius:var(--radius-card);
  padding:22px 24px; margin-bottom:20px; box-shadow:var(--shadow);
}
.subsection{margin:16px 0 16px 2px; padding-left:14px; border-left:2px solid var(--border);}
.section-note{color:var(--text-dim); font-size:12.5px; margin:-4px 0 12px;}

.msg{
  display:flex; gap:10px; align-items:flex-start; background:var(--surface-2); border:1px solid var(--border);
  border-radius:var(--radius-card); padding:12px 14px; font-style:normal; color:var(--text-dim);
}
.msg p{margin:0;}
.msg-icon{flex-shrink:0;}

.empty-state{
  text-align:center; padding:26px 14px; color:var(--text-dim); background:var(--surface-2);
  border:1px dashed var(--border-strong); border-radius:var(--radius-card);
}
.empty-state p{margin:6px 0 0;}
.empty-icon{font-size:22px;}
.empty-state-positive{
  background:var(--sev-ok-bg); border-style:solid; border-color:transparent; color:var(--sev-ok);
  text-align:left; display:flex; gap:12px; align-items:center; padding:16px 18px;
}
.empty-state-positive .empty-icon{font-size:24px;}
.empty-state-positive p{margin:0; color:var(--text);}

/* ---------- details / summary ---------- */
details.subsection-details{
  border:1px solid var(--border); border-radius:var(--radius-card); margin:10px 0; overflow:hidden;
  background:var(--surface);
}
details.subsection-details > summary{
  cursor:pointer; padding:10px 14px; font-weight:600; list-style:none; display:flex; align-items:center;
  gap:6px; background:var(--surface-2); user-select:none;
}
details.subsection-details > summary::-webkit-details-marker{display:none;}
details.subsection-details > summary::before{
  content:"\25B8"; display:inline-block; transition:transform .15s ease; color:var(--text-dim);
}
details.subsection-details[open] > summary::before{transform:rotate(90deg);}
.details-body{padding:14px;}

.footer-details{
  border:1px solid var(--border); border-radius:var(--radius-card); margin-top:10px; background:var(--surface);
}
.footer-details > summary{cursor:pointer; padding:9px 14px; font-weight:600; background:var(--surface-2);}
.footer-details > summary::-webkit-details-marker{display:none;}
.footer-details .details-body{padding:12px 14px;}

/* ---------- tables ---------- */
.table-filter-bar{display:flex; align-items:center; gap:10px; margin-bottom:8px;}
.table-filter-input{
  flex:1; max-width:320px; padding:6px 10px; border:1px solid var(--border); border-radius:var(--radius-control);
  background:var(--surface-2); color:var(--text); font-family:var(--font-sans); font-size:12.5px;
}
.table-filter-count{font-size:11.5px; color:var(--text-dim); white-space:nowrap;}

.table-wrap{overflow-x:auto; border:1px solid var(--border); border-radius:var(--radius-card); max-width:100%;}
.table-wrap.wide{max-width:100%;}
table.data-table{
  border-collapse:collapse; width:100%; font-size:12.5px; min-width:480px;
}
table.data-table thead th{
  position:sticky; top:0; background:var(--surface-3); text-align:left; padding:9px 12px;
  border-bottom:1px solid var(--border-strong); white-space:nowrap; cursor:pointer; user-select:none;
  font-weight:600; z-index:1;
}
table.data-table thead th:hover{background:var(--surface-2);}
.th-label{margin-right:4px;}
.sort-ind{display:inline-block; width:10px; color:var(--text-dim); font-size:10px;}
table.data-table thead th[aria-sort="ascending"] .sort-ind::after{content:"\25B2";}
table.data-table thead th[aria-sort="descending"] .sort-ind::after{content:"\25BC";}
table.data-table thead th[aria-sort="none"] .sort-ind::after{content:"\00b7";}
table.data-table tbody td{
  padding:8px 12px; border-bottom:1px solid var(--border); vertical-align:top;
}
table.data-table tbody tr:nth-child(even){background:var(--surface-2);}
table.data-table tbody tr:hover{background:var(--accent-soft);}
table.data-table tbody tr[hidden]{display:none;}
table.data-table tbody tr.search-hit{outline:2px solid var(--accent); outline-offset:-2px;}

.mono{font-family:var(--font-mono); font-size:12px;}
.dim{color:var(--text-dim);}
.truncate{cursor:pointer; border-bottom:1px dotted var(--border-strong);}
.truncate.is-expanded{white-space:normal; word-break:break-word; display:inline-block;}

.findings-card .table-wrap{max-width:100%;}
.findings-filterbar{display:flex; flex-wrap:wrap; gap:8px; margin:14px 0 16px;}
.chip-filter{
  border:1px solid var(--border); background:var(--surface-2); color:var(--text); border-radius:var(--radius-badge);
  padding:6px 14px; font-size:12.5px; cursor:pointer; display:flex; align-items:center; gap:6px;
}
.chip-filter.active{background:var(--accent); color:#fff; border-color:var(--accent);}
.chip-filter.sev-crit.active{background:var(--sev-crit); border-color:var(--sev-crit);}
.chip-filter.sev-warn.active{background:var(--sev-warn); border-color:var(--sev-warn);}
.chip-filter.sev-info.active{background:var(--sev-info); border-color:var(--sev-info);}
.chip-filter.sev-ok.active{background:var(--sev-ok); border-color:var(--sev-ok);}
.chip-count{background:rgba(0,0,0,.12); border-radius:999px; padding:0 6px; font-size:11px;}
.ver-link{white-space:nowrap; font-size:12px;}

/* ---------- footer ---------- */
.report-footer{margin-top:30px; padding-top:16px; border-top:1px solid var(--border); color:var(--text-dim); font-size:12px;}
.footer-meta{margin:0 0 10px;}
.log-block{
  font-family:var(--font-mono); font-size:11px; white-space:pre-wrap; word-break:break-word; margin:0;
  max-height:360px; overflow-y:auto;
}

/* ---------- print header (repeats via position:fixed while printing) ---------- */
.print-header{display:none;}

/* ---------- reduced motion ---------- */
@media (prefers-reduced-motion: reduce){
  *{transition:none !important; animation:none !important; scroll-behavior:auto !important;}
}

/* ---------- print ---------- */
@media print{
  .sidebar, .topbar, .search-panel, .toc-filter-input, .table-filter-bar, .findings-filterbar,
  .icon-btn{display:none !important;}
  .main-col{display:block;}
  .content{max-width:100%; padding:14px 18px 0; margin:0;}
  body{background:#fff; color:#000;}
  *{-webkit-print-color-adjust:exact; print-color-adjust:exact;}
  .print-header{
    display:flex !important; justify-content:space-between; position:fixed; top:0; left:0; right:0;
    font-size:9px; color:#555; border-bottom:1px solid #ccc; padding:4px 10px; background:#fff;
  }
  details, details > *{display:block !important;}
  details > summary::before{display:none;}
  .card, tr{page-break-inside:avoid;}
  .card{page-break-before:always;}
  .card:first-of-type{page-break-before:avoid;}
  table.data-table{font-size:10px;}
  a[href^="#"]::after{content:"";}
}
'@

#endregion Render-CSS

# =============================================================================
# REGION: JavaScript del reporte (script completo, reutilizable)
# =============================================================================
#region Render-JS

$script:ReportJs = @'
(function(){
  "use strict";

  function safeRun(fn){ try { fn(); } catch (e) { /* el reporte sigue siendo legible */ } }

  function prefersReducedMotion(){
    try { return !!(window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches); }
    catch (e) { return false; }
  }

  function scrollToId(id){
    var el = document.getElementById(id);
    if (!el) return;
    if (el.tagName && el.tagName.toLowerCase() === "details") { el.open = true; }
    if (el.querySelectorAll) {
      var inner = el.querySelectorAll("details");
      for (var i = 0; i < inner.length; i++) { inner[i].open = true; }
    }
    try {
      el.scrollIntoView({ behavior: prefersReducedMotion() ? "auto" : "smooth", block: "start" });
    } catch (e) {
      el.scrollIntoView();
    }
  }

  function initAnchorOpen(){
    var anchors = document.querySelectorAll('a[href^="#"]');
    for (var i = 0; i < anchors.length; i++) {
      (function(a){
        a.addEventListener("click", function(ev){
          var id = a.getAttribute("href").slice(1);
          if (!id) return;
          var el = document.getElementById(id);
          if (!el) return;
          ev.preventDefault();
          scrollToId(id);
          history.replaceState ? history.replaceState(null, "", "#" + id) : (window.location.hash = id);
        });
      })(anchors[i]);
    }
  }

  function initThemeToggle(){
    var btn = document.getElementById("theme-toggle");
    if (!btn) return;
    btn.addEventListener("click", function(){
      var current = "system";
      try { current = localStorage.getItem("asbuilt-theme") || "system"; } catch (e) {}
      var next = "dark";
      if (current === "light") { next = "dark"; }
      else if (current === "dark") { next = "system"; }
      else { next = "light"; }

      if (next === "system") { document.documentElement.removeAttribute("data-theme"); }
      else { document.documentElement.setAttribute("data-theme", next); }

      try { localStorage.setItem("asbuilt-theme", next); } catch (e) {}
    });
  }

  function initPrintButton(){
    var btn = document.getElementById("print-btn");
    if (!btn) return;
    btn.addEventListener("click", function(){ window.print(); });
  }

  function parseSortValue(cellText, type){
    var t = (cellText || "").replace(/\s+/g, " ").trim();
    if (type === "num") {
      var cleaned = t.replace(/,/g, "");
      var m = cleaned.match(/-?[0-9]+(\.[0-9]+)?/);
      return m ? parseFloat(m[0]) : -Infinity;
    }
    if (type === "date") {
      var iso = Date.parse(t);
      if (!isNaN(iso)) { return iso; }
      var m2 = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m2) { return new Date(parseInt(m2[3],10), parseInt(m2[2],10)-1, parseInt(m2[1],10)).getTime(); }
      return -Infinity;
    }
    if (type === "bool") {
      var tl = t.toLowerCase();
      return (tl === "true" || tl === "si") ? 1 : 0;
    }
    return t.toLowerCase();
  }

  function sortTableBy(table, th, allThs){
    var colIndex = parseInt(th.getAttribute("data-col"), 10);
    var type = th.getAttribute("data-type") || "text";
    var current = th.getAttribute("aria-sort");
    var dir = "ascending";
    if (current === "ascending") { dir = "descending"; }

    for (var i = 0; i < allThs.length; i++) { allThs[i].setAttribute("aria-sort", "none"); }
    th.setAttribute("aria-sort", dir);

    var tbody = table.querySelector("tbody");
    if (!tbody) return;
    var rows = Array.prototype.slice.call(tbody.querySelectorAll("tr"));
    var mult = dir === "ascending" ? 1 : -1;

    rows.sort(function(a, b){
      var ca = a.children[colIndex] ? a.children[colIndex].textContent : "";
      var cb = b.children[colIndex] ? b.children[colIndex].textContent : "";
      var va = parseSortValue(ca, type);
      var vb = parseSortValue(cb, type);
      if (va < vb) return -1 * mult;
      if (va > vb) return 1 * mult;
      return 0;
    });

    for (var j = 0; j < rows.length; j++) { tbody.appendChild(rows[j]); }
  }

  function initSortableTables(root){
    var tables = root.querySelectorAll("table.data-table");
    for (var i = 0; i < tables.length; i++) {
      (function(table){
        var ths = table.querySelectorAll("thead th[data-col]");
        for (var k = 0; k < ths.length; k++) {
          (function(th){
            th.addEventListener("click", function(){ sortTableBy(table, th, ths); });
            th.addEventListener("keydown", function(ev){
              if (ev.key === "Enter" || ev.key === " ") { ev.preventDefault(); sortTableBy(table, th, ths); }
            });
          })(ths[k]);
        }
      })(tables[i]);
    }
  }

  function initTableFilters(root){
    var inputs = root.querySelectorAll(".table-filter-input");
    for (var i = 0; i < inputs.length; i++) {
      (function(inp){
        inp.addEventListener("input", function(){
          var targetId = inp.getAttribute("data-filter-for");
          var table = document.getElementById(targetId);
          if (!table) return;
          var q = inp.value.trim().toLowerCase();
          var rows = table.querySelectorAll("tbody tr");
          var shown = 0;
          for (var r = 0; r < rows.length; r++) {
            var text = (rows[r].textContent || "").toLowerCase();
            var match = q === "" || text.indexOf(q) !== -1;
            rows[r].hidden = !match;
            if (match) shown++;
          }
          var counter = document.querySelector('.table-filter-count[data-count-for="' + targetId + '"]');
          if (counter) { counter.textContent = "mostrando " + shown + " de " + rows.length; }
        });
      })(inputs[i]);
    }
  }

  function initFindingsFilter(root){
    var bar = root.querySelector(".findings-filterbar");
    if (!bar) return;
    var buttons = bar.querySelectorAll(".chip-filter");
    var table = document.getElementById("tbl-hallazgos");
    for (var i = 0; i < buttons.length; i++) {
      (function(btn){
        btn.addEventListener("click", function(){
          for (var b = 0; b < buttons.length; b++) {
            buttons[b].classList.remove("active");
            buttons[b].setAttribute("aria-pressed", "false");
          }
          btn.classList.add("active");
          btn.setAttribute("aria-pressed", "true");
          var sev = btn.getAttribute("data-sev");
          if (!table) return;
          var rows = table.querySelectorAll("tbody tr");
          for (var r = 0; r < rows.length; r++) {
            var rsev = rows[r].getAttribute("data-sev");
            rows[r].hidden = (sev !== "ALL" && rsev !== sev);
          }
        });
      })(buttons[i]);
    }
  }

  function initTocFilter(){
    var input = document.getElementById("toc-filter");
    if (!input) return;
    input.addEventListener("input", function(){
      var q = input.value.trim().toLowerCase();
      var items = document.querySelectorAll(".toc-item");
      for (var i = 0; i < items.length; i++) {
        var li = items[i];
        var text = li.getAttribute("data-toc-text") || "";
        var selfMatch = q === "" || text.indexOf(q) !== -1;
        var childMatch = false;
        if (!selfMatch) {
          var childItems = li.querySelectorAll(".toc-item");
          for (var c = 0; c < childItems.length; c++) {
            var ctext = childItems[c].getAttribute("data-toc-text") || "";
            if (q === "" || ctext.indexOf(q) !== -1) { childMatch = true; break; }
          }
        }
        li.hidden = !(selfMatch || childMatch);
      }
    });
  }

  function initScrollspy(){
    if (typeof IntersectionObserver === "undefined") return;
    var targets = document.querySelectorAll('main.content [id^="sec-"]');
    var links = {};
    var tocLinks = document.querySelectorAll(".toc-link");
    for (var i = 0; i < tocLinks.length; i++) {
      var target = tocLinks[i].getAttribute("data-target");
      if (target) { links[target] = tocLinks[i]; }
    }
    var observer = new IntersectionObserver(function(entries){
      for (var e = 0; e < entries.length; e++) {
        var entry = entries[e];
        var link = links[entry.target.id];
        if (!link) continue;
        if (entry.isIntersecting) {
          var active = document.querySelectorAll(".toc-link.active");
          for (var a = 0; a < active.length; a++) { active[a].classList.remove("active"); }
          link.classList.add("active");
        }
      }
    }, { rootMargin: "-15% 0px -70% 0px", threshold: 0 });
    for (var t = 0; t < targets.length; t++) { observer.observe(targets[t]); }
  }

  var searchIndex = [];

  function buildSearchIndex(){
    var nodes = document.querySelectorAll('main.content [id^="sec-"]');
    for (var i = 0; i < nodes.length; i++) {
      var n = nodes[i];
      var headingEl = n.querySelector("h2, h3, h4");
      var title = headingEl ? headingEl.textContent.trim() : n.id;
      searchIndex.push({ id: n.id, title: title, text: (n.textContent || "").toLowerCase() });
    }
  }

  function clearSearchHighlights(){
    var hits = document.querySelectorAll(".data-table tbody tr.search-hit");
    for (var i = 0; i < hits.length; i++) { hits[i].classList.remove("search-hit"); }
  }

  function initGlobalSearch(){
    var input = document.getElementById("global-search");
    var panel = document.getElementById("search-panel");
    if (!input || !panel) return;

    input.addEventListener("input", function(){
      var q = input.value.trim().toLowerCase();
      clearSearchHighlights();

      if (q.length < 2) { panel.hidden = true; panel.innerHTML = ""; return; }

      var results = [];
      for (var i = 0; i < searchIndex.length; i++) {
        var item = searchIndex[i];
        if (item.text.indexOf(q) === -1) continue;
        var count = item.text.split(q).length - 1;
        results.push({ id: item.id, title: item.title, count: count });
      }
      results.sort(function(a,b){ return b.count - a.count; });

      if (results.length === 0) {
        panel.innerHTML = '<div class="search-empty">Sin coincidencias.</div>';
      } else {
        panel.innerHTML = "";
        var top = results.slice(0, 25);
        for (var r = 0; r < top.length; r++) {
          var btn = document.createElement("button");
          btn.type = "button";
          btn.className = "search-result";
          btn.setAttribute("data-target", top[r].id);
          var titleSpan = document.createElement("span");
          titleSpan.className = "search-result-title";
          titleSpan.textContent = top[r].title;
          var countSpan = document.createElement("span");
          countSpan.className = "search-result-count";
          countSpan.textContent = String(top[r].count);
          btn.appendChild(titleSpan);
          btn.appendChild(countSpan);
          panel.appendChild(btn);
        }
      }
      panel.hidden = false;

      for (var k = 0; k < results.length; k++) {
        var sec = document.getElementById(results[k].id);
        if (!sec) continue;
        var rows = sec.querySelectorAll("tbody tr");
        for (var rr = 0; rr < rows.length; rr++) {
          if ((rows[rr].textContent || "").toLowerCase().indexOf(q) !== -1) {
            rows[rr].classList.add("search-hit");
          }
        }
      }
    });

    panel.addEventListener("click", function(ev){
      var btn = ev.target && ev.target.closest ? ev.target.closest(".search-result") : null;
      if (!btn) return;
      var id = btn.getAttribute("data-target");
      scrollToId(id);
      panel.hidden = true;
      input.value = "";
      clearSearchHighlights();
    });

    document.addEventListener("click", function(ev){
      if (panel.hidden) return;
      if (!panel.contains(ev.target) && ev.target !== input) { panel.hidden = true; }
    });
  }

  function toggleTruncate(el){
    var full = el.getAttribute("data-full") || "";
    var expanded = el.getAttribute("aria-expanded") === "true";
    if (expanded) {
      var shortText = full.length > 160 ? (full.slice(0, 157) + "...") : full;
      el.textContent = shortText;
      el.setAttribute("aria-expanded", "false");
      el.classList.remove("is-expanded");
    } else {
      el.textContent = full;
      el.setAttribute("aria-expanded", "true");
      el.classList.add("is-expanded");
    }
  }

  function initTruncateHandlers(){
    document.addEventListener("click", function(ev){
      var t = ev.target && ev.target.closest ? ev.target.closest(".truncate") : null;
      if (!t) return;
      toggleTruncate(t);
    });
    document.addEventListener("keydown", function(ev){
      if (ev.key !== "Enter" && ev.key !== " ") return;
      var t = ev.target && ev.target.closest ? ev.target.closest(".truncate") : null;
      if (!t) return;
      ev.preventDefault();
      toggleTruncate(t);
    });
  }

  function boot(){
    safeRun(initThemeToggle);
    safeRun(initPrintButton);
    safeRun(function(){ initSortableTables(document); });
    safeRun(function(){ initTableFilters(document); });
    safeRun(function(){ initFindingsFilter(document); });
    safeRun(initTocFilter);
    safeRun(initScrollspy);
    safeRun(buildSearchIndex);
    safeRun(initGlobalSearch);
    safeRun(initAnchorOpen);
    safeRun(initTruncateHandlers);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
'@

#endregion Render-JS

# =============================================================================
# REGION: Helpers de escape, formato de celdas y metadatos de columna
# =============================================================================
#region Render-Helpers

$script:ColRegexMono   = 'Version|Thumbprint|Huella|Ruta|Path|SID|GUID|Hash|Serie|MAC|Direccion|IP|Puerto'
$script:ColRegexInvert = 'Riesgo|Sospechoso|Pendiente|Expirado|Vulnerab|Deshabilitad'
$script:ColRegexPct    = 'Pct$|Porcentaje$'
$script:ColRegexStatus = 'Estado|Status|State|Cumple|HealthStatus'

function ConvertTo-HtmlSafe {
    <# Escapa & < > " ' para insertar texto de forma segura en HTML. #>
    param([string]$Text)

    if ($null -eq $Text) { return '' }
    $s = [string]$Text
    $s = $s.Replace('&', '&amp;')
    $s = $s.Replace('<', '&lt;')
    $s = $s.Replace('>', '&gt;')
    $s = $s.Replace('"', '&quot;')
    $s = $s.Replace("'", '&#39;')
    return $s
}

function Get-SectionAnchorId {
    <# Genera un id de anclaje HTML estable a partir del Id de seccion (ej '1.11.3' -> 'sec-1-11-3'). #>
    param([string]$Id)

    if ([string]::IsNullOrEmpty($Id)) { return 'sec-x' }
    $limpio = [regex]::Replace($Id, '[^A-Za-z0-9\-]', '-')
    return 'sec-' + $limpio
}

function Get-MetaValue {
    <# Lee una clave de un hashtable de metadatos, con valor por defecto si falta o esta vacia. #>
    param($Meta, [string]$Key, $Default = 'N/D')

    if ($Meta -and $Meta.ContainsKey($Key) -and $null -ne $Meta[$Key] -and [string]$Meta[$Key] -ne '') {
        return $Meta[$Key]
    }
    return $Default
}

function ConvertTo-RenderSafeArray {
    <#
        Normaliza cualquier valor enumerable a un arreglo nativo, sin usar el
        operador @() directamente sobre el objeto de entrada. Algunos runtimes
        de PowerShell tienen un comportamiento inconsistente al envolver con
        @() una System.Collections.Generic.List[object] creada via New-Object
        (el wrap silenciosamente devuelve una coleccion vacia en vez de lanzar
        un error claro), asi que esta funcion recorre el enumerable a mano en
        vez de depender de ese operador para el caso de listas genericas.
    #>
    param($InputObject)

    # OJO: cada return lleva el operador coma unario a proposito. En PS 5.1
    # `return $arreglo` desenrolla un arreglo de UN elemento y el llamador
    # recibe el objeto suelto; como en 5.1 `$objetoSuelto[0]` devuelve $null
    # (en PS 7 devuelve el objeto), la tabla se quedaba sin columnas y toda
    # seccion de exactamente una fila se renderizaba como vacia.
    if ($null -eq $InputObject) { return ,@() }
    if ($InputObject -is [string]) { return ,@($InputObject) }
    if ($InputObject -is [array]) { return ,$InputObject }

    if ($InputObject -is [System.Collections.IEnumerable]) {
        $buffer = New-Object System.Collections.ArrayList
        foreach ($item in $InputObject) { [void]$buffer.Add($item) }
        return ,$buffer.ToArray()
    }

    return ,@($InputObject)
}

function Get-ColumnDataType {
    <# Muestrea hasta 50 filas de una columna para decidir su tipo de dato para el ordenamiento: num/date/bool/text. #>
    param($Rows, [string]$ColumnName)

    $muestras = 0
    $numOk = 0
    $dateOk = 0
    $boolOk = 0

    foreach ($fila in $Rows) {
        if ($muestras -ge 50) { break }
        $v = $null
        if ($fila -is [System.Collections.IDictionary]) {
            $v = $fila[$ColumnName]
        } elseif ($fila -and $fila.PSObject -and $fila.PSObject.Properties[$ColumnName]) {
            $v = $fila.$ColumnName
        }
        if ($null -eq $v) { continue }
        $s = [string]$v
        if ([string]::IsNullOrWhiteSpace($s)) { continue }

        $muestras++
        $sl = $s.ToLowerInvariant().Trim()

        if ($sl -eq 'true' -or $sl -eq 'false' -or $sl -eq 'si' -or $sl -eq 'no') { $boolOk++; continue }

        if ($s -match '^\s*\d{4}-\d{2}-\d{2}' -or $s -match '^\s*\d{1,2}/\d{1,2}/\d{4}') {
            $dt = Get-Date
            $parsedOk = $false
            try { $parsedOk = [datetime]::TryParse($s, [ref]$dt) } catch { $parsedOk = $false }
            if ($parsedOk) { $dateOk++; continue }
        }

        if ($s -match '^-?[0-9][0-9.,]*') { $numOk++; continue }
    }

    if ($muestras -eq 0) { return 'text' }
    if (($boolOk / $muestras) -ge 0.8) { return 'bool' }
    if (($dateOk / $muestras) -ge 0.7) { return 'date' }
    if (($numOk / $muestras) -ge 0.7) { return 'num' }
    return 'text'
}

function Format-CellValue {
    <#
        Formatea un valor de celda a HTML segun heuristicas del nombre de columna:
        booleanos (con inversion de semantica para columnas de riesgo), dias para
        expirar, columnas de porcentaje (mini-medidor), columnas de estado, columnas
        que deben ir en fuente monoespaciada, y truncado de texto largo. Siempre
        escapa el texto final con ConvertTo-HtmlSafe.
    #>
    param($Value, [string]$ColumnName)

    if ($Value -is [System.Array]) {
        $filtrado = @($Value | Where-Object { $null -ne $_ -and [string]$_ -ne '' })
        if ($filtrado.Count -eq 0) { return '<span class="dim">&mdash;</span>' }
        $Value = ($filtrado -join ', ')
    }

    if ($null -eq $Value -or ([string]$Value).Trim() -eq '') {
        return '<span class="dim">&mdash;</span>'
    }

    $text = [string]$Value
    if ($text -match '^System\.') {
        return '<span class="dim">&mdash;</span>'
    }
    $lower = $text.ToLowerInvariant().Trim()

    $isMono      = ($ColumnName -match $script:ColRegexMono)
    $isInvert    = ($ColumnName -match $script:ColRegexInvert)
    $isPct       = ($ColumnName -eq 'LibrePct') -or ($ColumnName -match $script:ColRegexPct)
    $isStatusCol = ($ColumnName -match $script:ColRegexStatus)
    $isExpiry    = ($ColumnName -match 'DiasParaExpirar')

    # --- Booleanos (True/False, Si/No), con inversion de semantica para columnas de riesgo ---
    if ($lower -eq 'true' -or $lower -eq 'false' -or $lower -eq 'si' -or $lower -eq 'no') {
        $esVerdadero = ($lower -eq 'true' -or $lower -eq 'si')
        $etiqueta = 'No'
        if ($esVerdadero) { $etiqueta = 'Si' }

        $clase = 'badge-neutral'
        if ($isInvert) {
            if ($esVerdadero) { $clase = 'badge-crit' } else { $clase = 'badge-ok' }
        } else {
            if ($esVerdadero) { $clase = 'badge-ok' } else { $clase = 'badge-neutral' }
        }
        return '<span class="badge ' + $clase + '">' + $etiqueta + '</span>'
    }

    # --- Dias para expirar ---
    if ($isExpiry) {
        $parsedNum = 0.0
        if ([double]::TryParse($text, [ref]$parsedNum)) {
            $clase = 'badge-neutral'
            if ($parsedNum -lt 30) { $clase = 'badge-crit' }
            elseif ($parsedNum -lt 90) { $clase = 'badge-warn' }
            return '<span class="badge ' + $clase + '">' + (ConvertTo-HtmlSafe $text) + '</span>'
        }
    }

    # --- Columnas de porcentaje: mini medidor + numero ---
    if ($isPct) {
        $m = [regex]::Match($text, '[-+]?[0-9]*\.?[0-9]+')
        if ($m.Success) {
            $pctVal = [double]$m.Value
            if ($pctVal -lt 0) { $pctVal = 0 }
            if ($pctVal -gt 100) { $pctVal = 100 }

            $esLibre = ($ColumnName -match 'Libre')
            $sevClass = 'meter-ok'
            if ($esLibre) {
                if ($pctVal -lt 10) { $sevClass = 'meter-crit' } elseif ($pctVal -lt 20) { $sevClass = 'meter-warn' }
            } else {
                if ($pctVal -ge 90) { $sevClass = 'meter-crit' } elseif ($pctVal -ge 75) { $sevClass = 'meter-warn' }
            }

            $safeText = ConvertTo-HtmlSafe $text
            $pctStr = $pctVal.ToString([System.Globalization.CultureInfo]::InvariantCulture)
            return '<span class="cell-meter"><span class="meter-track"><span class="meter-fill ' + $sevClass + '" style="width:' + $pctStr + '%"></span></span><span class="meter-label mono">' + $safeText + '</span></span>'
        }
    }

    # --- Columnas de estado ---
    if ($isStatusCol) {
        $buenos = @('running', 'healthy', 'ok', 'cumple', 'habilitado', 'enabled', 'activo', 'started', 'up')
        $malos  = @('stopped', 'degraded', 'disabled', 'deshabilitado', 'detenido', 'critical', 'failed', 'error', 'down')
        $clase = 'badge-neutral'
        if ($buenos -contains $lower) { $clase = 'badge-ok' }
        elseif ($malos -contains $lower -or $lower -eq 'no cumple') { $clase = 'badge-crit' }
        return '<span class="badge ' + $clase + '">' + (ConvertTo-HtmlSafe $text) + '</span>'
    }

    # --- Columnas monoespaciadas (rutas, hashes, IPs, GUIDs, etc.) ---
    if ($isMono) {
        if ($text.Length -gt 160) {
            $safeAttr = ConvertTo-HtmlSafe $text
            $truncado = ConvertTo-HtmlSafe ($text.Substring(0, 157) + '...')
            return '<span class="mono truncate" data-full="' + $safeAttr + '" tabindex="0" role="button" aria-expanded="false" title="' + $safeAttr + '">' + $truncado + '</span>'
        }
        return '<span class="mono">' + (ConvertTo-HtmlSafe $text) + '</span>'
    }

    # --- Texto largo generico: truncar con expansion al click ---
    if ($text.Length -gt 160) {
        $safeAttr = ConvertTo-HtmlSafe $text
        $truncado = ConvertTo-HtmlSafe ($text.Substring(0, 157) + '...')
        return '<span class="truncate" data-full="' + $safeAttr + '" tabindex="0" role="button" aria-expanded="false" title="' + $safeAttr + '">' + $truncado + '</span>'
    }

    return ConvertTo-HtmlSafe $text
}

function Build-EmptyState {
    <# Estado vacio discreto para secciones sin datos o roles no instalados. #>
    param([string]$Texto = 'Sin datos disponibles o rol no instalado.')

    $safe = ConvertTo-HtmlSafe $Texto
    return '<div class="empty-state"><span class="empty-icon" aria-hidden="true">&#8709;</span><p>' + $safe + '</p></div>'
}

function Build-MessageBlock {
    <# Bloque para secciones cuyo Data es un string (mensaje/nota/error de recoleccion). #>
    param([string]$Text)

    $safe = ConvertTo-HtmlSafe $Text
    return '<div class="msg"><span class="msg-icon" aria-hidden="true">&#8505;</span><p>' + $safe + '</p></div>'
}

function New-HtmlTable {
    <#
        Construye el fragmento HTML de una tabla a partir de un arreglo de
        PSCustomObject/hashtable. Las columnas salen de la union de propiedades
        de hasta 50 filas de muestra, preservando el orden de aparicion. Agrega
        encabezados ordenables y, si hay mas de 15 filas, un filtro de texto.
    #>
    param($Data, [string]$Id, [switch]$Wide)

    if ($null -eq $Data) { return (Build-EmptyState) }
    if ($Data -is [string]) { return (Build-MessageBlock -Text $Data) }

    $filas = ConvertTo-RenderSafeArray -InputObject $Data
    # Re-envoltura defensiva: si algun runtime desenrolla igual el arreglo de un
    # elemento, esto garantiza que $filas[0] sea indexable y no $null.
    if ($null -ne $filas -and -not ($filas -is [array])) { $filas = @($filas) }
    if ($null -eq $filas -or $filas.Count -eq 0) { return (Build-EmptyState) }

    if (-not $Id -or $Id.Trim() -eq '') {
        $Id = 'tbl-' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
    }

    $columnas = New-Object System.Collections.Generic.List[string]
    $vistas   = New-Object 'System.Collections.Generic.HashSet[string]'
    $limiteMuestra = [Math]::Min(50, $filas.Count)

    for ($i = 0; $i -lt $limiteMuestra; $i++) {
        $fila = $filas[$i]
        $props = $null
        if ($fila -is [System.Collections.IDictionary]) {
            $props = $fila.Keys
        } elseif ($fila -and $fila.PSObject) {
            $props = $fila.PSObject.Properties | ForEach-Object { $_.Name }
        }
        if ($props) {
            foreach ($p in $props) {
                if (-not $vistas.Contains($p)) {
                    [void]$vistas.Add($p)
                    $columnas.Add($p)
                }
            }
        }
    }

    if ($columnas.Count -eq 0) { return (Build-EmptyState) }

    $tipos = @{}
    foreach ($c in $columnas) { $tipos[$c] = Get-ColumnDataType -Rows $filas -ColumnName $c }

    $sb = New-Object System.Text.StringBuilder
    $totalFilas = $filas.Count
    $claseWide = ''
    if ($Wide) { $claseWide = ' wide' }

    if ($totalFilas -gt 15) {
        [void]$sb.Append('<div class="table-filter-bar"><input type="text" class="table-filter-input" data-filter-for="' + $Id + '" placeholder="Filtrar ' + $totalFilas + ' filas..." aria-label="Filtrar filas de la tabla"><span class="table-filter-count" data-count-for="' + $Id + '">mostrando ' + $totalFilas + ' de ' + $totalFilas + '</span></div>')
    }

    [void]$sb.Append('<div class="table-wrap' + $claseWide + '"><table class="data-table" id="' + $Id + '" data-total-rows="' + $totalFilas + '"><thead><tr>')

    $colIndex = 0
    foreach ($c in $columnas) {
        $tipo = $tipos[$c]
        $safeC = ConvertTo-HtmlSafe $c
        [void]$sb.Append('<th data-col="' + $colIndex + '" data-type="' + $tipo + '" aria-sort="none" tabindex="0" role="columnheader"><span class="th-label">' + $safeC + '</span><span class="sort-ind" aria-hidden="true"></span></th>')
        $colIndex++
    }
    [void]$sb.Append('</tr></thead><tbody>')

    foreach ($fila in $filas) {
        [void]$sb.Append('<tr>')
        foreach ($c in $columnas) {
            $v = $null
            if ($fila -is [System.Collections.IDictionary]) {
                $v = $fila[$c]
            } elseif ($fila -and $fila.PSObject -and $fila.PSObject.Properties[$c]) {
                $v = $fila.$c
            }
            $celda = Format-CellValue -Value $v -ColumnName $c
            [void]$sb.Append('<td>' + $celda + '</td>')
        }
        [void]$sb.Append('</tr>')
    }

    [void]$sb.Append('</tbody></table></div>')

    return $sb.ToString()
}

#endregion Render-Helpers

# =============================================================================
# REGION: Arbol de secciones, tabla de contenidos y bloques de la pagina
# =============================================================================
#region Render-Estructura

function Build-SectionTree {
    <# Agrupa $script:ReportSections en un arbol segun la profundidad de puntos del Id. #>
    param($Sections)

    $porId = @{}
    $raices = New-Object System.Collections.Generic.List[object]

    foreach ($s in $Sections) {
        $nodo = [PSCustomObject]@{
            Section  = $s
            Children = New-Object System.Collections.Generic.List[object]
        }
        $porId[$s.Id] = $nodo
    }

    foreach ($s in $Sections) {
        $nodo = $porId[$s.Id]
        $segmentos = $s.Id.Split('.')
        $agregado = $false
        if ($segmentos.Count -gt 1) {
            $idPadre = [string]::Join('.', $segmentos[0..($segmentos.Count - 2)])
            if ($porId.ContainsKey($idPadre)) {
                $porId[$idPadre].Children.Add($nodo)
                $agregado = $true
            }
        }
        if (-not $agregado) { $raices.Add($nodo) }
    }

    return $raices
}

function Build-TocItems {
    <# Genera recursivamente la lista anidada del TOC a partir del arbol de secciones. #>
    param($Nodos, [int]$Profundidad = 0)

    if (-not $Nodos -or $Nodos.Count -eq 0) { return '' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<ul class="toc-list" data-depth="' + $Profundidad + '">')

    foreach ($n in $Nodos) {
        $s = $n.Section
        $anchor = Get-SectionAnchorId -Id $s.Id
        $titulo = ConvertTo-HtmlSafe $s.Title
        $icono = ''
        if ($s.Icon) { $icono = (ConvertTo-HtmlSafe $s.Icon) + ' ' }
        $badge = ''
        if ($s.RowCount -gt 0) { $badge = '<span class="toc-count">' + $s.RowCount + '</span>' }
        $tocTexto = ((ConvertTo-HtmlSafe ($s.Id + ' ' + $s.Title))).ToLowerInvariant()

        [void]$sb.Append('<li class="toc-item" data-toc-text="' + $tocTexto + '">')
        [void]$sb.Append('<a href="#' + $anchor + '" class="toc-link" data-target="' + $anchor + '"><span class="toc-id">' + (ConvertTo-HtmlSafe $s.Id) + '</span><span class="toc-title">' + $icono + $titulo + '</span>' + $badge + '</a>')
        if ($n.Children.Count -gt 0) {
            [void]$sb.Append((Build-TocItems -Nodos $n.Children -Profundidad ($Profundidad + 1)))
        }
        [void]$sb.Append('</li>')
    }

    [void]$sb.Append('</ul>')
    return $sb.ToString()
}

function Build-Sidebar {
    <# Barra lateral fija: hostname, filtro de secciones y TOC generado automaticamente. #>
    param($ReportSections, $ComputerName)

    $arbol = Build-SectionTree -Sections $ReportSections
    $tocHtml = Build-TocItems -Nodos $arbol -Profundidad 0

    $hostname = $ComputerName
    if (-not $hostname) { $hostname = 'EQUIPO' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<nav class="sidebar" id="sidebar" aria-label="Tabla de contenidos">')
    [void]$sb.Append('<div class="sidebar-host">' + (ConvertTo-HtmlSafe $hostname) + '</div>')
    [void]$sb.Append('<input type="text" id="toc-filter" class="toc-filter-input" placeholder="Filtrar secciones..." aria-label="Filtrar secciones del indice">')
    [void]$sb.Append('<div class="toc-scroll">')
    [void]$sb.Append('<a href="#sec-resumen-hallazgos" class="toc-link toc-link-summary" data-target="sec-resumen-hallazgos">&#9873; Resumen de Hallazgos</a>')
    [void]$sb.Append($tocHtml)
    [void]$sb.Append('</div>')
    [void]$sb.Append('</nav>')
    return $sb.ToString()
}

function Build-Topbar {
    <# Barra superior fija: hostname, fecha, puntaje de salud, buscador global, tema e impresion. #>
    param($Meta, $ComputerName, $Puntaje)

    $hostname = $ComputerName
    if (-not $hostname) { $hostname = 'EQUIPO' }
    $fecha = Get-MetaValue -Meta $Meta -Key 'FechaGeneracion' -Default (Get-Date -Format 'yyyy-MM-dd HH:mm')

    $pClase = 'badge-crit'
    if ($Puntaje -ge 85) { $pClase = 'badge-ok' } elseif ($Puntaje -ge 60) { $pClase = 'badge-warn' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="topbar" id="topbar">')
    [void]$sb.Append('<div class="topbar-id"><span class="topbar-host">' + (ConvertTo-HtmlSafe $hostname) + '</span><span class="topbar-date">' + (ConvertTo-HtmlSafe ([string]$fecha)) + '</span><span class="badge ' + $pClase + ' topbar-score">Salud: ' + $Puntaje + '/100</span></div>')
    [void]$sb.Append('<div class="topbar-actions">')
    [void]$sb.Append('<div class="search-wrap"><input type="text" id="global-search" class="global-search-input" placeholder="Buscar en el reporte..." aria-label="Busqueda global"><div id="search-panel" class="search-panel" hidden></div></div>')
    [void]$sb.Append('<button type="button" id="theme-toggle" class="icon-btn" aria-label="Cambiar tema" title="Cambiar tema">&#9788;</button>')
    [void]$sb.Append('<button type="button" id="print-btn" class="icon-btn" aria-label="Imprimir reporte" title="Imprimir">&#128438;</button>')
    [void]$sb.Append('</div>')
    [void]$sb.Append('</div>')
    [void]$sb.Append('<div class="print-header"><span>' + (ConvertTo-HtmlSafe $hostname) + '</span><span>' + (ConvertTo-HtmlSafe ([string]$fecha)) + '</span></div>')
    return $sb.ToString()
}

function Build-RoleChips {
    <# Chips de rol activo, uno por capacidad presente en $script:Caps. #>
    param($Caps)

    $mapa = @(
        @{ Key = 'IIS'; Label = 'IIS' },
        @{ Key = 'DNS'; Label = 'DNS Server' },
        @{ Key = 'DHCP'; Label = 'DHCP Server' },
        @{ Key = 'FileServer'; Label = 'File Server' },
        @{ Key = 'RDSH'; Label = 'RDS / Terminal Services' },
        @{ Key = 'WSUS'; Label = 'WSUS' },
        @{ Key = 'SQL'; Label = 'SQL Server' },
        @{ Key = 'HyperV'; Label = 'Hyper-V' },
        @{ Key = 'Cluster'; Label = 'Cluster de Failover' },
        @{ Key = 'IsDC'; Label = 'Controlador de Dominio' }
    )

    $sb = New-Object System.Text.StringBuilder
    $huboAlguno = $false
    foreach ($m in $mapa) {
        if ($Caps -and $Caps.ContainsKey($m.Key) -and $Caps[$m.Key]) {
            [void]$sb.Append('<span class="role-chip">' + (ConvertTo-HtmlSafe $m.Label) + '</span>')
            $huboAlguno = $true
        }
    }
    if (-not $huboAlguno) {
        [void]$sb.Append('<span class="role-chip role-chip-muted">Sin roles de servidor adicionales detectados</span>')
    }
    return $sb.ToString()
}

function Build-KpiCard {
    <# Tarjeta KPI generica: label chico, valor grande, barra opcional y linea de contexto. #>
    param([string]$Label, [string]$Value, [string]$Context = '', [string]$BarHtml = '', [string]$ExtraClass = '')

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="kpi-card ' + $ExtraClass + '">')
    [void]$sb.Append('<div class="kpi-label">' + (ConvertTo-HtmlSafe $Label) + '</div>')
    [void]$sb.Append('<div class="kpi-value">' + $Value + '</div>')
    if ($BarHtml) { [void]$sb.Append($BarHtml) }
    if ($Context) { [void]$sb.Append('<div class="kpi-context">' + (ConvertTo-HtmlSafe $Context) + '</div>') }
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

function Build-ScoreBar {
    <# Barra de progreso del puntaje de salud: verde >=85, ambar 60-84, rojo <60. #>
    param([double]$Score)

    $s = $Score
    if ($s -lt 0) { $s = 0 }
    if ($s -gt 100) { $s = 100 }
    $clase = 'meter-crit'
    if ($s -ge 85) { $clase = 'meter-ok' } elseif ($s -ge 60) { $clase = 'meter-warn' }
    $sStr = $s.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    return '<div class="meter-track score-track"><div class="meter-fill ' + $clase + '" style="width:' + $sStr + '%"></div></div>'
}

function Get-HealthScore {
    <# Devuelve el puntaje de salud de $script:FindingsSummary, o lo calcula como respaldo. #>
    param($FindingsSummary, $Findings)

    if ($FindingsSummary -and $FindingsSummary.ContainsKey('PuntajeSalud') -and $null -ne $FindingsSummary['PuntajeSalud']) {
        return [Math]::Round([double]$FindingsSummary['PuntajeSalud'], 0)
    }

    $cCrit = 0
    $cWarn = 0
    if ($Findings) {
        $cCrit = @($Findings | Where-Object { $_.Severity -eq 'CRIT' }).Count
        $cWarn = @($Findings | Where-Object { $_.Severity -eq 'WARN' }).Count
    }
    $p = 100 - ($cCrit * 15) - ($cWarn * 5)
    if ($p -lt 0) { $p = 0 }
    return [Math]::Round($p, 0)
}

function Build-Hero {
    <# Portada del reporte: titulo, subtitulo, chips de rol y fila de tarjetas KPI. #>
    param($Meta, $Caps, [double]$Puntaje, $Findings, $ComputerName)

    $hostname = $ComputerName
    if (-not $hostname) { $hostname = Get-MetaValue -Meta $Meta -Key 'HostName' -Default 'EQUIPO' }

    $osCaption = 'Desconocido'
    $osBuild = ''
    $isVm = $false
    if ($Caps) {
        if ($Caps.ContainsKey('OSCaption') -and $Caps['OSCaption']) { $osCaption = $Caps['OSCaption'] }
        if ($Caps.ContainsKey('OSBuild') -and $Caps['OSBuild']) { $osBuild = $Caps['OSBuild'] }
        if ($Caps.ContainsKey('IsVM')) { $isVm = [bool]$Caps['IsVM'] }
    }
    $fabricanteModelo = Get-MetaValue -Meta $Meta -Key 'FabricanteModelo' -Default ''

    $subtitleParts = New-Object System.Collections.Generic.List[string]
    if ($osCaption -and $osCaption -ne 'Desconocido') {
        $bt = $osCaption
        if ($osBuild -and $osBuild -ne 'Desconocido') { $bt = $bt + ' (Build ' + $osBuild + ')' }
        $subtitleParts.Add((ConvertTo-HtmlSafe $bt))
    }
    if ($fabricanteModelo) { $subtitleParts.Add((ConvertTo-HtmlSafe $fabricanteModelo)) }
    if ($isVm) { $subtitleParts.Add('Maquina Virtual') } else { $subtitleParts.Add('Equipo Fisico') }
    $subtitle = [string]::Join(' &middot; ', $subtitleParts)

    $cCritTotal = @($Findings | Where-Object { $_.Severity -eq 'CRIT' }).Count
    $cWarnTotal = @($Findings | Where-Object { $_.Severity -eq 'WARN' }).Count

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<header class="hero">')
    [void]$sb.Append('<h1>Reporte de Configuraci&oacute;n &mdash; ' + (ConvertTo-HtmlSafe $hostname) + '</h1>')
    if ($subtitle) { [void]$sb.Append('<p class="hero-subtitle">' + $subtitle + '</p>') }
    [void]$sb.Append('<div class="role-chips">' + (Build-RoleChips -Caps $Caps) + '</div>')

    [void]$sb.Append('<div class="kpi-grid">')

    $scoreBar = Build-ScoreBar -Score $Puntaje
    [void]$sb.Append((Build-KpiCard -Label 'Puntaje de Salud' -Value ($Puntaje.ToString() + '<span class="kpi-unit">/100</span>') -BarHtml $scoreBar -Context 'Segun hallazgos de esta corrida' -ExtraClass 'kpi-score'))
    [void]$sb.Append((Build-KpiCard -Label 'Criticos' -Value $cCritTotal.ToString() -Context 'Hallazgos severidad CRIT' -ExtraClass 'kpi-crit'))
    [void]$sb.Append((Build-KpiCard -Label 'Advertencias' -Value $cWarnTotal.ToString() -Context 'Hallazgos severidad WARN' -ExtraClass 'kpi-warn'))

    $uptime = Get-MetaValue -Meta $Meta -Key 'UptimeDias' -Default $null
    if ($null -ne $uptime) {
        [void]$sb.Append((Build-KpiCard -Label 'Uptime' -Value ([string]$uptime + '<span class="kpi-unit"> dias</span>') -Context 'Desde el ultimo arranque'))
    } else {
        [void]$sb.Append((Build-KpiCard -Label 'Uptime' -Value 'N/D' -Context 'Dato no disponible'))
    }

    $ram = Get-MetaValue -Meta $Meta -Key 'RamTotalGB' -Default $null
    if ($null -ne $ram) {
        [void]$sb.Append((Build-KpiCard -Label 'RAM Total' -Value ([string]$ram + '<span class="kpi-unit"> GB</span>') -Context 'Memoria fisica instalada'))
    } else {
        [void]$sb.Append((Build-KpiCard -Label 'RAM Total' -Value 'N/D' -Context 'Dato no disponible'))
    }

    $discoPct = Get-MetaValue -Meta $Meta -Key 'DiscoMasOcupadoPct' -Default $null
    $discoNombre = Get-MetaValue -Meta $Meta -Key 'DiscoMasOcupadoNombre' -Default ''
    if ($null -ne $discoPct) {
        $dPct = [double]$discoPct
        $dClase = 'meter-ok'
        if ($dPct -ge 90) { $dClase = 'meter-crit' } elseif ($dPct -ge 75) { $dClase = 'meter-warn' }
        $dPctStr = $dPct.ToString([System.Globalization.CultureInfo]::InvariantCulture)
        $barra = '<div class="meter-track"><div class="meter-fill ' + $dClase + '" style="width:' + $dPctStr + '%"></div></div>'
        $ctx = 'Ocupacion mas alta detectada'
        if ($discoNombre) { $ctx = 'Unidad ' + $discoNombre }
        [void]$sb.Append((Build-KpiCard -Label 'Disco Mas Ocupado' -Value ($dPctStr + '<span class="kpi-unit">%</span>') -BarHtml $barra -Context $ctx))
    } else {
        [void]$sb.Append((Build-KpiCard -Label 'Disco Mas Ocupado' -Value 'N/D' -Context 'Dato no disponible'))
    }

    $parches = Get-MetaValue -Meta $Meta -Key 'ParchesPendientes' -Default $null
    if ($null -ne $parches) {
        [void]$sb.Append((Build-KpiCard -Label 'Parches Pendientes' -Value ([string]$parches) -Context 'Actualizaciones sin aplicar'))
    } else {
        [void]$sb.Append((Build-KpiCard -Label 'Parches Pendientes' -Value 'N/D' -Context 'No relevado en esta corrida'))
    }

    $ultimoParche = Get-MetaValue -Meta $Meta -Key 'UltimoParche' -Default 'N/D'
    [void]$sb.Append((Build-KpiCard -Label 'Ultimo Parche' -Value (ConvertTo-HtmlSafe ([string]$ultimoParche)) -Context 'Fecha de instalacion'))

    [void]$sb.Append('</div>')
    [void]$sb.Append('</header>')

    return $sb.ToString()
}

function Build-FindingsSection {
    <# Resumen ejecutivo de hallazgos: filtros por severidad y tabla ordenada CRIT>WARN>INFO>OK, por categoria. #>
    param($Findings)

    $lista = @()
    if ($Findings) { $lista = ConvertTo-RenderSafeArray -InputObject $Findings }

    $rank = @{ 'CRIT' = 0; 'WARN' = 1; 'INFO' = 2; 'OK' = 3 }
    $ordenados = $lista | Sort-Object @{Expression = { if ($rank.ContainsKey($_.Severity)) { $rank[$_.Severity] } else { 9 } } }, Category

    $cCrit = @($lista | Where-Object { $_.Severity -eq 'CRIT' }).Count
    $cWarn = @($lista | Where-Object { $_.Severity -eq 'WARN' }).Count
    $cInfo = @($lista | Where-Object { $_.Severity -eq 'INFO' }).Count
    $cOk   = @($lista | Where-Object { $_.Severity -eq 'OK' }).Count
    $cTotal = $lista.Count

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<section class="card findings-card" id="sec-resumen-hallazgos">')
    [void]$sb.Append('<h2>Resumen Ejecutivo de Hallazgos</h2>')

    if ($cTotal -eq 0 -or ($cCrit -eq 0 -and $cWarn -eq 0)) {
        [void]$sb.Append('<div class="empty-state empty-state-positive"><span class="empty-icon" aria-hidden="true">&#10003;</span><p><strong>Sin hallazgos criticos ni advertencias.</strong> La configuracion relevada no presenta desvios de severidad alta o media en esta corrida.</p></div>')
    }

    [void]$sb.Append('<div class="findings-filterbar" role="group" aria-label="Filtrar hallazgos por severidad">')
    [void]$sb.Append('<button type="button" class="chip-filter active" data-sev="ALL" aria-pressed="true">Todos <span class="chip-count">' + $cTotal + '</span></button>')
    [void]$sb.Append('<button type="button" class="chip-filter sev-crit" data-sev="CRIT" aria-pressed="false">Criticos <span class="chip-count">' + $cCrit + '</span></button>')
    [void]$sb.Append('<button type="button" class="chip-filter sev-warn" data-sev="WARN" aria-pressed="false">Advertencias <span class="chip-count">' + $cWarn + '</span></button>')
    [void]$sb.Append('<button type="button" class="chip-filter sev-info" data-sev="INFO" aria-pressed="false">Informativos <span class="chip-count">' + $cInfo + '</span></button>')
    [void]$sb.Append('<button type="button" class="chip-filter sev-ok" data-sev="OK" aria-pressed="false">Correcto <span class="chip-count">' + $cOk + '</span></button>')
    [void]$sb.Append('</div>')

    if ($cTotal -gt 0) {
        [void]$sb.Append('<div class="table-wrap"><table class="data-table findings-table" id="tbl-hallazgos"><caption class="sr-only">Listado de hallazgos de la auditoria, ordenados por severidad y categoria</caption><thead><tr>')
        [void]$sb.Append('<th data-col="0" data-type="text" aria-sort="none">Severidad</th><th data-col="1" data-type="text" aria-sort="none">Categoria</th><th data-col="2" data-type="text" aria-sort="none">Hallazgo</th><th data-col="3" data-type="text" aria-sort="none">Detalle</th><th data-col="4" data-type="text" aria-sort="none">Recomendacion</th><th data-col="5" data-type="text" aria-sort="none">Control</th><th data-col="6" data-type="text" aria-sort="none">&nbsp;</th>')
        [void]$sb.Append('</tr></thead><tbody>')

        foreach ($f in $ordenados) {
            $sevClase = 'badge-neutral'
            $sevTexto = [string]$f.Severity
            switch ($f.Severity) {
                'CRIT' { $sevClase = 'badge-crit' }
                'WARN' { $sevClase = 'badge-warn' }
                'INFO' { $sevClase = 'badge-info' }
                'OK'   { $sevClase = 'badge-ok' }
            }
            $verLink = '<span class="dim">&mdash;</span>'
            if ($f.SectionRef) {
                $anchor = Get-SectionAnchorId -Id $f.SectionRef
                $verLink = '<a href="#' + $anchor + '" class="ver-link">Ver &rarr;</a>'
            }
            [void]$sb.Append('<tr data-sev="' + (ConvertTo-HtmlSafe $sevTexto) + '">')
            [void]$sb.Append('<td><span class="badge ' + $sevClase + '">' + (ConvertTo-HtmlSafe $sevTexto) + '</span></td>')
            [void]$sb.Append('<td>' + (ConvertTo-HtmlSafe ([string]$f.Category)) + '</td>')
            [void]$sb.Append('<td>' + (ConvertTo-HtmlSafe ([string]$f.Item)) + '</td>')
            [void]$sb.Append('<td>' + (ConvertTo-HtmlSafe ([string]$f.Detail)) + '</td>')
            [void]$sb.Append('<td>' + (ConvertTo-HtmlSafe ([string]$f.Recommendation)) + '</td>')
            [void]$sb.Append('<td>' + (ConvertTo-HtmlSafe ([string]$f.Control)) + '</td>')
            [void]$sb.Append('<td>' + $verLink + '</td>')
            [void]$sb.Append('</tr>')
        }

        [void]$sb.Append('</tbody></table></div>')
    }

    [void]$sb.Append('</section>')
    return $sb.ToString()
}

function Build-SectionContent {
    <# Renderiza el contenido propio (Data) de un nodo: mensaje, tabla (en details) o estado vacio. #>
    param($Nodo)

    $s = $Nodo.Section
    $sb = New-Object System.Text.StringBuilder

    if ($s.IsMessage) {
        [void]$sb.Append((Build-MessageBlock -Text ([string]$s.Data)))
    } elseif ($s.RowCount -gt 0) {
        $tablaId = 'tbl-' + ($s.Id -replace '[^A-Za-z0-9]', '-')
        $tablaHtml = New-HtmlTable -Data $s.Data -Id $tablaId -Wide:$s.Wide
        $abierto = ''
        if ($s.RowCount -le 40) { $abierto = ' open' }
        $tituloResumen = ConvertTo-HtmlSafe $s.Title
        [void]$sb.Append('<details class="subsection-details"' + $abierto + '><summary>' + $tituloResumen + ' <span class="badge badge-neutral count-badge">' + $s.RowCount + ' filas</span></summary><div class="details-body">' + $tablaHtml + '</div></details>')
    } else {
        [void]$sb.Append((Build-EmptyState))
    }

    return $sb.ToString()
}

function Build-SectionNode {
    <# Renderiza recursivamente un nodo del arbol: seccion de primer nivel (card) o sub-seccion anidada. #>
    param($Nodo, [int]$Profundidad)

    $s = $Nodo.Section
    $anchor = Get-SectionAnchorId -Id $s.Id

    $tag = 'h4'
    if ($Profundidad -eq 0) { $tag = 'h2' } elseif ($Profundidad -eq 1) { $tag = 'h3' }

    $esRaiz = ($Profundidad -eq 0)
    $claseWrap = 'subsection'
    if ($esRaiz) { $claseWrap = 'card' }

    $tagName = 'div'
    if ($esRaiz) { $tagName = 'section' }

    $iconoTxt = ''
    if ($s.Icon) { $iconoTxt = (ConvertTo-HtmlSafe $s.Icon) + ' ' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<' + $tagName + ' class="' + $claseWrap + '" id="' + $anchor + '">')
    [void]$sb.Append('<' + $tag + '>' + $iconoTxt + (ConvertTo-HtmlSafe $s.Id) + ' &mdash; ' + (ConvertTo-HtmlSafe $s.Title) + '</' + $tag + '>')

    if ($s.Note) {
        [void]$sb.Append('<p class="section-note">' + (ConvertTo-HtmlSafe $s.Note) + '</p>')
    }

    $tieneDataPropia = $s.IsMessage -or $s.RowCount -gt 0 -or ($Nodo.Children.Count -eq 0)
    if ($tieneDataPropia) {
        [void]$sb.Append((Build-SectionContent -Nodo $Nodo))
    }

    foreach ($hijo in $Nodo.Children) {
        [void]$sb.Append((Build-SectionNode -Nodo $hijo -Profundidad ($Profundidad + 1)))
    }

    [void]$sb.Append('</' + $tagName + '>')

    return $sb.ToString()
}

function Build-Sections {
    <# Punto de entrada para renderizar todas las secciones a partir de $script:ReportSections. #>
    param($ReportSections)

    $arbol = Build-SectionTree -Sections $ReportSections
    $sb = New-Object System.Text.StringBuilder
    foreach ($raiz in $arbol) {
        [void]$sb.Append((Build-SectionNode -Nodo $raiz -Profundidad 0))
    }
    return $sb.ToString()
}

function Build-Footer {
    <# Pie de pagina: metadata de auditoria, tiempos por seccion y log de ejecucion (colapsados). #>
    param($Meta, $Timings, $LogLines, $ScriptVersion, $ComputerName)

    $usuario = Get-MetaValue -Meta $Meta -Key 'GeneradoPor' -Default $null
    if (-not $usuario) { $usuario = $env:USERNAME }
    if (-not $usuario) { $usuario = 'N/D' }

    $equipoGenerador = Get-MetaValue -Meta $Meta -Key 'EquipoGenerador' -Default $null
    if (-not $equipoGenerador) { $equipoGenerador = $env:COMPUTERNAME }
    if (-not $equipoGenerador) { $equipoGenerador = 'N/D' }

    $fecha = Get-MetaValue -Meta $Meta -Key 'FechaGeneracion' -Default (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $duracion = Get-MetaValue -Meta $Meta -Key 'DuracionTotal' -Default $null

    $tiemposArr = ConvertTo-RenderSafeArray -InputObject $Timings
    $logArr     = ConvertTo-RenderSafeArray -InputObject $LogLines

    if (-not $duracion -and $tiemposArr.Count -gt 0) {
        $totalSeg = 0.0
        foreach ($t in $tiemposArr) {
            if ($t.Segundos) { $totalSeg += [double]$t.Segundos }
        }
        $duracion = [Math]::Round($totalSeg, 1).ToString([System.Globalization.CultureInfo]::InvariantCulture) + ' s (suma de secciones medidas)'
    }
    if (-not $duracion) { $duracion = 'N/D' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<footer class="report-footer">')
    [void]$sb.Append('<p class="footer-meta">Generado por <strong>' + (ConvertTo-HtmlSafe ([string]$usuario)) + '</strong> desde <strong>' + (ConvertTo-HtmlSafe ([string]$equipoGenerador)) + '</strong> el ' + (ConvertTo-HtmlSafe ([string]$fecha)) + ' &middot; Script v' + (ConvertTo-HtmlSafe ([string]$ScriptVersion)) + ' &middot; Duracion total: ' + (ConvertTo-HtmlSafe ([string]$duracion)) + '</p>')

    if ($tiemposArr.Count -gt 0) {
        [void]$sb.Append('<details class="footer-details"><summary>Tiempos por seccion (' + $tiemposArr.Count + ')</summary><div class="details-body">')
        [void]$sb.Append((New-HtmlTable -Data $tiemposArr -Id 'tbl-timings'))
        [void]$sb.Append('</div></details>')
    }

    if ($logArr.Count -gt 0) {
        [void]$sb.Append('<details class="footer-details"><summary>Log de ejecucion (' + $logArr.Count + ' lineas)</summary><div class="details-body"><pre class="log-block">')
        foreach ($l in $logArr) { [void]$sb.Append((ConvertTo-HtmlSafe $l) + "`n") }
        [void]$sb.Append('</pre></div></details>')
    }

    [void]$sb.Append('</footer>')
    return $sb.ToString()
}

#endregion Render-Estructura

# =============================================================================
# REGION: Ensamblado del documento HTML final
# =============================================================================
#region Render-Main

function ConvertTo-ReportHtml {
    <#
        Ensambla el documento HTML completo a partir de $script:ReportSections,
        $script:Findings, $script:FindingsSummary, $script:Caps, $script:Timings,
        $script:LogLines, $script:ScriptVersion y $script:TargetName.

        $Meta es un hashtable opcional con datos que no tienen un origen fijo en
        el contrato compartido (se usan valores por defecto -N/D- si faltan):
          HostName, GeneradoPor, EquipoGenerador, FechaGeneracion, DuracionTotal,
          FabricanteModelo, UptimeDias, RamTotalGB, DiscoMasOcupadoPct,
          DiscoMasOcupadoNombre, ParchesPendientes, UltimoParche.
    #>
    param([hashtable]$Meta)

    if (-not $Meta) { $Meta = @{} }

    $secciones = $script:ReportSections
    if (-not $secciones) { $secciones = @() }

    $hallazgos = $script:Findings
    if (-not $hallazgos) { $hallazgos = @() }

    $resumenHallazgos = $script:FindingsSummary
    $capacidades = $script:Caps
    if (-not $capacidades) { $capacidades = @{} }

    $tiempos = $script:Timings
    $logLineas = $script:LogLines

    $version = $script:ScriptVersion
    if (-not $version) { $version = 'N/D' }

    $equipo = $script:TargetName
    if (-not $equipo) { $equipo = Get-MetaValue -Meta $Meta -Key 'HostName' -Default 'EQUIPO' }

    $puntaje = Get-HealthScore -FindingsSummary $resumenHallazgos -Findings $hallazgos

    $sb = New-Object System.Text.StringBuilder

    [void]$sb.Append('<!doctype html><html lang="es"><head><meta charset="utf-8">')
    [void]$sb.Append('<meta name="viewport" content="width=device-width, initial-scale=1">')
    $tituloPagina = 'Reporte de Configuracion - ' + $equipo
    [void]$sb.Append('<title>' + (ConvertTo-HtmlSafe $tituloPagina) + '</title>')
    [void]$sb.Append('<style>' + $script:ReportCss + '</style>')
    [void]$sb.Append('</head><body>')

    # Script anti-flash de tema: se ejecuta antes del primer paint para evitar
    # el destello de tema incorrecto al abrir el archivo desde disco.
    [void]$sb.Append('<script>try{var t=localStorage.getItem("asbuilt-theme");if(t==="dark"||t==="light"){document.documentElement.setAttribute("data-theme",t);}}catch(e){}</script>')

    [void]$sb.Append((Build-Sidebar -ReportSections $secciones -ComputerName $equipo))

    [void]$sb.Append('<div class="main-col">')
    [void]$sb.Append((Build-Topbar -Meta $Meta -ComputerName $equipo -Puntaje $puntaje))
    [void]$sb.Append('<main class="content" id="content">')

    [void]$sb.Append((Build-Hero -Meta $Meta -Caps $capacidades -Puntaje $puntaje -Findings $hallazgos -ComputerName $equipo))
    [void]$sb.Append((Build-FindingsSection -Findings $hallazgos))
    [void]$sb.Append((Build-Sections -ReportSections $secciones))
    [void]$sb.Append((Build-Footer -Meta $Meta -Timings $tiempos -LogLines $logLineas -ScriptVersion $version -ComputerName $equipo))

    [void]$sb.Append('</main>')
    [void]$sb.Append('</div>')

    [void]$sb.Append('<script>' + $script:ReportJs + '</script>')
    [void]$sb.Append('</body></html>')

    return $sb.ToString()
}

#endregion Render-Main

# #############################################################################
# ##  MOTOR DE HALLAZGOS, EXPORTADORES, DRIFT Y FLOTA
# #############################################################################

# =============================================================================
# FRAGMENTO 04 - MOTOR DE HALLAZGOS, EXPORTADORES, DRIFT Y MODO FLOTA
# Get-ServerFullReport v3.0
# =============================================================================
# Este fragmento consume la API compartida definida en CONTRACT.md (provista por
# el fragmento CORE): $script:ReportSections, $script:Findings, $script:Caps,
# Add-ReportSection, Add-Finding, Write-ReportLog, Test-SectionEnabled, etc.
# NO recolecta datos nuevos: solo analiza, exporta, compara y orquesta flota.
# =============================================================================

#region A - MOTOR DE HALLAZGOS
# =============================================================================
# REGION A: Invoke-FindingsEngine
# =============================================================================

# --- Helpers de acceso defensivo a $script:ReportSections ---------------------

function Get-SectionData {
    # Devuelve el .Data de la seccion cuyo Id coincide exactamente, o $null
    # si la seccion no existe o $script:ReportSections no esta poblado.
    param([string]$Id)
    try {
        if (-not $script:ReportSections) { return $null }
        $sec = $script:ReportSections | Where-Object { $_.Id -eq $Id } | Select-Object -First 1
        if (-not $sec) { return $null }
        return $sec.Data
    } catch { return $null }
}

function Get-Section {
    # Devuelve el objeto seccion completo (Id, Title, Data, Note) o $null.
    param([string]$Id)
    try {
        if (-not $script:ReportSections) { return $null }
        return ($script:ReportSections | Where-Object { $_.Id -eq $Id } | Select-Object -First 1)
    } catch { return $null }
}

function Find-SectionByTitle {
    # Busca una seccion por coincidencia de titulo (regex). Fallback cuando la
    # numeracion exacta de una subseccion puede variar entre fragmentos.
    param([string]$Pattern)
    try {
        if (-not $script:ReportSections -or -not $Pattern) { return $null }
        return ($script:ReportSections | Where-Object { $_.Title -and $_.Title -match $Pattern } | Select-Object -First 1)
    } catch { return $null }
}

function Resolve-SeccionDatos {
    # Intenta ubicar los datos de una seccion probando primero una lista de Ids
    # candidatos (numeracion esperada) y, si ninguno aparece, cae a busqueda por
    # titulo. Devuelve $null si no se encuentra nada (la regla que lo use debe
    # saltearse sin romper).
    param([string[]]$IdCandidates, [string]$TitlePattern)
    try {
        if ($IdCandidates) {
            foreach ($id in $IdCandidates) {
                $d = Get-SectionData -Id $id
                if ($null -ne $d) { return $d }
            }
        }
        if ($TitlePattern) {
            $s = Find-SectionByTitle -Pattern $TitlePattern
            if ($s) { return $s.Data }
        }
        return $null
    } catch { return $null }
}

function Test-HasRows {
    # $false para $null, string (mensaje de error/no instalado) o array vacio.
    # $true para cualquier otro dato (objeto unico o array con elementos).
    param($Data)
    if ($null -eq $Data) { return $false }
    if ($Data -is [string]) { return $false }
    if ($Data -is [System.Collections.IEnumerable]) {
        $count = 0
        foreach ($x in $Data) { $count++ ; if ($count -gt 0) { break } }
        return ($count -gt 0)
    }
    return $true
}

function ConvertTo-RowArray {
    # Normaliza $Data (objeto unico, array, $null o string) a un array de filas.
    param($Data)
    # Operador coma unario en cada return: en PS 5.1 `return $arreglo` desenrolla
    # los arreglos de un solo elemento, y las reglas que hacen $filas[0] o
    # .Count terminaban leyendo mal las secciones de una sola fila.
    if ($null -eq $Data) { return ,@() }
    if ($Data -is [string]) { return ,@() }
    if ($Data -is [array]) { return ,$Data }
    if ($Data -is [System.Collections.IEnumerable]) {
        $buf = New-Object System.Collections.ArrayList
        foreach ($it in $Data) { [void]$buf.Add($it) }
        return ,$buf.ToArray()
    }
    return ,@($Data)
}

function Get-Prop {
    # Devuelve el primer valor no nulo/no vacio entre varios nombres de
    # propiedad candidatos (para tolerar variaciones de nombre entre
    # fragmentos colectores distintos). $Default si ninguno aplica.
    param($Obj, [string[]]$Names, $Default = $null)
    if ($null -eq $Obj) { return $Default }
    foreach ($n in $Names) {
        try {
            $p = $Obj.PSObject.Properties[$n]
            if ($p -and $null -ne $p.Value -and "$($p.Value)" -ne '') { return $p.Value }
        } catch {}
    }
    return $Default
}

function ConvertTo-DoubleSafe {
    param($Value, [double]$Default = 0)
    try {
        if ($null -eq $Value) { return $Default }
        return [double]$Value
    } catch { return $Default }
}

function Test-ValorUtilizable {
    # Devuelve $true solo si $Valor es un dato que el colector realmente pudo
    # obtener. $false para $null, cadena vacia, o cualquiera de los
    # sentinels de "no se pudo leer" que usan los colectores ('N/D', 'N/A',
    # 'No disponible', 'No configurado', 'Desconocido', etc). Ninguna regla
    # que compare un valor contra un umbral debe hacerlo sin pasar antes por
    # aca: afirmar algo sobre un dato no disponible es peor que no decir nada.
    param($Valor)
    if ($null -eq $Valor) { return $false }
    if ($Valor -is [string]) {
        $t = $Valor.Trim()
        if ($t -eq '') { return $false }
        if ($t -match '^(?i:N/D|N/A|NA|No\s+disponible|No\s+configurado|No\s+se\s+pudo\s+(determinar|leer|obtener|consultar).*|Desconocido|Unknown|Not\s+Available|None|Sin\s+datos|Error.*)$') { return $false }
        return $true
    }
    return $true
}

function Get-NumeroSeguro {
    # Convierte $Valor a [double] SOLO si Test-ValorUtilizable lo acepta Y el
    # texto es realmente numerico. A diferencia de ConvertTo-DoubleSafe (que
    # devuelve 0 ante cualquier error de conversion), esta funcion devuelve
    # $null cuando el dato no se pudo leer, para que el llamador nunca
    # confunda un "N/D" con un cero real (el bug que genero el WARN falso de
    # "VigenciaMaxima = N/D dias" tratado como si fuera "vence cada 0 dias").
    # El llamador SIEMPRE debe chequear -ne $null antes de comparar contra un
    # umbral.
    param($Valor)
    if (-not (Test-ValorUtilizable $Valor)) { return $null }
    try {
        if ($Valor -is [double] -or $Valor -is [int] -or $Valor -is [long] -or $Valor -is [decimal] -or $Valor -is [single]) {
            return [double]$Valor
        }
        $texto = "$Valor".Trim()
        $out = 0.0
        if ([double]::TryParse($texto, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$out)) { return $out }
        if ([double]::TryParse($texto, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::CurrentCulture, [ref]$out)) { return $out }
        return $null
    } catch { return $null }
}

function ConvertTo-DateSafe {
    # Convierte texto/objeto a [datetime] de forma defensiva, PROBANDO
    # PRIMERO los formatos explicitos con dia antes que mes (dd-MM-yyyy,
    # dd/MM/yyyy, etc, con CurrentCulture) para no confundir "03-09-2026"
    # (3 de septiembre) con el 9 de marzo. Ese bug exacto convirtio un dato
    # de un dia de atraso en un WARN de firmas de antivirus con "179 dias".
    # Orden de intentos:
    #   1) Formatos explicitos dd-MM-yyyy / dd/MM/yyyy (con y sin hora, con
    #      hora de 1 o 2 digitos) via TryParseExact + CurrentCulture.
    #   2) ISO 8601 (yyyy-MM-ddTHH:mm:ss / yyyy-MM-dd HH:mm:ss).
    #   3) TryParse generico con CurrentCulture.
    #   4) TryParse generico con InvariantCulture (ultimo recurso).
    # Guarda de cordura: si el resultado queda mas de 1 dia en el futuro o
    # antes de 1990, se considera no parseable (mejor no emitir un hallazgo
    # que emitir uno inventado por una fecha mal interpretada) y se devuelve
    # $null.
    param($Value)
    try {
        if ($null -eq $Value -or "$Value" -eq '') { return $null }
        if ($Value -is [datetime]) { $resultado = $Value }
        else {
            $texto = "$Value".Trim()
            if (-not $texto) { return $null }
            $resultado = $null
            $out = [datetime]::MinValue

            $formatosExplicitos = @(
                'dd-MM-yyyy HH:mm:ss', 'dd/MM/yyyy HH:mm:ss',
                'd-M-yyyy H:mm:ss', 'd/M/yyyy H:mm:ss',
                'dd-MM-yyyy', 'dd/MM/yyyy', 'd-M-yyyy', 'd/M/yyyy'
            )
            foreach ($fmt in $formatosExplicitos) {
                if ([datetime]::TryParseExact($texto, $fmt, [System.Globalization.CultureInfo]::CurrentCulture, [System.Globalization.DateTimeStyles]::None, [ref]$out)) {
                    $resultado = $out
                    break
                }
            }

            if (-not $resultado) {
                $formatosIso = @('yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd')
                foreach ($fmt in $formatosIso) {
                    if ([datetime]::TryParseExact($texto, $fmt, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$out)) {
                        $resultado = $out
                        break
                    }
                }
            }

            if (-not $resultado -and [datetime]::TryParse($texto, [System.Globalization.CultureInfo]::CurrentCulture, [System.Globalization.DateTimeStyles]::None, [ref]$out)) {
                $resultado = $out
            }

            if (-not $resultado -and [datetime]::TryParse($texto, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$out)) {
                $resultado = $out
            }

            if (-not $resultado) { return $null }
        }

        # Guarda de cordura: nunca devolver una fecha claramente absurda.
        if ($resultado -gt (Get-Date).AddDays(1)) { return $null }
        if ($resultado -lt [datetime]::ParseExact('1990-01-01', 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)) { return $null }
        return $resultado
    } catch { return $null }
}

function Get-FilaParametro {
    # Varias secciones (por ejemplo 1.11.9, "Estado de Hardening del Sistema
    # Operativo") no exponen sus datos como propiedades sueltas del objeto,
    # sino como una TABLA de filas Parametro/ValorActual/ValorRecomendado/
    # Cumple. Una regla que necesite un dato de esas tablas debe buscar la
    # fila por el texto de su columna Parametro (regex, insensible a mayus/
    # minusculas) en lugar de intentar Get-Prop sobre el objeto contenedor.
    # Devuelve la primera fila cuyo Parametro matchea, o $null.
    param($Data, [string]$Patron)
    if ($null -eq $Data -or $Data -is [string]) { return $null }
    foreach ($row in (ConvertTo-RowArray $Data)) {
        $p = Get-Prop $row @('Parametro')
        if (-not $p -or "$p" -notmatch $Patron) { continue }
        # OJO: cuando el colector no pudo leer el parametro deja
        # ValorActual='No se pudo determinar' y Cumple=$false por defecto. Ese
        # $false NO significa "no cumple", significa "no se sabe". Devolver la
        # fila igual hacia que las reglas afirmaran incumplimientos inexistentes
        # (ej: "SMBv1 habilitado" en un equipo donde el dato no se pudo consultar).
        $valor = Get-Prop $row @('ValorActual')
        if (-not (Test-ValorUtilizable $valor)) {
            Set-TopicoNoEvaluable
            return $null
        }
        return $row
    }
    # Tampoco encontrar la fila significa que el topico no se pudo evaluar. El
    # nombre del parametro varia entre equipos (ej. 'UAC' vs 'UAC (EnableLUA)'),
    # y devolver $null en silencio hacia que el bloque cerrara con un OK.
    Set-TopicoNoEvaluable
    return $null
}

function Set-TopicoNoEvaluable {
    <#
        Marca el topico de cumplimiento en curso como "no evaluable": alguno de
        los datos que necesitaba no se pudo leer. Invoke-CheckBlock consulta esta
        marca y, si esta puesta, NO emite el mensaje OK del topico.

        Sin esto se daba el peor caso posible: la regla no encontraba el dato, no
        emitia hallazgo, y el bloque cerraba con un OK ("UAC esta habilitado",
        "RDP requiere NLA") sobre algo que nunca se pudo verificar.
    #>
    $script:TopicoNoEvaluable = $true
}

function Test-ValorVerdadero {
    # Interpreta un valor de la columna ValorActual (que puede venir como
    # $true/$false, 'True'/'False', '1'/'0', 'Enabled'/'Disabled',
    # 'Habilitado'/'Deshabilitado') como booleano. Devuelve $null si no se
    # puede determinar (el llamador debe tratar eso como "desconocido", no
    # como "false").
    param($Valor)
    if ($null -eq $Valor) { return $null }
    if ($Valor -is [bool]) { return $Valor }
    $txt = "$Valor".Trim()
    if ($txt -match '(?i)^(true|si|1|enabled|habilitado)$') { return $true }
    if ($txt -match '(?i)^(false|no|0|disabled|deshabilitado)$') { return $false }
    return $null
}

# --- Bookkeeping para hallazgos OK de resumen ---------------------------------

$script:ComplianceChecks = @{}

function Invoke-Rule {
    # Envuelve una regla individual en try/catch: un dato inesperado en UNA
    # regla nunca debe interrumpir el resto del motor de hallazgos.
    param([scriptblock]$Body)
    try { & $Body } catch {
        try { Write-ReportLog -Message "Regla de hallazgos omitida por error: $($_.Exception.Message)" -Level Debug } catch {}
    }
}

function Invoke-CheckBlock {
    # Ejecuta un grupo de reglas relacionadas (un "topico" de cumplimiento) y
    # registra si produjo algun hallazgo. Si el topico se evaluo y NO levanto
    # nada (ni CRIT, ni WARN, ni INFO), se emite una finding OK con $OkMessage
    # como resumen positivo, para que el reporte muestre explicitamente lo que
    # SI esta bien.
    #
    # IMPORTANTE: la condicion es "cero hallazgos", no "cero CRIT". Contar solo
    # los CRIT hacia que un topico con WARN o INFO emitiera igual su mensaje OK,
    # produciendo pares contradictorios en el resumen ejecutivo (por ejemplo
    # "Ningun agente EDR detectado" junto a "Se detecto un agente EDR").
    param(
        [Parameter(Mandatory)][string]$Categoria,
        [Parameter(Mandatory)][string]$Topico,
        [Parameter(Mandatory)][string]$OkMessage,
        [Parameter(Mandatory)][scriptblock]$Body,
        # Ids de $script:ReportSections de los que depende este topico. Antes
        # de correr el Body se verifica que AL MENOS UNA de estas secciones
        # exista y tenga datos usables (Test-HasRows sobre Get-SectionData;
        # una seccion cuyo -Data es un string de error/"no instalado" NO
        # cuenta como dato). Si ninguna tiene datos, el Body se ejecuta
        # igual (por si alguna regla tiene su propio fallback por titulo),
        # pero el topico queda marcado como NO EVALUADO y su mensaje OK no
        # se emite: el reporte queda en silencio sobre ese topico en vez de
        # afirmar falsamente que esta todo bien.
        #
        # IMPORTANTE: todo topico NUEVO debe declarar -RequiereSecciones con
        # los Ids reales de los que depende su Body (ver INVENTARIO_SECCIONES.md).
        # Si se omite este parametro se conserva el comportamiento historico
        # (topico siempre considerado "evaluado"), unicamente para no romper
        # topicos preexistentes que todavia no fueron auditados/alineados.
        [string[]]$RequiereSecciones
    )
    try {
        $antes = 0
        foreach ($f in $script:Findings) { if ($f.Severity -ne 'OK') { $antes++ } }

        $evaluado = $true
        if ($RequiereSecciones -and $RequiereSecciones.Count -gt 0) {
            $evaluado = $false
            foreach ($id in $RequiereSecciones) {
                if (Test-HasRows (Get-SectionData -Id $id)) { $evaluado = $true; break }
            }
        }

        $script:TopicoNoEvaluable = $false
        & $Body

        # Una regla pudo marcar el topico como no evaluable aunque la seccion
        # existiera: la seccion estaba, pero el valor puntual decia
        # 'No se pudo determinar'. En ese caso tampoco corresponde el OK.
        $noEvaluable = $false
        try { $noEvaluable = [bool]$script:TopicoNoEvaluable } catch {}
        $script:TopicoNoEvaluable = $false
        if ($noEvaluable) { $evaluado = $false }

        $despues = 0
        foreach ($f in $script:Findings) { if ($f.Severity -ne 'OK') { $despues++ } }
        $script:ComplianceChecks["$Categoria|$Topico"] = [PSCustomObject]@{
            Categoria = $Categoria
            Topico    = $Topico
            OkMessage = $OkMessage
            TuvoCrit  = ($despues -gt $antes)
            Evaluado  = $evaluado
        }
    } catch {
        try { Write-ReportLog -Message "Bloque de cumplimiento '$Topico' omitido por error: $($_.Exception.Message)" -Level Debug } catch {}
    }
}

# --- Listas de referencia usadas por varias reglas ----------------------------

$script:PuertosSensibles = @(135, 139, 445, 1433, 3389, 5985, 5986, 23, 21)

$script:ServiciosDemandaLegitima = @(
    'sppsvc','gpsvc','RemoteRegistry','MapsBroker','TrustedInstaller','BITS','WbioSrvc',
    'CDPSvc','DoSvc','wuauserv','edgeupdate','dbupdate','ShellHWDetection','tiledatamodelsvc',
    'CDPUserSvc','WpnService','SysMain'
)

$script:PatronCuentaGenerica = 'svc[_.-]?|[_.-]svc|servicio|service[_.-]|soporte|support|generic|compartid|shared|operador|backup[_.-]|app[_.-]'

# Grupos y cuentas integradas que NO deben marcarse como "genericas": son
# objetos estandar de Windows/AD y tienen sus propias reglas dedicadas.
$script:CuentasIntegradas = '(Domain|Enterprise|Schema)\s+Admins$|\\Administrators?$|\\Administradores?$|\\Administrator$|\\Administrador$|Domain Users$|Usuarios del dominio$'

# Emisores de CA raiz publicas conocidas (Microsoft Trusted Root Program).
# Cualquier raiz cuyo Sujeto/Emisor contenga uno de estos nombres es una CA
# comercial legitima que aparece de forma normal en cualquier Windows: NO
# debe marcarse como sospechosa solo por no estar en una lista corta. Una
# lista blanca demasiado chica fue la causa de 6-7 WARN falsos en una sola
# corrida (SSL.com, SecureTrust, SECOM, HARICA, etc, todas raices publicas
# reales). Coincidencia parcial, sin distinguir mayus/minusculas.
$script:EmisoresCAPublicasConocidas = @(
    'DigiCert', 'VeriSign', 'GlobalSign', 'Sectigo', 'Comodo', 'Entrust', 'Baltimore',
    'Thawte', 'GeoTrust', 'Go\s*Daddy', 'Starfield', 'Amazon', 'Let''s\s*Encrypt', 'ISRG',
    'USERTrust', 'AddTrust', 'Certum', 'Asseco', 'Unizeto', 'QuoVadis', 'Symantec',
    'Google\s*Trust\s*Services', '\bGTS\b', 'Apple', 'T-TeleSec', 'Staat der Nederlanden',
    'TUBITAK', 'GLOBALTRUST', 'Trustcor', 'Telefonica', 'Vodafone', 'Verizon', 'Cybertrust',
    'Sparkassen', 'Microsoft', 'SSL\.com', 'SSL\s*Corporation', 'SecureTrust', 'SECOM',
    'Security\s*Communication', 'HARICA', 'Buypass', 'D-TRUST', 'T-Systems',
    'Deutsche\s*Telekom', 'TeliaSonera', '\bTelia\b', 'Actalis', 'Camerfirma', '\bANF\b',
    'Firmaprofesional', 'Izenpe', 'Netlock', 'e-Szigno', 'Microsec', 'TWCA', 'Chunghwa',
    'HiPKI', 'GDCA', 'CFCA', 'emSign', 'IdenTrust', 'SecureSign', '\bJCSI\b', 'OISTE',
    'WISeKey', '\bAtos\b', 'Hongkong\s*Post', 'TrustAsia', 'vTrus', 'ZeroSSL', 'XRamp',
    'AffirmTrust', 'Trustwave', 'Network\s*Solutions', '\bSonera\b', 'Certigna',
    'Dhimyotis', 'LuxTrust', 'Disig', 'SwissSign'
)
$script:PatronEmisorCAConfiable = '(?i)(' + ($script:EmisoresCAPublicasConocidas -join '|') + ')'

# Patrones de herramientas de interceptacion/proxy TLS conocidas. Una raiz
# AUTOFIRMADA cuyo Sujeto/Emisor coincide con alguno de estos SI merece un
# WARN (a diferencia de una CA publica desconocida, que es solo INFO).
$script:PatronInterceptorTls = '(?i)(localhost|Fiddler|Charles\s*Proxy|Charles\s*SSL|mitmproxy|Burp|BurpSuite|OWASP\s*ZAP|\bZAP\b|Kaspersky|ESET\s*SSL\s*Filter|\bESET\b|Bitdefender|Avast|\bAVG\b|Sophos|Netskope|Zscaler|Forcepoint|Blue\s*Coat|Symantec\s*Web|Cisco\s*Umbrella|Fortinet|FortiGate|Palo\s*Alto|Check\s*Point|Sonicwall|Barracuda|McAfee\s*Web\s*Gateway|Trend\s*Micro)'

function Test-EmisorCAConfiable {
    # $true si el Sujeto/Emisor de una CA raiz coincide con una autoridad
    # publica comercial conocida (Microsoft Trusted Root Program). Sin dato
    # (texto vacio) se considera "no evaluable" -> $true, para no marcar
    # nada cuando el colector no pudo leer el certificado.
    param([string]$Sujeto, [string]$Emisor)
    $texto = "$Emisor $Sujeto".Trim()
    if (-not $texto) { return $true }
    return ($texto -match $script:PatronEmisorCAConfiable)
}

# --- Exclusion de volumenes vacios/sin formatear para las reglas de storage ---

function Get-VolumenesStorageExcluidos {
    # Devuelve el conjunto de letras de unidad (sin ':', en mayusculas) que
    # NINGUNA regla de storage (espacio critico, volumen de sistema,
    # BitLocker) debe evaluar: unidades opticas/removibles vacias, sin
    # sistema de archivos, o con tamano total nulo/cero/menor a 1 GB. Sin
    # esto, un lector de CD/DVD vacio (0 GB de 0 GB = "0% libre") se
    # reportaba como espacio critico.
    $excluidas = New-Object System.Collections.Generic.HashSet[string]
    try {
        $vols = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.5.2', '1.5.1') -TitlePattern 'Volumen|Volumes')
        foreach ($v in $vols) {
            $letraRaw = "$(Get-Prop $v @('DriveLetter', 'Unidad', 'Letra'))"
            if (-not $letraRaw) { continue }
            $letra = $letraRaw.Trim().TrimEnd(':').ToUpperInvariant()
            if (-not $letra) { continue }
            $total = Get-NumeroSeguro (Get-Prop $v @('TotalGB', 'TamanoGB', 'SizeGB'))
            $fs = "$(Get-Prop $v @('FileSystem', 'FileSystemLabel'))".Trim()
            $tipoUnidad = "$(Get-Prop $v @('DriveType', 'TipoUnidad', 'MediaType'))"
            $esRemovibleOptico = ($tipoUnidad -match '(?i)CD-?ROM|DVD|Removable|Removible|Optic')
            $sinTamano = ($null -eq $total -or $total -lt 1)
            $sinFileSystem = ($fs -eq '')
            # Respaldo por sistema de archivos: una ISO/DVD montada (CDFS o UDF,
            # p. ej. la ISO de instalacion de Windows Server, ~7 GB) tiene tamano
            # real y 0 GB libres, asi que ni el tamano ni el FileSystem vacio la
            # excluyen. DriveType lo resuelve, pero un JSON de una version anterior
            # o un colector sin esa columna no lo trae.
            $esOpticoPorFS = ($fs -match '^(?i:CDFS|UDF)$')
            if ($sinTamano -or $esRemovibleOptico -or $sinFileSystem -or $esOpticoPorFS) { [void]$excluidas.Add($letra) }
        }
    } catch {}
    # Operador coma unario: sin el, PowerShell ENUMERA el HashSet al
    # devolverlo, y un HashSet vacio (el caso normal: casi ningun equipo
    # tiene un volumen para excluir) se enumera a CERO objetos de salida,
    # asi que el llamador recibe $null en vez de un HashSet vacio -- y
    # "$null.Contains(...)" revienta el motor de hallazgos entero.
    return ,$excluidas
}

# =============================================================================
# FUNCION PRINCIPAL DEL MOTOR
# =============================================================================
function Invoke-FindingsEngine {
    [CmdletBinding()]
    param()

    try { Write-ReportLog -Message "Ejecutando motor de hallazgos sobre las secciones recolectadas..." -Level Info } catch {}

    $script:ComplianceChecks = @{}

    #region Storage
    Invoke-CheckBlock -Categoria 'Storage' -Topico 'EspacioLibre' -OkMessage 'Todos los volumenes tienen espacio libre suficiente (>= 20%).' -RequiereSecciones @('1.5.2','1.5.1') -Body {
        $vols = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.5.2','1.5.1') -TitlePattern 'Volumen|Volumes')
        $volsExcluidos = Get-VolumenesStorageExcluidos
        foreach ($v in $vols) {
            Invoke-Rule {
                $unidad = Get-Prop $v @('DriveLetter','Unidad','Letra') 'N/D'
                # Unidades opticas/removibles vacias o sin formatear (0 GB de
                # 0 GB totales) no son un problema de espacio: no hay nada
                # que evaluar.
                if ($volsExcluidos.Contains("$unidad".Trim().TrimEnd(':').ToUpperInvariant())) { return }
                $pct = Get-NumeroSeguro (Get-Prop $v @('LibrePct','PctLibre','PorcentajeLibre','FreePercent'))
                if ($null -eq $pct) { return }
                $libreGB = Get-Prop $v @('LibreGB','EspacioLibreGB','FreeGB')
                $totalGB = Get-Prop $v @('TotalGB','TamanoGB','SizeGB')
                $detail = "Unidad $unidad`: $libreGB GB libres de $totalGB GB totales ($pct% libre)"
                if ($pct -lt 10) {
                    Add-Finding -Severity 'CRIT' -Category 'Storage' -Item "Volumen $unidad con espacio critico" -Detail $detail -Recommendation 'Liberar espacio o expandir el volumen de inmediato.' -Control 'ISO 27001 A.8.6' -SectionRef '1.5.2'
                } elseif ($pct -lt 20) {
                    Add-Finding -Severity 'WARN' -Category 'Storage' -Item "Volumen $unidad con poco espacio libre" -Detail $detail -Recommendation 'Planificar liberacion de espacio o ampliacion.' -Control 'ISO 27001 A.8.6' -SectionRef '1.5.2'
                }
                $libreGBNum = Get-NumeroSeguro $libreGB
                if ($unidad -match '^C' -and $null -ne $libreGBNum -and $libreGBNum -lt 15) {
                    Add-Finding -Severity 'CRIT' -Category 'Storage' -Item 'Volumen de sistema (C:) con menos de 15 GB libres' -Detail $detail -Recommendation 'Liberar espacio en el volumen de sistema: sin espacio los parches de Windows no pueden instalarse.' -Control 'ISO 27001 A.8.6' -SectionRef '1.5.2'
                }
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Storage' -Topico 'SaludDiscos' -OkMessage 'Todos los discos fisicos reportan estado de salud normal.' -RequiereSecciones @('1.17.6') -Body {
        # OJO: 1.5.3 es BitLocker y 1.5.4 es Deduplicacion; la salud de discos
        # fisicos (HealthStatus/OperationalStatus) la publica el colector en
        # 1.17.6 ("Salud de Discos Fisicos").
        $discos = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.17.6') -TitlePattern 'Salud.*Disco|Discos Fisicos|PhysicalDisk')
        foreach ($d in $discos) {
            Invoke-Rule {
                $salud = Get-Prop $d @('HealthStatus','EstadoSalud','Salud')
                if ($null -eq $salud) { return }
                if ("$salud" -ne 'Healthy' -and "$salud" -ne 'Saludable') {
                    $nombre = Get-Prop $d @('FriendlyName','Modelo','DeviceId','DeviceID') 'N/D'
                    Add-Finding -Severity 'CRIT' -Category 'Storage' -Item "Disco fisico con estado de salud degradado" -Detail "$nombre reporta HealthStatus=$salud" -Recommendation 'Revisar el disco fisico y planificar su reemplazo antes de una falla.' -Control 'ISO 27001 A.8.6' -SectionRef '1.5.3'
                }
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Storage' -Topico 'BitLocker' -OkMessage 'Los volumenes de datos fijos estan cifrados con BitLocker.' -RequiereSecciones @('1.5.3') -Body {
        # OJO: 1.5.4 es Deduplicacion de Datos, no BitLocker. El estado real de
        # BitLocker (VolumeStatus/ProtectionStatus) esta en 1.5.3.
        $bl = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.5.3') -TitlePattern 'BitLocker')
        $hw = Get-SectionData -Id '1.1'
        $esPortatil = $false
        try {
            $chasis = Get-Prop $hw @('TipoChasis','ChassisType','FormFactor','TipoSistema')
            if ($chasis -and "$chasis" -match 'Laptop|Portatil|Notebook') { $esPortatil = $true }
        } catch {}
        $volsExcluidos = Get-VolumenesStorageExcluidos
        foreach ($v in $bl) {
            Invoke-Rule {
                $unidadCheck = "$(Get-Prop $v @('DriveLetter','MountPoint','Unidad'))".Trim().TrimEnd(':').ToUpperInvariant()
                if ($unidadCheck -and $volsExcluidos.Contains($unidadCheck)) { return }
                $estado = Get-Prop $v @('VolumeStatus','Estado','EstadoCifrado','ProtectionStatus')
                $tipo = Get-Prop $v @('VolumeType','Tipo') 'Fixed'
                if ($null -eq $estado) { return }
                $esDatos = ("$tipo" -notmatch 'OperatingSystem|Sistema')
                $sinCifrar = ("$estado" -match 'FullyDecrypted|Decrypted|Deshabilitado|Off')
                if (($esDatos -and $esPortatil -and $sinCifrar) -or ("$estado" -match 'FullyDecrypted')) {
                    $unidad = Get-Prop $v @('DriveLetter','MountPoint','Unidad') 'N/D'
                    # C: es el volumen del sistema, no un volumen "de datos":
                    # etiquetarlo mal confunde al leer el resumen ejecutivo.
                    $esSistema  = ("$unidad" -match '^C:?$')
                    $etiquetaVol = if ($esSistema) { "Volumen del sistema sin cifrar ($unidad)" } else { "Volumen de datos sin cifrar ($unidad)" }
                    $recomendacion = if ($esSistema) {
                        'Habilitar BitLocker en el volumen del sistema. Sin cifrado, cualquiera con acceso fisico al disco puede leer los datos arrancando desde otro medio.'
                    } else {
                        'Habilitar BitLocker en volumenes de datos fijos, especialmente en equipos portatiles.'
                    }
                    Add-Finding -Severity 'WARN' -Category 'Storage' -Item $etiquetaVol -Detail "Estado BitLocker: $estado" -Recommendation $recomendacion -Control 'CIS 18.10.9 / ISO 27001 A.8.24' -SectionRef '1.5.3'
                }
            }
        }
    }
    #endregion Storage

    #region Parches y Disponibilidad
    Invoke-CheckBlock -Categoria 'Disponibilidad' -Topico 'ReinicioYUptime' -OkMessage 'No hay reinicio pendiente y el uptime esta dentro de rangos normales.' -RequiereSecciones @('1.11.7','1.2.1','1.17.1') -Body {
        Invoke-Rule {
            $rp = Get-SectionData -Id '1.11.7'
            if ($null -eq $rp) { $rp = Resolve-SeccionDatos -TitlePattern 'Reinicio Pendiente|Reboot Pending' }
            $pendiente = $null
            if ($rp) { $pendiente = Get-Prop $rp @('ReinicioPendiente','PendingReboot') }
            # El uptime esta duplicado en dos secciones con nombres de columna
            # distintos: 1.2.1 (Configuracion del SO) expone UptimeDias, y
            # 1.17.1 (Uptime y Arranque) expone DiasEncendido. Se prueban
            # ambas antes de recurrir a calcular la diferencia contra la
            # fecha de UltimoArranque.
            $os = Resolve-SeccionDatos -IdCandidates @('1.2.1','1.17.1') -TitlePattern 'Sistema Operativo|OS Configuration|Uptime y Arranque'
            $uptimeDias = $null
            if ($os) {
                $uptimeDiasDirecto = Get-NumeroSeguro (Get-Prop $os @('UptimeDias','DiasEncendido'))
                if ($null -ne $uptimeDiasDirecto) {
                    $uptimeDias = $uptimeDiasDirecto
                } else {
                    $arranque = ConvertTo-DateSafe (Get-Prop $os @('UltimoArranque','LastBootUpTime'))
                    if ($arranque) { $uptimeDias = [math]::Round(((Get-Date) - $arranque).TotalDays, 1) }
                }
            }
            if ($pendiente -eq $true) {
                if ($uptimeDias -ne $null -and $uptimeDias -gt 60) {
                    Add-Finding -Severity 'CRIT' -Category 'Parches' -Item 'Reinicio pendiente con uptime mayor a 60 dias' -Detail "Uptime actual: $uptimeDias dias" -Recommendation 'Coordinar ventana de mantenimiento y reiniciar el servidor cuanto antes.' -Control 'CIS 18.9.108 / ISO 27001 A.8.8' -SectionRef '1.11.7'
                } else {
                    Add-Finding -Severity 'WARN' -Category 'Parches' -Item 'Reinicio pendiente' -Detail 'El servidor tiene un reinicio pendiente por actualizaciones o cambios de configuracion.' -Recommendation 'Programar reinicio en la proxima ventana de mantenimiento.' -Control 'CIS 18.9.108 / ISO 27001 A.8.8' -SectionRef '1.11.7'
                }
            }
            if ($uptimeDias -ne $null -and $uptimeDias -gt 180) {
                Add-Finding -Severity 'WARN' -Category 'Disponibilidad' -Item 'Uptime mayor a 180 dias' -Detail "Uptime actual: $uptimeDias dias. Indica que no se aplican parches con reinicio." -Recommendation 'Revisar el proceso de gestion de parches y programar un reinicio.' -Control 'CIS 18.9.108 / ISO 27001 A.8.8' -SectionRef '1.2.1'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Parches' -Topico 'Hotfixes' -OkMessage 'El ultimo hotfix se instalo dentro de los ultimos 60 dias.' -RequiereSecciones @('1.2.2') -Body {
        Invoke-Rule {
            $hotfixes = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.2.2') -TitlePattern 'Hotfix')
            if ($hotfixes.Count -eq 0) { return }
            $fechas = @()
            foreach ($h in $hotfixes) {
                $f = ConvertTo-DateSafe (Get-Prop $h @('InstalledOn','FechaInstalacion'))
                if ($f) { $fechas += $f }
            }
            if ($fechas.Count -eq 0) { return }
            $ultimo = ($fechas | Sort-Object -Descending | Select-Object -First 1)
            $dias = [math]::Round(((Get-Date) - $ultimo).TotalDays, 0)
            # Formato chileno dd-MM-yyyy en el texto del hallazgo.
            $ultimo = $ultimo.ToString('dd-MM-yyyy')
            if ($dias -gt 120) {
                Add-Finding -Severity 'CRIT' -Category 'Parches' -Item 'Ultimo hotfix instalado hace mas de 120 dias' -Detail "Ultimo hotfix: $ultimo ($dias dias atras)" -Recommendation 'Aplicar actualizaciones acumulativas pendientes de inmediato.' -Control 'CIS 18.9.108 / ISO 27001 A.8.8' -SectionRef '1.2.2'
            } elseif ($dias -gt 60) {
                Add-Finding -Severity 'WARN' -Category 'Parches' -Item 'Ultimo hotfix instalado hace mas de 60 dias' -Detail "Ultimo hotfix: $ultimo ($dias dias atras)" -Recommendation 'Revisar el ciclo de aplicacion de parches mensuales.' -Control 'CIS 18.9.108 / ISO 27001 A.8.8' -SectionRef '1.2.2'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Parches' -Topico 'ActualizacionesFaltantes' -OkMessage 'No hay actualizaciones criticas o importantes pendientes.' -RequiereSecciones @('1.2.3') -Body {
        Invoke-Rule {
            $missing = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.2.3') -TitlePattern 'Actualizaciones Faltantes|Missing.*Update')
            if ($missing.Count -eq 0) { return }
            $criticas = @($missing | Where-Object { (Get-Prop $_ @('Severidad','MsrcSeverity')) -eq 'Critical' })
            $importantes = @($missing | Where-Object { (Get-Prop $_ @('Severidad','MsrcSeverity')) -eq 'Important' })
            if ($criticas.Count -gt 0) {
                Add-Finding -Severity 'CRIT' -Category 'Parches' -Item 'Actualizaciones criticas de Windows pendientes' -Detail "$($criticas.Count) actualizacion(es) con severidad Critical sin instalar" -Recommendation 'Instalar las actualizaciones criticas a la brevedad.' -Control 'CIS 18.9.108 / ISO 27001 A.8.8' -SectionRef '1.2.3'
            }
            if ($importantes.Count -gt 0) {
                Add-Finding -Severity 'WARN' -Category 'Parches' -Item 'Actualizaciones importantes de Windows pendientes' -Detail "$($importantes.Count) actualizacion(es) con severidad Important sin instalar" -Recommendation 'Planificar la instalacion en la proxima ventana de mantenimiento.' -Control 'CIS 18.9.108 / ISO 27001 A.8.8' -SectionRef '1.2.3'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Disponibilidad' -Topico 'ApagadosInesperados' -OkMessage 'No se registraron apagados inesperados en los ultimos 30 dias.' -RequiereSecciones @('1.17.2') -Body {
        Invoke-Rule {
            $eventos = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.17.2') -TitlePattern 'Apagados Inesperados|Unexpected Shutdown')
            $count = 0
            foreach ($e in $eventos) {
                $id = Get-Prop $e @('EventoID','EventId','Id')
                if ("$id" -eq '6008' -or "$id" -eq '41') { $count++ }
            }
            if ($count -eq 0 -and $eventos.Count -gt 0) {
                $count = Get-NumeroSeguro (Get-Prop $eventos[0] @('Ocurrencias','Count','Cantidad'))
                if ($null -eq $count) { $count = 0 }
            }
            if ($count -gt 0) {
                Add-Finding -Severity 'WARN' -Category 'Disponibilidad' -Item 'Apagados inesperados en los ultimos 30 dias' -Detail "$count evento(s) 6008/41 detectados" -Recommendation 'Investigar la causa (corte de energia, falla de hardware, kernel panic).' -SectionRef '1.17.2'
            }
        }
    }
    #endregion Parches y Disponibilidad

    #region Seguridad
    Invoke-CheckBlock -Categoria 'Seguridad' -Topico 'Cifrado' -OkMessage 'Cifrado: TLS 1.2/1.3 unicamente, sin protocolos obsoletos.' -RequiereSecciones @('1.11.2') -Body {
        Invoke-Rule {
            $tls = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.11.2') -TitlePattern 'Protocolos.*Cifrados|SChannel|Protocolos y Cifrados')
            foreach ($t in $tls) {
                # El colector real llama a esta columna "Nombre" (con "Tipo" en
                # una columna separada indicando Protocolo/Cifrado/Hash), no
                # "Protocolo"; se conserva 'Protocolo' como candidato extra por
                # compatibilidad con datos de prueba/formatos anteriores.
                $tipoFila = Get-Prop $t @('Tipo')
                if ($tipoFila -and "$tipoFila" -ne 'Protocolo') { continue }
                $proto = Get-Prop $t @('Nombre','Protocolo')
                $lado = Get-Prop $t @('Lado')
                $estado = Get-Prop $t @('Estado')
                if (-not $proto -or -not $estado) { continue }
                if ($lado -and "$lado" -ne 'Server') { continue }
                if (($proto -match 'SSL 2\.0|SSL 3\.0|TLS 1\.0|TLS 1\.1') -and ($estado -eq 'Habilitado')) {
                    Add-Finding -Severity 'CRIT' -Category 'Seguridad' -Item "Protocolo obsoleto habilitado: $proto" -Detail "Lado $lado`: $estado" -Recommendation 'Deshabilitar el protocolo obsoleto en SCHANNEL (registro).' -Control 'CIS 18.x / PCI-DSS 4.1 / ISO 27001 A.8.24' -SectionRef '1.11.2'
                }
                if (($proto -eq 'TLS 1.2') -and ($estado -eq 'Deshabilitado')) {
                    Add-Finding -Severity 'CRIT' -Category 'Seguridad' -Item 'TLS 1.2 deshabilitado explicitamente' -Detail "Lado $lado`: $estado. Esto puede romper la conectividad de clientes y servicios." -Recommendation 'Habilitar TLS 1.2 salvo que TLS 1.3 este garantizado en todos los clientes.' -Control 'CIS 18.x / PCI-DSS 4.1 / ISO 27001 A.8.24' -SectionRef '1.11.2'
                }
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Seguridad' -Topico 'SMB' -OkMessage 'SMBv1 deshabilitado y firma SMB requerida.' -RequiereSecciones @('1.7.1','1.11.9') -Body {
        Invoke-Rule {
            # 1.7.1 (Get-SmbServerConfiguration) SOLO se recolecta si el equipo
            # tiene shares de datos (Caps.FileServer); en el resto de los
            # servidores el estado real de SMBv1/firma SMB solo esta
            # disponible en la tabla de hardening 1.11.9 (Parametro/
            # ValorActual/Cumple), que se recolecta siempre.
            $smb = Get-SectionData -Id '1.7.1'
            if (-not $smb) { $smb = Resolve-SeccionDatos -TitlePattern 'SMB Network Interface|Configuracion SMB' }
            $smb1 = $null
            $firma = $null
            if ($smb -and -not ($smb -is [string])) {
                $smb1 = Get-Prop $smb @('EnableSMB1Protocol')
                $firma = Get-Prop $smb @('RequireSecuritySignature','EnableSecuritySignature')
            }

            $hardening = Get-SectionData -Id '1.11.9'
            if (-not $hardening) { $hardening = Resolve-SeccionDatos -TitlePattern 'Hardening del Sistema Operativo' }

            if ($null -eq $smb1) {
                $filaSmb1 = Get-FilaParametro -Data $hardening -Patron 'SMBv1'
                if ($filaSmb1) {
                    $cumple = Get-Prop $filaSmb1 @('Cumple')
                    if ($cumple -eq $true -or "$cumple" -eq 'True') { $smb1 = $false }
                    elseif ($cumple -eq $false -or "$cumple" -eq 'False') { $smb1 = $true }
                    else {
                        $valorActual = Get-Prop $filaSmb1 @('ValorActual')
                        $vb = Test-ValorVerdadero $valorActual
                        if ($null -ne $vb) { $smb1 = $vb }
                        elseif ("$valorActual" -eq 'Enabled') { $smb1 = $true }
                        elseif ("$valorActual" -eq 'Disabled') { $smb1 = $false }
                    }
                }
            }
            if ($smb1 -eq $true) {
                Add-Finding -Severity 'CRIT' -Category 'Seguridad' -Item 'SMBv1 habilitado' -Detail 'EnableSMB1Protocol = True' -Recommendation 'Deshabilitar SMBv1 (Set-SmbServerConfiguration -EnableSMB1Protocol $false). Protocolo vulnerable a WannaCry/EternalBlue.' -Control 'CIS 18.4.3 / ISO 27001 A.8.8' -SectionRef '1.7.1'
            }

            if ($null -eq $firma) {
                $filaFirma = Get-FilaParametro -Data $hardening -Patron 'Firma SMB'
                if ($filaFirma) {
                    $cumple = Get-Prop $filaFirma @('Cumple')
                    if ($cumple -eq $true -or "$cumple" -eq 'True') { $firma = $true }
                    elseif ($cumple -eq $false -or "$cumple" -eq 'False') { $firma = $false }
                }
            }
            if ($firma -eq $false) {
                Add-Finding -Severity 'WARN' -Category 'Seguridad' -Item 'Firma SMB no requerida' -Detail 'RequireSecuritySignature = False' -Recommendation 'Requerir firma SMB para mitigar ataques de relay.' -Control 'CIS 18.4 / ISO 27001 A.8.24' -SectionRef '1.7.1'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Seguridad' -Topico 'RDP' -OkMessage 'RDP requiere NLA y no esta expuesto sin restricciones.' -RequiereSecciones @('1.11.9','1.4.6') -Body {
        Invoke-Rule {
            # Ninguna seccion de colector se titula literalmente "RDP"/"Remote
            # Desktop" con el estado de NLA (el titulo 'Grupo Remote Desktop
            # Users' matchea por accidente el patron de titulo, pero es una
            # lista de miembros, no de configuracion). El estado real de NLA
            # esta en la tabla de hardening 1.11.9.
            $nla = $null
            $rdp = Resolve-SeccionDatos -TitlePattern 'Escritorio Remoto|RDP|Remote Desktop'
            if ($rdp -and -not ($rdp -is [string])) {
                $nla = Get-Prop $rdp @('UserAuthenticationRequired','NLA','RequiereNLA')
            }
            if ($null -eq $nla) {
                $hardening = Get-SectionData -Id '1.11.9'
                if (-not $hardening) { $hardening = Resolve-SeccionDatos -TitlePattern 'Hardening del Sistema Operativo' }
                $filaNla = Get-FilaParametro -Data $hardening -Patron 'Network Level Authentication|NLA|Autenticacion a nivel de red'
                if ($filaNla) {
                    $cumple = Get-Prop $filaNla @('Cumple')
                    if ($cumple -eq $true -or "$cumple" -eq 'True') { $nla = 1 }
                    elseif ($cumple -eq $false -or "$cumple" -eq 'False') { $nla = 0 }
                    else {
                        # Get-NumeroSeguro (no ConvertTo-DoubleSafe): si
                        # ValorActual viene 'N/D', ConvertTo-DoubleSafe lo
                        # convertia en 0 y disparaba un CRIT falso de "NLA
                        # deshabilitado" sobre un dato que en realidad no se
                        # pudo leer.
                        $valorActual = Get-Prop $filaNla @('ValorActual')
                        $nla = Get-NumeroSeguro $valorActual
                    }
                }
            }
            if ($nla -eq $false -or $nla -eq 0) {
                Add-Finding -Severity 'CRIT' -Category 'Seguridad' -Item 'RDP con NLA deshabilitado' -Detail 'UserAuthenticationRequired = 0' -Recommendation 'Habilitar Network Level Authentication para RDP.' -Control 'CIS 18.x / ISO 27001 A.8.24' -SectionRef '1.10'
            }
            $puertos = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.4.6') -TitlePattern 'Puertos en Escucha|Listening')
            $rdpExpuesto = $false
            foreach ($p in $puertos) {
                $puerto = Get-Prop $p @('PuertoLocal','LocalPort')
                if ("$puerto" -eq '3389') { $rdpExpuesto = $true }
            }
            if ($rdpExpuesto) {
                $perfiles = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.3.1') -TitlePattern 'Perfiles de Firewall|Firewall Profile')
                $publicoPermite = $false
                foreach ($fp in $perfiles) {
                    $nombre = Get-Prop $fp @('Name','Nombre')
                    $accionDefault = Get-Prop $fp @('DefaultInboundAction')
                    $habilitado = Get-Prop $fp @('Enabled')
                    if ("$nombre" -match 'Public|Publico' -and ($habilitado -eq $true -or "$habilitado" -eq 'True') -and "$accionDefault" -eq 'Allow') { $publicoPermite = $true }
                }
                if ($publicoPermite) {
                    Add-Finding -Severity 'WARN' -Category 'Seguridad' -Item 'RDP (3389) expuesto con perfil publico permisivo' -Detail 'Puerto 3389 en escucha y el perfil de firewall Public permite trafico entrante por defecto.' -Recommendation 'Restringir el acceso RDP por IP de origen o VPN, o bloquear en el perfil publico.' -SectionRef '1.4.6'
                }
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Seguridad' -Topico 'UAC' -OkMessage 'UAC (User Account Control) esta habilitado.' -RequiereSecciones @('1.11.9') -Body {
        Invoke-Rule {
            # Ninguna seccion se titula "UAC"/"User Account Control"; el
            # estado real de EnableLUA esta en la tabla de hardening 1.11.9
            # (fila con Parametro que contiene "EnableLUA").
            $enableLua = $null
            $uac = Resolve-SeccionDatos -TitlePattern 'UAC|User Account Control'
            if ($uac -and -not ($uac -is [string])) { $enableLua = Get-Prop $uac @('EnableLUA') }
            if ($null -eq $enableLua) {
                $hardening = Get-SectionData -Id '1.11.9'
                if (-not $hardening) { $hardening = Resolve-SeccionDatos -TitlePattern 'Hardening del Sistema Operativo' }
                $filaUac = Get-FilaParametro -Data $hardening -Patron 'EnableLUA|^UAC'
                if ($filaUac) {
                    $cumple = Get-Prop $filaUac @('Cumple')
                    if ($cumple -eq $true -or "$cumple" -eq 'True') { $enableLua = $true }
                    elseif ($cumple -eq $false -or "$cumple" -eq 'False') { $enableLua = $false }
                    else { $enableLua = Test-ValorVerdadero (Get-Prop $filaUac @('ValorActual')) }
                }
            }
            if ($null -eq $enableLua) { Set-TopicoNoEvaluable; return }
            if ($enableLua -eq 0 -or $enableLua -eq $false) {
                Add-Finding -Severity 'CRIT' -Category 'Seguridad' -Item 'UAC deshabilitado' -Detail 'EnableLUA = 0' -Recommendation 'Habilitar UAC (EnableLUA = 1) para mantener la separacion de privilegios.' -SectionRef '1.11'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Seguridad' -Topico 'Firewall' -OkMessage 'Todos los perfiles de firewall estan habilitados.' -RequiereSecciones @('1.3.1','1.3.2') -Body {
        Invoke-Rule {
            $perfiles = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.3.1') -TitlePattern 'Perfiles de Firewall|Firewall Profile')
            foreach ($fp in $perfiles) {
                $nombre = Get-Prop $fp @('Name','Nombre') 'N/D'
                $habilitado = Get-Prop $fp @('Enabled')
                if ($habilitado -eq $false -or "$habilitado" -eq 'False') {
                    Add-Finding -Severity 'CRIT' -Category 'Seguridad' -Item "Perfil de firewall deshabilitado: $nombre" -Detail "Enabled = False" -Recommendation 'Habilitar el perfil de firewall Windows.' -SectionRef '1.3.1'
                }
            }
        }
        Invoke-Rule {
            $reglas = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.3.2') -TitlePattern 'Reglas de Firewall|Firewall Rule')
            $count = 0
            foreach ($r in $reglas) {
                $accion = Get-Prop $r @('Action')
                $direccion = Get-Prop $r @('Direction')
                $perfil = Get-Prop $r @('Profile')
                if ("$accion" -eq 'Allow' -and "$direccion" -eq 'Inbound' -and "$perfil" -match 'Public') { $count++ }
            }
            if ($count -gt 0) {
                Add-Finding -Severity 'INFO' -Category 'Seguridad' -Item 'Reglas de firewall entrantes permisivas en perfil publico' -Detail "$count regla(s) con Action=Allow, Direction=Inbound y Profile incluyendo Public" -Recommendation 'Revisar si esas reglas son necesarias en el perfil publico.' -SectionRef '1.3.2'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Seguridad' -Topico 'Antivirus' -OkMessage 'Antivirus/EDR activo con proteccion en tiempo real y firmas actualizadas.' -RequiereSecciones @('1.11.6') -Body {
        Invoke-Rule {
            $av = Get-SectionData -Id '1.11.6'
            if (-not $av) { $av = Resolve-SeccionDatos -TitlePattern 'Antivirus' }
            # Si la seccion no existe o vino como string de error, no hay nada
            # que evaluar: antes esta funcion asumia "sin deteccion" por
            # defecto y emitia un CRIT igual, incluso sin datos.
            if (-not $av -or ($av -is [string])) { return }
            $sinDeteccion = $true
            $filas = ConvertTo-RowArray $av
            foreach ($row in $filas) {
                # El colector real NO expone RealTimeProtectionEnabled ni
                # AntivirusSignatureLastUpdated como propiedades sueltas: las
                # vuelca dentro del texto formateado de la columna
                # WindowsDefender (Out-String de Get-MpComputerStatus). Se
                # intenta la propiedad directa primero (compatibilidad con
                # otros formatos) y se cae a extraerlo por regex del volcado.
                $wdTexto = "$(Get-Prop $row @('WindowsDefender'))"
                $rt = Get-Prop $row @('RealTimeProtectionEnabled')
                if ($null -eq $rt -and $wdTexto -match 'RealTimeProtectionEnabled\s*:\s*(\S+)') { $rt = $matches[1] }
                if ($null -ne $rt) {
                    $sinDeteccion = $false
                    if ($rt -eq $false -or "$rt" -eq 'False') {
                        Add-Finding -Severity 'CRIT' -Category 'Seguridad' -Item 'Proteccion en tiempo real de antivirus deshabilitada' -Detail 'RealTimeProtectionEnabled = False' -Recommendation 'Habilitar la proteccion en tiempo real del antivirus/EDR.' -SectionRef '1.11.6'
                    }
                }
                $fechaFirma = ConvertTo-DateSafe (Get-Prop $row @('AntivirusSignatureLastUpdated'))
                if (-not $fechaFirma -and $wdTexto -match 'AntivirusSignatureLastUpdated\s*:\s*([^\r\n]+)') { $fechaFirma = ConvertTo-DateSafe $matches[1].Trim() }
                if ($fechaFirma) {
                    $sinDeteccion = $false
                    $diasFirma = ((Get-Date) - $fechaFirma).TotalDays
                    if ($diasFirma -gt 7) {
                        Add-Finding -Severity 'WARN' -Category 'Seguridad' -Item 'Firmas de antivirus desactualizadas' -Detail "Ultima actualizacion de firmas: $fechaFirma ($([math]::Round($diasFirma,0)) dias atras)" -Recommendation 'Forzar actualizacion de firmas de antivirus.' -SectionRef '1.11.6'
                    }
                }
                if ($wdTexto -and $wdTexto -notmatch '(?i)no disponible|deshabilitado' -and $wdTexto -ne '') { $sinDeteccion = $false }
                $edr = Get-Prop $row @('AgentesEDRTerceros')
                if ($edr -and "$edr" -notmatch 'Ninguno detectado') { $sinDeteccion = $false }
            }
            if ($sinDeteccion) {
                Add-Finding -Severity 'CRIT' -Category 'Seguridad' -Item 'Ningun antivirus/EDR detectado' -Detail 'No se detecto Windows Defender activo ni agente EDR de terceros conocido.' -Recommendation 'Instalar y habilitar un antivirus/EDR corporativo.' -SectionRef '1.11.6'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Seguridad' -Topico 'CertificadosRaiz' -OkMessage 'No se detectaron autoridades certificadoras raiz sospechosas ni una CA raiz autofirmada local no reconocida.' -RequiereSecciones @('1.11.13') -Body {
        Invoke-Rule {
            # OJO: NO se confia en la propiedad 'Sospechoso' que trae el
            # colector: su lista blanca es demasiado corta y marca como
            # sospechosas raices publicas legitimas (SSL.com, SecureTrust,
            # SECOM, HARICA, etc, todas del Microsoft Trusted Root Program).
            # Este motor reclasifica cada CA raiz con su propio criterio,
            # mas amplio, en base a Sujeto/Emisor.
            $cas = ConvertTo-RowArray (Get-SectionData -Id '1.11.13')
            if ($cas.Count -eq 0) { $cas = ConvertTo-RowArray (Resolve-SeccionDatos -TitlePattern 'CA Raiz|Root CA|Autoridad Certificadora') }
            if ($cas.Count -eq 0) { return }

            # Nombre del equipo: una raiz AUTOFIRMADA cuyo CN coincide con el
            # hostname es un indicio fuerte de proxy TLS/interceptacion local
            # instalado a mano (ej. CN=MiEquipo en un equipo llamado MIEQUIPO).
            $hw = Get-SectionData -Id '1.1'
            $hostnameEquipo = "$(Get-Prop $hw @('Hostname'))"
            if (-not $hostnameEquipo) { $hostnameEquipo = "$script:TargetName" }
            if (-not $hostnameEquipo -or $hostnameEquipo -eq 'localhost') { $hostnameEquipo = "$env:COMPUTERNAME" }

            $peligrosas = New-Object System.Collections.Generic.List[string]
            $noReconocidas = New-Object System.Collections.Generic.List[string]

            foreach ($ca in $cas) {
                $sujeto = "$(Get-Prop $ca @('Sujeto','Subject','CN') 'N/D')"
                $emisor = "$(Get-Prop $ca @('Emisor','Issuer') $sujeto)"
                $cn = $sujeto
                if ($cn -match 'CN=([^,]+)') { $cn = $matches[1].Trim() }

                $esAutofirmado = ($sujeto -and $emisor -and $sujeto -eq $emisor)
                $matchHostname = ($hostnameEquipo -and $cn -and ($cn -eq $hostnameEquipo -or $cn -match [regex]::Escape($hostnameEquipo)))
                $matchInterceptor = ($sujeto -match $script:PatronInterceptorTls -or $emisor -match $script:PatronInterceptorTls)

                if ($esAutofirmado -and ($matchHostname -or $matchInterceptor)) {
                    $peligrosas.Add($sujeto)
                    continue
                }

                if (-not (Test-EmisorCAConfiable -Sujeto $sujeto -Emisor $emisor)) {
                    $noReconocidas.Add($sujeto)
                }
            }

            foreach ($p in ($peligrosas | Select-Object -Unique)) {
                Add-Finding -Severity 'WARN' -Category 'Seguridad' -Item "CA raiz autofirmada local sospechosa: $p" -Detail "$p es una raiz autofirmada cuyo nombre coincide con el equipo o con una herramienta de interceptacion TLS conocida." -Recommendation 'Verificar el origen de esta CA raiz local (proxy TLS corporativo autorizado vs. herramienta no autorizada) y removerla si no es legitima.' -Control 'ISO 27001 A.8.24' -SectionRef '1.11.13'
            }

            $noReconocidasUnicas = @($noReconocidas | Select-Object -Unique)
            if ($noReconocidasUnicas.Count -gt 0) {
                $lista = ($noReconocidasUnicas | Select-Object -First 10) -join '; '
                if ($noReconocidasUnicas.Count -gt 10) { $lista += " (y $($noReconocidasUnicas.Count - 10) mas)" }
                Add-Finding -Severity 'INFO' -Category 'Seguridad' -Item 'CA raiz no reconocida como CA publica comercial' -Detail "$lista -- verificar si se trata de una CA interna esperada." -Recommendation 'Confirmar si estas raices corresponden a una PKI corporativa propia o a un proxy TLS autorizado; documentarlas si son legitimas.' -Control 'ISO 27001 A.8.24' -SectionRef '1.11.13'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Seguridad' -Topico 'TpmSecureBoot' -OkMessage 'TPM y Secure Boot presentes/habilitados en hardware compatible.' -RequiereSecciones @('1.11.10') -Body {
        Invoke-Rule {
            $tpm = Resolve-SeccionDatos -IdCandidates @('1.11.10') -TitlePattern 'TPM|Secure Boot'
            if (-not $tpm -or ($tpm -is [string])) { Set-TopicoNoEvaluable; return }
            $filasTpm = ConvertTo-RowArray $tpm
            if ($filasTpm.Count -eq 0) { Set-TopicoNoEvaluable; return }
            # Se exige evidencia afirmativa: si ni TPM ni Secure Boot devolvieron
            # un booleano legible, no se puede sostener el mensaje OK del topico.
            $algoLegible = $false
            foreach ($row in $filasTpm) {
                $pv = Get-Prop $row @('TPM_Presente','TpmPresente','TpmPresent','Presente')
                $sv = Get-Prop $row @('SecureBootHabilitado','SecureBootEnabled')
                if ($pv -is [bool] -or $sv -is [bool]) { $algoLegible = $true }
            }
            if (-not $algoLegible) { Set-TopicoNoEvaluable; return }
            if ($tpm -and -not ($tpm -is [string])) {
                $filas = $filasTpm
                foreach ($row in $filas) {
                    # La propiedad real es 'TPM_Presente' (con guion bajo).
                    $presente = Get-Prop $row @('TPM_Presente','TpmPresente','TpmPresent','Presente')
                    $secureBoot = Get-Prop $row @('SecureBootHabilitado','SecureBootEnabled')
                    if ($presente -eq $false) {
                        Add-Finding -Severity 'INFO' -Category 'Seguridad' -Item 'TPM ausente' -Detail 'No se detecto modulo TPM en el equipo.' -Recommendation 'Evaluar si el hardware soporta TPM y habilitarlo en BIOS/UEFI.' -SectionRef '1.1'
                    }
                    if ($secureBoot -eq $false) {
                        Add-Finding -Severity 'INFO' -Category 'Seguridad' -Item 'Secure Boot deshabilitado' -Detail 'Secure Boot esta deshabilitado en un equipo que lo soporta.' -Recommendation 'Habilitar Secure Boot en BIOS/UEFI si es compatible.' -SectionRef '1.1'
                    }
                }
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Seguridad' -Topico 'LAPS' -OkMessage 'LAPS configurado en el equipo de dominio.' -RequiereSecciones @('1.11.11') -Body {
        Invoke-Rule {
            if (-not $script:Caps -or -not $script:Caps.DomainMember) { return }
            $laps = Resolve-SeccionDatos -IdCandidates @('1.11.11') -TitlePattern 'LAPS'
            if (-not $laps -or ($laps -is [string])) { return }
            $filas = ConvertTo-RowArray $laps
            if ($filas.Count -eq 0) { return }
            # El colector real no expone una propiedad booleana
            # LAPSConfigurado/Configurado/Instalado: expone 'LAPSDetectado'
            # (texto que dice que variante de LAPS se detecto, o que no se
            # detecto ninguna) y 'Habilitado' (booleano o texto segun la
            # variante, con 'Si (...)' cuando hay politica visible).
            $configurado = $false
            foreach ($row in $filas) {
                $detectado = "$(Get-Prop $row @('LAPSDetectado'))"
                if ($detectado -match '(?i)no se detecto') { continue }
                # Se detecto alguna variante de LAPS (legacy o moderno) en el
                # equipo, aunque la politica de GPO no sea visible localmente.
                $configurado = $true
                $habilitado = Get-Prop $row @('Habilitado','LAPSConfigurado','Configurado','Instalado')
                if ($habilitado -eq $false -or "$habilitado" -eq 'False') { $configurado = $false }
            }
            if (-not $configurado) {
                Add-Finding -Severity 'WARN' -Category 'Seguridad' -Item 'LAPS no configurado' -Detail 'El equipo pertenece al dominio pero no se detecto configuracion de LAPS (Local Administrator Password Solution).' -Recommendation 'Implementar LAPS para rotar automaticamente la contrasena del administrador local.' -Control 'ISO 27001 A.5.17' -SectionRef '1.14'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Seguridad' -Topico 'Activacion' -OkMessage 'Windows esta activado.' -RequiereSecciones @('1.11.12') -Body {
        Invoke-Rule {
            $lic = Resolve-SeccionDatos -IdCandidates @('1.11.12') -TitlePattern 'Licencia de Windows|Activacion|License'
            if ($lic -and -not ($lic -is [string])) {
                # SoftwareLicensingProduct devuelve UNA FILA POR PRODUCTO con
                # clave parcial instalado: el sistema operativo, y a veces
                # ademas SKUs de Office u otros complementos (ej. "Office 16,
                # Office16O365HomePremR_Grace edition"). Evaluar todas las
                # filas sin filtrar hace que un Office sin activar se
                # reporte como "Windows sin activar". Se filtra por
                # ApplicationID (GUID fijo de Windows) cuando el colector lo
                # expone; si no, por Nombre empezando con "Windows".
                $filasTodas = ConvertTo-RowArray $lic
                $filasWindows = @($filasTodas | Where-Object {
                    $appId = Get-Prop $_ @('ApplicationID','ApplicationId')
                    if ($appId) { return ("$appId" -eq '55c92734-d682-4d71-983e-d6ec3f16059f') }
                    $nombre = "$(Get-Prop $_ @('Nombre','Name'))"
                    return ($nombre -match '(?i)^\s*Windows')
                })
                foreach ($fila in $filasWindows) {
                    # El colector real mapea el codigo numerico a texto
                    # ('Licenciado', 'Sin licencia', 'Periodo de gracia...',
                    # etc.) en EstadoLicencia; no es el codigo '1' crudo.
                    $estado = Get-Prop $fila @('EstadoLicencia','LicenseStatus')
                    if ($null -ne $estado -and "$estado" -notmatch '(?i)^licenciado$|^licensed$|^1$') {
                        Add-Finding -Severity 'WARN' -Category 'Seguridad' -Item 'Windows sin activar' -Detail "EstadoLicencia = $estado" -Recommendation 'Activar la licencia de Windows.' -SectionRef '1.2'
                    }
                }
            }
        }
    }
    #endregion Seguridad

    #region Certificados
    Invoke-CheckBlock -Categoria 'Certificados' -Topico 'Vigencia' -OkMessage 'No hay certificados vencidos ni por vencer en menos de 90 dias.' -RequiereSecciones @('1.11.1') -Body {
        $certs = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.11.1') -TitlePattern 'Certificados SSL|Certificados')
        foreach ($c in $certs) {
            Invoke-Rule {
                # Get-NumeroSeguro: si DiasParaExpirar viene 'N/D' (el
                # colector no pudo calcularlo), ConvertTo-DoubleSafe lo
                # convertia en 0 y el certificado se reportaba como vencido.
                $dias = Get-NumeroSeguro (Get-Prop $c @('DiasParaExpirar'))
                if ($null -eq $dias) { return }
                $sujeto = Get-Prop $c @('Sujeto','Subject') 'N/D'
                $cn = "$sujeto"
                if ($cn -match 'CN=([^,]+)') { $cn = $matches[1] }
                $detail = "CN=$cn, dias restantes: $dias"
                if ($dias -lt 0) {
                    Add-Finding -Severity 'CRIT' -Category 'Certificados' -Item "Certificado expirado: $cn" -Detail $detail -Recommendation 'Renovar el certificado de inmediato.' -Control 'ISO 27001 A.8.24' -SectionRef '1.11.1'
                } elseif ($dias -lt 30) {
                    Add-Finding -Severity 'CRIT' -Category 'Certificados' -Item "Certificado por expirar en menos de 30 dias: $cn" -Detail $detail -Recommendation 'Renovar el certificado con urgencia.' -Control 'ISO 27001 A.8.24' -SectionRef '1.11.1'
                } elseif ($dias -lt 90) {
                    Add-Finding -Severity 'WARN' -Category 'Certificados' -Item "Certificado por expirar en menos de 90 dias: $cn" -Detail $detail -Recommendation 'Planificar la renovacion del certificado.' -Control 'ISO 27001 A.8.24' -SectionRef '1.11.1'
                }
            }
        }
        Invoke-Rule {
            $sitios = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.6.3','1.6.2.1') -TitlePattern 'Sites Configuration|Configuracion de Sitios')
            foreach ($s in $sitios) {
                $bindings = Get-Prop $s @('Bindings')
                if ($bindings -match 'https' -and $certs) {
                    foreach ($c in $certs) {
                        $auto = Get-Prop $c @('AutoFirmado','SelfSigned')
                        if ($auto -eq $true) {
                            Add-Finding -Severity 'WARN' -Category 'Certificados' -Item 'Sitio IIS usando certificado autofirmado' -Detail "Sitio: $(Get-Prop $s @('Sitio','Name'))" -Recommendation 'Reemplazar el certificado autofirmado por uno emitido por una CA confiable.' -Control 'ISO 27001 A.8.24' -SectionRef '1.6'
                            break
                        }
                    }
                }
            }
        }
    }
    #endregion Certificados

    #region Cuentas
    Invoke-CheckBlock -Categoria 'Cuentas' -Topico 'CuentasLocales' -OkMessage 'No hay cuentas locales habilitadas con contrasena sin expiracion, y la cuenta Administrador esta renombrada.' -RequiereSecciones @('1.14.3') -Body {
        # OJO: 1.14.1 es "Cuentas de Servicio con Privilegios Elevados"
        # (servicios corriendo con cuentas no-sistema), no cuentas de usuario
        # local. El listado de cuentas locales (Usuarios Locales) es 1.14.3.
        $cuentas = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.14.3','1.14.1') -TitlePattern 'Usuarios Locales|Cuentas Locales|Local Accounts|Local Users')
        foreach ($u in $cuentas) {
            Invoke-Rule {
                $nombre = Get-Prop $u @('Nombre','Name') 'N/D'
                $habilitada = Get-Prop $u @('Habilitado','Habilitada','Enabled')
                $sinExpirar = Get-Prop $u @('PasswordNeverExpires')
                if (($habilitada -eq $true -or "$habilitada" -eq 'True') -and ($sinExpirar -eq $true -or "$sinExpirar" -eq 'True')) {
                    Add-Finding -Severity 'WARN' -Category 'Cuentas' -Item "Cuenta local con contrasena sin expiracion: $nombre" -Detail 'PasswordNeverExpires = True' -Recommendation 'Configurar expiracion de contrasena o migrar a un mecanismo de rotacion (LAPS).' -Control 'CIS 1.1 / ISO 27001 A.5.17' -SectionRef '1.14.1'
                }
                if (($nombre -match '^(Administrator|Administrador)$') -and ($habilitada -eq $true -or "$habilitada" -eq 'True')) {
                    Add-Finding -Severity 'WARN' -Category 'Cuentas' -Item 'Cuenta Administrator/Administrador local habilitada sin renombrar' -Detail "Cuenta '$nombre' sigue habilitada con su nombre por defecto." -Recommendation 'Renombrar y/o deshabilitar la cuenta de administrador local por defecto.' -Control 'CIS 2.3.1' -SectionRef '1.14.1'
                }
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Cuentas' -Topico 'AdministradoresLocales' -OkMessage 'El grupo de Administradores locales tiene una membresia acotada y sin cuentas genericas.' -RequiereSecciones @('1.14.2') -Body {
        Invoke-Rule {
            # OJO: 1.14.3 es Usuarios Locales, no miembros del grupo
            # Administradores; no debe usarse como candidato aqui porque
            # tiene una forma de fila distinta (Nombre/Habilitado/... en vez
            # de Nombre/TipoObjeto/Origen/SID) y contaria mal la membresia.
            $admins = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.14.2') -TitlePattern 'Administradores Locales|Local Administrators')
            if ($admins.Count -eq 0) { return }
            if ($admins.Count -gt 5) {
                Add-Finding -Severity 'WARN' -Category 'Cuentas' -Item 'Mas de 5 miembros en el grupo Administradores locales' -Detail "$($admins.Count) miembros detectados" -Recommendation 'Revisar y reducir la membresia del grupo de administradores locales al minimo necesario.' -Control 'CIS 2.3.1' -SectionRef '1.14.2'
            }
            # Una sola finding agrupando las cuentas sospechosas, en vez de una
            # por cuenta: en un servidor con varias, el resumen ejecutivo se
            # llenaba de filas casi identicas.
            $genericas = @()
            foreach ($a in $admins) {
                $nombre = Get-Prop $a @('Nombre','Name','Miembro') 'N/D'
                if ("$nombre" -notmatch '\\') { continue }
                if ("$nombre" -match $script:CuentasIntegradas) { continue }
                if ("$nombre" -match $script:PatronCuentaGenerica) { $genericas += "$nombre" }
            }
            if ($genericas.Count -gt 0) {
                Add-Finding -Severity 'INFO' -Category 'Cuentas' -Item 'Cuentas de servicio o compartidas con acceso administrativo local' -Detail ("$($genericas.Count) cuenta(s): " + (($genericas | Select-Object -First 10) -join ', ')) -Recommendation 'Verificar que el acceso administrativo de estas cuentas este justificado y documentado, y que sus contrasenas se roten.' -Control 'ISO 27001 A.8.2' -SectionRef '1.14.2'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Cuentas' -Topico 'ServiciosDominio' -OkMessage 'Los servicios que corren con cuentas de dominio estan identificados sin hallazgos adicionales.' -RequiereSecciones @('1.14.1','1.2.7') -Body {
        Invoke-Rule {
            # OJO: 1.14.4 es el grupo "Remote Desktop Users" (Nombre/
            # TipoObjeto/Origen/SID), no cuentas de servicio. Las cuentas de
            # servicio con privilegios elevados estan en 1.14.1.
            $servicios = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.14.1') -TitlePattern 'Cuentas de Servicio|Privileged Service Accounts')
            if ($servicios.Count -eq 0) { $servicios = ConvertTo-RowArray (Get-SectionData -Id '1.2.7') }
            $porCuenta = @{}
            foreach ($s in $servicios) {
                $cuenta = Get-Prop $s @('StartName','EjecutarComo')
                if (-not $cuenta -or "$cuenta" -notmatch '\\') { continue }
                if ("$cuenta" -match '^(NT AUTHORITY|NT SERVICE)\\') { continue }
                $nombreServicio = Get-Prop $s @('DisplayName','Name','Nombre') 'N/D'
                if (-not $porCuenta.ContainsKey("$cuenta")) { $porCuenta["$cuenta"] = New-Object System.Collections.Generic.List[string] }
                $porCuenta["$cuenta"].Add("$nombreServicio")
            }
            foreach ($cuenta in $porCuenta.Keys) {
                $lista = ($porCuenta[$cuenta] | Select-Object -First 10) -join ', '
                Add-Finding -Severity 'INFO' -Category 'Cuentas' -Item "Servicios corriendo con cuenta de dominio: $cuenta" -Detail "Servicios: $lista" -Recommendation 'Verificar que la cuenta de servicio tenga solo los privilegios necesarios y su contrasena se gestione de forma segura.' -Control 'ISO 27001 A.8.2' -SectionRef '1.14.4'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Cuentas' -Topico 'PoliticaContrasenas' -OkMessage 'La politica de contrasenas y bloqueo cumple los minimos recomendados.' -RequiereSecciones @('1.14.5') -Body {
        Invoke-Rule {
            # OJO: 1.14.6 es Politica de Auditoria (Categoria/Subcategoria/
            # Configuracion), no politica de contrasenas. La politica de
            # contrasenas y bloqueo es 1.14.5.
            $pol = Resolve-SeccionDatos -IdCandidates @('1.14.5') -TitlePattern 'Politica de Contrasenas|Password Policy'
            if (-not $pol -or $pol -is [string]) { return }
            # Get-NumeroSeguro (no ConvertTo-DoubleSafe) es obligatorio aca: el
            # colector deja 'N/D' cuando no pudo leer la politica real, y
            # ConvertTo-DoubleSafe convertia ese 'N/D' en 0, disparando un
            # WARN falso de "vigencia maxima = 0 dias (nunca expira)".
            $longMin = Get-NumeroSeguro (Get-Prop $pol @('LongitudMinima','MinPasswordLength'))
            $vigMax  = Get-NumeroSeguro (Get-Prop $pol @('VigenciaMaxima','MaxPasswordAge'))
            $bloqueo = Get-NumeroSeguro (Get-Prop $pol @('UmbralBloqueo','LockoutThreshold'))
            if ($null -ne $longMin -and $longMin -lt 14) {
                Add-Finding -Severity 'WARN' -Category 'Cuentas' -Item 'Longitud minima de contrasena menor a 14 caracteres' -Detail "LongitudMinima = $longMin" -Recommendation 'Elevar la longitud minima de contrasena a al menos 14 caracteres.' -Control 'CIS 1.1.x' -SectionRef '1.14.5'
            }
            if ($null -ne $vigMax -and ($vigMax -eq 0 -or $vigMax -gt 365)) {
                Add-Finding -Severity 'WARN' -Category 'Cuentas' -Item 'Vigencia maxima de contrasena fuera de politica' -Detail "VigenciaMaxima = $vigMax dias (0 = nunca expira)" -Recommendation 'Configurar vencimiento de contrasena en un plazo razonable (<= 365 dias).' -Control 'CIS 1.1.x' -SectionRef '1.14.5'
            }
            if ($null -ne $bloqueo -and $bloqueo -eq 0) {
                Add-Finding -Severity 'WARN' -Category 'Cuentas' -Item 'Umbral de bloqueo de cuenta deshabilitado' -Detail 'UmbralBloqueo = 0 (sin bloqueo por intentos fallidos)' -Recommendation 'Configurar un umbral de bloqueo de cuenta (ej. 5-10 intentos).' -Control 'CIS 1.1.x' -SectionRef '1.14.5'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Cuentas' -Topico 'TareasProgramadas' -OkMessage 'No hay tareas programadas de alto privilegio corriendo con cuentas de dominio.' -RequiereSecciones @('1.11.5') -Body {
        Invoke-Rule {
            $tareas = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.11.5') -TitlePattern 'Tareas Programadas')
            foreach ($t in $tareas) {
                $cuenta = Get-Prop $t @('EjecutarComo')
                $nivel = Get-Prop $t @('NivelPrivilegio')
                if ("$cuenta" -match '\\' -and "$nivel" -match 'Highest') {
                    $nombre = Get-Prop $t @('Nombre') 'N/D'
                    Add-Finding -Severity 'INFO' -Category 'Cuentas' -Item "Tarea programada de alto privilegio con cuenta de dominio: $nombre" -Detail "EjecutarComo=$cuenta, NivelPrivilegio=$nivel" -Recommendation 'Verificar la necesidad de ejecutar esta tarea con privilegios elevados y cuenta de dominio.' -SectionRef '1.11.5'
                }
            }
        }
    }
    #endregion Cuentas

    #region Red
    Invoke-CheckBlock -Categoria 'Red' -Topico 'PuertosSensibles' -OkMessage 'No se identificaron puertos sensibles adicionales fuera de lo esperado en la exposicion de red.' -RequiereSecciones @('1.4.6') -Body {
        Invoke-Rule {
            $puertos = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.4.6') -TitlePattern 'Puertos en Escucha|Listening')
            $encontrados = New-Object System.Collections.Generic.List[string]
            foreach ($p in $puertos) {
                $puerto = Get-Prop $p @('PuertoLocal','LocalPort')
                if ($puerto -and ($script:PuertosSensibles -contains ([int]([string]$puerto -replace '\D',''))) ) {
                    if (-not $encontrados.Contains("$puerto")) { $encontrados.Add("$puerto") }
                }
            }
            if ($encontrados.Count -gt 0) {
                Add-Finding -Severity 'INFO' -Category 'Red' -Item 'Puertos sensibles en escucha' -Detail "Puertos: $($encontrados -join ', ')" -Recommendation 'Verificar que estos puertos no esten expuestos innecesariamente fuera de la red confiable.' -SectionRef '1.4.6'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Red' -Topico 'Gateways' -OkMessage 'Un unico default gateway activo.' -RequiereSecciones @('1.4.7') -Body {
        Invoke-Rule {
            $gws = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.4.7') -TitlePattern 'Gateway|Puerta de Enlace')
            if ($gws.Count -gt 1) {
                Add-Finding -Severity 'WARN' -Category 'Red' -Item 'Mas de un default gateway activo' -Detail "$($gws.Count) gateways detectados" -Recommendation 'Verificar la configuracion de rutas; multiples gateways activos pueden causar comportamiento de red inconsistente.' -SectionRef '1.4'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Red' -Topico 'ArchivoHosts' -OkMessage 'No hay entradas personalizadas relevantes en el archivo hosts.' -RequiereSecciones @('1.4.8') -Body {
        Invoke-Rule {
            $hosts = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.4.8') -TitlePattern 'Hosts')
            if ($hosts.Count -gt 0) {
                Add-Finding -Severity 'INFO' -Category 'Red' -Item 'Entradas personalizadas en el archivo hosts' -Detail "$($hosts.Count) entrada(s) detectadas" -Recommendation 'Revisar que las entradas del archivo hosts sean legitimas y esten documentadas.' -SectionRef '1.4'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Red' -Topico 'MTU' -OkMessage 'Todos los adaptadores usan el MTU estandar (1500).' -RequiereSecciones @('1.4.5') -Body {
        Invoke-Rule {
            $mtus = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.4.5') -TitlePattern 'MTU')
            foreach ($m in $mtus) {
                $mtu = Get-NumeroSeguro (Get-Prop $m @('NlMtu','MTU'))
                if ($null -ne $mtu -and $mtu -ne 1500 -and $mtu -ne 0) {
                    $iface = Get-Prop $m @('InterfaceAlias') 'N/D'
                    Add-Finding -Severity 'INFO' -Category 'Red' -Item "Adaptador con MTU distinto de 1500: $iface" -Detail "MTU=$mtu" -Recommendation 'Confirmar que el MTU no estandar sea intencional (jumbo frames, VPN, etc).' -SectionRef '1.4.5'
                }
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Red' -Topico 'DNSPublico' -OkMessage 'Los servidores DNS configurados son internos, no publicos.' -RequiereSecciones @('1.4.4') -Body {
        Invoke-Rule {
            if (-not $script:Caps -or -not $script:Caps.DomainMember) { return }
            $dns = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.4.4') -TitlePattern 'Servidores DNS|DNS Servers')
            foreach ($d in $dns) {
                $servers = Get-Prop $d @('ServerAddresses')
                if ($servers -and ("$servers" -match '8\.8\.8\.8|1\.1\.1\.1')) {
                    $iface = Get-Prop $d @('InterfaceAlias') 'N/D'
                    Add-Finding -Severity 'WARN' -Category 'Red' -Item 'DNS apuntando a un servidor publico en equipo de dominio' -Detail "Interfaz $iface`: $servers" -Recommendation 'Usar los DNS internos del dominio para resolucion correcta de recursos AD.' -SectionRef '1.4.4'
                }
            }
        }
    }
    #endregion Red

    #region Backup y Monitoreo
    Invoke-CheckBlock -Categoria 'Backup' -Topico 'AgenteBackup' -OkMessage 'Se detecto un agente de backup instalado y en ejecucion.' -RequiereSecciones @('1.11.8') -Body {
        Invoke-Rule {
            $bk = Get-SectionData -Id '1.11.8'
            if (-not $bk) { $bk = Resolve-SeccionDatos -TitlePattern 'Backup' }
            if (-not $bk -or ($bk -is [string])) {
                Add-Finding -Severity 'WARN' -Category 'Backup' -Item 'Ningun agente de backup detectado' -Detail 'No se detecto agente de backup (Veeam/Veritas/NetBackup/Backup Exec) instalado como servicio.' -Recommendation 'Confirmar que el servidor este cubierto por la solucion de backup corporativa.' -Control 'ISO 27001 A.8.13' -SectionRef '1.11.8'
                return
            }
            $filas = ConvertTo-RowArray $bk
            foreach ($f in $filas) {
                $estado = Get-Prop $f @('Status','Estado')
                if ($estado -and "$estado" -ne 'Running') {
                    $nombre = Get-Prop $f @('DisplayName','Name') 'N/D'
                    Add-Finding -Severity 'CRIT' -Category 'Backup' -Item "Servicio de backup instalado pero detenido: $nombre" -Detail "Status=$estado" -Recommendation 'Iniciar el servicio de backup y verificar que los trabajos programados se ejecuten.' -Control 'ISO 27001 A.8.13' -SectionRef '1.11.8'
                }
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Backup' -Topico 'Monitoreo' -OkMessage 'Se detecto un agente de monitoreo/EDR conocido.' -RequiereSecciones @('1.11.6') -Body {
        Invoke-Rule {
            $av = Get-SectionData -Id '1.11.6'
            if (-not $av) { $av = Resolve-SeccionDatos -TitlePattern 'Antivirus|Monitoreo|EDR' }
            if (-not $av -or ($av -is [string])) { return }
            $detectado = $false
            if ($av -and -not ($av -is [string])) {
                foreach ($f in (ConvertTo-RowArray $av)) {
                    $edr = Get-Prop $f @('AgentesEDRTerceros')
                    if ($edr -and "$edr" -notmatch 'Ninguno detectado') { $detectado = $true }
                }
            }
            if (-not $detectado) {
                Add-Finding -Severity 'INFO' -Category 'Backup' -Item 'Ningun agente de monitoreo/EDR conocido detectado' -Detail 'No se identifico un agente de monitoreo (SCOM, Zabbix, Datadog, etc) ni EDR de terceros.' -Recommendation 'Confirmar la cobertura de monitoreo de este servidor.' -SectionRef '1.11.6'
            }
        }
    }
    #endregion Backup y Monitoreo

    #region Servicios y Salud
    Invoke-CheckBlock -Categoria 'Disponibilidad' -Topico 'ServiciosDetenidos' -OkMessage 'No hay servicios en modo automatico detenidos de forma anomala.' -RequiereSecciones @('1.2.7') -Body {
        Invoke-Rule {
            $servicios = ConvertTo-RowArray (Get-SectionData -Id '1.2.7')
            $anomalos = New-Object System.Collections.Generic.List[string]
            foreach ($s in $servicios) {
                $modo = Get-Prop $s @('StartMode')
                $estado = Get-Prop $s @('State')
                $nombre = Get-Prop $s @('Name')
                if ("$modo" -eq 'Auto' -and "$estado" -ne 'Running' -and $nombre -and ($script:ServiciosDemandaLegitima -notcontains $nombre)) {
                    $anomalos.Add("$nombre")
                }
            }
            if ($anomalos.Count -gt 0) {
                $lista = ($anomalos | Select-Object -First 10) -join ', '
                Add-Finding -Severity 'WARN' -Category 'Disponibilidad' -Item 'Servicios en modo automatico que no estan corriendo' -Detail "$($anomalos.Count) servicio(s), primeros 10: $lista" -Recommendation 'Investigar por que estos servicios automaticos no estan en ejecucion.' -SectionRef '1.2.7'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Disponibilidad' -Topico 'EventLog' -OkMessage 'No hay eventos criticos repetidos mas de 50 veces en la ventana analizada.' -RequiereSecciones @('1.17.3') -Body {
        Invoke-Rule {
            # OJO: 1.17.2 es "Apagados Inesperados y Reinicios", no errores
            # criticos del Event Log. La tabla real de errores/criticos
            # agrupados (con la columna de cantidad de ocurrencias) es 1.17.3.
            $eventos = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.17.3') -TitlePattern 'Errores Criticos del Event Log|Registro de Eventos|Event Log')
            # Get-NumeroSeguro: una fila con Ocurrencias no numerico/N-D no
            # debe colarse en el "top" (con ConvertTo-DoubleSafe hubiera
            # entrado como 0, que igual la excluye del ">50", pero se prefiere
            # explicito).
            $top = @($eventos | Where-Object { $n = Get-NumeroSeguro (Get-Prop $_ @('Ocurrencias','Count','Cantidad')); $null -ne $n -and $n -gt 50 } |
                Sort-Object { Get-NumeroSeguro (Get-Prop $_ @('Ocurrencias','Count','Cantidad')) } -Descending | Select-Object -First 3)
            foreach ($e in $top) {
                $id = Get-Prop $e @('EventoID','EventId','Id') 'N/D'
                $origen = Get-Prop $e @('Proveedor','Origen','Source','Mensaje') 'N/D'
                $cant = Get-Prop $e @('Ocurrencias','Count','Cantidad')
                Add-Finding -Severity 'WARN' -Category 'Disponibilidad' -Item "Evento critico repetido: ID $id" -Detail "Origen/mensaje: $origen. Ocurrencias: $cant" -Recommendation 'Investigar la causa raiz de este evento recurrente.' -SectionRef '1.17.3'
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Disponibilidad' -Topico 'Performance' -OkMessage 'Memoria, cola de disco y CPU dentro de rangos normales durante el muestreo.' -RequiereSecciones @('1.17.4') -Body {
        Invoke-Rule {
            # OJO: la seccion real es 1.17.4 ("Baseline de Rendimiento"), no
            # 1.17.3 (que es la tabla de errores del Event Log). Ademas su
            # forma es distinta a la asumida: no es un objeto unico con
            # MemoriaDisponiblePct/ColaDiscoPromedio/CPUPromedio, sino UNA FILA
            # POR CONTADOR (Contador/Promedio/Maximo/Minimo/Muestras). Hay que
            # buscar la fila del contador de interes por nombre.
            $perf = ConvertTo-RowArray (Resolve-SeccionDatos -IdCandidates @('1.17.4') -TitlePattern 'Baseline de Rendimiento|Performance|Rendimiento')
            if ($perf.Count -eq 0) { return }

            $filaMem = $perf | Where-Object { "$(Get-Prop $_ @('Contador'))" -match 'Available MBytes' } | Select-Object -First 1
            $filaDisco = $perf | Where-Object { "$(Get-Prop $_ @('Contador'))" -match 'Disk Queue Length' } | Select-Object -First 1
            $filaCpu = $perf | Where-Object { "$(Get-Prop $_ @('Contador'))" -match '% Processor Time' } | Select-Object -First 1

            # El objeto ya no trae el total de memoria fisica para calcular un
            # porcentaje: se conserva la intencion original (memoria
            # disponible critica) sobre el valor absoluto en MB que si expone
            # el colector.
            if ($filaMem) {
                $memMB = Get-NumeroSeguro (Get-Prop $filaMem @('Promedio'))
                if ($null -ne $memMB -and $memMB -gt 0 -and $memMB -lt 500) {
                    Add-Finding -Severity 'CRIT' -Category 'Disponibilidad' -Item 'Memoria disponible critica durante el muestreo' -Detail "Memoria disponible promedio: $memMB MB" -Recommendation 'Investigar consumo de memoria y considerar ampliar RAM.' -SectionRef '1.17.4'
                }
            }
            if ($filaDisco) {
                $colaDisco = Get-NumeroSeguro (Get-Prop $filaDisco @('Promedio'))
                if ($null -ne $colaDisco -and $colaDisco -gt 2) {
                    Add-Finding -Severity 'WARN' -Category 'Disponibilidad' -Item 'Cola de disco promedio elevada' -Detail "Cola de disco promedio: $colaDisco" -Recommendation 'Revisar el subsistema de almacenamiento; puede haber cuello de botella de I/O.' -SectionRef '1.17.4'
                }
            }
            if ($filaCpu) {
                $cpuProm = Get-NumeroSeguro (Get-Prop $filaCpu @('Promedio'))
                if ($null -ne $cpuProm -and $cpuProm -gt 85) {
                    Add-Finding -Severity 'WARN' -Category 'Disponibilidad' -Item 'CPU promedio elevado durante el muestreo' -Detail "CPU promedio: $cpuProm%" -Recommendation 'Investigar procesos consumidores de CPU y evaluar capacidad.' -SectionRef '1.17.4'
                }
            }
        }
    }

    Invoke-CheckBlock -Categoria 'Seguridad' -Topico 'NTP' -OkMessage 'El servicio de hora de Windows esta sincronizado con una fuente valida.' -RequiereSecciones @('1.11.4') -Body {
        Invoke-Rule {
            $ntp = Get-SectionData -Id '1.11.4'
            if (-not $ntp) { $ntp = Resolve-SeccionDatos -TitlePattern 'NTP|Windows Time' }
            if (-not $ntp -or $ntp -is [string]) { return }
            $estadoTxt = "$(Get-Prop $ntp @('Estado'))"
            $fuentesTxt = "$(Get-Prop $ntp @('Fuentes'))"
            $sinSincronizar = ($estadoTxt -match 'Unsynchronized|no sincronizado')
            $reloJLocal = ($fuentesTxt -match 'Local CMOS Clock')
            if ($sinSincronizar) {
                Add-Finding -Severity 'WARN' -Category 'Seguridad' -Item 'Servicio de hora (NTP) no sincronizado' -Detail $estadoTxt -Recommendation 'Verificar el servicio Windows Time y la conectividad hacia el servidor NTP.' -Control 'ISO 27001 A.8.15' -SectionRef '1.11.4'
            }
            if ($reloJLocal -and $script:Caps -and $script:Caps.DomainMember) {
                Add-Finding -Severity 'WARN' -Category 'Seguridad' -Item 'Fuente de hora es el reloj local (CMOS) en equipo de dominio' -Detail $fuentesTxt -Recommendation 'Configurar el equipo para sincronizar contra un controlador de dominio o servidor NTP corporativo.' -Control 'ISO 27001 A.8.15' -SectionRef '1.11.4'
            }
        }
    }
    #endregion Servicios y Salud

    #region Hallazgos OK de resumen por topico de cumplimiento
    foreach ($key in $script:ComplianceChecks.Keys) {
        $chk = $script:ComplianceChecks[$key]
        $evaluado = $true
        try { if ($chk.PSObject.Properties['Evaluado']) { $evaluado = $chk.Evaluado } } catch {}
        if ($evaluado -and -not $chk.TuvoCrit) {
            try { Add-Finding -Severity 'OK' -Category $chk.Categoria -Item $chk.OkMessage -Detail '' } catch {}
        }
    }
    #endregion

    Update-FindingsSummary
    try { Write-ReportLog -Message "Motor de hallazgos finalizado: $($script:FindingsSummary.Total) hallazgo(s) (CRIT=$($script:FindingsSummary.Crit), WARN=$($script:FindingsSummary.Warn), INFO=$($script:FindingsSummary.Info), OK=$($script:FindingsSummary.Ok)). Puntaje de salud: $($script:FindingsSummary.PuntajeSalud)/100." -Level Success } catch {}

    return $script:FindingsSummary
}

function Update-FindingsSummary {
    # Recalcula $script:FindingsSummary a partir del contenido actual de
    # $script:Findings. Se llama al final de Invoke-FindingsEngine y de nuevo
    # al final de Compare-ReportBaseline (que agrega hallazgos de drift), para
    # que el resumen siempre refleje el total real independientemente del
    # orden de ejecucion.
    try {
        $crit = 0; $warn = 0; $info = 0; $ok = 0
        $porCategoria = @{}
        if ($script:Findings) {
            foreach ($f in $script:Findings) {
                switch ($f.Severity) {
                    'CRIT' { $crit++ }
                    'WARN' { $warn++ }
                    'INFO' { $info++ }
                    'OK'   { $ok++ }
                }
                if ($f.Severity -eq 'CRIT' -or $f.Severity -eq 'WARN') {
                    if (-not $porCategoria.ContainsKey($f.Category)) { $porCategoria[$f.Category] = 0 }
                    $porCategoria[$f.Category] = $porCategoria[$f.Category] + 1
                }
            }
        }
        $total = $crit + $warn + $info + $ok
        # Formula de PuntajeSalud (0-100): se parte de 100 y se descuenta segun
        # severidad: cada CRIT resta 12 puntos y cada WARN resta 4. Los INFO
        # no descuentan (son observaciones para documentar, no defectos) y los
        # OK tampoco suman. El resultado nunca baja de 0. Un servidor con
        # muchos INFO pero sin CRIT/WARN conserva puntaje alto.
        $puntaje = 100 - ($crit * 12) - ($warn * 4) - ($info * 0)
        if ($puntaje -lt 0) { $puntaje = 0 }
        $puntaje = [math]::Round($puntaje, 0)

        $script:FindingsSummary = @{
            Crit         = $crit
            Warn         = $warn
            Info         = $info
            Ok           = $ok
            Total        = $total
            PorCategoria = $porCategoria
            PuntajeSalud = $puntaje
        }
    } catch {
        $script:FindingsSummary = @{ Crit=0; Warn=0; Info=0; Ok=0; Total=0; PorCategoria=@{}; PuntajeSalud=100 }
    }
    return $script:FindingsSummary
}

#endregion A

#region B - EXPORTADORES
# =============================================================================
# REGION B: Export-ReportJson / Export-ReportCsv / Export-ReportMarkdown
# =============================================================================

function Get-FileSlug {
    # Convierte un titulo de seccion en un slug seguro para nombre de archivo.
    param([string]$Text)
    if (-not $Text) { return 'Seccion' }
    try {
        $invalidos = [System.IO.Path]::GetInvalidFileNameChars() -join ''
        $patron = '[' + [regex]::Escape($invalidos) + ']'
        $slug = ($Text -replace $patron, '_') -replace '\s+', '_'
        $slug = $slug -replace '_+', '_'
        if ($slug.Length -gt 60) { $slug = $slug.Substring(0, 60) }
        return $slug
    } catch { return 'Seccion' }
}

function Protect-MarkdownPipe {
    # Escapa pipes y saltos de linea para no romper tablas Markdown.
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $t = "$Text" -replace '\|', '\|'
    $t = $t -replace "`r?`n", ' '
    return $t
}

function Get-ParametrosUsados {
    # Intenta capturar los parametros del script como se invoco (si el
    # fragmento CORE expone $script:BoundParameters); si no existe, devuelve
    # una tabla vacia en lugar de fallar.
    try {
        if (-not $script:BoundParameters) { return @{} }
        # Nunca volcar credenciales al JSON: -Credential se reduce al nombre de
        # usuario (enmascarado si corre con -Redact) y jamas a la contrasena.
        $limpio = @{}
        foreach ($k in $script:BoundParameters.Keys) {
            $v = $script:BoundParameters[$k]
            if ($k -match 'Credential|Password|Secret|Token|ApiKey') {
                $usuario = $null
                try { if ($v -is [pscredential]) { $usuario = $v.UserName } } catch {}
                if (-not $usuario) { $usuario = '(omitido)' }
                try { $usuario = Protect-Sensitive -Value $usuario -Kind 'Account' } catch {}
                $limpio[$k] = "$usuario (contrasena no almacenada)"
            } elseif ($v -is [pscredential] -or $v -is [securestring]) {
                $limpio[$k] = '(omitido)'
            } else {
                $limpio[$k] = $v
            }
        }
        return $limpio
    } catch {}
    return @{}
}

function ConvertTo-JsonSafeValue {
    # Devuelve una COPIA de $Value apta para ConvertTo-Json: cualquier
    # [datetime] encontrado (a cualquier profundidad) se convierte a texto
    # ISO 8601 ('yyyy-MM-ddTHH:mm:ss'). Sin esto, ConvertTo-Json de PS 5.1
    # (que serializa fechas via JavaScriptSerializer) las deja como
    # '/Date(1743109017000)/', ilegible para cualquier consumidor.
    # NO muta $Value: arma objetos/arreglos nuevos, asi que es seguro
    # llamarla sobre $script:ReportSections sin romper el orden de tablas
    # del renderer HTML (que necesita los [datetime] reales).
    # Las filas de los colectores son PSCustomObject de propiedades
    # escalares (o arreglos de esas filas), asi que no hace falta soportar
    # tipos .NET arbitrarios en profundidad; el limite de recursion es solo
    # una salvaguarda contra una estructura inesperada.
    param($Value, [int]$Profundidad = 0)
    if ($null -eq $Value) { return $null }
    if ($Profundidad -gt 12) { return $Value }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-ddTHH:mm:ss') }
    if ($Value -is [string]) {
        # Un string devuelto por una sesion remota (Invoke-Remote) llega
        # decorado con NoteProperties de PSRemoting (PSComputerName,
        # RunspaceId, PSShowComputerName) pegadas sobre el mismo objeto
        # string. Sigue siendo -is [string] (el tipo CLR no cambia), pero
        # ConvertTo-Json SI mira esas propiedades extra y lo serializa como
        # un objeto -- {"value":"...", "PSComputerName":"...", ...} -- en
        # vez del texto plano que espera cualquier consumidor (incluida
        # esta misma seccion cuando se recarga desde el JSON: el motor de
        # hallazgos usa "-is [string]" para saber que una seccion es un
        # mensaje, y esa comprobacion fallaba sobre el objeto corrupto).
        # La interpolacion fuerza un string nuevo y limpio, sin esas
        # propiedades pegadas.
        return "$Value"
    }
    # Los enums salen como su NOMBRE ('True', 'Allow'), igual en local y en
    # remoto. Un enum remoto deserializado se serializaria como
    # {"value":1,"Value":"True"} (ver Get-NombreDeEnum) y uno local como un
    # numero, lo que ademas haria distintos entre si un baseline local y uno remoto.
    $nombreEnum = Get-NombreDeEnum $Value
    if ($null -ne $nombreEnum) { return $nombreEnum }
    if ($Value -is [System.Collections.IDictionary]) {
        $nuevo = [ordered]@{}
        foreach ($k in $Value.Keys) { $nuevo[$k] = ConvertTo-JsonSafeValue -Value $Value[$k] -Profundidad ($Profundidad + 1) }
        return $nuevo
    }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        $nuevo = New-Object PSObject
        foreach ($p in $Value.PSObject.Properties) {
            $nuevo | Add-Member -MemberType NoteProperty -Name $p.Name -Value (ConvertTo-JsonSafeValue -Value $p.Value -Profundidad ($Profundidad + 1))
        }
        return $nuevo
    }
    if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string])) {
        $lista = New-Object System.Collections.Generic.List[object]
        foreach ($item in $Value) { $lista.Add((ConvertTo-JsonSafeValue -Value $item -Profundidad ($Profundidad + 1))) }
        return ,$lista.ToArray()
    }
    return $Value
}

# Texto (regex) que identifica entradas de $script:Timings que miden el
# PROCESO COMPLETO de recoleccion (no una seccion puntual). Sumarlas junto
# con las entradas de seccion individuales cuenta el mismo tiempo dos veces:
# en una corrida real de ~126s, el JSON exportado mostraba ~256s si algo
# sumaba todas las filas de Timings a ciegas.
$script:PatronTimingAgregado = '(?i)^Recoleccion de (todas las secciones|todo)\s*$'

function Get-TimingsResumen {
    # Devuelve @{ Marcados=[array]; TotalSegundos=[double]; SumaSeccionesSegundos=[double] }
    #   Marcados              : copia de $script:Timings con la propiedad
    #                           EsAgregado agregada (para que cualquier
    #                           consumidor pueda excluir esas filas al sumar).
    #   TotalSegundos         : tiempo real de recoleccion (de la entrada
    #                           agregada si existe; si no, la suma de las
    #                           secciones individuales).
    #   SumaSeccionesSegundos : suma de las entradas NO agregadas (solo
    #                           informativo -- se solapa con TotalSegundos
    #                           cuando hay entrada agregada, no sumar ambas).
    $marcados = New-Object System.Collections.Generic.List[object]
    $sumaSecciones = 0.0
    $totalAgregado = $null
    try {
        foreach ($t in (ConvertTo-SafeArray $script:Timings)) {
            $nombre = "$(Get-Prop $t @('Seccion'))"
            $segundos = Get-NumeroSeguro (Get-Prop $t @('Segundos'))
            $esAgregado = [bool]($nombre -match $script:PatronTimingAgregado)
            if ($esAgregado) {
                if ($null -ne $segundos -and $null -eq $totalAgregado) { $totalAgregado = $segundos }
            } elseif ($null -ne $segundos) {
                $sumaSecciones += $segundos
            }
            $marcados.Add([PSCustomObject]@{
                Seccion    = $nombre
                Segundos   = (Get-Prop $t @('Segundos'))
                Estado     = (Get-Prop $t @('Estado'))
                EsAgregado = $esAgregado
            })
        }
    } catch {}
    $total = if ($null -ne $totalAgregado) { $totalAgregado } else { $sumaSecciones }
    return @{
        Marcados              = $marcados.ToArray()
        TotalSegundos         = [math]::Round($total, 1)
        SumaSeccionesSegundos = [math]::Round($sumaSecciones, 1)
    }
}

function Export-ReportJson {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    try {
        $secciones = @()
        if ($script:ReportSections) {
            foreach ($sec in $script:ReportSections) {
                $rowCount = 0
                try {
                    if ($sec.Data -is [string] -or $null -eq $sec.Data) { $rowCount = 0 }
                    elseif ($sec.Data -is [System.Collections.IEnumerable]) { $rowCount = @($sec.Data).Count }
                    else { $rowCount = 1 }
                } catch { $rowCount = 0 }
                $secciones += [PSCustomObject]@{
                    Id       = $sec.Id
                    Title    = $sec.Title
                    Note     = $sec.Note
                    RowCount = $rowCount
                    # ConvertTo-JsonSafeValue trabaja sobre una COPIA: no se
                    # muta $sec.Data, que el renderer HTML sigue necesitando
                    # con sus [datetime] reales para ordenar tablas.
                    Data     = (ConvertTo-JsonSafeValue -Value $sec.Data)
                }
            }
        }

        $resumenTimings = Get-TimingsResumen

        $metadata = [PSCustomObject]@{
            ScriptVersion              = $script:ScriptVersion
            ComputerName               = $script:TargetName
            FechaGeneracion            = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
            GeneradoPor                = "$(Get-UsuarioActual) en $(Get-EquipoActual)"
            Formato                    = 'JSON'
            ParametrosUsados           = (Get-ParametrosUsados)
            HashSHA256                 = $null
            # Tiempo real de recoleccion: NO sumar TiempoSumaSeccionesSegundos
            # sobre este valor, se cuenta el mismo trabajo dos veces (una fila
            # de Timings agrega el proceso completo ademas de las filas por
            # seccion que ya suman lo mismo).
            TiempoTotalSegundos        = $resumenTimings.TotalSegundos
            TiempoSumaSeccionesSegundos = $resumenTimings.SumaSeccionesSegundos
        }

        $raiz = [PSCustomObject]@{
            Metadata        = $metadata
            Capabilities    = $script:Caps
            # Timestamp de cada hallazgo es [datetime]: pasa por
            # ConvertTo-JsonSafeValue igual que las secciones.
            Findings        = (ConvertTo-JsonSafeValue -Value (ConvertTo-SafeArray $script:Findings))
            FindingsSummary = $script:FindingsSummary
            Timings         = $resumenTimings.Marcados
            Sections        = (ConvertTo-SafeArray $secciones)
        }

        # PS 5.1: ConvertTo-Json trunca la profundidad por default, asi que se
        # fuerza -Depth alto. Los arreglos se normalizan con ConvertTo-SafeArray
        # (no con el operador coma unario, que anidaria un nivel de mas).
        $json = $raiz | ConvertTo-Json -Depth 10 -Compress:$false
        # Se escribe UTF-8 SIN BOM a mano: Out-File -Encoding UTF8 en PS 5.1
        # siempre antepone BOM, lo que rompe parsers JSON estrictos (p.ej.
        # json.load de Python) y muchos ingestores de CMDB.
        [System.IO.File]::WriteAllText($Path, $json, (New-Object System.Text.UTF8Encoding($false)))

        $hash = $null
        try { $hash = (Get-FileHash -Path $Path -Algorithm SHA256 -ErrorAction Stop).Hash } catch {}
        if ($hash) {
            try { Write-ReportLog -Message "JSON generado en '$Path' (SHA256: $hash)" -Level Success } catch {}
        } else {
            try { Write-ReportLog -Message "JSON generado en '$Path'" -Level Success } catch {}
        }
        return $Path
    } catch {
        try { Write-ReportLog -Message "Error al exportar JSON: $($_.Exception.Message)" -Level Error } catch {}
        return $null
    }
}

function Export-ReportCsv {
    # OJO -Encoding UTF8: en PS 5.1, Export-Csv -Encoding UTF8 SI antepone
    # BOM. A diferencia del JSON (donde el BOM rompe parsers estrictos), en
    # CSV el BOM es DESEADO aca a proposito: es la senal que usa Excel en
    # Windows para reconocer UTF-8 y mostrar bien los acentos/enes; sin BOM,
    # Excel asume ANSI/Windows-1252 y las tildes salen mal. Se deja tal cual
    # (no se cambia a WriteAllText sin BOM como en el JSON).
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    try {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $dir = Split-Path -Path $Path -Parent
        if (-not $dir) { $dir = '.' }
        $csvDir = Join-Path $dir ($baseName + '_csv')
        if (-not (Test-Path $csvDir)) { New-Item -Path $csvDir -ItemType Directory -Force | Out-Null }

        if ($script:ReportSections) {
            foreach ($sec in $script:ReportSections) {
                try {
                    if ($sec.Data -is [string] -or $null -eq $sec.Data) { continue }
                    $filas = ConvertTo-RowArray $sec.Data
                    if ($filas.Count -eq 0) { continue }
                    $slug = Get-FileSlug -Text $sec.Title
                    $nombreArchivo = "$($sec.Id)_$slug.csv" -replace '[\\/:*?"<>|]', '_'
                    $rutaArchivo = Join-Path $csvDir $nombreArchivo
                    $filas | Export-Csv -Path $rutaArchivo -NoTypeInformation -Encoding UTF8 -Force
                } catch {
                    try { Write-ReportLog -Message "No se pudo exportar a CSV la seccion $($sec.Id): $($_.Exception.Message)" -Level Debug } catch {}
                }
            }
        }

        if ($script:Findings -and $script:Findings.Count -gt 0) {
            try { $script:Findings | Export-Csv -Path (Join-Path $csvDir '_hallazgos.csv') -NoTypeInformation -Encoding UTF8 -Force } catch {}
        }

        if ($script:FindingsSummary) {
            # TiempoTotalSegundos NO se suma con las filas de Timings por
            # seccion: ya cuenta el proceso completo (ver Get-TimingsResumen).
            $resumenTimings = Get-TimingsResumen
            $resumen = @(
                [PSCustomObject]@{ Metrica = 'Criticos';                   Valor = $script:FindingsSummary.Crit }
                [PSCustomObject]@{ Metrica = 'Advertencias';               Valor = $script:FindingsSummary.Warn }
                [PSCustomObject]@{ Metrica = 'Informativos';               Valor = $script:FindingsSummary.Info }
                [PSCustomObject]@{ Metrica = 'OK';                         Valor = $script:FindingsSummary.Ok }
                [PSCustomObject]@{ Metrica = 'Total';                      Valor = $script:FindingsSummary.Total }
                [PSCustomObject]@{ Metrica = 'PuntajeSalud';               Valor = $script:FindingsSummary.PuntajeSalud }
                [PSCustomObject]@{ Metrica = 'TiempoTotalSegundos';        Valor = $resumenTimings.TotalSegundos }
                [PSCustomObject]@{ Metrica = 'TiempoSumaSeccionesSegundos'; Valor = $resumenTimings.SumaSeccionesSegundos }
            )
            try { $resumen | Export-Csv -Path (Join-Path $csvDir '_resumen.csv') -NoTypeInformation -Encoding UTF8 -Force } catch {}
        }

        try { Write-ReportLog -Message "CSV generados en '$csvDir'" -Level Success } catch {}
        return $csvDir
    } catch {
        try { Write-ReportLog -Message "Error al exportar CSV: $($_.Exception.Message)" -Level Error } catch {}
        return $null
    }
}

function Export-ReportMarkdown {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    try {
        $sb = New-Object System.Text.StringBuilder

        [void]$sb.AppendLine("# Reporte As-Built - $($script:TargetName)")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("- Version del script: $($script:ScriptVersion)")
        [void]$sb.AppendLine("- Fecha de generacion: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))")
        [void]$sb.AppendLine("- Generado por: $env:USERNAME en $env:COMPUTERNAME")
        if ($script:FindingsSummary) {
            [void]$sb.AppendLine("- Puntaje de salud: $($script:FindingsSummary.PuntajeSalud)/100 (Criticos: $($script:FindingsSummary.Crit), Advertencias: $($script:FindingsSummary.Warn), Informativos: $($script:FindingsSummary.Info), OK: $($script:FindingsSummary.Ok))")
        }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('## Resumen de Hallazgos')
        [void]$sb.AppendLine('')

        if ($script:Findings -and $script:Findings.Count -gt 0) {
            [void]$sb.AppendLine('| Severidad | Categoria | Hallazgo | Detalle | Control |')
            [void]$sb.AppendLine('|---|---|---|---|---|')
            foreach ($f in $script:Findings) {
                $emoji = switch ($f.Severity) {
                    'CRIT' { [System.Char]::ConvertFromUtf32(0x1F534) }
                    'WARN' { [System.Char]::ConvertFromUtf32(0x1F7E1) }
                    'INFO' { [System.Char]::ConvertFromUtf32(0x1F535) }
                    'OK'   { [System.Char]::ConvertFromUtf32(0x1F7E2) }
                    default { '' }
                }
                $item = Protect-MarkdownPipe -Text $f.Item
                $detail = Protect-MarkdownPipe -Text $f.Detail
                $control = Protect-MarkdownPipe -Text $f.Control
                [void]$sb.AppendLine("| $emoji $($f.Severity) | $($f.Category) | $item | $detail | $control |")
            }
        } else {
            [void]$sb.AppendLine('_Sin hallazgos registrados._')
        }
        [void]$sb.AppendLine('')

        if ($script:ReportSections) {
            foreach ($sec in $script:ReportSections) {
                try {
                    $profundidad = ($sec.Id -split '\.').Count
                    $nivel = [Math]::Min(2 + [Math]::Max($profundidad - 1, 0), 6)
                    $numerales = '#' * $nivel
                    [void]$sb.AppendLine("$numerales $($sec.Id) $($sec.Title)")
                    if ($sec.Note) { [void]$sb.AppendLine("_$($sec.Note)_") }
                    [void]$sb.AppendLine('')

                    if ($null -eq $sec.Data -or ($sec.Data -is [array] -and $sec.Data.Count -eq 0)) {
                        [void]$sb.AppendLine('_Sin datos disponibles._')
                    } elseif ($sec.Data -is [string]) {
                        [void]$sb.AppendLine("_$($sec.Data)_")
                    } else {
                        $filas = ConvertTo-RowArray $sec.Data
                        if ($filas.Count -gt 0) {
                            $propiedades = @($filas[0].PSObject.Properties.Name)
                            $omitidas = 0
                            $limitadas = $filas
                            if ($filas.Count -gt 60) {
                                $omitidas = $filas.Count - 60
                                $limitadas = $filas[0..59]
                            }
                            [void]$sb.AppendLine('| ' + ($propiedades -join ' | ') + ' |')
                            [void]$sb.AppendLine('|' + (($propiedades | ForEach-Object { '---' }) -join '|') + '|')
                            foreach ($fila in $limitadas) {
                                $valores = foreach ($p in $propiedades) { Protect-MarkdownPipe -Text ([string]$fila.$p) }
                                [void]$sb.AppendLine('| ' + ($valores -join ' | ') + ' |')
                            }
                            if ($omitidas -gt 0) {
                                [void]$sb.AppendLine('')
                                [void]$sb.AppendLine("_Se omitieron $omitidas fila(s) adicionales por longitud. Ver el reporte HTML o el CSV completo para el detalle._")
                            }
                        } else {
                            [void]$sb.AppendLine('_Sin datos disponibles._')
                        }
                    }
                    [void]$sb.AppendLine('')
                } catch {
                    try { Write-ReportLog -Message "No se pudo renderizar en Markdown la seccion $($sec.Id): $($_.Exception.Message)" -Level Debug } catch {}
                }
            }
        }

        # UTF-8 sin BOM (igual que el JSON): a diferencia del CSV, el
        # Markdown se suele consumir con herramientas de linea de comandos y
        # visores web que no necesitan el BOM y a veces lo muestran como un
        # caracter suelto al inicio del archivo.
        [System.IO.File]::WriteAllText($Path, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
        try { Write-ReportLog -Message "Markdown generado en '$Path'" -Level Success } catch {}
        return $Path
    } catch {
        try { Write-ReportLog -Message "Error al exportar Markdown: $($_.Exception.Message)" -Level Error } catch {}
        return $null
    }
}

#endregion B

#region C - COMPARACION CONTRA BASELINE (DRIFT)
# =============================================================================
# REGION C: Compare-ReportBaseline
# =============================================================================

# Mapa Id de seccion -> propiedad(es) clave para emparejar filas entre el
# baseline y el reporte actual. Para secciones no listadas aqui se usa la
# primera propiedad de tipo string de la primera fila disponible.
$script:DriftKeys = @{
    '1.2.6'  = @('DisplayName')
    '1.2.7'  = @('Name')
    '1.14.3' = @('Nombre')
    '1.14.2' = @('Nombre')
    '1.4.6'  = @('PuertoLocal', 'Protocolo')
    # 1.3.2 ("Reglas de Firewall - Resumen por Perfil") no tiene una columna
    # DisplayName: sus filas son Perfil/Direccion/Accion/Cantidad. Se usa la
    # combinacion Perfil+Direccion+Accion como clave (Cantidad es el valor
    # que cambia, no la clave).
    '1.3.2'  = @('Perfil', 'Direccion', 'Accion')
}

function Get-NoisyPropertyPattern {
    # Nombres de propiedad que cambian por naturaleza (contadores, fechas de
    # muestreo, uptime, etc.) y que NO deben reportarse como drift.
    return 'Uptime|Ultim|Proxim|Fecha|Libre|Disponible|Memoria|CPU|Segundos|Tiempo|Hora|Contador|Duracion'
}

function Get-DriftKeyValue {
    param($Row, [string[]]$KeyProps)
    $valores = foreach ($k in $KeyProps) {
        $p = $Row.PSObject.Properties[$k]
        if ($p) { "$($p.Value)" } else { '' }
    }
    return ($valores -join '|')
}

function Add-DriftFinding {
    param(
        [string]$SeccionId,
        [string]$Titulo,
        [ValidateSet('Agregado', 'Eliminado', 'Modificado')][string]$Tipo,
        [string]$Elemento,
        [string]$Propiedad = ''
    )
    $severidad = 'INFO'
    if ($SeccionId -eq '1.14.2' -or $SeccionId -eq '1.14.3') {
        # Cuenta nueva en Administradores locales -> CRIT segun especificacion.
        $severidad = if ($Tipo -eq 'Agregado') { 'CRIT' } else { 'WARN' }
    } elseif ($SeccionId -eq '1.2.6') {
        $severidad = 'INFO'
    } elseif ($SeccionId -eq '1.2.7') {
        $severidad = if ($Tipo -eq 'Eliminado') { 'WARN' } else { 'INFO' }
    } elseif ($SeccionId -eq '1.4.6') {
        $severidad = if ($Tipo -eq 'Agregado') { 'WARN' } else { 'INFO' }
    } elseif ($SeccionId -eq '1.3.2' -or ($Titulo -and $Titulo -match 'Firewall')) {
        $severidad = if ($Tipo -eq 'Agregado') { 'WARN' } else { 'INFO' }
    } elseif ($Titulo -and $Titulo -match 'Politica de Contrasenas|Auditoria') {
        $severidad = 'WARN'
    } else {
        $severidad = if ($Tipo -eq 'Eliminado') { 'WARN' } else { 'INFO' }
    }
    $item = "$Tipo en '$Titulo': $Elemento"
    $detail = if ($Propiedad) { "Propiedad modificada: $Propiedad" } else { '' }
    try { Add-Finding -Severity $severidad -Category 'Cambios' -Item $item -Detail $detail -SectionRef '1.18' } catch {}
}

function Read-ReportJsonFile {
    <#
        Lee un JSON generado por este script de forma tolerante:
        - descarta el BOM si esta presente (los reportes de versiones previas
          se escribieron con Out-File -Encoding UTF8, que en PS 5.1 pone BOM);
        - si ConvertFrom-Json falla por claves que solo difieren en mayusculas
          (caso tipico: un enum serializado como {"value":0,"Value":"Disabled"}),
          colapsa esos pares a un unico valor de texto y reintenta.
        Devuelve $null si no se pudo leer, nunca lanza.
    #>
    param([Parameter(Mandatory)][string]$Path)

    try {
        if (-not (Test-Path -LiteralPath $Path)) { return $null }
        $raw = [System.IO.File]::ReadAllText($Path)
        if ($raw.Length -gt 0 -and $raw[0] -eq [char]0xFEFF) { $raw = $raw.Substring(1) }

        try {
            return ($raw | ConvertFrom-Json -ErrorAction Stop)
        } catch {
            # No se filtra por el TEXTO del error: PowerShell lo localiza
            # ('different casing' en ingles, 'distintas mayusculas' en PS 7 en
            # espanol, 'claves duplicadas' en PS 5.1), y con un filtro por texto
            # el reintento solo funcionaba en un SO en ingles. Se reintenta ante
            # cualquier fallo; si la normalizacion no cambia nada, el JSON esta
            # roto por otra causa y se relanza el error original.
            $normalizado = [regex]::Replace(
                $raw,
                '\{\s*"value"\s*:\s*[^,}]+,\s*"Value"\s*:\s*("(?:[^"\\]|\\.)*"|[^}]+)\s*\}',
                { param($m) $m.Groups[1].Value })
            if ($normalizado -ceq $raw) { throw }
            Write-ReportLog -Message 'El JSON contiene claves que solo difieren en mayusculas (enum serializado por una version anterior); se normaliza y se reintenta.' -Level Debug
            return ($normalizado | ConvertFrom-Json -ErrorAction Stop)
        }
    } catch {
        try { Write-ReportLog -Message "No se pudo leer el JSON '$Path': $($_.Exception.Message)" -Level Warn } catch {}
        return $null
    }
}

function Compare-ReportBaseline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BaselinePath,
        [Parameter(Mandatory)]$Current
    )
    $filasResultado = New-Object System.Collections.Generic.List[object]
    try {
        if (-not (Test-Path $BaselinePath)) {
            try { Write-ReportLog -Message "No se encontro el archivo de baseline: $BaselinePath" -Level Warn } catch {}
            Add-ReportSection -Id '1.18' -Title 'Cambios respecto al baseline' -Data "No se pudo comparar: no se encontro el archivo de baseline '$BaselinePath'" -Note 'Comparacion contra baseline no realizada'
            return @()
        }

        $baseline = Read-ReportJsonFile -Path $BaselinePath
        if ($null -eq $baseline) { throw "No se pudo leer el baseline '$BaselinePath'." }
        $baselineFecha = 'desconocida'
        try { if ($baseline.Metadata -and $baseline.Metadata.FechaGeneracion) { $baselineFecha = $baseline.Metadata.FechaGeneracion } } catch {}

        $seccionesActuales = $null
        try {
            if ($Current -and $Current.PSObject.Properties['Sections'] -and $Current.Sections) { $seccionesActuales = $Current.Sections }
        } catch {}
        if (-not $seccionesActuales) { $seccionesActuales = $script:ReportSections }
        $seccionesBaseline = $baseline.Sections

        if (-not $seccionesActuales -or -not $seccionesBaseline) {
            try { Write-ReportLog -Message 'Baseline o reporte actual sin secciones utilizables; se omite la comparacion.' -Level Warn } catch {}
            Add-ReportSection -Id '1.18' -Title 'Cambios respecto al baseline' -Data 'No hay secciones comparables entre el baseline y el reporte actual.' -Note "Comparado contra '$BaselinePath' (generado: $baselineFecha)"
            return @()
        }

        $ruidosas = Get-NoisyPropertyPattern

        foreach ($secActual in $seccionesActuales) {
            try {
                $secId = $secActual.Id
                if (-not $secId) { continue }
                if ($secId -match '^1\.17(\.|$)') { continue }

                $secBase = $seccionesBaseline | Where-Object { $_.Id -eq $secId } | Select-Object -First 1
                if (-not $secBase) { continue }
                if ($secActual.Data -is [string] -or $secBase.Data -is [string]) { continue }

                $filasActuales = ConvertTo-RowArray $secActual.Data
                $filasBase = ConvertTo-RowArray $secBase.Data
                if ($filasActuales.Count -eq 0 -and $filasBase.Count -eq 0) { continue }

                $propsClave = $null
                if ($script:DriftKeys.ContainsKey($secId)) {
                    $propsClave = $script:DriftKeys[$secId]
                } else {
                    $muestra = $null
                    if ($filasActuales.Count -gt 0) { $muestra = $filasActuales[0] } elseif ($filasBase.Count -gt 0) { $muestra = $filasBase[0] }
                    if ($muestra) {
                        $primeraPropString = $muestra.PSObject.Properties | Where-Object { $_.Value -is [string] } | Select-Object -First 1
                        if ($primeraPropString) { $propsClave = @($primeraPropString.Name) }
                    }
                }
                if (-not $propsClave) { continue }

                $mapaActual = @{}
                foreach ($r in $filasActuales) { $k = Get-DriftKeyValue -Row $r -KeyProps $propsClave; if ($k) { $mapaActual[$k] = $r } }
                $mapaBase = @{}
                foreach ($r in $filasBase) { $k = Get-DriftKeyValue -Row $r -KeyProps $propsClave; if ($k) { $mapaBase[$k] = $r } }

                foreach ($k in $mapaActual.Keys) {
                    if (-not $mapaBase.ContainsKey($k)) {
                        $filasResultado.Add([PSCustomObject]@{
                            Seccion = $secId; Titulo = $secActual.Title; Tipo = 'Agregado'
                            Elemento = $k; Propiedad = ($propsClave -join '+'); ValorAnterior = ''; ValorActual = $k
                        })
                        Add-DriftFinding -SeccionId $secId -Titulo $secActual.Title -Tipo 'Agregado' -Elemento $k
                    } else {
                        $filaActual = $mapaActual[$k]; $filaBase = $mapaBase[$k]
                        $nombresProps = @($filaActual.PSObject.Properties.Name) + @($filaBase.PSObject.Properties.Name) | Select-Object -Unique
                        foreach ($pName in $nombresProps) {
                            if ($pName -match $ruidosas) { continue }
                            $valActual = $null; $valBase = $null
                            $cp = $filaActual.PSObject.Properties[$pName]; if ($cp) { $valActual = $cp.Value }
                            $bp = $filaBase.PSObject.Properties[$pName]; if ($bp) { $valBase = $bp.Value }
                            if ("$valActual" -ne "$valBase") {
                                $filasResultado.Add([PSCustomObject]@{
                                    Seccion = $secId; Titulo = $secActual.Title; Tipo = 'Modificado'
                                    Elemento = $k; Propiedad = $pName; ValorAnterior = $valBase; ValorActual = $valActual
                                })
                                Add-DriftFinding -SeccionId $secId -Titulo $secActual.Title -Tipo 'Modificado' -Elemento $k -Propiedad $pName
                            }
                        }
                    }
                }
                foreach ($k in $mapaBase.Keys) {
                    if (-not $mapaActual.ContainsKey($k)) {
                        $filasResultado.Add([PSCustomObject]@{
                            Seccion = $secId; Titulo = $secActual.Title; Tipo = 'Eliminado'
                            Elemento = $k; Propiedad = ($propsClave -join '+'); ValorAnterior = $k; ValorActual = ''
                        })
                        Add-DriftFinding -SeccionId $secId -Titulo $secActual.Title -Tipo 'Eliminado' -Elemento $k
                    }
                }
            } catch {
                try { Write-ReportLog -Message "Error comparando la seccion $($secActual.Id) contra el baseline: $($_.Exception.Message)" -Level Debug } catch {}
            }
        }

        $nota = "Comparado contra '$BaselinePath' (generado: $baselineFecha)"
        if ($filasResultado.Count -gt 0) {
            Add-ReportSection -Id '1.18' -Title 'Cambios respecto al baseline' -Data (, @($filasResultado.ToArray())) -Note $nota -Wide
        } else {
            Add-ReportSection -Id '1.18' -Title 'Cambios respecto al baseline' -Data 'No se detectaron cambios respecto al baseline.' -Note $nota
        }

        Update-FindingsSummary
        return $filasResultado.ToArray()
    } catch {
        try { Write-ReportLog -Message "Error al comparar contra el baseline: $($_.Exception.Message)" -Level Error } catch {}
        try { Add-ReportSection -Id '1.18' -Title 'Cambios respecto al baseline' -Data "Error al procesar el baseline: $($_.Exception.Message)" -Note 'Comparacion fallida' } catch {}
        return @()
    }
}

#endregion C

#region D - MODO FLOTA
# =============================================================================
# REGION D: Invoke-FleetReport
# =============================================================================

function New-FleetIndexHtml {
    param([object[]]$Resultados, [string]$OutputPath)
    try {
        # OJO: $script:ReportCss y $script:ReportJs traen el CSS/JS CRUDO, sin
        # etiquetas. Hay que envolverlos en <style>/<script>; inyectarlos pelados
        # deja el indice sin estilos y con el JavaScript visible como texto.
        $cssCrudo = if ($script:ReportCss) { $script:ReportCss } else {
            'body{font-family:"Segoe UI",Arial,sans-serif;background:#f7f7f7;color:#222;margin:20px;}' +
            'table{border-collapse:collapse;width:100%;background:#fff;}th{background:#0f3d63;color:#fff;padding:6px 10px;text-align:left;}' +
            'td{padding:5px 10px;border-bottom:1px solid #ddd;}'
        }
        $cssCrudo += "`n.barra{background:#d8e0ea;border-radius:4px;overflow:hidden;height:14px;width:120px;display:inline-block;vertical-align:middle;}" +
                     "`n.barra-interna{height:100%;}" +
                     "`nbody{display:block;}#contenido{max-width:1400px;margin:0 auto;padding:28px 36px;}"
        $css = "<style>`n$cssCrudo`n</style>"
        $js  = if ($script:ReportJs) { "<script>`n$($script:ReportJs)`n</script>" } else { '' }

        $ordenados = $Resultados | Sort-Object -Property @{Expression = 'PuntajeSalud'; Descending = $true }
        $filasHtml = ''
        foreach ($r in $ordenados) {
            $colorBarra = if ($r.PuntajeSalud -ge 80) { '#2e7d32' } elseif ($r.PuntajeSalud -ge 50) { '#f9a825' } else { '#c62828' }
            $link = 'N/D'
            try {
                if ($r.ReportePath -and (Test-Path $r.ReportePath)) {
                    $link = "<a href='$([System.IO.Path]::GetFileName($r.ReportePath))'>Ver reporte</a>"
                }
            } catch {}
            $filasHtml += "<tr><td>$($r.Equipo)</td><td>$($r.Estado)</td><td>$($r.SO)</td><td>$($r.IP)</td>" +
                "<td><div class='barra'><div class='barra-interna' style='width:$($r.PuntajeSalud)%;background:$colorBarra;'></div></div> $($r.PuntajeSalud)</td>" +
                "<td>$($r.Crit)</td><td>$($r.Warn)</td><td>$($r.EspacioLibreMinGB)</td><td>$($r.Uptime)</td><td>$($r.UltimoParche)</td><td>$link</td></tr>`n"
        }

        $html = @"
<!DOCTYPE html>
<html lang='es'><head><meta charset='UTF-8'><title>Indice de Reportes de Flota</title>$css</head>
<body>
<h1>Indice de Reportes de Flota</h1>
<p>Generado: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) - Equipos: $($Resultados.Count)</p>
<table id='tablaFlota'>
<thead><tr><th>Equipo</th><th>Estado</th><th>SO</th><th>IP</th><th>Puntaje de Salud</th><th>CRIT</th><th>WARN</th><th>Espacio Libre Minimo (GB)</th><th>Uptime</th><th>Ultimo Parche</th><th>Reporte</th></tr></thead>
<tbody>
$filasHtml
</tbody>
</table>
$js
</body></html>
"@
        $ruta = Join-Path $OutputPath '_flota_index.html'
        $html | Out-File -FilePath $ruta -Encoding UTF8 -Force
        return $ruta
    } catch {
        try { Write-ReportLog -Message "Error al generar el indice de flota: $($_.Exception.Message)" -Level Error } catch {}
        return $null
    }
}

# Cuerpo ejecutado por cada equipo de la flota (proceso separado o runspace).
# Recibe todos los datos por parametro (no depende de $script: del proceso
# padre, ya que corre aislado en runspace o proceso hijo).
$script:FleetWorker = {
    param($Equipo, $ScriptPathWorker, $OutDir, $TimeoutSecWorker, $ExtraParamsWorker, $PsExeName)

    $resultado = [PSCustomObject]@{
        Equipo = $Equipo; Estado = 'Error'; SO = ''; IP = ''; PuntajeSalud = 0
        Crit = 0; Warn = 0; EspacioLibreMinGB = $null; Uptime = ''; UltimoParche = ''
        ReportePath = ''; Mensaje = ''
    }
    try {
        $alcanzable = $true
        try { $alcanzable = [bool](Test-Connection -ComputerName $Equipo -Count 1 -Quiet -ErrorAction SilentlyContinue) } catch { $alcanzable = $true }
        if (-not $alcanzable) {
            $resultado.Estado = 'Inalcanzable'
            $resultado.Mensaje = 'No responde a ping ICMP'
            return $resultado
        }

        # Start-Process pasa los argumentos a la linea de comandos tal cual, asi
        # que TODO valor debe ir entrecomillado: sin esto, cualquier ruta con
        # espacios ('C:\Program Files', el perfil del usuario) parte el comando
        # en dos y el proceso hijo falla. Las comillas dobles internas se
        # escapan duplicandolas, que es lo que espera el parser de Windows.
        function Format-ArgumentoCitado {
            param([string]$Valor)
            $limpio = "$Valor" -replace '"', '""'
            return '"' + $limpio + '"'
        }

        $argumentos = @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass',
            '-File',         (Format-ArgumentoCitado $ScriptPathWorker),
            '-ComputerName', (Format-ArgumentoCitado $Equipo),
            '-OutputPath',   (Format-ArgumentoCitado $OutDir),
            '-Quiet'
        )
        if ($ExtraParamsWorker) {
            foreach ($k in $ExtraParamsWorker.Keys) {
                $v = $ExtraParamsWorker[$k]
                if ($v -is [switch]) { if ($v.IsPresent) { $argumentos += "-$k" } }
                elseif ($v -eq $true) { $argumentos += "-$k" }
                elseif ($v -eq $false) { continue }
                elseif ($v -is [array]) { $argumentos += "-$k"; $argumentos += (Format-ArgumentoCitado ($v -join ',')) }
                else { $argumentos += "-$k"; $argumentos += (Format-ArgumentoCitado $v) }
            }
        }

        $proceso = Start-Process -FilePath $PsExeName -ArgumentList $argumentos -PassThru -WindowStyle Hidden
        $termino = $proceso.WaitForExit($TimeoutSecWorker * 1000)
        if (-not $termino) {
            try { $proceso.Kill() } catch {}
            $resultado.Estado = 'Timeout'
            $resultado.Mensaje = "Supero el timeout de $TimeoutSecWorker segundos"
            return $resultado
        }
        if ($proceso.ExitCode -ne 0) {
            $resultado.Estado = 'Error'
            $resultado.Mensaje = "El proceso finalizo con codigo de salida $($proceso.ExitCode)"
            return $resultado
        }

        $jsonFile = Get-ChildItem -Path $OutDir -Filter "*$Equipo*.json" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $jsonFile) {
            $resultado.Estado = 'Error'
            $resultado.Mensaje = 'No se genero el archivo JSON de salida esperado'
            return $resultado
        }

        $data = Read-ReportJsonFile -Path $jsonFile.FullName
        $resultado.Estado = 'OK'
        $resultado.ReportePath = ($jsonFile.FullName -replace '\.json$', '.html')
        try { $resultado.PuntajeSalud = $data.FindingsSummary.PuntajeSalud } catch {}
        try { $resultado.Crit = $data.FindingsSummary.Crit } catch {}
        try { $resultado.Warn = $data.FindingsSummary.Warn } catch {}
        try {
            $osSec = $data.Sections | Where-Object { $_.Id -eq '1.2.1' } | Select-Object -First 1
            if ($osSec -and $osSec.Data) {
                $resultado.SO = $osSec.Data.SistemaOperativo
                $resultado.Uptime = $osSec.Data.UltimoArranque
            }
        } catch {}
        try {
            $ipSec = $data.Sections | Where-Object { $_.Id -eq '1.4.2' } | Select-Object -First 1
            if ($ipSec -and $ipSec.Data) {
                $primeraIp = @($ipSec.Data) | Select-Object -First 1
                if ($primeraIp) { $resultado.IP = $primeraIp.IPAddress }
            }
        } catch {}
        try {
            $volSec = $data.Sections | Where-Object { $_.Id -eq '1.5.2' } | Select-Object -First 1
            if ($volSec -and $volSec.Data) {
                # Sin unidades opticas / ISO montadas (0 GB libres por naturaleza).
                $volsFijos = @($volSec.Data | Where-Object {
                    "$($_.DriveType)" -notmatch 'CD-?ROM|DVD|Removable|Removible|Optic' -and
                    "$($_.FileSystem)" -notmatch '^(?i:CDFS|UDF)$' -and
                    "$($_.FileSystem)".Trim() -ne '' -and
                    [double]$_.TotalGB -ge 1 })
                if ($volsFijos.Count -gt 0) {
                    $resultado.EspacioLibreMinGB = ($volsFijos | Measure-Object -Property LibreGB -Minimum).Minimum
                }
            }
        } catch {}
        try {
            $hfSec = $data.Sections | Where-Object { $_.Id -eq '1.2.2' } | Select-Object -First 1
            if ($hfSec -and $hfSec.Data) {
                $ultimo = (@($hfSec.Data) | Sort-Object InstalledOn -Descending | Select-Object -First 1)
                if ($ultimo) { $resultado.UltimoParche = $ultimo.InstalledOn }
            }
        } catch {}
    } catch {
        $resultado.Estado = 'Error'
        $resultado.Mensaje = $_.Exception.Message
    }
    return $resultado
}

function Invoke-FleetReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$ComputerName,
        [string]$ScriptPath = $PSCommandPath,
        [string]$OutputPath = '.',
        [int]$ThrottleLimit = 8,
        [int]$TimeoutSec = 900,
        [hashtable]$ExtraParams = @{}
    )

    if (-not $ScriptPath) {
        try { Write-ReportLog -Message 'No se puede ejecutar el modo flota: el script fue pegado directamente en la consola (PSCommandPath vacio). Guardelo como archivo .ps1 y vuelva a ejecutarlo, o indique -ScriptPath explicitamente.' -Level Error } catch {}
        return $null
    }
    if (-not (Test-Path $ScriptPath)) {
        try { Write-ReportLog -Message "No se encontro el script en la ruta indicada: $ScriptPath" -Level Error } catch {}
        return $null
    }
    if (-not (Test-Path $OutputPath)) {
        try { New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null } catch {}
    }

    $equipos = @($ComputerName | Where-Object { $_ } | Select-Object -Unique)
    if ($equipos.Count -lt 2) {
        try { Write-ReportLog -Message "El modo flota esta pensado para 2 o mas equipos; se recibieron $($equipos.Count)." -Level Warn } catch {}
    }

    $psExeName = if ($script:IsPS7) { 'pwsh' } else { 'powershell.exe' }
    $resultados = New-Object System.Collections.Generic.List[object]

    try { Write-ReportLog -Message "Iniciando modo flota para $($equipos.Count) equipo(s), ThrottleLimit=$ThrottleLimit, TimeoutSec=$TimeoutSec" -Level Info } catch {}

    if ($script:IsPS7) {
        # PS7: paralelismo nativo con ForEach-Object -Parallel.
        $worker = $script:FleetWorker
        $resultadosParalelos = $equipos | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            $eq = $_
            $wk = $using:worker
            & $wk $eq $using:ScriptPath $using:OutputPath $using:TimeoutSec $using:ExtraParams $using:psExeName
        }
        foreach ($r in $resultadosParalelos) { $resultados.Add($r) }
    } else {
        # PS 5.1: pool de runspaces manual, con timeout por equipo y cancelacion.
        $pool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
        $pool.Open()
        $pendientes = New-Object System.Collections.Generic.List[object]
        try {
            foreach ($equipo in $equipos) {
                $ps = [powershell]::Create()
                $ps.RunspacePool = $pool
                [void]$ps.AddScript($script:FleetWorker).AddArgument($equipo).AddArgument($ScriptPath).AddArgument($OutputPath).AddArgument($TimeoutSec).AddArgument($ExtraParams).AddArgument($psExeName)
                $handle = $ps.BeginInvoke()
                $pendientes.Add([PSCustomObject]@{ PS = $ps; Handle = $handle; Equipo = $equipo; Inicio = Get-Date })
            }

            while ($pendientes.Count -gt 0) {
                Start-Sleep -Milliseconds 500
                for ($i = $pendientes.Count - 1; $i -ge 0; $i--) {
                    $h = $pendientes[$i]
                    $transcurridoSeg = ((Get-Date) - $h.Inicio).TotalSeconds
                    if ($h.Handle.IsCompleted) {
                        $r = $null
                        try { $r = $h.PS.EndInvoke($h.Handle) } catch {}
                        if ($r -and $r.Count -gt 0) {
                            $resultados.Add($r[0])
                        } else {
                            $resultados.Add([PSCustomObject]@{ Equipo = $h.Equipo; Estado = 'Error'; SO = ''; IP = ''; PuntajeSalud = 0; Crit = 0; Warn = 0; EspacioLibreMinGB = $null; Uptime = ''; UltimoParche = ''; ReportePath = ''; Mensaje = 'Sin resultado del runspace' })
                        }
                        try { $h.PS.Dispose() } catch {}
                        $pendientes.RemoveAt($i)
                    } elseif ($transcurridoSeg -gt $TimeoutSec) {
                        try { $h.PS.Stop() } catch {}
                        try { $h.PS.Dispose() } catch {}
                        $resultados.Add([PSCustomObject]@{ Equipo = $h.Equipo; Estado = 'Timeout'; SO = ''; IP = ''; PuntajeSalud = 0; Crit = 0; Warn = 0; EspacioLibreMinGB = $null; Uptime = ''; UltimoParche = ''; ReportePath = ''; Mensaje = "Supero el timeout de $TimeoutSec segundos" })
                        $pendientes.RemoveAt($i)
                    }
                }
            }
        } finally {
            try { $pool.Close() } catch {}
            try { $pool.Dispose() } catch {}
        }
    }

    # --- Consolidacion de hallazgos e inalcanzables de toda la flota ----------
    $todosHallazgos = New-Object System.Collections.Generic.List[object]
    $inalcanzables = New-Object System.Collections.Generic.List[object]

    foreach ($r in $resultados) {
        if ($r.Estado -eq 'OK' -and $r.ReportePath) {
            $jsonPath = $r.ReportePath -replace '\.html$', '.json'
            if (Test-Path $jsonPath) {
                try {
                    $d = Read-ReportJsonFile -Path $jsonPath
                    foreach ($f in @($d.Findings)) {
                        if (-not $f) { continue }
                        $todosHallazgos.Add([PSCustomObject]@{
                            Equipo = $r.Equipo; Severidad = $f.Severity; Categoria = $f.Category
                            Item = $f.Item; Detail = $f.Detail; Recommendation = $f.Recommendation; Control = $f.Control
                        })
                    }
                } catch {}
            }
        } else {
            $inalcanzables.Add([PSCustomObject]@{ Equipo = $r.Equipo; Estado = $r.Estado; Mensaje = $r.Mensaje })
        }
    }

    try {
        if ($todosHallazgos.Count -gt 0) {
            $todosHallazgos | Export-Csv -Path (Join-Path $OutputPath '_flota_hallazgos.csv') -NoTypeInformation -Encoding UTF8 -Force
        }
    } catch {}
    try {
        if ($inalcanzables.Count -gt 0) {
            $inalcanzables | Export-Csv -Path (Join-Path $OutputPath '_flota_inalcanzables.csv') -NoTypeInformation -Encoding UTF8 -Force
        } else {
            [PSCustomObject]@{ Equipo = ''; Estado = ''; Mensaje = 'Ningun equipo inalcanzable' } | Export-Csv -Path (Join-Path $OutputPath '_flota_inalcanzables.csv') -NoTypeInformation -Encoding UTF8 -Force
        }
    } catch {}

    $indexPath = New-FleetIndexHtml -Resultados $resultados.ToArray() -OutputPath $OutputPath

    $ok = @($resultados | Where-Object { $_.Estado -eq 'OK' }).Count
    $conError = @($resultados | Where-Object { $_.Estado -ne 'OK' }).Count

    try { Write-ReportLog -Message "Modo flota finalizado: $ok equipo(s) OK, $conError equipo(s) con error/timeout/inalcanzable." -Level Success } catch {}

    $topCriticos = @($todosHallazgos | Where-Object { $_.Severidad -eq 'CRIT' } | Group-Object Item | Sort-Object Count -Descending | Select-Object -First 5)
    if ($topCriticos.Count -gt 0) {
        try { Write-ReportLog -Message 'Top 5 hallazgos criticos mas repetidos en el parque:' -Level Info } catch {}
        foreach ($t in $topCriticos) {
            try { Write-ReportLog -Message ("  - {0} (x{1} equipos)" -f $t.Name, $t.Count) -Level Info } catch {}
        }
    }

    return [PSCustomObject]@{
        Resultados    = $resultados.ToArray()
        IndexPath     = $indexPath
        Ok            = $ok
        ConError      = $conError
        TopCriticos   = $topCriticos
        Hallazgos     = $todosHallazgos.ToArray()
        Inalcanzables = $inalcanzables.ToArray()
    }
}

#endregion D

# #############################################################################
# ##  BLOQUE PRINCIPAL
# #############################################################################

# =============================================================================
# FRAGMENTO 06 - BLOQUE PRINCIPAL (MAIN)
# Get-ServerFullReport v3.0
# =============================================================================
# Este fragmento es el "pegamento" del script: NO recolecta datos nuevos ni
# define reglas de hallazgos, y NO redefine ninguna funcion de los fragmentos
# CORE (01), ANALISIS (04) ni RENDER (05). Unicamente:
#   - envuelve los colectores sueltos (fragmentos 02 y 03) en una funcion
#     invocable una vez por equipo (Invoke-AllCollectors);
#   - orquesta el flujo completo por equipo (Invoke-SingleTargetReport);
#   - arma los metadatos que necesita el renderer (Get-ReportMeta);
#   - decide entre modo equipo unico y modo flota, e imprime el resumen final
#     por consola.
#
# CODIGOS DE SALIDA DEL PROCESO (utiles para un scheduler o pipeline):
#   0 = Reporte generado sin errores de recoleccion en ninguna seccion.
#   1 = El reporte se genero, pero una o mas secciones tuvieron error durante
#       la recoleccion (ver Write-ReportLog -Level Error / seccion de log
#       incluida en el reporte), o el modo flota reporto equipos con error.
#       Tambien se usa para errores de uso (por ejemplo, modo flota invocado
#       sin que el script este guardado como archivo .ps1).
#   2 = Al menos un equipo objetivo fue inalcanzable (no se pudo abrir sesion
#       remota) y no se genero reporte para el. En modo flota, si CUALQUIER
#       equipo quedo inalcanzable el codigo final es 2 aunque el resto de la
#       flota haya salido bien (es la condicion mas grave a senalizar).
# =============================================================================

# Copia de los parametros usados, para que Get-ParametrosUsados (fragmento 04)
# pueda incluirlos en el JSON exportado. Es solo informativo: si no se usara
# en ningun lado, no rompe nada.
$script:BoundParameters = $PSBoundParameters

#region MAIN-01 - Envoltorio de colectores
function Invoke-AllCollectors {
    <#
        Envoltorio vacio a proposito. Los fragmentos 02 (colectores base) y 03
        (colectores nuevos) estan escritos como codigo suelto a nivel de
        archivo (arriba a abajo, sin funcion contenedora), para poder
        ejecutarlos en cualquier orden durante el ensamblado. Esta funcion los
        agrupa en una unica unidad invocable, de forma que
        Invoke-SingleTargetReport pueda llamarlos una vez por equipo (y en el
        orden correcto) tanto en modo equipo unico como dentro de cada proceso
        hijo del modo flota.

        El script de ensamblado reemplaza el marcador de abajo por el
        contenido de frag_02_collectors_base.ps1 seguido de
        frag_03_collectors_new.ps1 (en ese orden), indentado dentro de esta
        funcion. Este fragmento (06) no escribe ningun colector: solo provee
        el marcador.
    #>
# =============================================================================
# COLECTORES BASE (secciones portadas del script original)
# =============================================================================

# =============================================================================
# FRAGMENTO 02 - COLECTORES BASE (PORTADOS DEL ORIGINAL)
# Hardware (1.1), Sistema Operativo (1.2.x), Firewall (1.3.x), Red (1.4.x),
# Storage (1.5.x), IIS (1.6.x), File Server (1.7.x), DHCP (1.8.x),
# DNS (1.9.x), Terminal Services / RD Licensing (1.10.x),
# Seguridad y Cumplimiento base (1.11.1-1.11.8), WSUS (1.12), Group Policy
# (1.13), Cuentas de servicio privilegiadas (1.14.1) y Entornos Python/Git
# (1.15.x). Portado 1:1 desde original.ps1, sin perder columnas, agregando
# manejo de errores, timeouts explicitos y notas explicativas.
#
# Consume la API compartida del CORE: Invoke-Remote, Add-ReportSection,
# Test-SectionEnabled, Measure-Section, Write-ReportLog, $script:Caps,
# $ScanGitRepos, $GitScanPaths, $PamBrokerEndpoint, $IncludeMissingUpdates.
#
# Este fragmento NO define variables $script: ni funciones globales nuevas.
# Unas pocas variables de trabajo (sin prefijo de scope) se declaran a nivel
# de script para reutilizar datos ya recolectados entre sub-secciones (por
# ejemplo, la lista de shares para iterar permisos por share, o los
# interpretes de Python detectados para la seccion de paquetes pip). No
# colisionan con nombres usados en otros fragmentos.
# =============================================================================

#region 1.1 HARDWARE DEL HOST
if (Test-SectionEnabled '1.1') {

Measure-Section '1.1 Hardware del host' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $cs   = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
            $bios = Get-CimInstance Win32_BIOS -ErrorAction Stop
            $cpu  = Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1
            [PSCustomObject]@{
                Hostname       = $cs.Name
                Fabricante     = $cs.Manufacturer
                Modelo         = $cs.Model
                NumeroSerie    = $bios.SerialNumber
                CPU            = $cpu.Name
                NucleosFisicos = $cpu.NumberOfCores
                NucleosLogicos = $cpu.NumberOfLogicalProcessors
                MemoriaTotalGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
                TipoSistema    = $cs.SystemType
                MaquinaVirtual = ($cs.Model -match 'Virtual|VMware|KVM')
            }
        } catch {
            "No se pudo obtener la informacion de hardware (Win32_ComputerSystem/Win32_BIOS/Win32_Processor): $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.1' -Title 'Hardware del Host' -Data $data -Note 'Identificacion fisica del equipo (fabricante, modelo, numero de serie) y capacidad de CPU/memoria. Es la base para contrastar contra el inventario de activos.'
}

}
#endregion

#region 1.2 SISTEMA OPERATIVO
if (Test-SectionEnabled '1.2') {

# -----------------------------------------------------------------------------
# 1.2.1 Configuracion del sistema operativo
# -----------------------------------------------------------------------------
Measure-Section '1.2.1 Configuracion del sistema operativo' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
            $tz = Get-TimeZone -ErrorAction Stop
            $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
            $uptimeDias = 'N/D'
            if ($os.LastBootUpTime) { $uptimeDias = [math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 1) }
            [PSCustomObject]@{
                SistemaOperativo  = $os.Caption
                Version           = $os.Version
                Build             = $os.BuildNumber
                Arquitectura      = $os.OSArchitecture
                FechaInstalacion  = $os.InstallDate
                UltimoArranque    = $os.LastBootUpTime
                UptimeDias        = $uptimeDias
                DirectorioWindows = $os.WindowsDirectory
                ZonaHoraria       = $tz.Id
                NombreDominio     = $cs.Domain
                RolDominio        = $cs.DomainRole
            }
        } catch {
            "No se pudo obtener la configuracion del sistema operativo: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.2.1' -Title 'Configuracion del Sistema Operativo' -Data $data -Note 'Version/build exacta del SO, zona horaria y rol en el dominio. UptimeDias alto sin ventanas de mantenimiento recientes suele anticipar parches pendientes de aplicar via reinicio.'
}

# -----------------------------------------------------------------------------
# 1.2.2 Hotfixes instalados
# -----------------------------------------------------------------------------
Measure-Section '1.2.2 Hotfixes instalados' {
    $data = Invoke-Remote -TimeoutSec 90 -ScriptBlock {
        try {
            Get-HotFix -ErrorAction Stop | Select-Object HotFixID, Description, InstalledBy, InstalledOn |
                Sort-Object InstalledOn -Descending
        } catch { "No se pudo consultar el historial de hotfixes: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.2.2' -Title 'Hotfixes Instalados' -Data $data -Note 'Actualizaciones aplicadas al sistema operativo, de la mas reciente a la mas antigua.'
}

# -----------------------------------------------------------------------------
# 1.2.3 Actualizaciones de Windows faltantes (opt-in via -IncludeMissingUpdates)
# -----------------------------------------------------------------------------
Measure-Section '1.2.3 Actualizaciones de Windows faltantes' {
    if (-not $IncludeMissingUpdates) {
        Add-ReportSection -Id '1.2.3' -Title 'Actualizaciones de Windows Faltantes' -Data 'Seccion omitida por rendimiento: la consulta al COM de Windows Update (Microsoft.Update.Session) puede tardar varios minutos por servidor, lo cual no escala en modo flota. IMPORTANTE: sin este dato el reporte NO puede evaluar actualizaciones de seguridad pendientes ni generar hallazgos de parches faltantes. Para incluirlo, volver a ejecutar el script agregando el switch -IncludeMissingUpdates (ejemplo: .\Get-ServerFullReport.ps1 -IncludeMissingUpdates).' -Note 'Requiere el switch -IncludeMissingUpdates (opt-in). Sin el, esta seccion se omite para no demorar corridas masivas contra muchos servidores, y el motor de hallazgos no puede evaluar parches faltantes.'
    } else {
        $data = Invoke-Remote -TimeoutSec 300 -ScriptBlock {
            try {
                $session  = New-Object -ComObject Microsoft.Update.Session
                $searcher = $session.CreateUpdateSearcher()
                $result   = $searcher.Search("IsInstalled=0 and IsHidden=0")
                $result.Updates | ForEach-Object {
                    [PSCustomObject]@{
                        Titulo      = $_.Title
                        KB          = ($_.KBArticleIDs -join ', ')
                        Severidad   = $_.MsrcSeverity
                        FechaPublic = $_.LastDeploymentChangeTime
                    }
                }
            } catch {
                "No se pudo consultar Windows Update (servicio detenido, WSUS sin contacto, o permisos insuficientes): $($_.Exception.Message)"
            }
        }
        Add-ReportSection -Id '1.2.3' -Title 'Actualizaciones de Windows Faltantes' -Data $data -Note 'Actualizaciones pendientes de instalar segun el agente local de Windows Update (o el WSUS configurado como origen, si aplica). Parches criticos/de seguridad faltantes son el hallazgo de compliance mas comun.'
    }
}

# -----------------------------------------------------------------------------
# 1.2.4 Drivers instalados
# -----------------------------------------------------------------------------
Measure-Section '1.2.4 Drivers' {
    $data = Invoke-Remote -TimeoutSec 120 -ScriptBlock {
        try {
            Get-CimInstance Win32_PnPSignedDriver -ErrorAction Stop | Where-Object { $_.DeviceName } |
                Select-Object DeviceName, DriverVersion, Manufacturer, DriverDate, DeviceClass, IsSigned |
                Sort-Object DeviceName -Unique
        } catch { "No se pudo enumerar los drivers instalados (Win32_PnPSignedDriver): $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.2.4' -Title 'Drivers Instalados' -Data $data -Note 'Version y estado de firma de cada driver PnP. Util para detectar drivers sin firmar o desactualizados tras un cambio de hardware o de hipervisor.' -Wide
}

# -----------------------------------------------------------------------------
# 1.2.5 Roles y caracteristicas instaladas
# -----------------------------------------------------------------------------
Measure-Section '1.2.5 Roles y caracteristicas instaladas' {
    $data = Invoke-Remote -TimeoutSec 120 -ScriptBlock {
        try {
            Import-Module ServerManager -ErrorAction Stop
            Get-WindowsFeature -ErrorAction Stop | Where-Object { $_.InstallState -eq 'Installed' } |
                Select-Object DisplayName, Name, FeatureType, Path
        } catch {
            "Modulo ServerManager no disponible (edicion Desktop/Client, no Windows Server) o error al consultar: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.2.5' -Title 'Roles y Caracteristicas Instaladas' -Data $data -Note 'Roles y features de Windows Server presentes en el equipo. En Windows 10/11 (edicion cliente) esta seccion no aplica porque el modulo ServerManager no existe.' -Wide
}

# -----------------------------------------------------------------------------
# 1.2.6 Aplicaciones instaladas
# -----------------------------------------------------------------------------
Measure-Section '1.2.6 Aplicaciones instaladas' {
    $data = Invoke-Remote -TimeoutSec 120 -ScriptBlock {
        try {
            $paths = @(
                @{ RegPath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'; Arq = 'x64' },
                @{ RegPath = 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'; Arq = 'x86' },
                @{ RegPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'; Arq = 'x64 (usuario actual)' }
            )
            $rows = New-Object System.Collections.Generic.List[PSObject]
            $vistos = New-Object System.Collections.Generic.HashSet[string]
            foreach ($p in $paths) {
                Get-ItemProperty -Path $p.RegPath -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName } | ForEach-Object {
                    $clave = "$($_.DisplayName)|$($_.DisplayVersion)|$($p.Arq)"
                    if ($vistos.Add($clave)) {
                        $esActualizacion = ($_.SystemComponent -eq 1) -or [bool]$_.ParentKeyName
                        $rows.Add([PSCustomObject]@{
                            DisplayName                          = $_.DisplayName
                            DisplayVersion                       = $_.DisplayVersion
                            Publisher                            = $_.Publisher
                            InstallDate                          = $_.InstallDate
                            InstallLocation                      = $_.InstallLocation
                            Arquitectura                         = $p.Arq
                            EsComponenteDelSistemaOActualizacion = $esActualizacion
                        })
                    }
                }
            }
            $rows | Sort-Object DisplayName
        } catch { "No se pudo enumerar las aplicaciones instaladas via registro: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.2.6' -Title 'Aplicaciones Instaladas' -Data $data -Note 'Inventario de software via registro (claves Uninstall, 64/32 bits). EsComponenteDelSistemaOActualizacion=true marca parches o componentes que normalmente no son software "de negocio" (SystemComponent=1 o poseen ParentKeyName); se conservan en la lista para no perder informacion, filtrar por esa columna si se busca solo software instalado explicitamente.' -Wide
}

# -----------------------------------------------------------------------------
# 1.2.7 Servicios
# -----------------------------------------------------------------------------
Measure-Section '1.2.7 Servicios' {
    $data = Invoke-Remote -TimeoutSec 120 -ScriptBlock {
        try {
            Get-CimInstance Win32_Service -ErrorAction Stop | ForEach-Object {
                $ruta = $_.PathName
                $sinComillas = $false
                if ($ruta -and -not ($ruta.TrimStart().StartsWith('"'))) {
                    if ($ruta -match '^([^"]*?\.exe)') {
                        $exePart = $matches[1]
                        if ($exePart -match '\s') { $sinComillas = $true }
                    }
                }
                [PSCustomObject]@{
                    Name            = $_.Name
                    DisplayName     = $_.DisplayName
                    State           = $_.State
                    StartMode       = $_.StartMode
                    StartName       = $_.StartName
                    PathName        = $ruta
                    RutaSinComillas = $sinComillas
                }
            } | Sort-Object State, DisplayName
        } catch { "No se pudo enumerar los servicios (Win32_Service): $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.2.7' -Title 'Servicios' -Data $data -Note 'Todos los servicios de Windows con su ruta de ejecutable. RutaSinComillas=true marca un "unquoted service path": una ruta con espacios sin comillas donde un atacante local con permiso de escritura en una carpeta intermedia puede plantar un ejecutable y escalar privilegios.' -Wide
}

}
#endregion

#region 1.3 FIREWALL DE WINDOWS
if (Test-SectionEnabled '1.3') {

# -----------------------------------------------------------------------------
# 1.3.1 Perfiles de firewall
# -----------------------------------------------------------------------------
Measure-Section '1.3.1 Perfiles de firewall' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            Get-NetFirewallProfile -ErrorAction Stop | Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction, LogFileName, LogAllowed, LogBlocked
        } catch { "No se pudo consultar los perfiles de firewall: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.3.1' -Title 'Perfiles de Firewall' -Data $data -Note 'Estado (Habilitado/Deshabilitado) y accion por defecto de cada perfil de red (Dominio/Privado/Publico). Un perfil deshabilitado deja al equipo sin filtrado alguno en esa red.'
}

# -----------------------------------------------------------------------------
# 1.3.2 Reglas de firewall - resumen por perfil/direccion/accion
# -----------------------------------------------------------------------------
Measure-Section '1.3.2 Reglas de firewall - resumen' {
    $data = Invoke-Remote -TimeoutSec 90 -ScriptBlock {
        try {
            Get-NetFirewallRule -ErrorAction Stop | Where-Object { $_.Enabled -eq 'True' } |
                Group-Object Profile, Direction, Action | ForEach-Object {
                    $partes = $_.Name -split ', '
                    [PSCustomObject]@{
                        Perfil    = $partes[0]
                        Direccion = $partes[1]
                        Accion    = $partes[2]
                        Cantidad  = $_.Count
                    }
                } | Sort-Object Perfil, Direccion, Accion
        } catch { "No se pudo obtener el resumen de reglas de firewall: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.3.2' -Title 'Reglas de Firewall - Resumen por Perfil' -Data $data -Note 'Conteo de reglas habilitadas agrupadas por perfil, direccion y accion. Sirve para detectar de un vistazo un numero anormal de reglas de bloqueo/permiso frente a otros servidores de la misma familia.'
}

# -----------------------------------------------------------------------------
# 1.3.3 Reglas de firewall entrantes - detalle con puertos y direcciones
# -----------------------------------------------------------------------------
Measure-Section '1.3.3 Reglas de firewall entrantes - detalle' {
    $data = Invoke-Remote -TimeoutSec 120 -ScriptBlock {
        try {
            $rules = @(Get-NetFirewallRule -ErrorAction Stop | Where-Object { $_.Enabled -eq 'True' -and $_.Direction -eq 'Inbound' })
            if ($rules.Count -eq 0) { return 'No hay reglas de firewall entrantes habilitadas en este equipo.' }

            # Traer los filtros de puerto/direccion UNA sola vez e indexarlos por
            # InstanceID en un hashtable, en lugar de consultar
            # Get-NetFirewallPortFilter/-AddressFilter por cada regla (con 500+
            # reglas eso puede tardar varios minutos).
            $portIndex = @{}
            foreach ($pf in (Get-NetFirewallPortFilter -ErrorAction SilentlyContinue)) { $portIndex[$pf.InstanceID] = $pf }
            $addrIndex = @{}
            foreach ($af in (Get-NetFirewallAddressFilter -ErrorAction SilentlyContinue)) { $addrIndex[$af.InstanceID] = $af }

            $rows = New-Object System.Collections.Generic.List[PSObject]
            foreach ($r in $rules) {
                $pf = $portIndex[$r.InstanceID]
                $af = $addrIndex[$r.InstanceID]
                $rows.Add([PSCustomObject]@{
                    DisplayName     = $r.DisplayName
                    Direccion       = $r.Direction
                    Accion          = $r.Action
                    Perfil          = $r.Profile
                    Protocolo       = if ($pf) { $pf.Protocol } else { 'N/D' }
                    PuertoLocal     = if ($pf) { ($pf.LocalPort -join ', ') } else { 'N/D' }
                    PuertoRemoto    = if ($pf) { ($pf.RemotePort -join ', ') } else { 'N/D' }
                    DireccionLocal  = if ($af) { ($af.LocalAddress -join ', ') } else { 'N/D' }
                    DireccionRemota = if ($af) { ($af.RemoteAddress -join ', ') } else { 'N/D' }
                })
            }
            $rows | Sort-Object Direccion, DisplayName
        } catch { "No se pudo obtener el detalle de reglas de firewall entrantes: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.3.3' -Title 'Reglas de Firewall Entrantes - Detalle' -Data $data -Note 'Puertos y direcciones IP locales/remotas de cada regla entrante habilitada (join con Get-NetFirewallPortFilter/-AddressFilter, indexado por InstanceID para evitar demoras). Limitado a reglas entrantes para mantener el reporte legible; ver 1.3.2 para el panorama completo por perfil.' -Wide
}

}
#endregion

#region 1.4 RED DEL HOST
if (Test-SectionEnabled '1.4') {

Measure-Section '1.4.1 Adaptadores de red' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-NetAdapter -ErrorAction Stop | Select-Object Name, InterfaceDescription, Status, LinkSpeed, MacAddress, MediaType, InterfaceIndex }
        catch { "No se pudieron enumerar los adaptadores de red: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.4.1' -Title 'Adaptadores de Red' -Data $data -Note 'Estado fisico y velocidad negociada de cada NIC. Un adaptador "Disconnected" que deberia estar activo suele explicar perdida de conectividad redundante.'
}

Measure-Section '1.4.2 Direcciones IP' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop | Select-Object InterfaceAlias, IPAddress, PrefixLength, PrefixOrigin, SuffixOrigin, AddressState }
        catch { "No se pudieron enumerar las direcciones IPv4: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.4.2' -Title 'Direcciones IP (IPv4)' -Data $data -Note 'PrefixOrigin=Dhcp en un servidor suele ser en si mismo un hallazgo: los servidores normalmente deben operar con IP estatica o reserva DHCP.'
}

Measure-Section '1.4.3 Cliente DNS' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DnsClient -ErrorAction Stop | Select-Object InterfaceAlias, ConnectionSpecificSuffix, RegisterThisConnectionsAddress, UseSuffixWhenRegistering }
        catch { "No se pudo obtener la configuracion del cliente DNS: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.4.3' -Title 'Configuracion del Cliente DNS' -Data $data -Note 'Sufijo DNS por interfaz y si esa interfaz registra su propia direccion en DNS dinamico.'
}

Measure-Section '1.4.4 Servidores DNS configurados' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction Stop | Select-Object InterfaceAlias, ServerAddresses }
        catch { "No se pudieron obtener los servidores DNS configurados: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.4.4' -Title 'Servidores DNS Configurados por Interfaz' -Data $data -Note 'DNS configurados manualmente o recibidos por DHCP en cada interfaz. La ausencia de un DNS secundario es un punto unico de falla de resolucion.'
}

Measure-Section '1.4.5 MTU de adaptadores' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-NetIPInterface -AddressFamily IPv4 -ErrorAction Stop | Select-Object InterfaceAlias, NlMtu, Dhcp, ConnectionState }
        catch { "No se pudo obtener el MTU de los adaptadores: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.4.5' -Title 'MTU de Adaptadores de Red' -Data $data -Note 'Un MTU distinto entre nodos que se comunican entre si (por ejemplo, storage iSCSI con jumbo frames) causa fragmentacion o perdida silenciosa de paquetes grandes.'
}

}
#endregion

#region 1.5 STORAGE DEL HOST
if (Test-SectionEnabled '1.5') {

Measure-Section '1.5.1 Discos locales' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            Get-CimInstance Win32_DiskDrive -ErrorAction Stop | Select-Object DeviceID, Model, InterfaceType, SerialNumber, MediaType,
                @{N='TamanoGB';E={[math]::Round($_.Size/1GB,2)}}, Partitions
        } catch { "No se pudieron enumerar los discos fisicos: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.5.1' -Title 'Discos Locales' -Data $data -Note 'Inventario fisico de discos (modelo, interfaz, tamano, numero de particiones). Es la base para planificar ampliaciones de storage.'
}

Measure-Section '1.5.2 Volumenes del host' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            Get-Volume -ErrorAction Stop | Where-Object { $_.DriveLetter } | Select-Object DriveLetter, FileSystemLabel, FileSystem, DriveType, HealthStatus,
                @{N='TotalGB';E={[math]::Round($_.Size/1GB,2)}},
                @{N='LibreGB';E={[math]::Round($_.SizeRemaining/1GB,2)}},
                @{N='LibrePct';E={ if ($_.Size -gt 0) { [math]::Round(($_.SizeRemaining/$_.Size)*100,1) } else { 0 } }}
        } catch { "No se pudieron enumerar los volumenes: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.5.2' -Title 'Volumenes del Host' -Data $data -Note 'Espacio libre por unidad. LibrePct por debajo del 10-15% es el hallazgo de storage mas comun y de resolucion mas urgente.' -Wide
}

}
#endregion

#region 1.6 IIS (solo si el rol Web Server esta instalado)
if (Test-SectionEnabled '1.6') {

if (-not $script:Caps.IIS) {
    Add-ReportSection -Id '1.6' -Title 'Servidor IIS' -Data 'El rol de servidor Web (IIS) no esta instalado en este equipo.' -Note 'Se documenta explicitamente la ausencia del rol para que el as-built quede completo.'
} else {

Measure-Section '1.6.1 IIS - Application Pools' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            Import-Module WebAdministration -ErrorAction Stop
            Get-ChildItem IIS:\AppPools -ErrorAction Stop | Select-Object Name,
                @{N='EstadoActual';E={(Get-WebAppPoolState -Name $_.Name -ErrorAction SilentlyContinue).Value}},
                managedRuntimeVersion, managedPipelineMode,
                @{N='IdentityType';E={$_.processModel.identityType}},
                @{N='IdentityUser';E={$_.processModel.userName}},
                @{N='IdleTimeoutMin';E={$_.processModel.idleTimeout.TotalMinutes}},
                @{N='PeriodicRestartMin';E={$_.recycling.periodicRestart.time.TotalMinutes}}
        } catch { "No se pudo consultar los Application Pools de IIS: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.6.1' -Title 'Application Pools de IIS' -Data $data -Note 'Estado y configuracion de identidad/reciclado de cada pool. Un pool detenido (Stopped) explica sitios web caidos aunque el servicio W3SVC este corriendo con normalidad.' -Wide
}

Measure-Section '1.6.2 IIS - sitios (resumen)' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            Import-Module WebAdministration -ErrorAction Stop
            Get-Website -ErrorAction Stop | Select-Object Name, ID, State, PhysicalPath,
                @{N='Bindings';E={ ($_.bindings.Collection | ForEach-Object { "$($_.protocol) $($_.bindingInformation)" }) -join ' | ' }},
                @{N='ApplicationPool';E={$_.applicationPool}}
        } catch { "No se pudo consultar los sitios de IIS: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.6.2' -Title 'Sitios de IIS - Resumen' -Data $data -Note 'Listado de sitios, bindings y pool asignado a cada uno.' -Wide
}

Measure-Section '1.6.2.1 IIS - sitios (detalle y aplicaciones)' {
    $data = Invoke-Remote -TimeoutSec 90 -ScriptBlock {
        try {
            Import-Module WebAdministration -ErrorAction Stop
            Get-Website -ErrorAction Stop | ForEach-Object {
                $site = $_
                $apps = Get-WebApplication -Site $site.Name -ErrorAction SilentlyContinue |
                    Select-Object Path, ApplicationPool, PhysicalPath
                [PSCustomObject]@{
                    Sitio        = $site.Name
                    Estado       = $site.State
                    RutaFisica   = $site.PhysicalPath
                    Bindings     = ($site.bindings.Collection | ForEach-Object { "$($_.protocol) $($_.bindingInformation)" }) -join ' | '
                    Aplicaciones = if ($apps) { ($apps | ForEach-Object { "$($_.Path) -> Pool:$($_.ApplicationPool)" }) -join ' ; ' } else { '(ninguna adicional a la raiz)' }
                }
            }
        } catch { "No se pudo obtener el detalle de sitios de IIS: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.6.2.1' -Title 'Sitios de IIS - Detalle y Aplicaciones Virtuales' -Data $data -Note 'Incluye las aplicaciones virtuales dentro de cada sitio y a que pool esta asignada cada una, importante cuando un sitio corre codigo con distintos niveles de aislamiento.' -Wide
}

}

}
#endregion

#region 1.7 FILE SERVER
if (Test-SectionEnabled '1.7') {

if (-not $script:Caps.FileServer) {
    Add-ReportSection -Id '1.7' -Title 'Servidor de Archivos (File Server)' -Data 'No se detectaron recursos compartidos de datos (fuera de los administrativos) en este equipo.' -Note 'Se documenta explicitamente la ausencia de shares de datos para que el as-built quede completo.'
} else {

Measure-Section '1.7.1 SMB - configuracion del servidor' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            Get-SmbServerConfiguration -ErrorAction Stop | Select-Object EnableSMB1Protocol, EnableSMB2Protocol, ServerHidden, AnnounceServer, EncryptData, RejectUnencryptedAccess, AuditSmb1Access
        } catch { "No se pudo obtener la configuracion SMB del servidor: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.7.1' -Title 'Configuracion SMB del Servidor' -Data $data -Note 'EnableSMB1Protocol=true es un hallazgo critico de seguridad: SMB1 es un protocolo obsoleto y vulnerable, deshabilitado por defecto desde Windows Server 2016/2019.'
}

# 1.7.2 se guarda en una variable de script para reutilizarla en el detalle
# de permisos por share que sigue a continuacion.
$fileShares = Measure-Section '1.7.2 Recursos compartidos (shares)' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            Get-SmbShare -ErrorAction Stop | Where-Object { $_.Name -notmatch '^[A-Z]\$$|^ADMIN\$$|^IPC\$$|^PRINT\$$' } |
                Select-Object Name, Path, Description, ShareType, CurrentUsers, FolderEnumerationMode, CachingMode
        } catch { "No se pudo enumerar los recursos compartidos (Get-SmbShare): $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.7.2' -Title 'Recursos Compartidos (File Shares)' -Data $data -Note 'Shares de datos publicados por el servidor, excluyendo los administrativos (C$, ADMIN$, IPC$, PRINT$).' -Wide
    $data
}

if ($fileShares -and ($fileShares -isnot [string]) -and (@($fileShares).Count -gt 0)) {
    $i = 1
    foreach ($share in @($fileShares)) {
        $perm = Invoke-Remote -TimeoutSec 45 -ArgumentList @($share.Name) -ScriptBlock {
            param($n)
            try { Get-SmbShareAccess -Name $n -ErrorAction Stop | Select-Object AccountName, AccessControlType, AccessRight }
            catch { "No se pudo obtener los permisos del share '$n': $($_.Exception.Message)" }
        }
        Add-ReportSection -Id "1.7.2.$i" -Title "Permisos del Share: $($share.Name)" -Data $perm -Note "Ruta: $($share.Path). ACL a nivel de recurso compartido (no confundir con los permisos NTFS del sistema de archivos subyacente)."
        $i++
    }
}

}

}
#endregion

#region 1.8 DHCP SERVER (solo si el rol DHCP esta instalado)
if (Test-SectionEnabled '1.8') {

if (-not $script:Caps.DHCP) {
    Add-ReportSection -Id '1.8' -Title 'Servidor DHCP' -Data 'El rol de servidor DHCP no esta instalado en este equipo.' -Note 'Se documenta explicitamente la ausencia del rol para que el as-built quede completo.'
} else {

Measure-Section '1.8.1 DHCP - estadisticas del servicio' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DhcpServerv4Statistics -ErrorAction Stop | Select-Object TotalScopes, TotalAddresses, AddressesInUse, AddressesAvailable, PercentageInUse }
        catch { "Modulo DhcpServer no disponible o error al consultar estadisticas: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.8.1' -Title 'Estadisticas del Servicio DHCP' -Data $data -Note 'PercentageInUse cercano a 100 anticipa agotamiento del rango de direcciones y clientes que no podran obtener IP nueva.'
}

$dhcpScopes = Measure-Section '1.8.2 DHCP - scopes' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DhcpServerv4Scope -ErrorAction Stop | Select-Object ScopeId, Name, StartRange, EndRange, SubnetMask, State, LeaseDuration }
        catch { "No se pudieron obtener los scopes DHCP: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.8.2' -Title 'Scopes DHCP' -Data $data -Note 'Rangos configurados, mascara de subred y duracion de lease de cada scope.' -Wide
    $data
}

Measure-Section '1.8.2.1 DHCP - estadisticas por scope' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DhcpServerv4ScopeStatistics -ErrorAction Stop | Select-Object ScopeId, Free, InUse, Reserved, PercentageInUse }
        catch { "No se pudieron obtener las estadisticas por scope: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.8.2.1' -Title 'Estadisticas por Scope' -Data $data -Note 'Direcciones libres/en uso/reservadas por scope. Un scope con Free=0 deja de poder asignar IPs nuevas a clientes de ese rango.'
}

Measure-Section '1.8.2.2 DHCP - failover' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DhcpServerv4Failover -ErrorAction Stop | Select-Object Name, PartnerServer, Mode, State, ScopeId }
        catch { "Sin configuracion de Failover DHCP en este servidor (o error al consultar): $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.8.2.2' -Title 'Configuracion de Failover DHCP' -Data $data -Note 'Relacion de alta disponibilidad con un servidor DHCP par. Su ausencia implica que este servidor es punto unico de falla para la asignacion de direcciones IP.'
}

Measure-Section '1.8.2.3 DHCP - binding de interfaces' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DhcpServerv4Binding -ErrorAction Stop | Select-Object InterfaceAlias, IPAddress, State }
        catch { "No se pudo obtener el binding de interfaces DHCP: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.8.2.3' -Title 'Binding de Interfaces de Red' -Data $data -Note 'Interfaces de red por las que el servicio DHCP esta escuchando solicitudes.'
}

Measure-Section '1.8.3 DHCP - opciones globales del servidor' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DhcpServerv4OptionValue -ErrorAction Stop | Select-Object OptionId, Name, Value, VendorClass }
        catch { "No se pudieron obtener las opciones globales de DHCP: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.8.3' -Title 'Opciones Globales del Servidor DHCP' -Data $data -Note 'Opciones (DNS, gateway, dominio, etc.) aplicadas a todos los scopes salvo que un scope puntual las sobrescriba.' -Wide
}

Measure-Section '1.8.3.1 DHCP - configuracion de DNS dinamico' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DhcpServerv4DnsSetting -ErrorAction Stop | Select-Object DynamicUpdates, DeleteDnsRRonLeaseExpiry, UpdateDnsRRForOlderClients }
        catch { "No se pudo obtener la configuracion de DNS dinamico de DHCP: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.8.3.1' -Title 'Configuracion de Actualizacion DNS Dinamica' -Data $data -Note 'Define como el servidor DHCP actualiza (o limpia) los registros DNS de los clientes al asignar o expirar un lease.'
}

if ($dhcpScopes -and ($dhcpScopes -isnot [string]) -and (@($dhcpScopes).Count -gt 0)) {
    $j = 1
    foreach ($scope in @($dhcpScopes)) {
        $sid = $scope.ScopeId
        $opts = Invoke-Remote -TimeoutSec 45 -ArgumentList @($sid) -ScriptBlock {
            param($s)
            try { Get-DhcpServerv4OptionValue -ScopeId $s -ErrorAction Stop | Select-Object OptionId, Name, Value }
            catch { "No se pudieron obtener las opciones del scope $s : $($_.Exception.Message)" }
        }
        Add-ReportSection -Id "1.8.4.$j" -Title "Opciones del Scope $sid ($($scope.Name))" -Data $opts -Note 'Opciones especificas de este scope, que sobrescriben las opciones globales del servidor (seccion 1.8.3) para los clientes de este rango en particular.'
        $j++
    }
}

}

}
#endregion

#region 1.9 DNS SERVER (solo si el rol DNS esta instalado)
if (Test-SectionEnabled '1.9') {

if (-not $script:Caps.DNS) {
    Add-ReportSection -Id '1.9' -Title 'Servidor DNS' -Data 'El rol de servidor DNS no esta instalado en este equipo.' -Note 'Se documenta explicitamente la ausencia del rol para que el as-built quede completo.'
} else {

Measure-Section '1.9.1 DNS - configuracion IP del servicio' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DnsServerSetting -ErrorAction Stop | Select-Object ListenAddresses, BootMethod, EnableDnsSec }
        catch { "Modulo DnsServer no disponible o error al consultar: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.9.1' -Title 'Configuracion IP del Servicio DNS' -Data $data -Note 'Direcciones IP en las que escucha el servicio y si DNSSEC esta habilitado.'
}

Measure-Section '1.9.2 DNS - scavenging' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DnsServerScavenging -ErrorAction Stop | Select-Object ScavengingState, ScavengingInterval, LastScavengeTime, RefreshInterval, NoRefreshInterval }
        catch { "No se pudo obtener la configuracion de scavenging: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.9.2' -Title 'Opciones de Scavenging (Limpieza de Registros)' -Data $data -Note 'Scavenging deshabilitado en un entorno con muchos clientes DHCP suele acumular registros huerfanos de equipos dados de baja que siguen resolviendo.'
}

Measure-Section '1.9.3 DNS - forwarders' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DnsServerForwarder -ErrorAction Stop | Select-Object IPAddress, Timeout, UseRootHint }
        catch { "No se pudo obtener la configuracion de forwarders: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.9.3' -Title 'Forwarders Configurados' -Data $data -Note 'Servidores DNS externos a los que se reenvian las consultas que este servidor no puede resolver localmente.'
}

$dnsZones = Measure-Section '1.9.4 DNS - zonas configuradas' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DnsServerZone -ErrorAction Stop | Select-Object ZoneName, ZoneType, IsAutoCreated, IsDsIntegrated, IsReverseLookupZone, DynamicUpdate }
        catch { "No se pudieron enumerar las zonas DNS: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.9.4' -Title 'Zonas DNS Configuradas' -Data $data -Note 'Listado completo de zonas alojadas (directas, inversas, integradas en AD o basadas en archivo).' -Wide
    $data
}

Measure-Section '1.9.4.1 DNS - politicas de transferencia de zona' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            Get-DnsServerZone -ErrorAction Stop | Where-Object { -not $_.IsAutoCreated } | ForEach-Object {
                $xfr = Get-DnsServerZoneTransferPolicy -ZoneName $_.ZoneName -ErrorAction SilentlyContinue
                [PSCustomObject]@{ Zona = $_.ZoneName; TransferPolicy = if ($xfr) { ($xfr | Out-String).Trim() } else { 'Segun configuracion de zona (SecureOnly/Any/None)' } }
            }
        } catch { "No se pudo obtener las politicas de transferencia de zona: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.9.4.1' -Title 'Politicas de Transferencia de Zona' -Data $data -Note 'Transferencias de zona sin restriccion (Any) permiten a cualquier host volcar el contenido completo de la zona: una fuga de informacion clasica en el reconocimiento de red.' -Wide
}

Measure-Section '1.9.4.2 DNS - zonas de resolucion inversa' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DnsServerZone -ErrorAction Stop | Where-Object { $_.IsReverseLookupZone } | Select-Object ZoneName, ZoneType, DynamicUpdate, IsDsIntegrated }
        catch { "No se pudieron obtener las zonas de resolucion inversa: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.9.4.2' -Title 'Zonas de Resolucion Inversa (PTR)' -Data $data -Note 'Zonas in-addr.arpa configuradas. Su ausencia impide resolver IP->nombre, lo cual rompe herramientas de diagnostico y algunos controles de seguridad.'
}

Measure-Section '1.9.4.3 DNS - forwarders condicionales' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try { Get-DnsServerZone -ErrorAction Stop | Where-Object { $_.ZoneType -eq 'Forwarder' } | Select-Object ZoneName, MasterServers, IsDsIntegrated }
        catch { "No se pudieron obtener los forwarders condicionales: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.9.4.3' -Title 'Forwarders Condicionales' -Data $data -Note 'Reglas de reenvio por dominio especifico, tipicas de confianza entre bosques/dominios de AD distintos.'
}

Measure-Section '1.9.4.4 DNS - aging por zona' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            Get-DnsServerZone -ErrorAction Stop | ForEach-Object {
                try {
                    $aging = Get-DnsServerZoneAging -Name $_.ZoneName -ErrorAction Stop
                    [PSCustomObject]@{ Zona = $_.ZoneName; AgingEnabled = $aging.AgingEnabled; ScavengeServers = ($aging.ScavengeServers -join ', ') }
                } catch { $null }
            } | Where-Object { $_ }
        } catch { "No se pudo obtener la configuracion de aging por zona: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.9.4.4' -Title 'Propiedades de Aging por Zona' -Data $data -Note 'Aging habilitado por zona es requisito para que el scavenging global (seccion 1.9.2) efectivamente limpie registros de esa zona.'
}

}

}
#endregion

#region 1.10 TERMINAL SERVICES / RD LICENSING Y ACCESOS RDP
if (Test-SectionEnabled '1.10') {

# El servicio TermService (base de $script:Caps.RDSH) corre en TODO Windows,
# incluido un Windows 10/11 cliente con solo el acceso RDP administrativo
# habilitado. Para no reportar licenciamiento RDS (CALs, RD Licensing) en un
# equipo que en realidad no tiene el rol de Session Host, se exige evidencia
# real del rol ademas del capability flag: Get-WindowsFeature RDS-RD-Server
# instalado, o Win32_TerminalServiceSetting.TerminalServerMode = 1.
$esRdsSessionHost = $false
try {
    $rdshCheck = Invoke-Remote -TimeoutSec 30 -ScriptBlock {
        try {
            $instalado = $false
            try {
                if (Get-Command Get-WindowsFeature -ErrorAction SilentlyContinue) {
                    $feat = Get-WindowsFeature -Name RDS-RD-Server -ErrorAction Stop
                    if ($feat -and $feat.Installed) { $instalado = $true }
                }
            } catch {}
            if (-not $instalado) {
                try {
                    $tsSetting = Get-CimInstance -Namespace root\cimv2\TerminalServices -ClassName Win32_TerminalServiceSetting -ErrorAction Stop
                    if ($tsSetting -and [int]$tsSetting.TerminalServerMode -eq 1) { $instalado = $true }
                } catch {}
            }
            [bool]$instalado
        } catch { $false }
    }
    if ($rdshCheck -is [bool]) { $esRdsSessionHost = $rdshCheck }
} catch { $esRdsSessionHost = $false }

$mostrarLicenciamientoRDS = $script:Caps.RDSLicensing -or $esRdsSessionHost

if (-not $script:Caps.RDSH -and -not $script:Caps.RDSLicensing) {
    Add-ReportSection -Id '1.10' -Title 'Terminal Services / RD Licensing' -Data 'Los roles Remote Desktop Session Host y RD Licensing no estan instalados en este equipo.' -Note 'Se documenta explicitamente la ausencia de estos roles para que el as-built quede completo. Se mantiene igualmente el historial de accesos RDP administrativos (1.10.4), util en cualquier equipo con RDP habilitado.'
} elseif ($mostrarLicenciamientoRDS) {
    Add-ReportSection -Id '1.10' -Title 'Terminal Services / RD Licensing' -Data 'Este equipo tiene instalado el rol de Escritorio Remoto (Session Host) y/o RD Licensing.' -Note 'Se detallan a continuacion la configuracion de licenciamiento RDS y las licencias (CALs) emitidas.'

Measure-Section '1.10.1 RDS - configuracion de licenciamiento' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            Get-RDLicenseConfiguration -ErrorAction Stop | Select-Object LicenseServer, Mode, ConnectionBroker
        } catch {
            try {
                $tsSetting = Get-CimInstance -Namespace root\cimv2\TerminalServices -ClassName Win32_TerminalServiceSetting -ErrorAction Stop
                [PSCustomObject]@{
                    ModoLicenciamiento = switch ($tsSetting.LicensingType) {
                        1 {'Per Device'} 2 {'Per User'} 4 {'Per User (temporal)'} default {"Desconocido ($($tsSetting.LicensingType))"}
                    }
                    ServidoresLicencia = ($tsSetting.GetSpecifiedLicenseServerList().SpecifiedLSList -join ', ')
                }
            } catch {
                "No se pudo determinar la configuracion de licenciamiento RDS: $($_.Exception.Message)"
            }
        }
    }
    Add-ReportSection -Id '1.10.1' -Title 'Configuracion de Licenciamiento RDS' -Data $data -Note 'Modo de licenciamiento (Per Device/Per User) y servidor de licencias configurado. Un modo mal configurado bloquea nuevas conexiones RDP pasado el periodo de gracia de 120 dias.'
}

Measure-Section '1.10.2 RDS - paquetes de licencias (CALs)' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            Get-CimInstance -Namespace root\cimv2\TerminalServices -ClassName Win32_TSLicenseKeyPack -ErrorAction Stop |
                Select-Object Description, ProductVersion, TypeAndModel, KeyPackType,
                    TotalLicenses, IssuedLicenses,
                    @{N='LicenciasDisponibles';E={$_.TotalLicenses - $_.IssuedLicenses}},
                    ExpirationDate
        } catch { "No se pudo consultar Win32_TSLicenseKeyPack (rol RD Licensing no instalado en este servidor): $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.10.2' -Title 'Paquetes de Licencias (CALs) - Uso por Tipo' -Data $data -Note 'LicenciasDisponibles bajo o negativo anticipa que usuarios o equipos nuevos no podran conectarse por RDP.' -Wide
}

Measure-Section '1.10.3 RDS - licencias emitidas (detalle)' {
    $data = Invoke-Remote -TimeoutSec 90 -ScriptBlock {
        try {
            Get-CimInstance -Namespace root\cimv2\TerminalServices -ClassName Win32_TSIssuedLicense -ErrorAction Stop |
                Select-Object sIssuedToUser, sIssuedToComputer,
                    @{N='FechaEmision';E={ if ($_.dtIssueDate) { [Management.ManagementDateTimeConverter]::ToDateTime($_.dtIssueDate) } }},
                    @{N='FechaExpiracion';E={ if ($_.dtExpirationDate) { [Management.ManagementDateTimeConverter]::ToDateTime($_.dtExpirationDate) } }},
                    sKeyPackId | Sort-Object FechaEmision -Descending
        } catch { "No se pudo consultar Win32_TSIssuedLicense: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.10.3' -Title 'Licencias Emitidas (Detalle por Usuario/Equipo)' -Data $data -Note 'Detalle de cada CAL emitida, a quien y cuando vence, para auditar el consumo real contra lo comprado (seccion 1.10.2).' -Wide
}

} else {
    Add-ReportSection -Id '1.10' -Title 'Terminal Services / RD Licensing' -Data 'Este equipo no tiene el rol de Escritorio Remoto (Session Host) instalado; solo se documenta el acceso RDP administrativo.' -Note 'El servicio Terminal Services (TermService) corre en todo Windows, incluidos equipos cliente con RDP administrativo habilitado; sin el rol RDS-RD-Server (o TerminalServerMode=1) no aplica licenciamiento RDS y no se consultan Win32_TSLicenseKeyPack/Win32_TSIssuedLicense. Ver 1.10.4 para el historial de accesos RDP.'
}

# El historial de accesos RDP (1.10.4) corre siempre que la seccion 1.10 este
# habilitada, sin depender de que el equipo tenga el rol de Session Host: en
# un equipo cliente/administrativo documenta igualmente quien se conecto por
# RDP administrativo.
Measure-Section '1.10.4 Accesos RDP - historial de logon' {
    $diasHistorial = 90
    $data = Invoke-Remote -TimeoutSec 90 -ArgumentList @($diasHistorial) -ScriptBlock {
        param($dias)
        try {
            $startDate = (Get-Date).AddDays(-$dias)
            $events = Get-WinEvent -FilterHashtable @{
                LogName   = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'
                Id        = 21
                StartTime = $startDate
            } -MaxEvents 10000 -ErrorAction Stop

            $parsed = $events | ForEach-Object {
                try {
                    $xmlEvt = [xml]$_.ToXml()
                    $eventData = $xmlEvt.Event.UserData.EventXML
                    [PSCustomObject]@{ Usuario = $eventData.User; IPOrigen = $eventData.Address; FechaHora = $_.TimeCreated }
                } catch { $null }
            } | Where-Object { $_ -and $_.Usuario }

            if (-not $parsed) { return "No se encontraron eventos de logon RDS validos en los ultimos $dias dias." }

            $parsed | Group-Object Usuario | ForEach-Object {
                $ultimo = $_.Group | Sort-Object FechaHora -Descending | Select-Object -First 1
                [PSCustomObject]@{
                    Usuario             = $_.Name
                    UltimoLogon         = $ultimo.FechaHora
                    UltimaIPOrigen      = $ultimo.IPOrigen
                    ConexionesEnVentana = $_.Count
                }
            } | Sort-Object UltimoLogon -Descending
        } catch {
            "No se encontraron eventos de logon RDS en los ultimos $dias dias, o el log 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' esta deshabilitado, vacio o fue rotado/purgado. Detalle: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.10.4' -Title "Accesos RDP - Historial de Logon (Ultimos $diasHistorial Dias)" -Data $data -Note 'Usuarios que iniciaron sesion via RDP (Escritorio Remoto), con fecha/IP del ultimo logon y cantidad de conexiones en la ventana. Se registra independientemente de si el equipo tiene el rol de Session Host: en un equipo cliente/administrativo documenta quien se conecto por RDP administrativo. Ventana fija de 90 dias, independiente del parametro -EventLogDays (pensado para logs de eventos generales, mas volatiles que este log especifico de RDS).' -Wide
}

}
#endregion

#region 1.11 SEGURIDAD Y CUMPLIMIENTO (base: certificados, SChannel, PAM, NTP, tareas, AV/EDR, reboot, backups)
if (Test-SectionEnabled '1.11') {

# -----------------------------------------------------------------------------
# 1.11.1 Certificados SSL/TLS instalados
# -----------------------------------------------------------------------------
Measure-Section '1.11.1 Certificados SSL/TLS instalados' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $ahora = Get-Date
            $certs = @(Get-ChildItem -Path Cert:\LocalMachine\My -ErrorAction Stop)
            if ($certs.Count -eq 0) {
                return "El almacen personal de certificados de la maquina (LocalMachine\My) no contiene certificados."
            }
            $certs | ForEach-Object {
                $cn = $_.Subject
                if ($_.Subject -match 'CN=([^,]+)') { $cn = $matches[1] }
                $usoMejorado = 'Todos los usos (sin restriccion EKU)'
                try {
                    if ($_.EnhancedKeyUsageList -and $_.EnhancedKeyUsageList.Count -gt 0) {
                        $nombresEku = ($_.EnhancedKeyUsageList | ForEach-Object { $_.FriendlyName } | Where-Object { $_ })
                        if ($nombresEku) { $usoMejorado = ($nombresEku -join ', ') }
                    }
                } catch {}
                [PSCustomObject]@{
                    CN                = $cn
                    Sujeto            = $_.Subject
                    Emisor            = $_.Issuer
                    HuellaDigital     = $_.Thumbprint
                    FechaExpiracion   = $_.NotAfter
                    DiasParaExpirar   = [math]::Round(($_.NotAfter - $ahora).TotalDays, 0)
                    Expirado          = ($_.NotAfter -lt $ahora)
                    AutoFirmado       = ($_.Subject -eq $_.Issuer)
                    UsoMejorado       = $usoMejorado
                    TieneClavePrivada = $_.HasPrivateKey
                }
            } | Sort-Object FechaExpiracion
        } catch { "No se pudo enumerar los certificados de Cert:\LocalMachine\My: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.11.1' -Title 'Certificados SSL/TLS Instalados' -Data $data -Note 'Certificados del almacen personal de la maquina (LocalMachine\My), ordenados por fecha de expiracion. Expirado=true o DiasParaExpirar bajo/negativo requieren renovacion inmediata; TieneClavePrivada=false indica un certificado que no puede usarse para autenticar el servicio que lo referencia.' -Wide
}

# -----------------------------------------------------------------------------
# 1.11.2 Protocolos, cifrados y hashes SChannel
# -----------------------------------------------------------------------------
Measure-Section '1.11.2 Protocolos y cifrados SChannel' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $rows = New-Object System.Collections.Generic.List[PSObject]

            $protocolPaths = @('SSL 2.0','SSL 3.0','TLS 1.0','TLS 1.1','TLS 1.2','TLS 1.3')
            $baseProto = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols'
            foreach ($proto in $protocolPaths) {
                foreach ($side in @('Server','Client')) {
                    $regPath = Join-Path (Join-Path $baseProto $proto) $side
                    $estado = 'No configurado (default del SO)'
                    if (Test-Path $regPath) {
                        $val = Get-ItemProperty -Path $regPath -Name Enabled -ErrorAction SilentlyContinue
                        if ($null -ne $val) { $estado = if ($val.Enabled -eq 0) { 'Deshabilitado' } else { 'Habilitado' } }
                    }
                    $rows.Add([PSCustomObject]@{ Tipo = 'Protocolo'; Nombre = $proto; Lado = $side; Estado = $estado })
                }
            }

            $baseCiphers = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers'
            $cipherNames = @('NULL','DES 56/56','RC2 40/128','RC2 56/128','RC2 128/128','RC4 40/128','RC4 56/128','RC4 64/128','RC4 128/128','Triple DES 168','AES 128/128','AES 256/256')
            foreach ($c in $cipherNames) {
                $regPath = Join-Path $baseCiphers $c
                $estado = 'No configurado (default del SO)'
                if (Test-Path $regPath) {
                    $val = Get-ItemProperty -Path $regPath -Name Enabled -ErrorAction SilentlyContinue
                    if ($null -ne $val) { $estado = if ($val.Enabled -eq 0) { 'Deshabilitado' } else { 'Habilitado' } }
                }
                $rows.Add([PSCustomObject]@{ Tipo = 'Cifrado'; Nombre = $c; Lado = 'N/A'; Estado = $estado })
            }

            $baseHashes = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Hashes'
            $hashNames = @('MD5','SHA','SHA256','SHA384','SHA512')
            foreach ($h in $hashNames) {
                $regPath = Join-Path $baseHashes $h
                $estado = 'No configurado (default del SO)'
                if (Test-Path $regPath) {
                    $val = Get-ItemProperty -Path $regPath -Name Enabled -ErrorAction SilentlyContinue
                    if ($null -ne $val) { $estado = if ($val.Enabled -eq 0) { 'Deshabilitado' } else { 'Habilitado' } }
                }
                $rows.Add([PSCustomObject]@{ Tipo = 'Hash'; Nombre = $h; Lado = 'N/A'; Estado = $estado })
            }

            $rows
        } catch { "No se pudo consultar la configuracion de SChannel: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.11.2' -Title 'Protocolos, Cifrados y Hashes SChannel' -Data $data -Note 'Protocolos SSL/TLS (incluye TLS 1.3), suites de cifrado y algoritmos hash habilitados o deshabilitados via registro SCHANNEL. Cuando no hay clave configurada, el sistema usa el default del SO (varia segun build). TLS 1.0/1.1 o SSL habilitados son hallazgos de cumplimiento tipicos (por ejemplo PCI-DSS).' -Wide
}

# -----------------------------------------------------------------------------
# 1.11.3 Estado PAM / BeyondTrust (usa -PamBrokerEndpoint, sin IP hardcodeada)
# -----------------------------------------------------------------------------
Measure-Section '1.11.3 Estado PAM / BeyondTrust' {
    $pamHostVal = $null
    $pamPortVal = $null
    if ($PamBrokerEndpoint) {
        $partesEndpoint = $PamBrokerEndpoint -split ':', 2
        $pamHostVal = $partesEndpoint[0]
        if ($partesEndpoint.Count -gt 1 -and $partesEndpoint[1]) {
            $parsedPort = 0
            if ([int]::TryParse($partesEndpoint[1], [ref]$parsedPort)) { $pamPortVal = $parsedPort } else { $pamPortVal = 3389 }
        } else {
            $pamPortVal = 3389
        }
    }

    $data = Invoke-Remote -TimeoutSec 45 -ArgumentList @($pamHostVal, $pamPortVal) -ScriptBlock {
        param($h, $p)
        try {
            $pamServices = Get-Service -ErrorAction SilentlyContinue | Where-Object {
                $_.DisplayName -match 'BeyondTrust|Bomgar|PMM|Privilege Management' -or $_.Name -match 'BeyondTrust|PMM'
            }
            $resultado = [PSCustomObject]@{
                AgentePAMInstalado         = [bool]$pamServices
                ServiciosPAMDetectados     = if ($pamServices) { ($pamServices.DisplayName -join ', ') } else { 'Ninguno detectado' }
                ConectividadResourceBroker = if ($h) { "$h`:$p" } else { 'No configurado (parametro -PamBrokerEndpoint vacio)' }
                ConectividadOK             = 'No se probo (no se especifico -PamBrokerEndpoint)'
            }
            if ($h) {
                try {
                    $connTest = Test-NetConnection -ComputerName $h -Port $p -WarningAction SilentlyContinue -ErrorAction Stop -InformationLevel Quiet
                    $resultado.ConectividadOK = $connTest
                } catch {
                    $resultado.ConectividadOK = "No se pudo probar: $($_.Exception.Message)"
                }
            }
            $resultado
        } catch { "No se pudo determinar el estado del agente PAM: $($_.Exception.Message)" }
    }

    $notaPam = if ($PamBrokerEndpoint) {
        "Verifica la presencia del agente PAM local y realiza un test de conectividad TCP (timeout corto) al Resource Broker configurado via -PamBrokerEndpoint ($PamBrokerEndpoint)."
    } else {
        'Verifica unicamente la presencia del agente PAM local instalado como servicio. No se realiza test de conectividad al Resource Broker porque no se especifico el parametro -PamBrokerEndpoint (formato host:puerto).'
    }
    Add-ReportSection -Id '1.11.3' -Title 'Estado PAM / BeyondTrust' -Data $data -Note $notaPam
}

# -----------------------------------------------------------------------------
# 1.11.4 Configuracion NTP
# -----------------------------------------------------------------------------
Measure-Section '1.11.4 Configuracion NTP' {
    $data = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        try {
            $w32tmStatus = w32tm /query /status 2>&1 | Out-String
            $w32tmPeers  = w32tm /query /peers 2>&1 | Out-String
            [PSCustomObject]@{
                Estado  = (($w32tmStatus -split "`n") | Select-Object -First 8) -join ' | '
                Fuentes = $w32tmPeers.Trim()
            }
        } catch { "No se pudo consultar el servicio Windows Time (w32tm): $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.11.4' -Title 'Configuracion NTP' -Data $data -Note 'Fuente de hora y estado de sincronizacion. Un desfase horario grande entre servidores rompe Kerberos (tolerancia tipica de 5 minutos) y dificulta correlacionar logs entre equipos.'
}

# -----------------------------------------------------------------------------
# 1.11.5 Tareas programadas activas
# -----------------------------------------------------------------------------
Measure-Section '1.11.5 Tareas programadas activas' {
    $data = Invoke-Remote -TimeoutSec 120 -ScriptBlock {
        try {
            $tareas = @(Get-ScheduledTask -ErrorAction Stop | Where-Object { $_.State -ne 'Disabled' })

            # Get-ScheduledTaskInfo se pide en UNA sola tuberia sobre todas las tareas
            # no deshabilitadas, en lugar de invocarlo una vez por tarea dentro de un
            # ForEach-Object (con 200+ tareas, una llamada por tarea puede tardar varios
            # segundos por el overhead de reabrir el Task Scheduler en cada invocacion).
            $infoPorTarea = @{}
            if ($tareas.Count -gt 0) {
                $tareas | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue | ForEach-Object {
                    $infoPorTarea["$($_.TaskPath)|$($_.TaskName)"] = $_
                }
            }

            $tareas | ForEach-Object {
                $info = $infoPorTarea["$($_.TaskPath)|$($_.TaskName)"]
                [PSCustomObject]@{
                    Nombre           = $_.TaskName
                    Ruta             = $_.TaskPath
                    Estado           = $_.State
                    EjecutarComo     = $_.Principal.UserId
                    NivelPrivilegio  = $_.Principal.RunLevel
                    UltimaEjecucion  = if ($info) { $info.LastRunTime } else { 'N/D' }
                    ProximaEjecucion = if ($info) { $info.NextRunTime } else { 'N/D' }
                    UltimoResultado  = if ($info) { $info.LastTaskResult } else { 'N/D' }
                }
            } | Sort-Object Nombre
        } catch { "No se pudieron enumerar las tareas programadas: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.11.5' -Title 'Tareas Programadas Activas' -Data $data -Note 'Tareas no deshabilitadas. UltimoResultado distinto de 0 indica que la ultima ejecucion fallo; conviene revisar con prioridad las tareas de backup/mantenimiento en ese estado.' -Wide
}

# -----------------------------------------------------------------------------
# 1.11.6 Antivirus / EDR detallado
# -----------------------------------------------------------------------------
Measure-Section '1.11.6 Antivirus / EDR detallado' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $defender = try {
                Get-MpComputerStatus -ErrorAction Stop | Select-Object AMServiceEnabled, AntivirusEnabled, RealTimeProtectionEnabled, AntivirusSignatureLastUpdated, QuickScanAge, FullScanAge
            } catch { $null }
            $edrServices = Get-Service -ErrorAction SilentlyContinue | Where-Object {
                $_.DisplayName -match 'Carbon Black|CB Defense|Trend Micro|Cortex|CrowdStrike|SentinelOne|Deep Security'
            } | Select-Object DisplayName, Status, StartType
            [PSCustomObject]@{
                WindowsDefender    = if ($defender) { ($defender | Out-String).Trim() } else { 'No disponible / deshabilitado (Get-MpComputerStatus fallo o Defender no esta activo)' }
                AgentesEDRTerceros = if ($edrServices) { ($edrServices | ForEach-Object { "$($_.DisplayName): $($_.Status)" }) -join ' | ' } else { 'Ninguno detectado' }
            }
        } catch { "No se pudo obtener el estado de Antivirus/EDR: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.11.6' -Title 'Antivirus / EDR Detallado' -Data $data -Note 'Estado de Windows Defender y presencia de agentes EDR de terceros conocidos. Un equipo sin ninguno de los dos protegiendo el endpoint es un hallazgo critico.'
}

# -----------------------------------------------------------------------------
# 1.11.7 Reinicio pendiente
# -----------------------------------------------------------------------------
Measure-Section '1.11.7 Reinicio pendiente' {
    $data = Invoke-Remote -TimeoutSec 30 -ScriptBlock {
        try {
            $pendingRename = Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations'
            $cbsReboot     = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
            $wuReboot      = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
            [PSCustomObject]@{
                PendingFileRename       = $pendingRename
                ComponentBasedServicing = $cbsReboot
                WindowsUpdateReboot     = $wuReboot
                ReinicioPendiente       = ($pendingRename -or $cbsReboot -or $wuReboot)
            }
        } catch { "No se pudo determinar si hay un reinicio pendiente: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.11.7' -Title 'Reinicio Pendiente' -Data $data -Note 'ReinicioPendiente=true significa que hay parches o cambios de componentes aplicados que no terminan de tomar efecto hasta el proximo reinicio del equipo.'
}

# -----------------------------------------------------------------------------
# 1.11.8 Estado de backups (agentes locales)
# -----------------------------------------------------------------------------
Measure-Section '1.11.8 Estado de backups' {
    $data = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        try {
            $backupServices = Get-Service -ErrorAction SilentlyContinue | Where-Object {
                $_.DisplayName -match 'Veeam|Veritas|NetBackup|Backup Exec'
            } | Select-Object DisplayName, Status, StartType
            if ($backupServices) { $backupServices }
            else { "No se detecto agente de backup (Veeam/Veritas/NetBackup/Backup Exec) instalado como servicio en este servidor." }
        } catch { "No se pudo determinar el estado de los agentes de backup: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.11.8' -Title 'Estado de Backups (Agentes Locales)' -Data $data -Note 'Presencia y estado del servicio del agente de backup local. Un agente detenido en un servidor que deberia respaldarse a diario es un hallazgo critico de continuidad.'
}

}
#endregion

#region 1.12 WSUS (solo si el rol WSUS esta instalado)
if (Test-SectionEnabled '1.12') {

if (-not $script:Caps.WSUS) {
    Add-ReportSection -Id '1.12' -Title 'Servidor WSUS' -Data 'El rol WSUS no esta instalado en este equipo.' -Note 'Se documenta explicitamente la ausencia del rol para que el as-built quede completo.'
} else {

Measure-Section '1.12.1 WSUS - configuracion del servidor' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            Import-Module UpdateServices -ErrorAction Stop
            $wsus = Get-WsusServer
            $config = $wsus.GetConfiguration()
            [PSCustomObject]@{
                Servidor           = $wsus.Name
                Puerto             = $wsus.PortNumber
                SincronizacionMode = $config.SyncFromMicrosoftUpdate
                UpstreamServer     = $config.UpstreamWsusServerName
                ProductosAprobados = ($config.GetEnabledUpdateLanguages() -join ', ')
            }
        } catch { "Modulo UpdateServices no disponible o error al consultar: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.12.1' -Title 'Configuracion del Servidor WSUS' -Data $data -Note 'Modo de sincronizacion (directo con Microsoft o desde un WSUS upstream) y puerto configurado.'
}

Measure-Section '1.12.2 WSUS - historial de sincronizacion' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $wsus = Get-WsusServer
            $wsus.GetSubscription().GetSynchronizationHistory() | Select-Object -First 10 | Select-Object StartTime, EndTime, Result, Error
        } catch { "No se pudo obtener el historial de sincronizacion WSUS: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.12.2' -Title 'Historial de Sincronizacion (Ultimas 10)' -Data $data -Note 'Sincronizaciones fallidas repetidas indican que el catalogo de parches esta desactualizado para todos los clientes de este WSUS.'
}

}

}
#endregion

#region 1.13 GROUP POLICY
if (Test-SectionEnabled '1.13') {

Measure-Section '1.13 Group Policy aplicadas (RSOP)' {
    $data = Invoke-Remote -TimeoutSec 90 -ScriptBlock {
        $rsopFile = Join-Path $env:TEMP "rsop_$($env:COMPUTERNAME)_$([guid]::NewGuid().ToString('N')).xml"
        try {
            gpresult /f /x $rsopFile *>$null
            if (Test-Path $rsopFile) {
                [xml]$rsop = Get-Content $rsopFile
                $rsop.Rsop.ComputerResults.GPO | Where-Object { $_.Name -and $_.Enabled -eq 'true' } |
                    Select-Object @{N='NombreGPO';E={$_.Name}}, @{N='Habilitada';E={$_.Enabled}}, @{N='Version';E={$_.VersionDirectory}}
            } else { "No se pudo generar el reporte RSOP (gpresult)." }
        } catch {
            "Error al obtener las GPOs aplicadas (verificar permisos o si el equipo pertenece al dominio): $($_.Exception.Message)"
        } finally {
            if (Test-Path $rsopFile) { Remove-Item $rsopFile -ErrorAction SilentlyContinue }
        }
    }
    Add-ReportSection -Id '1.13' -Title 'Group Policy Aplicadas (RSOP)' -Data $data -Note 'GPOs de computadora efectivamente aplicadas, segun Resultant Set of Policy. El archivo temporal se genera y se elimina en el propio equipo auditado.'
}

}
#endregion

#region 1.14.1 CUENTAS DE SERVICIO CON PRIVILEGIOS ELEVADOS
if (Test-SectionEnabled '1.14') {

Measure-Section '1.14.1 Cuentas de servicio con privilegios elevados' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $rows = @(Get-CimInstance Win32_Service -ErrorAction Stop | Where-Object {
                $_.StartName -and $_.StartName -notmatch '^(LocalSystem|NT AUTHORITY\\(NetworkService|LocalService)|NT SERVICE\\.*)$'
            } | Select-Object Name, DisplayName, StartName, State,
                @{N='TipoCuenta';E={ if ($_.StartName -match '\\') { 'Dominio' } else { 'Local' } }} |
                Sort-Object TipoCuenta, StartName)
            if ($rows.Count -eq 0) {
                return 'Ningun servicio corre con una cuenta de dominio o de usuario; todos usan cuentas integradas del sistema (LocalSystem/NetworkService/LocalService/NT SERVICE).'
            }
            $rows
        } catch { "No se pudo enumerar las cuentas de servicio con privilegios elevados: $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.14.1' -Title 'Cuentas de Servicio con Privilegios Elevados' -Data $data -Note 'Servicios que no corren bajo cuentas de sistema (LocalSystem/NetworkService/LocalService/NT SERVICE). Las cuentas de dominio aqui suelen tener contrasenas que nunca expiran, un riesgo si se filtran o si el titular se desvincula de la empresa.' -Wide
}

}
#endregion

#region 1.15 ENTORNOS DE DESARROLLO PYTHON / GIT
if (Test-SectionEnabled '1.15') {

# 1.15.1 se guarda en una variable de script para reutilizarla al armar el
# listado de interpretes "reales" (excluye el alias stub de Microsoft Store)
# que consume la seccion de paquetes pip (1.15.2).
$pythonInterpreters = Measure-Section '1.15.1 Interpretes de Python detectados' {
    $data = Invoke-Remote -TimeoutSec 90 -ScriptBlock {
        try {
            $found = New-Object System.Collections.Generic.List[PSObject]
            $seenPaths = New-Object System.Collections.Generic.HashSet[string]

            try {
                $pyList = & py -0p 2>$null
                foreach ($line in $pyList) {
                    if ($line -match '^\s*-\S*\s+(\S+)\s+(.+)$') {
                        $rutaEncontrada = $matches[2].Trim()
                        if ($rutaEncontrada -and (Test-Path $rutaEncontrada) -and $seenPaths.Add($rutaEncontrada)) {
                            $found.Add([PSCustomObject]@{ Version = 'N/D (via py launcher)'; Ruta = $rutaEncontrada; Origen = 'Python Launcher (py -0p)' })
                        }
                    }
                }
            } catch {}

            $regRoots = @('HKLM:\SOFTWARE\Python\PythonCore','HKLM:\SOFTWARE\WOW6432Node\Python\PythonCore','HKCU:\SOFTWARE\Python\PythonCore')
            foreach ($root in $regRoots) {
                if (Test-Path $root) {
                    Get-ChildItem $root -ErrorAction SilentlyContinue | ForEach-Object {
                        $ver = $_.PSChildName
                        $ip = Get-ItemProperty -Path "$($_.PSPath)\InstallPath" -ErrorAction SilentlyContinue
                        if ($ip -and $ip.'(default)') {
                            $exe = Join-Path $ip.'(default)' 'python.exe'
                            if ((Test-Path $exe) -and $seenPaths.Add($exe)) {
                                $found.Add([PSCustomObject]@{ Version = $ver; Ruta = $exe; Origen = 'Registro PythonCore' })
                            }
                        }
                    }
                }
            }

            try {
                $wherePython = & where.exe python python3 2>$null
                foreach ($p in $wherePython) {
                    if ($p -and (Test-Path $p) -and $seenPaths.Add($p)) {
                        $verOut = try { (& $p --version 2>&1) -join ' ' } catch { 'N/D' }
                        $found.Add([PSCustomObject]@{ Version = $verOut; Ruta = $p; Origen = 'PATH' })
                    }
                }
            } catch {}

            $commonPaths = @()
            $commonPaths += Get-ChildItem 'C:\Users\*\AppData\Local\Programs\Python\Python*\python.exe' -ErrorAction SilentlyContinue
            $commonPaths += Get-ChildItem 'C:\Python*\python.exe' -ErrorAction SilentlyContinue
            $commonPaths += Get-ChildItem 'C:\ProgramData\Anaconda3\python.exe' -ErrorAction SilentlyContinue
            $commonPaths += Get-ChildItem 'C:\ProgramData\Miniconda3\python.exe' -ErrorAction SilentlyContinue
            $commonPaths += Get-ChildItem 'C:\Users\*\Anaconda3\python.exe' -ErrorAction SilentlyContinue
            $commonPaths += Get-ChildItem 'C:\Users\*\Miniconda3\python.exe' -ErrorAction SilentlyContinue

            foreach ($p in $commonPaths) {
                if ($p.FullName -and $seenPaths.Add($p.FullName)) {
                    $verOut = try { (& $p.FullName --version 2>&1) -join ' ' } catch { 'Desconocida' }
                    $found.Add([PSCustomObject]@{ Version = $verOut; Ruta = $p.FullName; Origen = 'Ruta comun (no en PATH/registro)' })
                }
            }

            Get-ChildItem 'C:\Users\*\AppData\Local\Microsoft\WindowsApps\python3*.exe' -ErrorAction SilentlyContinue | ForEach-Object {
                if ($seenPaths.Add($_.FullName)) {
                    $esStub = $_.Length -lt 100KB
                    $found.Add([PSCustomObject]@{
                        Version = if ($esStub) { 'Alias de Microsoft Store (no instalado realmente)' } else { 'Instalado via Microsoft Store' }
                        Ruta    = $_.FullName
                        Origen  = 'WindowsApps'
                    })
                }
            }

            if ($found.Count -gt 0) { $found } else { 'No se detecto ningun interprete de Python en este equipo (Python Launcher, registro, PATH, rutas comunes y alias de Microsoft Store).' }
        } catch {
            "Error al buscar interpretes de Python: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.15.1' -Title 'Interpretes de Python Detectados' -Data $data -Note 'Deteccion via Python Launcher (py -0p), registro PythonCore, PATH y rutas comunes de Anaconda/Miniconda/AppData. Es normal encontrar multiples versiones instaladas en paralelo.' -Wide
    $data
}

$interpretesReales = @()
if ($pythonInterpreters -and ($pythonInterpreters -isnot [string])) {
    $interpretesReales = @($pythonInterpreters | Where-Object { $_.Version -notmatch 'Alias de Microsoft Store' })
}

if ($interpretesReales.Count -eq 0) {
    Add-ReportSection -Id '1.15.2' -Title 'Librerias Python Instaladas (pip freeze)' -Data 'No se detecto ningun interprete de Python con ejecucion real en este equipo (ver 1.15.1); no tiene sentido consultar pip sin un interprete que lo ejecute.' -Note 'Depende de que 1.15.1 haya encontrado al menos un interprete real (no solo el alias de Microsoft Store).'
} else {
    Measure-Section '1.15.2 Librerias Python instaladas (pip)' {
        $data = Invoke-Remote -TimeoutSec 120 -ArgumentList (,$interpretesReales) -ScriptBlock {
            param($interps)
            $results = @()
            foreach ($py in $interps) {
                try {
                    $freeze = & $py.Ruta -m pip list --disable-pip-version-check --format=freeze 2>$null
                    $count  = ($freeze | Measure-Object).Count
                    $results += [PSCustomObject]@{
                        Interprete    = $py.Ruta
                        TotalPaquetes = $count
                        Paquetes      = if ($count -gt 0) { ($freeze -join '; ') } else { '(sin paquetes o pip no disponible)' }
                    }
                } catch {
                    $results += [PSCustomObject]@{ Interprete = $py.Ruta; TotalPaquetes = 0; Paquetes = "pip no disponible o error al consultar: $($_.Exception.Message)" }
                }
            }
            $results
        }
        Add-ReportSection -Id '1.15.2' -Title 'Librerias Python Instaladas (pip freeze)' -Data $data -Note 'Paquetes pip por cada interprete real detectado (excluye el alias de Microsoft Store). Util para inventariar dependencias en equipos sin herramienta de gestion centralizada.' -Wide
    }
}

Measure-Section '1.15.3 Entornos virtuales Python (venv)' {
    $data = Invoke-Remote -TimeoutSec 300 -ArgumentList (,$GitScanPaths) -ScriptBlock {
        param($paths)
        try {
            $limiteResultados = 500
            $profundidadMax   = 4
            $results = New-Object System.Collections.Generic.List[PSObject]
            # Los patrones se evaluan contra la ruta completa de cada CARPETA antes de
            # descender en ella (poda real de la recursion), en lugar de filtrar items
            # despues de que Get-ChildItem -Recurse ya recorrio carpetas completas y
            # potencialmente enormes (node_modules, cache de AppData\Local, la papelera
            # de reciclaje, el arbol completo de C:\Windows). Esto es lo que evita el
            # escaneo lento cuando -GitScanPaths incluye el perfil del usuario.
            $excludePatterns = @(
                'node_modules', '__pycache__', '\\\.git(\\|$)',
                'AppData\\Local\\Temp', 'AppData\\Local\\Packages',
                '\$Recycle\.Bin', '\\Windows(\\|$)'
            )
            $truncado = $false

            function Find-PyvenvConPoda {
                param([string]$RaizInicial, [string[]]$PatronesExcluidos, [int]$ProfMax)
                $pila = New-Object System.Collections.Generic.Stack[PSObject]
                $pila.Push([PSCustomObject]@{ Ruta = $RaizInicial; Nivel = 0 })
                while ($pila.Count -gt 0) {
                    $actual = $pila.Pop()
                    $items = $null
                    try { $items = Get-ChildItem -LiteralPath $actual.Ruta -Force -ErrorAction SilentlyContinue } catch { continue }
                    if (-not $items) { continue }
                    foreach ($item in $items) {
                        $excluido = $false
                        foreach ($pat in $PatronesExcluidos) { if ($item.FullName -match $pat) { $excluido = $true; break } }
                        if ($excluido) { continue }
                        if ($item.PSIsContainer) {
                            if ($actual.Nivel -lt $ProfMax) {
                                $pila.Push([PSCustomObject]@{ Ruta = $item.FullName; Nivel = $actual.Nivel + 1 })
                            }
                        } elseif ($item.Name -eq 'pyvenv.cfg') {
                            $item
                        }
                    }
                }
            }

            foreach ($base in $paths) {
                if ($results.Count -ge $limiteResultados) { $truncado = $true; break }
                if ($base -and (Test-Path $base)) {
                    Find-PyvenvConPoda -RaizInicial $base -PatronesExcluidos $excludePatterns -ProfMax $profundidadMax |
                        ForEach-Object {
                            if ($results.Count -lt $limiteResultados) {
                                $cfg = Get-Content $_.FullName -ErrorAction SilentlyContinue
                                $verLine = $cfg | Where-Object { $_ -match '^version' } | Select-Object -First 1
                                $results.Add([PSCustomObject]@{
                                    Ruta               = $_.Directory.FullName
                                    VersionPython      = ($verLine -replace '^version\s*=\s*', '')
                                    UltimaModificacion = $_.LastWriteTime
                                })
                            } else { $truncado = $true }
                        }
                }
            }
            if ($truncado -or $results.Count -ge $limiteResultados) {
                $results.Add([PSCustomObject]@{ Ruta = "(truncado: se alcanzo el limite de $limiteResultados resultados o la profundidad maxima de $profundidadMax niveles)"; VersionPython = ''; UltimaModificacion = $null })
            }
            $results
        } catch { "Error al buscar entornos virtuales (venv): $($_.Exception.Message)" }
    }
    Add-ReportSection -Id '1.15.3' -Title 'Entornos Virtuales Python (venv) Detectados' -Data $data -Note 'Busqueda de pyvenv.cfg acotada a las rutas de -GitScanPaths, con poda real de la recursion (no desciende a node_modules, cache de AppData\Local, la papelera de reciclaje ni el arbol de Windows), profundidad maxima de 4 niveles y un tope de 500 resultados, para evitar un escaneo lento en un file server o un perfil de usuario con muchas carpetas.' -Wide
}

Measure-Section '1.15.4 Entornos Conda detectados' {
    $condaExe = Invoke-Remote -TimeoutSec 30 -ScriptBlock {
        try { (Get-Command conda -ErrorAction Stop | Select-Object -First 1).Source } catch { $null }
    }
    $data = if (-not $condaExe) {
        'No se detecto conda.exe en el PATH de este equipo.'
    } else {
        Invoke-Remote -TimeoutSec 120 -ArgumentList @($condaExe) -ScriptBlock {
            param($conda)
            try {
                $raw = & $conda env list --json 2>$null | ConvertFrom-Json
                $raw.envs | ForEach-Object {
                    $envPath = $_
                    $pkgCount = try { (& $conda list -p $envPath --json 2>$null | ConvertFrom-Json).Count } catch { 'N/D' }
                    [PSCustomObject]@{ Entorno = $envPath; TotalPaquetes = $pkgCount }
                }
            } catch { "No se pudo listar entornos conda: $($_.Exception.Message)" }
        }
    }
    Add-ReportSection -Id '1.15.4' -Title 'Entornos Conda Detectados' -Data $data -Note 'Entornos administrados por Conda/Anaconda/Miniconda y cantidad de paquetes en cada uno.' -Wide
}

if (-not $ScanGitRepos) {
    Add-ReportSection -Id '1.15.5' -Title 'Repositorios Git Detectados' -Data 'Escaneo de repositorios Git deshabilitado por defecto (usar el parametro -ScanGitRepos para habilitarlo).' -Note 'Desactivado por defecto porque el escaneo recursivo de carpetas .git puede ser lento en discos grandes o file servers con muchos proyectos.'
} else {
    Measure-Section '1.15.5 Repositorios Git detectados' {
        $data = Invoke-Remote -TimeoutSec 300 -ArgumentList (,$GitScanPaths) -ScriptBlock {
            param($paths)
            try {
                $limite         = 500
                $profundidadMax = 4
                $results = New-Object System.Collections.Generic.List[PSObject]
                # Poda real de la recursion (ver 1.15.3): se evalua cada carpeta antes de
                # descender en ella, en vez de filtrar despues de recorrerla por completo.
                $excludePatterns = @(
                    'node_modules', '__pycache__',
                    'AppData\\Local\\Temp', 'AppData\\Local\\Packages',
                    '\$Recycle\.Bin', '\\Windows(\\|$)'
                )

                function Find-DotGitConPoda {
                    param([string]$RaizInicial, [string[]]$PatronesExcluidos, [int]$ProfMax)
                    $pila = New-Object System.Collections.Generic.Stack[PSObject]
                    $pila.Push([PSCustomObject]@{ Ruta = $RaizInicial; Nivel = 0 })
                    while ($pila.Count -gt 0) {
                        $actual = $pila.Pop()
                        $items = $null
                        try { $items = Get-ChildItem -LiteralPath $actual.Ruta -Force -Directory -ErrorAction SilentlyContinue } catch { continue }
                        if (-not $items) { continue }
                        foreach ($item in $items) {
                            if ($item.Name -eq '.git') { $item; continue }
                            $excluido = $false
                            foreach ($pat in $PatronesExcluidos) { if ($item.FullName -match $pat) { $excluido = $true; break } }
                            if ($excluido) { continue }
                            if ($actual.Nivel -lt $ProfMax) {
                                $pila.Push([PSCustomObject]@{ Ruta = $item.FullName; Nivel = $actual.Nivel + 1 })
                            }
                        }
                    }
                }

                foreach ($base in $paths) {
                    if ($results.Count -ge $limite) { break }
                    if ($base -and (Test-Path $base)) {
                        Find-DotGitConPoda -RaizInicial $base -PatronesExcluidos $excludePatterns -ProfMax $profundidadMax |
                            ForEach-Object {
                                if ($results.Count -lt $limite) {
                                    $repoPath   = $_.Parent.FullName
                                    $configFile = Join-Path $_.FullName 'config'
                                    $remote = 'N/D'
                                    if (Test-Path $configFile) {
                                        $urlLine = Get-Content $configFile -ErrorAction SilentlyContinue | Where-Object { $_ -match '^\s*url\s*=' } | Select-Object -First 1
                                        if ($urlLine) {
                                            $remote = ($urlLine -replace '^\s*url\s*=\s*', '').Trim()
                                            $remote = $remote -replace '://[^@/]+@', '://***:***@'
                                        }
                                    }
                                    $branch = 'N/D'
                                    $headFile = Join-Path $_.FullName 'HEAD'
                                    if (Test-Path $headFile) {
                                        $headContent = Get-Content $headFile -ErrorAction SilentlyContinue
                                        if ($headContent -match 'ref:\s*refs/heads/(.+)$') { $branch = $matches[1] }
                                    }
                                    $results.Add([PSCustomObject]@{
                                        Ruta               = $repoPath
                                        RemoteOrigin       = $remote
                                        RamaActual         = $branch
                                        UltimaModificacion = (Get-Item $repoPath -ErrorAction SilentlyContinue).LastWriteTime
                                    })
                                }
                            }
                    }
                }
                if ($results.Count -ge $limite) {
                    $results.Add([PSCustomObject]@{ Ruta = "(truncado: se alcanzo el limite de $limite repositorios)"; RemoteOrigin = ''; RamaActual = ''; UltimaModificacion = $null })
                }
                $results
            } catch { "Error al buscar repositorios Git: $($_.Exception.Message)" }
        }
        Add-ReportSection -Id '1.15.5' -Title 'Repositorios Git Detectados' -Data $data -Note 'Repositorios .git encontrados dentro de las rutas configuradas (-GitScanPaths), con poda real de la recursion (no desciende a node_modules, cache de AppData\Local, la papelera de reciclaje ni el arbol de Windows), profundidad maxima de 4 niveles y tope de 500 resultados. Toda credencial embebida en la URL del remoto se enmascara automaticamente (usuario:password reemplazado por ***:***).' -Wide
    }
}

}
#endregion

# =============================================================================
# COLECTORES EXTENDIDOS (secciones nuevas v3.0)
# =============================================================================

# =============================================================================
# FRAGMENTO 03 - COLECTORES NUEVOS
# Cuentas y Accesos avanzado (1.14.2-1.14.7), Plataforma y Aplicaciones (1.16.x),
# Salud/Eventos/Rendimiento (1.17.x), complementos de Red/Storage/File Server
# (1.4.6-1.4.9, 1.5.3-1.5.4, 1.7.3) y complementos de Seguridad (1.11.9-1.11.13).
#
# Consume la API compartida del CORE: Invoke-Remote, Add-ReportSection,
# Add-Finding, Test-SectionEnabled, Measure-Section, Write-ReportLog,
# Protect-Sensitive, $script:Caps, $script:IsPS7, $EventLogDays, $PerfSampleSeconds.
#
# Este fragmento NO define variables $script: ni funciones globales nuevas:
# todo el estado vive dentro de cada scriptblock de Invoke-Remote o dentro del
# cuerpo de cada Measure-Section.
# =============================================================================

#region 1.14 CUENTAS Y ACCESOS (ampliacion)
if (Test-SectionEnabled '1.14') {

# -----------------------------------------------------------------------------
# 1.14.2 Administradores locales
# -----------------------------------------------------------------------------
Measure-Section '1.14.2 Administradores locales' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $rows = New-Object System.Collections.Generic.List[PSObject]
            $members = $null
            try {
                $members = Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop
            } catch { $members = $null }

            if ($members) {
                foreach ($m in $members) {
                    $origen = 'Local'
                    if ($m.PrincipalSource) { $origen = $m.PrincipalSource.ToString() }
                    elseif ($m.Name -match '^[^\\]+\\') { $origen = 'Dominio' }
                    $sidVal = 'N/D'
                    try { $sidVal = $m.SID.Value } catch {}
                    $rows.Add([PSCustomObject]@{
                        Nombre     = $m.Name
                        TipoObjeto = $m.ObjectClass
                        Origen     = $origen
                        SID        = $sidVal
                    })
                }
            } else {
                # Fallback ADSI: Get-LocalGroupMember falla cuando el grupo tiene SIDs
                # huerfanos de un dominio de confianza que ya no existe. Se resuelve
                # el nombre localizado del grupo Administradores via el SID conocido
                # S-1-5-32-544, para no depender del idioma del sistema operativo.
                $groupName = 'Administrators'
                try {
                    $sidObj = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')
                    $groupName = (($sidObj.Translate([System.Security.Principal.NTAccount])).Value -split '\\')[-1]
                } catch {}
                $group = [ADSI]"WinNT://$env:COMPUTERNAME/$groupName,group"
                $adsiMembers = @($group.Invoke('Members'))
                foreach ($m in $adsiMembers) {
                    $clase  = try { $m.GetType().InvokeMember('Class', 'GetProperty', $null, $m, $null) } catch { 'Desconocido' }
                    $path   = try { $m.GetType().InvokeMember('ADsPath', 'GetProperty', $null, $m, $null) } catch { '' }
                    $nombre = try { $m.GetType().InvokeMember('Name', 'GetProperty', $null, $m, $null) } catch { 'Desconocido' }
                    $sidStr = 'No se pudo resolver (posible SID huerfano de dominio)'
                    try {
                        $sidBytes = $m.GetType().InvokeMember('objectSID', 'GetProperty', $null, $m, $null)
                        if ($sidBytes) { $sidStr = (New-Object System.Security.Principal.SecurityIdentifier($sidBytes, 0)).Value }
                    } catch {}
                    $origenAdsi = 'Desconocido'
                    if ($path -match '^WinNT://([^/]+)/') {
                        $origenAdsi = if ($matches[1] -eq $env:COMPUTERNAME) { 'Local' } else { 'Dominio' }
                    }
                    $rows.Add([PSCustomObject]@{
                        Nombre     = $nombre
                        TipoObjeto = $clase
                        Origen     = $origenAdsi
                        SID        = $sidStr
                    })
                }
            }
            $rows
        } catch {
            "No se pudo obtener los miembros del grupo Administradores locales: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.14.2' -Title 'Administradores Locales' -Data $data -Note 'Miembros del grupo Administradores (SID S-1-5-32-544, incluye SIDs huerfanos de dominio via fallback ADSI). Toda cuenta aqui tiene control total del equipo.'
}

# -----------------------------------------------------------------------------
# 1.14.3 Usuarios locales
# -----------------------------------------------------------------------------
Measure-Section '1.14.3 Usuarios locales' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $rows = New-Object System.Collections.Generic.List[PSObject]
            $users = $null
            try { $users = Get-LocalUser -ErrorAction Stop } catch { $users = $null }

            if ($users) {
                foreach ($u in $users) {
                    $dias = 'N/D'
                    if ($u.PasswordLastSet) { $dias = [math]::Round(((Get-Date) - $u.PasswordLastSet).TotalDays, 0) }
                    $rows.Add([PSCustomObject]@{
                        Nombre                  = $u.Name
                        Habilitado              = $u.Enabled
                        PasswordNeverExpires    = $u.PasswordNeverExpires
                        PasswordRequired        = $u.PasswordRequired
                        UltimoLogon             = $u.LastLogon
                        PasswordLastSet         = $u.PasswordLastSet
                        DiasDesdeCambioPassword = $dias
                        Descripcion             = $u.Description
                    })
                }
            } else {
                # Fallback CIM: disponible incluso sin el modulo Microsoft.PowerShell.LocalAccounts
                $cimUsers = Get-CimInstance Win32_UserAccount -Filter "LocalAccount='True'" -ErrorAction Stop
                foreach ($u in $cimUsers) {
                    $rows.Add([PSCustomObject]@{
                        Nombre                  = $u.Name
                        Habilitado              = (-not $u.Disabled)
                        PasswordNeverExpires    = ($u.PasswordExpires -eq $false)
                        PasswordRequired        = $u.PasswordRequired
                        UltimoLogon             = 'N/D (requiere modulo LocalAccounts, no disponible en este equipo)'
                        PasswordLastSet         = 'N/D (requiere modulo LocalAccounts, no disponible en este equipo)'
                        DiasDesdeCambioPassword = 'N/D'
                        Descripcion             = $u.Description
                    })
                }
            }
            $rows
        } catch {
            "No se pudo obtener los usuarios locales: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.14.3' -Title 'Usuarios Locales' -Data $data -Note 'Cuentas de usuario locales (SAM) del equipo. No incluye cuentas de dominio.' -Wide
}

# -----------------------------------------------------------------------------
# 1.14.4 Grupo Remote Desktop Users
# -----------------------------------------------------------------------------
Measure-Section '1.14.4 Grupo Remote Desktop Users' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $rows = New-Object System.Collections.Generic.List[PSObject]
            $members = $null
            try { $members = Get-LocalGroupMember -SID 'S-1-5-32-555' -ErrorAction Stop } catch { $members = $null }

            if ($members) {
                foreach ($m in $members) {
                    $origen = 'Local'
                    if ($m.PrincipalSource) { $origen = $m.PrincipalSource.ToString() }
                    elseif ($m.Name -match '^[^\\]+\\') { $origen = 'Dominio' }
                    $sidVal = 'N/D'
                    try { $sidVal = $m.SID.Value } catch {}
                    $rows.Add([PSCustomObject]@{
                        Nombre     = $m.Name
                        TipoObjeto = $m.ObjectClass
                        Origen     = $origen
                        SID        = $sidVal
                    })
                }
                if ($rows.Count -eq 0) { 'El grupo local Remote Desktop Users existe pero no tiene miembros explicitos asignados en este equipo.' } else { $rows }
            } else {
                $groupName = 'Remote Desktop Users'
                try {
                    $sidObj = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-555')
                    $groupName = (($sidObj.Translate([System.Security.Principal.NTAccount])).Value -split '\\')[-1]
                } catch {}
                $group = [ADSI]"WinNT://$env:COMPUTERNAME/$groupName,group"
                if ($group.Path) {
                    $adsiMembers = @($group.Invoke('Members'))
                    foreach ($m in $adsiMembers) {
                        $clase  = try { $m.GetType().InvokeMember('Class', 'GetProperty', $null, $m, $null) } catch { 'Desconocido' }
                        $path   = try { $m.GetType().InvokeMember('ADsPath', 'GetProperty', $null, $m, $null) } catch { '' }
                        $nombre = try { $m.GetType().InvokeMember('Name', 'GetProperty', $null, $m, $null) } catch { 'Desconocido' }
                        $sidStr = 'No se pudo resolver (posible SID huerfano de dominio)'
                        try {
                            $sidBytes = $m.GetType().InvokeMember('objectSID', 'GetProperty', $null, $m, $null)
                            if ($sidBytes) { $sidStr = (New-Object System.Security.Principal.SecurityIdentifier($sidBytes, 0)).Value }
                        } catch {}
                        $origenAdsi = 'Desconocido'
                        if ($path -match '^WinNT://([^/]+)/') {
                            $origenAdsi = if ($matches[1] -eq $env:COMPUTERNAME) { 'Local' } else { 'Dominio' }
                        }
                        $rows.Add([PSCustomObject]@{
                            Nombre     = $nombre
                            TipoObjeto = $clase
                            Origen     = $origenAdsi
                            SID        = $sidStr
                        })
                    }
                    if ($rows.Count -eq 0) { 'El grupo local Remote Desktop Users existe pero no tiene miembros explicitos asignados en este equipo.' } else { $rows }
                } else {
                    "El grupo local Remote Desktop Users no existe o no tiene miembros en este equipo (comun en Domain Controllers)."
                }
            }
        } catch {
            "No se pudo obtener los miembros del grupo Remote Desktop Users: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.14.4' -Title 'Grupo Remote Desktop Users' -Data $data -Note 'Cuentas autorizadas a iniciar sesion remota via RDP sin ser Administradores (SID S-1-5-32-555).'
}

# -----------------------------------------------------------------------------
# 1.14.5 Politica de contrasenas y bloqueo
# -----------------------------------------------------------------------------
Measure-Section '1.14.5 Politica de contrasenas y bloqueo' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $result = [PSCustomObject]@{
                LongitudMinima     = 'N/D'
                VigenciaMaxima     = 'N/D'
                VigenciaMinima     = 'N/D'
                HistorialPasswords = 'N/D'
                UmbralBloqueo      = 'N/D'
                DuracionBloqueo    = 'N/D'
            }
            $viaWmi = $false
            try {
                $pol = Get-CimInstance Win32_AccountPolicy -ErrorAction Stop | Select-Object -First 1
                if ($pol) {
                    $result.LongitudMinima     = $pol.MinPasswordLength
                    $result.VigenciaMaxima     = $pol.MaxPasswordAge
                    $result.VigenciaMinima     = $pol.MinPasswordAge
                    $result.HistorialPasswords = $pol.PasswordHistoryLength
                    $result.UmbralBloqueo      = $pol.LockoutThreshold
                    $result.DuracionBloqueo    = $pol.LockoutDuration
                    $viaWmi = $true
                }
            } catch { $viaWmi = $false }

            if (-not $viaWmi) {
                # Win32_AccountPolicy no esta registrado por defecto en la mayoria de
                # los equipos modernos; se recurre a parsear "net accounts", tolerando
                # encabezados en ingles o en espanol.
                $lineas = net accounts 2>&1

                function Get-ValorPoliticaNet {
                    param([string[]]$Lineas, [string]$PatronClave)
                    foreach ($l in $Lineas) {
                        if ($l -match $PatronClave) {
                            if ($l -match '(\d+)\s*$') { return $matches[1] }
                            if ($l -match '(?i)never|nunca') { return 'Nunca' }
                        }
                    }
                    return 'N/D'
                }

                $result.LongitudMinima     = Get-ValorPoliticaNet -Lineas $lineas -PatronClave '(?i)minimum.*password.*length|longitud.*m.nima.*contrase|longitud minima'
                $result.VigenciaMaxima     = Get-ValorPoliticaNet -Lineas $lineas -PatronClave '(?i)maximum.*password.*age|vigencia.*m.xima.*contrase|vigencia maxima'
                $result.VigenciaMinima     = Get-ValorPoliticaNet -Lineas $lineas -PatronClave '(?i)minimum.*password.*age|vigencia.*m.nima.*contrase|vigencia minima'
                $result.HistorialPasswords = Get-ValorPoliticaNet -Lineas $lineas -PatronClave '(?i)password.*history|historial.*contrase'
                $result.UmbralBloqueo      = Get-ValorPoliticaNet -Lineas $lineas -PatronClave '(?i)lockout.*threshold|umbral.*bloqueo'
                $result.DuracionBloqueo    = Get-ValorPoliticaNet -Lineas $lineas -PatronClave '(?i)lockout.*duration|duracion.*bloqueo'
            }
            $result
        } catch {
            "No se pudo obtener la politica de contrasenas y bloqueo: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.14.5' -Title 'Politica de Contrasenas y Bloqueo' -Data $data -Note 'Politica de cuentas local (Win32_AccountPolicy con fallback a net accounts). En equipos miembro de dominio puede estar sobrescrita por GPO de dominio (ver seccion 1.13).'
}

# -----------------------------------------------------------------------------
# 1.14.6 Politica de auditoria
# -----------------------------------------------------------------------------
Measure-Section '1.14.6 Politica de auditoria' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $csvRaw = auditpol /get /category:* /r 2>&1
            if (-not $csvRaw) { throw "auditpol no devolvio salida" }
            $parsed = $csvRaw | ConvertFrom-Csv -ErrorAction Stop
            if (-not $parsed) { throw "ConvertFrom-Csv no genero filas" }

            $headers = ($parsed | Select-Object -First 1 | Get-Member -MemberType NoteProperty).Name
            $subCol  = $headers | Where-Object { $_ -match '(?i)subcategory|subcategoria' -and $_ -notmatch '(?i)guid' } | Select-Object -First 1
            $incCol  = $headers | Where-Object { $_ -match '(?i)inclusion' } | Select-Object -First 1
            $excCol  = $headers | Where-Object { $_ -match '(?i)exclusion' } | Select-Object -First 1

            # auditpol /r no expone la categoria por fila (solo la subcategoria);
            # se infiere via un mapeo de las subcategorias estandar de Windows en ingles.
            $mapaCategoria = @{
                'Security State Change'='Cambio de estado del sistema'; 'Security System Extension'='Extension del sistema de seguridad'
                'System Integrity'='Integridad del sistema'; 'IPsec Driver'='Controlador IPsec'; 'Other System Events'='Otros eventos del sistema'
                'Logon'='Inicio y cierre de sesion'; 'Logoff'='Inicio y cierre de sesion'; 'Account Lockout'='Inicio y cierre de sesion'
                'IPsec Main Mode'='Inicio y cierre de sesion'; 'IPsec Quick Mode'='Inicio y cierre de sesion'; 'IPsec Extended Mode'='Inicio y cierre de sesion'
                'Special Logon'='Inicio y cierre de sesion'; 'Other Logon/Logoff Events'='Inicio y cierre de sesion'; 'Network Policy Server'='Inicio y cierre de sesion'
                'User / Device Claims'='Inicio y cierre de sesion'; 'Group Membership'='Inicio y cierre de sesion'
                'File System'='Acceso a objetos'; 'Registry'='Acceso a objetos'; 'Kernel Object'='Acceso a objetos'; 'SAM'='Acceso a objetos'
                'Certification Services'='Acceso a objetos'; 'Application Generated'='Acceso a objetos'; 'Handle Manipulation'='Acceso a objetos'
                'File Share'='Acceso a objetos'; 'Filtering Platform Packet Drop'='Acceso a objetos'; 'Filtering Platform Connection'='Acceso a objetos'
                'Other Object Access Events'='Acceso a objetos'; 'Detailed File Share'='Acceso a objetos'; 'Removable Storage'='Acceso a objetos'
                'Central Policy Staging'='Acceso a objetos'
                'Sensitive Privilege Use'='Uso de privilegios'; 'Non Sensitive Privilege Use'='Uso de privilegios'; 'Other Privilege Use Events'='Uso de privilegios'
                'Process Creation'='Seguimiento detallado'; 'Process Termination'='Seguimiento detallado'; 'DPAPI Activity'='Seguimiento detallado'
                'RPC Events'='Seguimiento detallado'; 'Plug and Play Events'='Seguimiento detallado'; 'Token Right Adjusted Events'='Seguimiento detallado'
                'Audit Policy Change'='Cambio de directivas'; 'Authentication Policy Change'='Cambio de directivas'; 'Authorization Policy Change'='Cambio de directivas'
                'MPSSVC Rule-Level Policy Change'='Cambio de directivas'; 'Filtering Platform Policy Change'='Cambio de directivas'; 'Other Policy Change Events'='Cambio de directivas'
                'User Account Management'='Administracion de cuentas'; 'Computer Account Management'='Administracion de cuentas'; 'Security Group Management'='Administracion de cuentas'
                'Distribution Group Management'='Administracion de cuentas'; 'Application Group Management'='Administracion de cuentas'; 'Other Account Management Events'='Administracion de cuentas'
                'Directory Service Access'='Acceso DS'; 'Directory Service Changes'='Acceso DS'; 'Directory Service Replication'='Acceso DS'; 'Detailed Directory Service Replication'='Acceso DS'
                'Credential Validation'='Inicio de sesion de cuenta'; 'Kerberos Service Ticket Operations'='Inicio de sesion de cuenta'
                'Other Account Logon Events'='Inicio de sesion de cuenta'; 'Kerberos Authentication Service'='Inicio de sesion de cuenta'
            }

            $parsed | ForEach-Object {
                $sub = if ($subCol) { $_.$subCol } else { 'N/D' }
                $cat = if ($mapaCategoria.ContainsKey($sub)) { $mapaCategoria[$sub] } else { 'N/D (SO no ingles o subcategoria no reconocida)' }
                $inc = if ($incCol) { $_.$incCol } else { 'N/D' }
                $exc = if ($excCol) { $_.$excCol } else { '' }
                $cfg = if ($exc -and $exc -ne 'No Auditing' -and $exc -ne '') { "$inc (Exclusion: $exc)" } else { $inc }
                [PSCustomObject]@{
                    Categoria     = $cat
                    Subcategoria  = $sub
                    Configuracion = $cfg
                }
            }
        } catch {
            "No se pudo obtener la politica de auditoria (auditpol): $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.14.6' -Title 'Politica de Auditoria' -Data $data -Note 'Configuracion de auditoria avanzada por subcategoria (auditpol /get /category:* /r). La categoria se infiere por nombre de subcategoria en ingles; en SO localizados puede figurar como no reconocida.' -Wide
}

# -----------------------------------------------------------------------------
# 1.14.7 Derechos de usuario sensibles
# -----------------------------------------------------------------------------
Measure-Section '1.14.7 Derechos de usuario sensibles' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        $tempFile = Join-Path $env:TEMP "secedit_userrights_$([guid]::NewGuid().ToString('N')).cfg"
        try {
            $derechosDeInteres = @(
                'SeDebugPrivilege', 'SeBackupPrivilege', 'SeTakeOwnershipPrivilege',
                'SeServiceLogonRight', 'SeNetworkLogonRight', 'SeDenyNetworkLogonRight',
                'SeRemoteInteractiveLogonRight'
            )
            $nombresAmigables = @{
                'SeDebugPrivilege'             = 'Depurar programas (SeDebugPrivilege)'
                'SeBackupPrivilege'            = 'Realizar copias de seguridad (SeBackupPrivilege)'
                'SeTakeOwnershipPrivilege'     = 'Tomar posesion de archivos u objetos (SeTakeOwnershipPrivilege)'
                'SeServiceLogonRight'          = 'Iniciar sesion como servicio (SeServiceLogonRight)'
                'SeNetworkLogonRight'          = 'Acceso a este equipo desde la red (SeNetworkLogonRight)'
                'SeDenyNetworkLogonRight'      = 'Denegar acceso desde la red (SeDenyNetworkLogonRight)'
                'SeRemoteInteractiveLogonRight'= 'Permitir inicio de sesion mediante Escritorio Remoto (SeRemoteInteractiveLogonRight)'
            }

            $exportOutput = secedit /export /areas USER_RIGHTS /cfg $tempFile /quiet 2>&1
            if (-not (Test-Path $tempFile)) {
                throw "secedit no genero el archivo de exportacion: $exportOutput"
            }

            $contenido = Get-Content -Path $tempFile -ErrorAction Stop
            $rows = New-Object System.Collections.Generic.List[PSObject]

            foreach ($derecho in $derechosDeInteres) {
                $linea = $contenido | Where-Object { $_ -match "^\s*$derecho\s*=" } | Select-Object -First 1
                $titulo = if ($nombresAmigables.ContainsKey($derecho)) { $nombresAmigables[$derecho] } else { $derecho }
                if (-not $linea) {
                    $rows.Add([PSCustomObject]@{ Derecho = $titulo; AsignadoA = '(sin asignaciones)' })
                    continue
                }
                $valor = ($linea -split '=', 2)[1].Trim()
                $sidsCrudos = $valor -split ',' | Where-Object { $_ } | ForEach-Object { $_.Trim().TrimStart('*') }
                $nombresResueltos = foreach ($sidTxt in $sidsCrudos) {
                    try {
                        $sidObj = New-Object System.Security.Principal.SecurityIdentifier($sidTxt)
                        try { $sidObj.Translate([System.Security.Principal.NTAccount]).Value } catch { "$sidTxt (SID no resoluble, posiblemente huerfano)" }
                    } catch {
                        $sidTxt
                    }
                }
                $asignado = if ($nombresResueltos) { ($nombresResueltos -join '; ') } else { '(sin asignaciones)' }
                $rows.Add([PSCustomObject]@{ Derecho = $titulo; AsignadoA = $asignado })
            }
            $rows
        } catch {
            "No se pudo obtener los derechos de usuario (secedit): $($_.Exception.Message)"
        } finally {
            if (Test-Path $tempFile) { Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue }
        }
    }
    Add-ReportSection -Id '1.14.7' -Title 'Derechos de Usuario Sensibles' -Data $data -Note 'Asignacion de derechos de usuario criticos (secedit /export /areas USER_RIGHTS). El archivo temporal generado se elimina automaticamente al finalizar.' -Wide
}

} # fin Test-SectionEnabled '1.14'
#endregion

#region 1.16 PLATAFORMA Y APLICACIONES
if (Test-SectionEnabled '1.16') {

# -----------------------------------------------------------------------------
# 1.16.1 Dominio y ubicacion en AD
# -----------------------------------------------------------------------------
Measure-Section '1.16.1 Dominio y ubicacion en AD' {
    $data = Invoke-Remote -TimeoutSec 90 -ScriptBlock {
        try {
            $cs = Get-CimInstance Win32_ComputerSystem
            if (-not $cs.PartOfDomain) {
                return "El equipo no es miembro de un dominio (grupo de trabajo: $($cs.Workgroup))."
            }

            $dominio = $cs.Domain
            $forest = $dominio
            try {
                $forest = ([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()).Forest.Name
            } catch {}

            $sitio = 'N/D'
            try {
                $dsgetsiteOut = nltest /dsgetsite 2>&1 | Out-String
                if ($dsgetsiteOut -notmatch '(?i)error|no se pudo|command failed|no se reconoce') {
                    $sitio = (($dsgetsiteOut -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -First 1).Trim()
                } else {
                    $sitio = "No se pudo determinar: $($dsgetsiteOut.Trim())"
                }
            } catch { $sitio = 'No se pudo ejecutar nltest /dsgetsite' }

            $dcAutenticante = if ($env:LOGONSERVER) { $env:LOGONSERVER -replace '\\\\', '' } else { 'N/D' }

            $dcDominio = 'N/D'
            try {
                $dsgetdcOut = nltest /dsgetdc:$dominio 2>&1 | Out-String
                $lineaDc = ($dsgetdcOut -split "`r?`n") | Where-Object { $_ -match 'DC:|Nombre del PDC|\\\\' } | Select-Object -First 1
                if ($lineaDc) { $dcDominio = $lineaDc.Trim() } else { $dcDominio = (($dsgetdcOut.Trim() -split "`r?`n") | Select-Object -First 1) }
            } catch { $dcDominio = 'No se pudo ejecutar nltest /dsgetdc' }

            $dnCompleto = 'N/D (modulo ActiveDirectory no disponible en este equipo)'
            try {
                if (Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue) {
                    Import-Module ActiveDirectory -ErrorAction Stop
                    $adComputer = Get-ADComputer -Identity $env:COMPUTERNAME -ErrorAction Stop
                    $dnCompleto = $adComputer.DistinguishedName
                }
            } catch { $dnCompleto = "No se pudo consultar AD: $($_.Exception.Message)" }

            $canalSeguro = 'N/D'
            try { $canalSeguro = Test-ComputerSecureChannel -ErrorAction Stop } catch { $canalSeguro = "No se pudo probar: $($_.Exception.Message)" }

            [PSCustomObject]@{
                Dominio        = $dominio
                Forest         = $forest
                SitioAD        = $sitio
                DCAutenticante = $dcAutenticante
                DCDelDominio   = $dcDominio
                DNCompleto     = $dnCompleto
                CanalSeguroOK  = $canalSeguro
            }
        } catch {
            "No se pudo obtener la informacion de dominio/AD: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.16.1' -Title 'Dominio y Ubicacion en Active Directory' -Data $data -Note 'Pertenencia a dominio, forest, sitio de AD, DC autenticante y canal seguro con el DC.'
}

# -----------------------------------------------------------------------------
# 1.16.2 Replicacion AD (solo si el equipo es Controlador de Dominio)
# -----------------------------------------------------------------------------
if ($script:Caps.IsDC) {
Measure-Section '1.16.2 Replicacion AD' {
    $raw = Invoke-Remote -TimeoutSec 180 -ScriptBlock {
        try {
            $rows = New-Object System.Collections.Generic.List[PSObject]
            $csvRaw = repadmin /showrepl /csv 2>&1
            $csvTxt = ($csvRaw | Out-String)
            if ($csvTxt -notmatch '(?i)no se pudo|is not recognized|no se reconoce') {
                $parsed = $csvRaw | ConvertFrom-Csv -ErrorAction SilentlyContinue
                foreach ($p in $parsed) {
                    $rows.Add([PSCustomObject]@{
                        SocioReplicacion = $p.'Source DSA'
                        NamingContext    = $p.'Naming Context'
                        UltimoIntento    = $p.'Last Success Time'
                        UltimoResultado  = $p.'Last Failure Status'
                        NumeroDeFallos   = $p.'Number of Failures'
                    })
                }
            }
            $dcdiagRaw = dcdiag /test:Replications /test:Advertising /test:FSMOCheck 2>&1 | Out-String
            $resumenLineas = ($dcdiagRaw -split "`r?`n") | Where-Object { $_ -match '(?i)passed|failed|pas.|fall.' }
            $resumen = if ($resumenLineas) { ($resumenLineas -join ' | ').Trim() } else { 'No se pudo obtener resumen de dcdiag' }

            [PSCustomObject]@{ Filas = $rows; Resumen = $resumen }
        } catch {
            [PSCustomObject]@{ Filas = $null; Resumen = "Error: $($_.Exception.Message)" }
        }
    }

    if ($raw -is [string]) {
        Add-ReportSection -Id '1.16.2' -Title 'Replicacion de Active Directory' -Data $raw -Note 'Estado de replicacion entre controladores de dominio.'
    } else {
        $filas = $raw.Filas
        $dataFinal = if ($filas -and @($filas).Count -gt 0) { $filas } else { 'repadmin no devolvio filas de replicacion (verificar permisos de Domain Admins / Enterprise Admins).' }
        Add-ReportSection -Id '1.16.2' -Title 'Replicacion de Active Directory' -Data $dataFinal -Note "Detalle por socio de replicacion (repadmin /showrepl). Resumen dcdiag: $($raw.Resumen)" -Wide
    }
}
}

# -----------------------------------------------------------------------------
# 1.16.3 Capa de virtualizacion
# -----------------------------------------------------------------------------
Measure-Section '1.16.3 Capa de virtualizacion' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $cs = Get-CimInstance Win32_ComputerSystem
            $modelo = "$($cs.Manufacturer) $($cs.Model)"
            $hipervisor = 'Fisico / no detectado'
            if ($modelo -match 'VMware') { $hipervisor = 'VMware' }
            elseif ($modelo -match 'Virtual Machine' -and $cs.Manufacturer -match 'Microsoft') { $hipervisor = 'Hyper-V' }
            elseif ($modelo -match 'KVM|QEMU') { $hipervisor = 'KVM/QEMU' }
            elseif ($modelo -match 'Xen') { $hipervisor = 'Xen' }

            $detalle = [ordered]@{
                Hipervisor   = $hipervisor
                FabricanteHW = $cs.Manufacturer
                ModeloHW     = $cs.Model
            }

            if ($hipervisor -eq 'VMware') {
                $vmToolsKey = 'HKLM:\SOFTWARE\VMware, Inc.\VMware Tools'
                if (Test-Path $vmToolsKey) {
                    $vmToolsProps = Get-ItemProperty -Path $vmToolsKey -ErrorAction SilentlyContinue
                    $detalle.VMwareToolsInstallPath = $vmToolsProps.InstallPath
                    $detalle.VMwareToolsVersion      = $vmToolsProps.CurrentVersion
                } else {
                    $detalle.VMwareTools = 'No se detecto la clave de registro de VMware Tools (no instalado).'
                }
                $svc = Get-Service -Name VMTools -ErrorAction SilentlyContinue
                $detalle.VMwareToolsServicioEstado = if ($svc) { $svc.Status } else { 'Servicio VMTools no encontrado.' }
            }
            elseif ($hipervisor -eq 'Hyper-V') {
                $integSvcs = Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^vmic' }
                $detalle.IntegrationServices = if ($integSvcs) {
                    ($integSvcs | ForEach-Object { "$($_.DisplayName): $($_.Status)" }) -join ' | '
                } else { 'No se detectaron servicios de Integration Services (vmic*).' }
            }

            [PSCustomObject]$detalle
        } catch {
            "No se pudo determinar la capa de virtualizacion: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.16.3' -Title 'Capa de Virtualizacion' -Data $data -Note 'Deteccion de hipervisor y estado de las herramientas de integracion (VMware Tools / Hyper-V Integration Services).'

    if ($script:Caps.HyperV) {
        $vmsHost = Invoke-Remote -TimeoutSec 90 -ScriptBlock {
            try {
                Import-Module Hyper-V -ErrorAction Stop
                $vms = @(Get-VM -ErrorAction Stop | Select-Object @{N='Nombre';E={$_.Name}},
                    @{N='Estado';E={$_.State}},
                    @{N='CPUs';E={$_.ProcessorCount}},
                    @{N='MemoriaGB';E={[math]::Round($_.MemoryAssigned/1GB,2)}},
                    @{N='Uptime';E={$_.Uptime}},
                    @{N='Version';E={$_.Version}})
                if ($vms.Count -eq 0) {
                    'El rol Hyper-V esta presente pero no hay ninguna maquina virtual creada en este host.'
                } else {
                    $vms
                }
            } catch {
                "No se pudo obtener la lista de VMs del host Hyper-V (el rol Hyper-V esta presente segun deteccion de capacidades, pero la consulta fallo): $($_.Exception.Message)"
            }
        }
        Add-ReportSection -Id '1.16.3.1' -Title 'Maquinas Virtuales Alojadas (Host Hyper-V)' -Data $vmsHost -Note 'Este equipo actua como host de Hyper-V. Listado de VMs alojadas; se distingue explicitamente el caso de rol presente sin VMs creadas de una falla de consulta.' -Wide
    }
}

# -----------------------------------------------------------------------------
# 1.16.4 Instancias SQL Server (solo si se detecto el rol/motor)
# -----------------------------------------------------------------------------
if ($script:Caps.SQL) {
Measure-Section '1.16.4 Instancias SQL Server' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $rows = New-Object System.Collections.Generic.List[PSObject]
            $instKey = 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL'
            if (-not (Test-Path $instKey)) { $instKey = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL' }
            if (-not (Test-Path $instKey)) { throw "No se encontro la clave de instancias de SQL Server en el registro." }

            $instancias = Get-ItemProperty -Path $instKey -ErrorAction Stop
            $nombres = $instancias.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' }

            foreach ($inst in $nombres) {
                $nombreInstancia = $inst.Name
                $idInstancia = $inst.Value
                $verEdicionPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$idInstancia\Setup"
                $verInfo = Get-ItemProperty -Path $verEdicionPath -ErrorAction SilentlyContinue

                $servicioNombre = if ($nombreInstancia -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { "MSSQL`$$nombreInstancia" }
                $servicio = Get-CimInstance Win32_Service -Filter "Name='$servicioNombre'" -ErrorAction SilentlyContinue
                $agenteNombre = if ($nombreInstancia -eq 'MSSQLSERVER') { 'SQLSERVERAGENT' } else { "SQLAgent`$$nombreInstancia" }
                $agente = Get-CimInstance Win32_Service -Filter "Name='$agenteNombre'" -ErrorAction SilentlyContinue

                $puertoTcp = 'N/D'
                try {
                    $tcpPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$idInstancia\MSSQLServer\SuperSocketNetLib\Tcp\IPAll"
                    if (Test-Path $tcpPath) {
                        $tcpProps = Get-ItemProperty -Path $tcpPath -ErrorAction SilentlyContinue
                        if ($tcpProps.TcpPort) { $puertoTcp = $tcpProps.TcpPort }
                        elseif ($tcpProps.TcpDynamicPorts) { $puertoTcp = "$($tcpProps.TcpDynamicPorts) (dinamico)" }
                    }
                } catch {}

                $rows.Add([PSCustomObject]@{
                    NombreInstancia = $nombreInstancia
                    IdInstancia     = $idInstancia
                    Version         = $verInfo.Version
                    Edicion         = $verInfo.Edition
                    ServicioMotor   = if ($servicio) { "$($servicio.State) (Cuenta: $($servicio.StartName))" } else { 'No encontrado' }
                    ServicioAgente  = if ($agente) { "$($agente.State) (Cuenta: $($agente.StartName))" } else { 'No instalado' }
                    PuertoTCP       = $puertoTcp
                })
            }
            $rows
        } catch {
            "No se pudo obtener las instancias de SQL Server: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.16.4' -Title 'Instancias de SQL Server' -Data $data -Note 'Instancias detectadas via registro (nombre, version, edicion, puerto TCP, cuentas de servicio). No se realizan conexiones a las bases de datos.' -Wide
}
}

# -----------------------------------------------------------------------------
# 1.16.5 Roles adicionales presentes
# -----------------------------------------------------------------------------
Measure-Section '1.16.5 Roles adicionales presentes' {
    $dfsData = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        try {
            Import-Module DFSN -ErrorAction Stop
            $roots = Get-DfsnRoot -ErrorAction Stop
            if ($roots) { $roots | Select-Object Path, Type, State, TimeToLiveSec } else { 'No hay namespaces DFS configurados en este equipo.' }
        } catch { 'Rol DFS Namespaces no instalado o modulo DFSN no disponible en este equipo.' }
    }
    Add-ReportSection -Id '1.16.5.1' -Title 'DFS Namespaces' -Data $dfsData -Note 'Raices de namespace DFS alojadas en este servidor.'

    $printersData = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        try {
            $shared = Get-Printer -ErrorAction Stop | Where-Object { $_.Shared }
            if ($shared) { $shared | Select-Object Name, ShareName, DriverName, PortName, PublishedToAD } else { 'No hay impresoras compartidas en este equipo.' }
        } catch { 'No se pudo consultar el rol de impresion (Get-Printer no disponible).' }
    }
    Add-ReportSection -Id '1.16.5.2' -Title 'Impresoras Compartidas' -Data $printersData -Note 'Impresoras publicadas como recurso compartido de red desde este servidor.'

    if ($script:Caps.Cluster) {
        $clusterData = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
            try {
                Import-Module FailoverClusters -ErrorAction Stop
                $cluster = Get-Cluster -ErrorAction Stop
                $nodos = Get-ClusterNode -ErrorAction SilentlyContinue | Select-Object Name, State, Type
                $recursos = Get-ClusterResource -ErrorAction SilentlyContinue | Select-Object Name, State, OwnerGroup, ResourceType
                [PSCustomObject]@{ NombreCluster = $cluster.Name; Nodos = $nodos; Recursos = $recursos }
            } catch { "No se pudo obtener informacion del Failover Cluster: $($_.Exception.Message)" }
        }
        if ($clusterData -is [string]) {
            Add-ReportSection -Id '1.16.5.3' -Title 'Failover Cluster' -Data $clusterData -Note 'Estado del rol de Cluster de Conmutacion por Error.'
        } else {
            Add-ReportSection -Id '1.16.5.3' -Title 'Failover Cluster - Nodos' -Data $clusterData.Nodos -Note "Cluster: $($clusterData.NombreCluster)."
            Add-ReportSection -Id '1.16.5.4' -Title 'Failover Cluster - Recursos' -Data $clusterData.Recursos -Note "Cluster: $($clusterData.NombreCluster)." -Wide
        }
    }

    if ($script:Caps.Docker) {
        $dockerData = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
            try {
                $psOut = docker ps -a --format "{{.Names}}|{{.Image}}|{{.Status}}|{{.Ports}}" 2>&1
                if ($LASTEXITCODE -ne 0) { throw ($psOut | Out-String) }
                $psOut | Where-Object { $_ } | ForEach-Object {
                    $partes = $_ -split '\|'
                    [PSCustomObject]@{
                        Nombre  = $partes[0]
                        Imagen  = $partes[1]
                        Estado  = $partes[2]
                        Puertos = if ($partes.Count -gt 3) { $partes[3] } else { '' }
                    }
                }
            } catch { "No se pudo consultar Docker (docker ps): $($_.Exception.Message)" }
        }
        Add-ReportSection -Id '1.16.5.5' -Title 'Contenedores Docker' -Data $dockerData -Note 'Contenedores Docker presentes en el equipo (docker ps -a).' -Wide
    }
}

# -----------------------------------------------------------------------------
# 1.16.6 NIC Teaming / LBFO
# -----------------------------------------------------------------------------
Measure-Section '1.16.6 NIC Teaming / LBFO' {
    $data = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        try {
            $teams = Get-NetLbfoTeam -ErrorAction Stop
            if (-not $teams) { return 'No hay equipos NIC Teaming (LBFO) configurados en este servidor.' }
            $teams | ForEach-Object {
                $teamName = $_.Name
                $miembros = Get-NetLbfoTeamMember -Team $teamName -ErrorAction SilentlyContinue
                [PSCustomObject]@{
                    NombreTeam   = $teamName
                    ModoTeaming  = $_.TeamingMode
                    ModoBalanceo = $_.LoadBalancingAlgorithm
                    Estado       = $_.Status
                    Miembros     = if ($miembros) { ($miembros | ForEach-Object { "$($_.Name) ($($_.AdministrativeMode)/$($_.OperationalStatus))" }) -join ' | ' } else { 'N/D' }
                }
            }
        } catch {
            'NIC Teaming (LBFO) no disponible en este SO (requiere Windows Server; no soportado en Windows 10/11 cliente) o no hay equipos configurados.'
        }
    }
    Add-ReportSection -Id '1.16.6' -Title 'NIC Teaming / LBFO' -Data $data -Note 'Adaptadores de red agrupados en equipos de balanceo de carga y conmutacion por error (LBFO).' -Wide
}

# -----------------------------------------------------------------------------
# 1.16.7 iSCSI y MPIO
# -----------------------------------------------------------------------------
Measure-Section '1.16.7 iSCSI y MPIO' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        $resultado = [ordered]@{}

        try {
            $targets = Get-IscsiTarget -ErrorAction Stop
            $resultado.TargetsISCSI = if ($targets) {
                ($targets | ForEach-Object { "$($_.NodeAddress) (Conectado: $($_.IsConnected))" }) -join ' | '
            } else { 'No hay targets iSCSI registrados.' }
        } catch { $resultado.TargetsISCSI = 'Servicio iSCSI Initiator no disponible o no iniciado.' }

        try {
            $conns = Get-IscsiConnection -ErrorAction Stop
            $resultado.ConexionesISCSI = if ($conns) {
                ($conns | ForEach-Object { "$($_.TargetAddress):$($_.TargetPortNumber) <- $($_.InitiatorAddress)" }) -join ' | '
            } else { 'No hay conexiones iSCSI activas.' }
        } catch { $resultado.ConexionesISCSI = 'No se pudo consultar conexiones iSCSI.' }

        try {
            $mpio = Get-MPIOSetting -ErrorAction Stop
            $resultado.ConfiguracionMPIO = "PathVerificationState=$($mpio.PathVerificationState); PDORemovePeriod=$($mpio.PDORemovePeriod); RetryCount=$($mpio.RetryCount); RetryInterval=$($mpio.RetryInterval)"
        } catch { $resultado.ConfiguracionMPIO = 'Caracteristica MPIO no instalada.' }

        try {
            $mpclaimOut = mpclaim -s -d 2>&1 | Out-String
            $resultado.DiscosMPIO = if ($mpclaimOut -and $mpclaimOut.Trim()) { $mpclaimOut.Trim() } else { 'mpclaim no devolvio informacion de discos.' }
        } catch { $resultado.DiscosMPIO = 'Utilidad mpclaim no disponible (requiere caracteristica Multipath I/O instalada).' }

        [PSCustomObject]$resultado
    }
    Add-ReportSection -Id '1.16.7' -Title 'iSCSI y MPIO' -Data $data -Note 'Iniciador iSCSI y Multipath I/O. Cada campo indica explicitamente si el componente no esta instalado.'
}

# -----------------------------------------------------------------------------
# 1.16.8 Configuracion de WinRM y PowerShell Remoting
# -----------------------------------------------------------------------------
Measure-Section '1.16.8 Configuracion de WinRM y PowerShell Remoting' {
    $data = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        try {
            # Se lee todo via el proveedor WSMan: y el registro, en lugar de invocar el
            # binario winrm.cmd (un script de cmd.exe): winrm puede quedar esperando
            # entrada interactiva en ciertos estados del servicio, lo que hace fallar
            # Wait-Job con "uno o varios trabajos estan bloqueados esperando la entrada
            # del usuario". El proveedor WSMan: expone la misma informacion sin invocar
            # ningun proceso externo.
            if (-not (Test-Path 'WSMan:\localhost')) {
                throw "El proveedor WSMan: no esta disponible en este equipo (servicio WinRM detenido, no configurado, o el modulo Microsoft.WSMan.Management no esta cargado)."
            }

            $allowUnencrypted = 'N/D'; $basicAuth = 'N/D'; $kerberosAuth = 'N/D'; $negotiateAuth = 'N/D'; $certAuth = 'N/D'
            try { $allowUnencrypted = (Get-Item 'WSMan:\localhost\Service\AllowUnencrypted' -ErrorAction Stop).Value } catch {}
            try { $basicAuth        = (Get-Item 'WSMan:\localhost\Service\Auth\Basic' -ErrorAction Stop).Value } catch {}
            try { $kerberosAuth     = (Get-Item 'WSMan:\localhost\Service\Auth\Kerberos' -ErrorAction Stop).Value } catch {}
            try { $negotiateAuth    = (Get-Item 'WSMan:\localhost\Service\Auth\Negotiate' -ErrorAction Stop).Value } catch {}
            try { $certAuth         = (Get-Item 'WSMan:\localhost\Service\Auth\Certificate' -ErrorAction Stop).Value } catch {}

            $trustedHosts = 'N/D'
            try {
                $thVal = (Get-Item 'WSMan:\localhost\Client\TrustedHosts' -ErrorAction Stop).Value
                $trustedHosts = if ($thVal) { $thVal } else { '(vacio)' }
            } catch {}

            $listeners = Get-ChildItem WSMan:\localhost\Listener -ErrorAction SilentlyContinue | ForEach-Object {
                $props = Get-ChildItem $_.PSPath -ErrorAction SilentlyContinue
                $transporte = ($props | Where-Object { $_.Name -eq 'Transport' }).Value
                $puerto     = ($props | Where-Object { $_.Name -eq 'Port' }).Value
                $direccion  = ($props | Where-Object { $_.Name -eq 'Address' }).Value
                "$transporte $direccion`:$puerto"
            }
            $listenersStr = if ($listeners) { $listeners -join ' | ' } else { 'Sin listeners configurados' }

            $execPolicies = Get-ExecutionPolicy -List -ErrorAction SilentlyContinue
            $execPolStr = if ($execPolicies) { ($execPolicies | ForEach-Object { "$($_.Scope)=$($_.ExecutionPolicy)" }) -join ' | ' } else { 'N/D' }

            [PSCustomObject]@{
                AllowUnencrypted       = $allowUnencrypted
                AutenticacionBasic     = $basicAuth
                AutenticacionKerberos  = $kerberosAuth
                AutenticacionNegotiate = $negotiateAuth
                AutenticacionCert      = $certAuth
                TrustedHosts           = $trustedHosts
                Listeners              = $listenersStr
                PoliticasEjecucion     = $execPolStr
            }
        } catch {
            "No se pudo obtener la configuracion de WinRM via el proveedor WSMan: (servicio detenido, sin permisos suficientes, o WinRM no configurado en este equipo): $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.16.8' -Title 'Configuracion de WinRM y PowerShell Remoting' -Data $data -Note 'Servicio WinRM (autenticacion, cifrado, TrustedHosts), listeners activos y politica de ejecucion de PowerShell por ambito. Se consulta via el proveedor WSMan: (no invoca winrm.cmd, que puede quedar esperando entrada interactiva).' -Wide
}

} # fin Test-SectionEnabled '1.16'
#endregion

#region 1.17 SALUD, EVENTOS Y RENDIMIENTO
if (Test-SectionEnabled '1.17') {

# -----------------------------------------------------------------------------
# 1.17.1 Uptime y arranque
# -----------------------------------------------------------------------------
Measure-Section '1.17.1 Uptime y arranque' {
    $data = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        try {
            $os = Get-CimInstance Win32_OperatingSystem
            $ultimoArranque = $os.LastBootUpTime
            $dias = [math]::Round(((Get-Date) - $ultimoArranque).TotalDays, 2)

            $fastStartup = 'N/D'
            try {
                $val = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' -Name 'HiberbootEnabled' -ErrorAction SilentlyContinue
                $fastStartup = if ($null -ne $val.HiberbootEnabled) { [bool]$val.HiberbootEnabled } else { 'No configurado' }
            } catch { $fastStartup = 'No se pudo determinar' }

            $pageFiles = Get-CimInstance Win32_PageFileUsage -ErrorAction SilentlyContinue
            $pageFileStr = if ($pageFiles) {
                ($pageFiles | ForEach-Object { "$($_.Name): AsignadoMB=$($_.AllocatedBaseSize) UsoActualMB=$($_.CurrentUsage) PicoMB=$($_.PeakUsage)" }) -join ' | '
            } else { 'No hay archivos de paginacion configurados (o son gestionados en RAM).' }

            $autoManaged = 'N/D'
            try { $autoManaged = [bool](Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).AutomaticManagedPagefile } catch {}

            [PSCustomObject]@{
                UltimoArranque        = $ultimoArranque
                DiasEncendido         = $dias
                FastStartupHabilitado = $fastStartup
                PageFileGestionAuto   = $autoManaged
                ArchivosPaginacion    = $pageFileStr
            }
        } catch {
            "No se pudo obtener el estado de arranque/paginacion: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.17.1' -Title 'Uptime y Arranque' -Data $data -Note 'Tiempo desde el ultimo arranque, estado de Fast Startup y configuracion del archivo de paginacion.'
}

# -----------------------------------------------------------------------------
# 1.17.2 Apagados inesperados y reinicios (ultimos 90 dias)
# -----------------------------------------------------------------------------
Measure-Section '1.17.2 Apagados inesperados y reinicios (ultimos 90 dias)' {
    $data = Invoke-Remote -TimeoutSec 90 -ScriptBlock {
        try {
            $desde = (Get-Date).AddDays(-90)
            $filtro = @{ LogName = 'System'; Id = 6008, 1074, 41, 6013; StartTime = $desde }
            $eventos = Get-WinEvent -FilterHashtable $filtro -MaxEvents 500 -ErrorAction Stop
            $eventos | Sort-Object TimeCreated -Descending | ForEach-Object {
                $tipo = switch ($_.Id) {
                    6008 { 'Apagado inesperado' }
                    1074 { 'Apagado/reinicio iniciado por usuario o proceso' }
                    41   { 'Kernel-Power (reinicio sin apagado limpio)' }
                    6013 { 'Reporte de uptime del sistema' }
                    default { 'Otro' }
                }
                $msg = $_.Message
                if ($msg -and $msg.Length -gt 300) { $msg = $msg.Substring(0, 300) + '...' }
                [PSCustomObject]@{
                    Fecha    = $_.TimeCreated
                    EventoID = $_.Id
                    Tipo     = $tipo
                    Origen   = $_.ProviderName
                    Mensaje  = $msg
                }
            }
        } catch {
            "No se encontraron eventos de apagado/reinicio en los ultimos 90 dias, o el log System esta vacio/inaccesible: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.17.2' -Title 'Apagados Inesperados y Reinicios (90 dias)' -Data $data -Note 'Event IDs 6008 (apagado inesperado), 1074 (apagado/reinicio solicitado), 41 (Kernel-Power) y 6013 (uptime).' -Wide
}

# -----------------------------------------------------------------------------
# 1.17.3 Errores criticos del Event Log
# -----------------------------------------------------------------------------
Measure-Section '1.17.3 Errores criticos del Event Log' {
    $dias = 7
    try { if ($EventLogDays -and $EventLogDays -gt 0) { $dias = $EventLogDays } } catch {}

    $data = Invoke-Remote -TimeoutSec 120 -ArgumentList $dias -ScriptBlock {
        param($diasAtras)
        try {
            $desde = (Get-Date).AddDays(-1 * $diasAtras)
            $filtro = @{ LogName = 'System', 'Application'; Level = 1, 2; StartTime = $desde }
            $eventos = Get-WinEvent -FilterHashtable $filtro -MaxEvents 5000 -ErrorAction Stop

            $grupos = $eventos | Group-Object LogName, Id, ProviderName
            $filas = foreach ($g in $grupos) {
                $ordenados = $g.Group | Sort-Object TimeCreated
                $primero = $ordenados | Select-Object -First 1
                $ultimo  = $ordenados | Select-Object -Last 1
                $msg = $ultimo.Message
                if ($msg -and $msg.Length -gt 300) { $msg = $msg.Substring(0, 300) + '...' }
                elseif (-not $msg) { $msg = '(sin mensaje / proveedor no resoluble localmente)' }
                [PSCustomObject]@{
                    LogName           = $primero.LogName
                    EventoID          = $primero.Id
                    Proveedor         = $primero.ProviderName
                    Nivel             = $primero.LevelDisplayName
                    Cantidad          = $g.Count
                    PrimeraOcurrencia = $primero.TimeCreated
                    UltimaOcurrencia  = $ultimo.TimeCreated
                    MensajeEjemplo    = $msg
                }
            }
            $filas | Sort-Object Cantidad -Descending | Select-Object -First 25
        } catch {
            "No se encontraron eventos de Error/Critico en System o Application en los ultimos $diasAtras dias, o los logs estan vacios: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.17.3' -Title 'Errores Criticos del Event Log' -Data $data -Note "Eventos de nivel Error/Critico en System y Application de los ultimos $dias dias, agrupados por LogName+EventoID+Proveedor. Top 25 por cantidad de ocurrencias." -Wide
}

# -----------------------------------------------------------------------------
# 1.17.4 Baseline de rendimiento (opt-in via -PerfSampleSeconds)
# -----------------------------------------------------------------------------
if ($PerfSampleSeconds -and $PerfSampleSeconds -gt 0) {
Measure-Section '1.17.4 Baseline de rendimiento' {
    $data = Invoke-Remote -TimeoutSec ($PerfSampleSeconds + 60) -ArgumentList $PerfSampleSeconds -ScriptBlock {
        param($segundos)
        $contadoresIngles = @(
            '\Processor(_Total)\% Processor Time',
            '\Memory\Available MBytes',
            '\PhysicalDisk(_Total)\Avg. Disk Queue Length',
            '\System\Processor Queue Length'
        )

        function Get-MuestrasContador {
            param([string[]]$Rutas, [int]$Segundos)
            $muestras = Get-Counter -Counter $Rutas -SampleInterval 1 -MaxSamples $Segundos -ErrorAction Stop
            $porContador = @{}
            foreach ($m in $muestras) {
                foreach ($cs in $m.CounterSamples) {
                    if (-not $porContador.ContainsKey($cs.Path)) { $porContador[$cs.Path] = New-Object System.Collections.Generic.List[double] }
                    $porContador[$cs.Path].Add($cs.CookedValue)
                }
            }
            $porContador
        }

        try {
            $muestrasSampleadas = Get-MuestrasContador -Rutas $contadoresIngles -Segundos $segundos
            foreach ($ruta in $muestrasSampleadas.Keys) {
                $vals = $muestrasSampleadas[$ruta]
                [PSCustomObject]@{
                    Contador = $ruta
                    Promedio = [math]::Round((($vals | Measure-Object -Average).Average), 2)
                    Maximo   = [math]::Round((($vals | Measure-Object -Maximum).Maximum), 2)
                    Minimo   = [math]::Round((($vals | Measure-Object -Minimum).Minimum), 2)
                    Muestras = $vals.Count
                }
            }
        } catch {
            # Fallback por localizacion: SO no ingles. Se traducen los nombres via el
            # indice numerico compartido entre Perflib\009 (ingles) y CurrentLanguage.
            try {
                $engPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\009'
                $locPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\CurrentLanguage'
                $engCounters = (Get-ItemProperty -Path $engPath -Name Counter -ErrorAction Stop).Counter
                $locCounters = (Get-ItemProperty -Path $locPath -Name Counter -ErrorAction Stop).Counter

                $engIndexByName = @{}
                for ($i = 0; $i -lt $engCounters.Count - 1; $i += 2) { $engIndexByName[$engCounters[$i + 1]] = $engCounters[$i] }
                $locNameByIndex = @{}
                for ($i = 0; $i -lt $locCounters.Count - 1; $i += 2) { $locNameByIndex[$locCounters[$i]] = $locCounters[$i + 1] }

                function Get-NombreLocalizado {
                    param([string]$NombreIngles)
                    if ($engIndexByName.ContainsKey($NombreIngles)) {
                        $idx = $engIndexByName[$NombreIngles]
                        if ($locNameByIndex.ContainsKey($idx)) { return $locNameByIndex[$idx] }
                    }
                    return $null
                }

                $objProcesador = Get-NombreLocalizado 'Processor'
                $ctrPctCpu     = Get-NombreLocalizado '% Processor Time'
                $objMemoria    = Get-NombreLocalizado 'Memory'
                $ctrMemDisp    = Get-NombreLocalizado 'Available MBytes'
                $objDiscoFis   = Get-NombreLocalizado 'PhysicalDisk'
                $ctrColaDisco  = Get-NombreLocalizado 'Avg. Disk Queue Length'
                $objSistema    = Get-NombreLocalizado 'System'
                $ctrColaCpu    = Get-NombreLocalizado 'Processor Queue Length'

                if (-not ($objProcesador -and $ctrPctCpu -and $objMemoria -and $ctrMemDisp -and $objDiscoFis -and $ctrColaDisco -and $objSistema -and $ctrColaCpu)) {
                    throw "No se pudieron traducir todos los nombres de contador via Perflib."
                }

                $contadoresLocalizados = @(
                    "\$objProcesador(_Total)\$ctrPctCpu",
                    "\$objMemoria\$ctrMemDisp",
                    "\$objDiscoFis(_Total)\$ctrColaDisco",
                    "\$objSistema\$ctrColaCpu"
                )
                $muestrasSampleadas = Get-MuestrasContador -Rutas $contadoresLocalizados -Segundos $segundos
                foreach ($ruta in $muestrasSampleadas.Keys) {
                    $vals = $muestrasSampleadas[$ruta]
                    [PSCustomObject]@{
                        Contador = $ruta
                        Promedio = [math]::Round((($vals | Measure-Object -Average).Average), 2)
                        Maximo   = [math]::Round((($vals | Measure-Object -Maximum).Maximum), 2)
                        Minimo   = [math]::Round((($vals | Measure-Object -Minimum).Minimum), 2)
                        Muestras = $vals.Count
                    }
                }
            } catch {
                "No se pudo tomar el baseline de rendimiento: el sistema operativo esta en un idioma distinto al ingles y no se pudieron traducir los nombres de contador ($($_.Exception.Message))."
            }
        }
    }
    Add-ReportSection -Id '1.17.4' -Title 'Baseline de Rendimiento' -Data $data -Note "Muestreo de $PerfSampleSeconds segundos (1 muestra/seg): CPU total, memoria disponible, cola de disco fisico y cola de procesador." -Wide
}
}

# -----------------------------------------------------------------------------
# 1.17.5 Top procesos por memoria y CPU
# -----------------------------------------------------------------------------
Measure-Section '1.17.5 Top procesos por memoria y CPU' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            # El usuario propietario se resuelve con Get-Process -IncludeUserName (sin
            # tocar CIM/WMI) cuando el proceso corre elevado; si eso falla (requiere
            # privilegios de administrador), se recurre a UNA sola consulta CIM con
            # todos los PID de interes en un filtro OR, en lugar de una consulta
            # Win32_Process por proceso (15 llamadas WMI separadas, el costo real de
            # esta seccion).
            $usaIncludeUserName = $true
            try {
                $procesos = Get-Process -IncludeUserName -ErrorAction Stop | Sort-Object WorkingSet -Descending | Select-Object -First 15
            } catch {
                $usaIncludeUserName = $false
                $procesos = Get-Process -ErrorAction Stop | Sort-Object WorkingSet -Descending | Select-Object -First 15
            }

            $cimPorPid = @{}
            if (-not $usaIncludeUserName -and $procesos) {
                $filtroWql = ($procesos | ForEach-Object { "ProcessId=$($_.Id)" }) -join ' OR '
                if ($filtroWql) {
                    Get-CimInstance Win32_Process -Filter $filtroWql -ErrorAction SilentlyContinue | ForEach-Object { $cimPorPid[[int]$_.ProcessId] = $_ }
                }
            }

            $procesos | ForEach-Object {
                $proc = $_
                $usuario = 'N/D'
                if ($usaIncludeUserName) {
                    $usuario = if ($proc.UserName) { $proc.UserName } else { 'No se pudo determinar (proceso de sistema o sin permisos)' }
                } else {
                    try {
                        $cimProc = $cimPorPid[[int]$proc.Id]
                        if ($cimProc) {
                            $ownerInfo = Invoke-CimMethod -InputObject $cimProc -MethodName GetOwner -ErrorAction SilentlyContinue
                            if ($ownerInfo -and $ownerInfo.ReturnValue -eq 0) {
                                $usuario = if ($ownerInfo.Domain) { "$($ownerInfo.Domain)\$($ownerInfo.User)" } else { $ownerInfo.User }
                            }
                        }
                    } catch { $usuario = 'No se pudo determinar (proceso de sistema o sin permisos)' }
                }

                $ruta = 'No accesible (permisos)'
                try { if ($proc.Path) { $ruta = $proc.Path } } catch {}

                $cpuSeg = 'N/D'
                try { if ($null -ne $proc.CPU) { $cpuSeg = [math]::Round($proc.CPU, 2) } } catch {}

                $inicio = 'No accesible'
                try { if ($proc.StartTime) { $inicio = $proc.StartTime } } catch {}

                [PSCustomObject]@{
                    Nombre         = $proc.ProcessName
                    PID            = $proc.Id
                    MemoriaMB      = [math]::Round($proc.WorkingSet / 1MB, 2)
                    CPUSegundos    = $cpuSeg
                    Usuario        = $usuario
                    RutaEjecutable = $ruta
                    Inicio         = $inicio
                }
            }
        } catch {
            "No se pudo obtener el listado de procesos: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.17.5' -Title 'Top 15 Procesos por Memoria' -Data $data -Note 'Ordenado por memoria de trabajo (WorkingSet). El usuario propietario y la ruta pueden no resolverse para procesos de sistema por permisos.' -Wide
}

# -----------------------------------------------------------------------------
# 1.17.6 Estado de salud de discos fisicos
# -----------------------------------------------------------------------------
Measure-Section '1.17.6 Estado de salud de discos fisicos' {
    $raw = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        $out = @{}
        try {
            $physDisks = Get-PhysicalDisk -ErrorAction Stop
            $out.Discos = $physDisks | Select-Object FriendlyName, MediaType, OperationalStatus, HealthStatus,
                @{N='DesgastePct';E={ if ($null -ne $_.Wear) { $_.Wear } else { 'N/D' } }},
                @{N='TamanoGB';E={[math]::Round($_.Size/1GB,2)}}
        } catch { $out.Discos = 'Get-PhysicalDisk no disponible (requiere el modulo Storage, tipicamente Windows Server 2012+).' }

        try {
            $pools = Get-StoragePool -ErrorAction Stop | Where-Object { -not $_.IsPrimordial }
            $out.Pools = if ($pools) { ($pools | ForEach-Object { "$($_.FriendlyName): $($_.HealthStatus)/$($_.OperationalStatus)" }) -join ' | ' } else { 'Sin Storage Pools configurados.' }
        } catch { $out.Pools = 'Get-StoragePool no disponible en este equipo.' }

        try {
            $smart = Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction Stop
            $out.Smart = if ($smart) { ($smart | ForEach-Object { "$($_.InstanceName): PredictFailure=$($_.PredictFailure)" }) -join ' | ' } else { 'Sin datos SMART expuestos por el controlador.' }
        } catch { $out.Smart = 'Namespace SMART (root\wmi) no expuesto por el controlador (comun en discos detras de RAID/SAN/virtuales).' }

        [PSCustomObject]$out
    }
    $notaExtra = "Storage Pools: $($raw.Pools) | SMART: $($raw.Smart)"
    Add-ReportSection -Id '1.17.6' -Title 'Salud de Discos Fisicos' -Data $raw.Discos -Note "Estado reportado por Storage Spaces (HealthStatus/OperationalStatus) y SMART basico. $notaExtra" -Wide
}

} # fin Test-SectionEnabled '1.17'
#endregion

#region COMPLEMENTOS DE RED (1.4.6-1.4.9)
if (Test-SectionEnabled '1.4') {

# -----------------------------------------------------------------------------
# 1.4.6 Puertos TCP/UDP en escucha con proceso dueno
# -----------------------------------------------------------------------------
Measure-Section '1.4.6 Puertos TCP/UDP en escucha con proceso dueno' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $filas = New-Object System.Collections.Generic.List[PSObject]
            $procCache = @{}
            $svcByPid = @{}
            try {
                Get-CimInstance Win32_Service -ErrorAction SilentlyContinue | Where-Object { $_.ProcessId -gt 0 } | ForEach-Object {
                    $svcByPid[$_.ProcessId] = $_.Name
                }
            } catch {}

            $tcpListen = Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue
            foreach ($c in $tcpListen) {
                if (-not $procCache.ContainsKey($c.OwningProcess)) { $procCache[$c.OwningProcess] = Get-Process -Id $c.OwningProcess -ErrorAction SilentlyContinue }
                $p = $procCache[$c.OwningProcess]
                $ruta = 'N/D'
                try { if ($p -and $p.Path) { $ruta = $p.Path } } catch {}
                $svc = if ($svcByPid.ContainsKey($c.OwningProcess)) { $svcByPid[$c.OwningProcess] } else { '' }
                $filas.Add([PSCustomObject]@{
                    Protocolo      = 'TCP'
                    DireccionLocal = $c.LocalAddress
                    PuertoLocal    = $c.LocalPort
                    PID            = $c.OwningProcess
                    Proceso        = if ($p) { $p.ProcessName } else { 'N/D' }
                    RutaProceso    = $ruta
                    Servicio       = $svc
                })
            }

            $udpListen = Get-NetUDPEndpoint -ErrorAction SilentlyContinue
            foreach ($c in $udpListen) {
                if (-not $procCache.ContainsKey($c.OwningProcess)) { $procCache[$c.OwningProcess] = Get-Process -Id $c.OwningProcess -ErrorAction SilentlyContinue }
                $p = $procCache[$c.OwningProcess]
                $ruta = 'N/D'
                try { if ($p -and $p.Path) { $ruta = $p.Path } } catch {}
                $svc = if ($svcByPid.ContainsKey($c.OwningProcess)) { $svcByPid[$c.OwningProcess] } else { '' }
                $filas.Add([PSCustomObject]@{
                    Protocolo      = 'UDP'
                    DireccionLocal = $c.LocalAddress
                    PuertoLocal    = $c.LocalPort
                    PID            = $c.OwningProcess
                    Proceso        = if ($p) { $p.ProcessName } else { 'N/D' }
                    RutaProceso    = $ruta
                    Servicio       = $svc
                })
            }

            $filas | Sort-Object PuertoLocal
        } catch {
            "No se pudo obtener los puertos en escucha: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.4.6' -Title 'Puertos TCP/UDP en Escucha' -Data $data -Note 'Puertos en escucha con el proceso propietario, ruta del ejecutable y, si corresponde, el servicio de Windows asociado.' -Wide
}

# -----------------------------------------------------------------------------
# 1.4.7 Tabla de rutas persistentes y por defecto
# -----------------------------------------------------------------------------
Measure-Section '1.4.7 Tabla de rutas persistentes y por defecto' {
    $data = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        try {
            Get-NetRoute -ErrorAction Stop | Where-Object {
                $_.DestinationPrefix -eq '0.0.0.0/0' -or $_.PolicyStore -eq 'PersistentStore'
            } | Select-Object @{N='Destino';E={($_.DestinationPrefix -split '/')[0]}},
                @{N='MascaraPrefijo';E={($_.DestinationPrefix -split '/')[1]}},
                @{N='SiguienteSalto';E={$_.NextHop}},
                @{N='InterfaceAlias';E={$_.InterfaceAlias}},
                @{N='Metrica';E={$_.RouteMetric}},
                @{N='Persistente';E={$_.PolicyStore -eq 'PersistentStore'}} |
                Sort-Object Destino
        } catch {
            "No se pudo obtener la tabla de rutas: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.4.7' -Title 'Rutas Persistentes y Puerta de Enlace por Defecto' -Data $data -Note 'Incluye la ruta por defecto (0.0.0.0/0) y toda ruta estatica persistente (route -p / New-NetRoute -PolicyStore PersistentStore).' -Wide
}

# -----------------------------------------------------------------------------
# 1.4.8 Archivo hosts
# -----------------------------------------------------------------------------
Measure-Section '1.4.8 Archivo hosts' {
    $data = Invoke-Remote -TimeoutSec 30 -ScriptBlock {
        try {
            $rutaHosts = Join-Path $env:SystemRoot 'System32\drivers\etc\hosts'
            if (-not (Test-Path $rutaHosts)) { return "No se encontro el archivo hosts en $rutaHosts" }
            $lineas = Get-Content -Path $rutaHosts -ErrorAction Stop
            $entradas = $lineas | Where-Object { $_.Trim() -and -not $_.Trim().StartsWith('#') }
            if ($entradas) {
                $entradas | ForEach-Object {
                    $partes = ($_ -split '\s+') | Where-Object { $_ }
                    [PSCustomObject]@{
                        IP            = $partes[0]
                        NombreHost    = if ($partes.Count -gt 1) { $partes[1] } else { '' }
                        LineaCompleta = $_.Trim()
                    }
                }
            } else {
                "El archivo hosts no tiene entradas personalizadas (sin entradas, solo comentarios o esta vacio)."
            }
        } catch {
            "No se pudo leer el archivo hosts: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.4.8' -Title 'Archivo Hosts' -Data $data -Note 'Entradas de resolucion de nombres estaticas en %SystemRoot%\System32\drivers\etc\hosts (se excluyen comentarios y lineas vacias).'
}

# -----------------------------------------------------------------------------
# 1.4.9 Proxy configurado
# -----------------------------------------------------------------------------
Measure-Section '1.4.9 Proxy configurado' {
    $data = Invoke-Remote -TimeoutSec 30 -ScriptBlock {
        try {
            # netsh puede emitir su salida de texto en UTF-8 mientras la consola sigue
            # decodificando con la pagina de codigos OEM, lo que produce caracteres
            # acentuados ilegibles (mojibake). Se fuerza la codificacion de salida de la
            # consola a UTF-8 solo durante esta llamada puntual, y se restaura despues.
            $codigoConsolaPrevio = $null
            try { $codigoConsolaPrevio = [Console]::OutputEncoding } catch {}
            try { if ($codigoConsolaPrevio) { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } } catch {}
            try {
                $winhttpRaw = netsh winhttp show proxy 2>&1 | Out-String
            } finally {
                try { if ($codigoConsolaPrevio) { [Console]::OutputEncoding = $codigoConsolaPrevio } } catch {}
            }

            $userProxyEnable = 'N/D'; $userProxyServer = 'N/D'; $userProxyOverride = 'N/D'
            try {
                $regPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
                $props = Get-ItemProperty -Path $regPath -ErrorAction Stop
                $userProxyEnable = [bool]$props.ProxyEnable
                $userProxyServer = if ($props.ProxyServer) { $props.ProxyServer } else { '(no configurado)' }
                $userProxyOverride = if ($props.ProxyOverride) { $props.ProxyOverride } else { '(no configurado)' }
            } catch { $userProxyEnable = 'No se pudo leer HKCU (puede requerir ejecutarse como el usuario interactivo)' }

            [PSCustomObject]@{
                ProxyWinHTTP            = $winhttpRaw.Trim()
                ProxyUsuarioHabilitado  = $userProxyEnable
                ProxyUsuarioServidor    = $userProxyServer
                ProxyUsuarioExcepciones = $userProxyOverride
            }
        } catch {
            "No se pudo obtener la configuracion de proxy: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.4.9' -Title 'Proxy Configurado' -Data $data -Note 'Proxy WinHTTP a nivel de sistema (usado por Windows Update y servicios) y proxy de WinINet del usuario interactivo bajo el cual corre el script (HKCU puede no reflejar el proxy de otros usuarios).'
}

} # fin Test-SectionEnabled '1.4'
#endregion

#region COMPLEMENTOS DE STORAGE (1.5.3-1.5.4)
if (Test-SectionEnabled '1.5') {

# -----------------------------------------------------------------------------
# 1.5.3 Estado de BitLocker
# -----------------------------------------------------------------------------
Measure-Section '1.5.3 Estado de BitLocker' {
    $data = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        try {
            $vols = Get-BitLockerVolume -ErrorAction Stop
            if (-not $vols) { return "No hay volumenes reportados por BitLocker en este equipo." }
            $vols | ForEach-Object {
                $tipos = if ($_.KeyProtector) { (($_.KeyProtector | ForEach-Object { $_.KeyProtectorType }) -join ', ') } else { 'Ninguno' }
                [PSCustomObject]@{
                    MountPoint           = $_.MountPoint
                    VolumeStatus         = $_.VolumeStatus
                    ProtectionStatus     = $_.ProtectionStatus
                    EncryptionPercentage = $_.EncryptionPercentage
                    EncryptionMethod     = $_.EncryptionMethod
                    KeyProtectorTypes    = $tipos
                }
            }
        } catch {
            "Modulo BitLocker no disponible o caracteristica no instalada en este equipo (Get-BitLockerVolume fallo): $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.5.3' -Title 'Estado de BitLocker' -Data $data -Note 'Estado de cifrado de volumenes con BitLocker Drive Encryption y tipos de protector de clave configurados.' -Wide
}

# -----------------------------------------------------------------------------
# 1.5.4 Deduplicacion y cuotas
# -----------------------------------------------------------------------------
Measure-Section '1.5.4 Deduplicacion y cuotas' {
    $dedupData = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        try {
            Import-Module Deduplication -ErrorAction Stop
            $vols = Get-DedupVolume -ErrorAction Stop
            if ($vols) {
                $vols | Select-Object Volume, Enabled, Capacity,
                    @{N='AhorroGB';E={[math]::Round($_.SavedSpace/1GB,2)}},
                    SavingsPercent, OptimizedFilesCount
            } else { 'No hay volumenes con deduplicacion habilitada.' }
        } catch { 'Rol/caracteristica de Deduplicacion de Datos no instalado en este servidor.' }
    }
    Add-ReportSection -Id '1.5.4' -Title 'Deduplicacion de Datos' -Data $dedupData -Note 'Ahorro de espacio por deduplicacion (Data Deduplication), por volumen.' -Wide

    $fsrmData = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        try {
            Import-Module FileServerResourceManager -ErrorAction Stop
            $cuotas = Get-FsrmQuota -ErrorAction Stop
            if ($cuotas) {
                $cuotas | Select-Object Path, Size, Usage, Description,
                    @{N='PorcentajeUsado';E={ if ($_.Size -gt 0) { [math]::Round(($_.Usage/$_.Size)*100,1) } else { 0 } }}
            } else { 'No hay cuotas FSRM configuradas.' }
        } catch { 'Rol File Server Resource Manager (FSRM) no instalado en este servidor.' }
    }
    Add-ReportSection -Id '1.5.4.1' -Title 'Cuotas FSRM' -Data $fsrmData -Note 'Cuotas de disco configuradas via File Server Resource Manager.' -Wide
}

} # fin Test-SectionEnabled '1.5'
#endregion

#region COMPLEMENTOS DE FILE SERVER (1.7.3)
if (Test-SectionEnabled '1.7') {

# -----------------------------------------------------------------------------
# 1.7.3 Permisos NTFS efectivos de los shares
# -----------------------------------------------------------------------------
Measure-Section '1.7.3 Permisos NTFS efectivos de los shares' {
    $data = Invoke-Remote -TimeoutSec 90 -ScriptBlock {
        try {
            $shares = Get-SmbShare -ErrorAction Stop | Where-Object { $_.Name -notmatch '^[A-Z]\$$|^ADMIN\$$|^IPC\$$|^PRINT\$$' }
            if (-not $shares) { return "No hay recursos compartidos no administrativos en este equipo." }

            $riesgosas = @('Everyone', 'Todos', 'Authenticated Users', 'Usuarios autentificados', 'BUILTIN\Users', 'BUILTIN\Usuarios')
            $permisosEscritura = @('FullControl', 'Modify', 'Write', 'Control total', 'Modificar', 'Escritura')

            $listaShares = @($shares)
            $truncado = $false
            if ($listaShares.Count -gt 10) { $truncado = $true; $listaShares = $listaShares | Select-Object -First 10 }

            $filas = New-Object System.Collections.Generic.List[PSObject]
            foreach ($sh in $listaShares) {
                try {
                    $acl = Get-Acl -Path $sh.Path -ErrorAction Stop
                    foreach ($ace in $acl.Access) {
                        $identidad = $ace.IdentityReference.Value
                        $esRiesgoso = $false
                        foreach ($idRiesgosa in $riesgosas) {
                            if ($identidad -match [regex]::Escape($idRiesgosa)) {
                                foreach ($permEscritura in $permisosEscritura) {
                                    if ($ace.FileSystemRights -match $permEscritura) { $esRiesgoso = $true }
                                }
                            }
                        }
                        $filas.Add([PSCustomObject]@{
                            Recurso    = $sh.Name
                            RutaLocal  = $sh.Path
                            Identidad  = $identidad
                            Permisos   = $ace.FileSystemRights.ToString()
                            TipoAcceso = $ace.AccessControlType.ToString()
                            Heredado   = $ace.IsInherited
                            EsRiesgoso = $esRiesgoso
                        })
                    }
                } catch {
                    $filas.Add([PSCustomObject]@{
                        Recurso = $sh.Name; RutaLocal = $sh.Path; Identidad = 'N/D'
                        Permisos = "No se pudo leer el ACL: $($_.Exception.Message)"
                        TipoAcceso = ''; Heredado = ''; EsRiesgoso = $false
                    })
                }
            }

            if ($truncado) {
                $filas.Add([PSCustomObject]@{
                    Recurso = '(AVISO)'; RutaLocal = ''
                    Identidad = "Se trunco el analisis a los primeros 10 shares de $($shares.Count) totales."
                    Permisos = ''; TipoAcceso = ''; Heredado = ''; EsRiesgoso = $false
                })
            }
            $filas
        } catch {
            "No se pudo obtener los permisos NTFS de los shares: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.7.3' -Title 'Permisos NTFS Efectivos de los Shares' -Data $data -Note 'ACL de la ruta local de cada recurso compartido no administrativo (limitado a los primeros 10 shares). EsRiesgoso=true cuando Everyone/Todos/Authenticated Users/BUILTIN Users tiene permisos de escritura o control total.' -Wide
}

} # fin Test-SectionEnabled '1.7'
#endregion

#region COMPLEMENTOS DE SEGURIDAD (1.11.9-1.11.13)
if (Test-SectionEnabled '1.11') {

# -----------------------------------------------------------------------------
# 1.11.9 Estado de hardening del sistema operativo
# -----------------------------------------------------------------------------
Measure-Section '1.11.9 Estado de hardening del sistema operativo' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            $filas = New-Object System.Collections.Generic.List[PSObject]

            $smb1Estado = 'N/D'
            try {
                $feat = Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -ErrorAction Stop
                # Se fuerza a string: el enum de State puede volver serializado como un
                # objeto {value,Value} al pasar por una sesion remota (WinRM), en vez de
                # como texto plano.
                $smb1Estado = [string]$feat.State
            } catch {
                try {
                    $smbCfg = Get-SmbServerConfiguration -ErrorAction Stop
                    $smb1Estado = if ($smbCfg.EnableSMB1Protocol) { 'Enabled' } else { 'Disabled' }
                } catch { $smb1Estado = 'No se pudo determinar' }
            }
            $filas.Add([PSCustomObject]@{ Parametro='SMBv1 habilitado'; ValorActual=$smb1Estado; ValorRecomendado='Disabled'; Cumple=($smb1Estado -eq 'Disabled') })

            try {
                $smbCfg2 = Get-SmbServerConfiguration -ErrorAction Stop
                $filas.Add([PSCustomObject]@{ Parametro='Firma SMB requerida (servidor)'; ValorActual=$smbCfg2.RequireSecuritySignature; ValorRecomendado='True'; Cumple=($smbCfg2.RequireSecuritySignature -eq $true) })
            } catch {
                $filas.Add([PSCustomObject]@{ Parametro='Firma SMB requerida (servidor)'; ValorActual='No se pudo determinar'; ValorRecomendado='True'; Cumple=$false })
            }

            try {
                $llmnrPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient'
                $llmnrVal = (Get-ItemProperty -Path $llmnrPath -Name EnableMulticast -ErrorAction SilentlyContinue).EnableMulticast
                $llmnrEstado = if ($null -eq $llmnrVal) { 'No configurado (habilitado por default)' } elseif ($llmnrVal -eq 0) { 'Disabled' } else { 'Enabled' }
                $filas.Add([PSCustomObject]@{ Parametro='LLMNR'; ValorActual=$llmnrEstado; ValorRecomendado='Disabled'; Cumple=($llmnrEstado -eq 'Disabled') })
            } catch {
                $filas.Add([PSCustomObject]@{ Parametro='LLMNR'; ValorActual='No se pudo determinar'; ValorRecomendado='Disabled'; Cumple=$false })
            }

            try {
                $nbtCfgs = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction Stop
                $nbtEstado = ($nbtCfgs | ForEach-Object {
                    switch ($_.TcpipNetbiosOptions) { 0 { 'Default (via DHCP)' } 1 { 'Habilitado' } 2 { 'Deshabilitado' } default { 'N/D' } }
                }) -join ', '
                $cumpleNbt = -not ($nbtEstado -match 'Habilitado|Default')
                $filas.Add([PSCustomObject]@{ Parametro='NetBIOS sobre TCP/IP'; ValorActual=$nbtEstado; ValorRecomendado='Deshabilitado'; Cumple=$cumpleNbt })
            } catch {
                $filas.Add([PSCustomObject]@{ Parametro='NetBIOS sobre TCP/IP'; ValorActual='No se pudo determinar'; ValorRecomendado='Deshabilitado'; Cumple=$false })
            }

            try {
                $uacPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
                $uacProps = Get-ItemProperty -Path $uacPath -ErrorAction Stop
                $enableLua = [bool]$uacProps.EnableLUA
                $consentPrompt = $uacProps.ConsentPromptBehaviorAdmin
                $filas.Add([PSCustomObject]@{ Parametro='UAC habilitado (EnableLUA)'; ValorActual=$enableLua; ValorRecomendado='True'; Cumple=($enableLua -eq $true) })
                $filas.Add([PSCustomObject]@{ Parametro='UAC ConsentPromptBehaviorAdmin'; ValorActual=$consentPrompt; ValorRecomendado='2 (Pedir consentimiento en escritorio seguro)'; Cumple=($consentPrompt -eq 2) })
            } catch {
                $filas.Add([PSCustomObject]@{ Parametro='UAC'; ValorActual='No se pudo determinar'; ValorRecomendado='Habilitado'; Cumple=$false })
            }

            try {
                $tsPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'
                $tsWinstationsPath = "$tsPath\WinStations\RDP-Tcp"
                $fDeny = (Get-ItemProperty -Path $tsPath -Name fDenyTSConnections -ErrorAction SilentlyContinue).fDenyTSConnections
                $rdpHabilitado = ($fDeny -eq 0)
                $filas.Add([PSCustomObject]@{ Parametro='RDP habilitado'; ValorActual=$rdpHabilitado; ValorRecomendado='Solo si es necesario'; Cumple='N/A (depende del uso)' })

                if ($rdpHabilitado) {
                    $puertoRdp = (Get-ItemProperty -Path $tsWinstationsPath -Name PortNumber -ErrorAction SilentlyContinue).PortNumber
                    if (-not $puertoRdp) { $puertoRdp = 3389 }
                    $filas.Add([PSCustomObject]@{ Parametro='Puerto RDP'; ValorActual=$puertoRdp; ValorRecomendado='Restringido por firewall (opcionalmente distinto de 3389)'; Cumple='Revisar manualmente' })

                    $nla = (Get-ItemProperty -Path $tsWinstationsPath -Name UserAuthentication -ErrorAction SilentlyContinue).UserAuthentication
                    $filas.Add([PSCustomObject]@{ Parametro='RDP - Network Level Authentication'; ValorActual=$nla; ValorRecomendado='1 (Habilitado)'; Cumple=($nla -eq 1) })

                    $cifradoRdp = (Get-ItemProperty -Path $tsWinstationsPath -Name MinEncryptionLevel -ErrorAction SilentlyContinue).MinEncryptionLevel
                    $filas.Add([PSCustomObject]@{ Parametro='RDP - Nivel minimo de cifrado'; ValorActual=$cifradoRdp; ValorRecomendado='3 (Alto) o superior'; Cumple=($cifradoRdp -ge 3) })
                }
            } catch {
                $filas.Add([PSCustomObject]@{ Parametro='RDP'; ValorActual='No se pudo determinar'; ValorRecomendado='N/A'; Cumple=$false })
            }

            $filas
        } catch {
            "No se pudo obtener el estado de hardening del sistema operativo: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.11.9' -Title 'Estado de Hardening del Sistema Operativo' -Data $data -Note 'Comparacion contra valores de referencia orientativos (tipo CIS Benchmark). Verificar contra la politica de seguridad interna vigente.' -Wide
}

# -----------------------------------------------------------------------------
# 1.11.10 TPM y Secure Boot
# -----------------------------------------------------------------------------
Measure-Section '1.11.10 TPM y Secure Boot' {
    $data = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        $resultado = [ordered]@{}

        try {
            $tpm = Get-Tpm -ErrorAction Stop
            $resultado.TPM_Presente          = $tpm.TpmPresent
            $resultado.TPM_Listo             = $tpm.TpmReady
            $resultado.TPM_Habilitado        = $tpm.TpmEnabled
            # Algunos fabricantes de TPM devuelven ManufacturerVersion con caracteres
            # nulos de relleno al final (cadena de longitud fija a nivel de firmware);
            # se recortan para que no queden bytes nulos incrustados en el reporte.
            $resultado.TPM_VersionFabricante = if ($tpm.ManufacturerVersion) { $tpm.ManufacturerVersion.TrimEnd([char]0) } else { $tpm.ManufacturerVersion }
        } catch {
            $resultado.TPM_Presente = 'Cmdlet Get-Tpm no disponible en este equipo (modulo TrustedPlatformModule ausente o sin TPM).'
        }

        try {
            $sb = Confirm-SecureBootUEFI -ErrorAction Stop
            $resultado.SecureBootHabilitado = $sb
        } catch {
            $resultado.SecureBootHabilitado = "No se pudo determinar (equipo con BIOS legacy / no UEFI, o sin permisos): $($_.Exception.Message)"
        }

        try {
            $dg = Get-CimInstance -Namespace 'root\Microsoft\Windows\DeviceGuard' -ClassName Win32_DeviceGuard -ErrorAction Stop
            if ($dg) {
                $vbsEstado = switch ($dg.VirtualizationBasedSecurityStatus) { 0 { 'Deshabilitado' } 1 { 'Habilitado sin bloqueo' } 2 { 'Habilitado y en ejecucion' } default { 'N/D' } }
                $resultado.VBS_Estado = $vbsEstado
                $resultado.CredentialGuard_Configurado = if ($dg.SecurityServicesConfigured -contains 1) { 'Si' } else { 'No' }
                $resultado.CredentialGuard_EnEjecucion  = if ($dg.SecurityServicesRunning -contains 1) { 'Si' } else { 'No' }
                $resultado.HVCI_Configurado = if ($dg.SecurityServicesConfigured -contains 2) { 'Si' } else { 'No' }
                $resultado.HVCI_EnEjecucion = if ($dg.SecurityServicesRunning -contains 2) { 'Si' } else { 'No' }
            }
        } catch {
            $resultado.DeviceGuard = 'No se pudo consultar Win32_DeviceGuard (requiere Windows 10/Server 2016+ con soporte de Device Guard/Credential Guard).'
        }

        [PSCustomObject]$resultado
    }
    Add-ReportSection -Id '1.11.10' -Title 'TPM y Secure Boot' -Data $data -Note 'Estado del modulo TPM, Secure Boot (UEFI) y capas de Device Guard / Credential Guard / HVCI (Virtualization Based Security).'
}

# -----------------------------------------------------------------------------
# 1.11.11 LAPS
# -----------------------------------------------------------------------------
Measure-Section '1.11.11 LAPS' {
    $data = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        try {
            $resultado = [ordered]@{}
            $legacyDll = Test-Path "$env:SystemRoot\System32\AdmPwd.dll"
            $legacyPolicyPath = 'HKLM:\Software\Policies\Microsoft Services\AdmPwd'
            $legacyPolicy = Test-Path $legacyPolicyPath

            $modernPolicyPath = 'HKLM:\Software\Microsoft\Policies\LAPS'
            $modernPolicy = Test-Path $modernPolicyPath
            $modernServicio = Get-Service -Name LAPS -ErrorAction SilentlyContinue
            $modernModulo = Get-Module -ListAvailable -Name LAPS -ErrorAction SilentlyContinue

            if ($modernPolicy -or $modernServicio -or $modernModulo) {
                $resultado.LAPSDetectado = 'Windows LAPS (moderno, integrado en el SO)'
                $props = if ($modernPolicy) { Get-ItemProperty -Path $modernPolicyPath -ErrorAction SilentlyContinue } else { $null }
                $resultado.Habilitado = if ($props -and $null -ne $props.BackupDirectory) { "Si (BackupDirectory=$($props.BackupDirectory))" } else { 'Politica no configurada via GPO local visible (puede estar aplicada por GPO de dominio no reflejada en este equipo, o el rol no esta configurado).' }
                $resultado.PoliticaComplejidad = if ($props) { $props.PasswordComplexity } else { 'N/D' }
                $resultado.EdadMaximaPassword  = if ($props) { $props.PasswordAgeDays } else { 'N/D' }
            }
            elseif ($legacyDll -or $legacyPolicy) {
                $resultado.LAPSDetectado = 'LAPS Legacy (Microsoft LAPS clasico, AdmPwd.dll)'
                $resultado.AdmPwdDllPresente = $legacyDll
                if ($legacyPolicy) {
                    $props = Get-ItemProperty -Path $legacyPolicyPath -ErrorAction SilentlyContinue
                    $resultado.Habilitado = [bool]$props.AdmPwdEnabled
                    $resultado.PoliticaComplejidad = $props.PasswordComplexity
                    $resultado.EdadMaximaPassword = $props.PasswordAgeDays
                } else {
                    $resultado.Habilitado = 'DLL presente pero sin politica GPO visible localmente.'
                }
            }
            else {
                $resultado.LAPSDetectado = 'No se detecto LAPS (ni legacy ni Windows LAPS) en este equipo.'
            }

            [PSCustomObject]$resultado
        } catch {
            "No se pudo determinar el estado de LAPS: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.11.11' -Title 'LAPS (Local Administrator Password Solution)' -Data $data -Note 'Detecta LAPS Legacy (AdmPwd.dll) y Windows LAPS moderno, y su politica de complejidad/edad de contrasena si es visible localmente.'
}

# -----------------------------------------------------------------------------
# 1.11.12 Activacion y licenciamiento de Windows
# -----------------------------------------------------------------------------
Measure-Section '1.11.12 Activacion y licenciamiento de Windows' {
    $data = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        try {
            $productos = Get-CimInstance SoftwareLicensingProduct -Filter "PartialProductKey IS NOT NULL" -ErrorAction Stop
            if (-not $productos) { return "No se encontraron productos con clave parcial (SoftwareLicensingProduct)." }

            $mapaEstado = @{
                0 = 'Sin licencia'; 1 = 'Licenciado'; 2 = 'Periodo de gracia inicial'
                3 = 'Periodo de gracia adicional'; 4 = 'Gracia por reinstalacion'
                5 = 'Notificacion'; 6 = 'Gracia extendida'
            }

            $productos | ForEach-Object {
                $canal = 'N/D'
                if ($_.ProductKeyChannel -match 'MAK') { $canal = 'MAK' }
                elseif ($_.ProductKeyChannel -match 'KMS') { $canal = 'KMS' }
                elseif ($_.ProductKeyChannel) { $canal = $_.ProductKeyChannel }

                $diasRestantes = 'N/A'
                if ($_.GracePeriodRemaining) { $diasRestantes = [math]::Round($_.GracePeriodRemaining / 1440, 1) }

                [PSCustomObject]@{
                    Nombre          = $_.Name
                    Descripcion     = $_.Description
                    EstadoLicencia  = if ($mapaEstado.ContainsKey([int]$_.LicenseStatus)) { $mapaEstado[[int]$_.LicenseStatus] } else { "Codigo desconocido ($($_.LicenseStatus))" }
                    CanalActivacion = $canal
                    ServidorKMS     = if ($_.KeyManagementServiceMachine) { $_.KeyManagementServiceMachine } else { 'N/A' }
                    DiasRestantes   = $diasRestantes
                }
            }
        } catch {
            "No se pudo obtener el estado de activacion de Windows: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.11.12' -Title 'Activacion y Licenciamiento de Windows' -Data $data -Note 'Estado de licenciamiento (SoftwareLicensingProduct). DiasRestantes solo aplica durante un periodo de gracia.' -Wide
}

# -----------------------------------------------------------------------------
# 1.11.13 Almacenes de certificados adicionales
# -----------------------------------------------------------------------------
Measure-Section '1.11.13 Almacenes de certificados adicionales' {
    $data = Invoke-Remote -TimeoutSec 45 -ScriptBlock {
        try {
            $emisoresConfiables = @(
                'Microsoft', 'DigiCert', 'VeriSign', 'GlobalSign', 'Sectigo', 'Comodo', 'Entrust',
                'Baltimore', 'Thawte', 'GeoTrust', 'Go Daddy', 'GoDaddy', 'Amazon', "Let's Encrypt",
                'ISRG', 'USERTrust', 'AddTrust', 'Certum', 'QuoVadis', 'SwissSign', 'Starfield', 'Symantec'
            )

            function Test-EsEmisorConfiable {
                param([string]$Emisor)
                foreach ($e in $emisoresConfiables) { if ($Emisor -match [regex]::Escape($e)) { return $true } }
                return $false
            }

            $filas = New-Object System.Collections.Generic.List[PSObject]

            $rootCerts = Get-ChildItem -Path 'Cert:\LocalMachine\Root' -ErrorAction SilentlyContinue | Where-Object { $_.Issuer -notmatch 'Microsoft' }
            foreach ($c in $rootCerts) {
                $filas.Add([PSCustomObject]@{
                    Almacen         = 'LocalMachine\Root'
                    Sujeto          = $c.Subject
                    Emisor          = $c.Issuer
                    HuellaDigital   = $c.Thumbprint
                    FechaExpiracion = $c.NotAfter
                    Sospechoso      = (-not (Test-EsEmisorConfiable $c.Issuer))
                })
            }

            $webHostingCerts = Get-ChildItem -Path 'Cert:\LocalMachine\WebHosting' -ErrorAction SilentlyContinue
            foreach ($c in $webHostingCerts) {
                $filas.Add([PSCustomObject]@{
                    Almacen         = 'LocalMachine\WebHosting'
                    Sujeto          = $c.Subject
                    Emisor          = $c.Issuer
                    HuellaDigital   = $c.Thumbprint
                    FechaExpiracion = $c.NotAfter
                    Sospechoso      = (-not (Test-EsEmisorConfiable $c.Issuer))
                })
            }

            if ($filas.Count -gt 0) { $filas } else { "No hay CAs de terceros en Cert:\LocalMachine\Root (fuera de las de Microsoft) ni certificados en Cert:\LocalMachine\WebHosting." }
        } catch {
            "No se pudo enumerar los almacenes de certificados adicionales: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.11.13' -Title 'Almacenes de Certificados Adicionales (Root de Terceros / WebHosting)' -Data $data -Note 'CAs raiz no-Microsoft instaladas (posible indicio de MITM/proxy TLS corporativo o malware) y certificados del almacen WebHosting (usado por IIS con Centralized Certificate Store). Sospechoso=true cuando el emisor no esta en la lista blanca conocida.' -Wide
}

} # fin Test-SectionEnabled '1.11'
#endregion

}
#endregion MAIN-01 - Envoltorio de colectores

#region MAIN-02 - Utilidades locales de orquestacion
function Format-DurationTexto {
    <# Da formato corto y legible a un TimeSpan: '45s', '2m 14s', '1h 05m'. #>
    param([Parameter(Mandatory)][timespan]$TimeSpan)
    try {
        $totalHoras = [math]::Floor($TimeSpan.TotalHours)
        if ($totalHoras -ge 1) {
            return ('{0}h {1:D2}m' -f $totalHoras, $TimeSpan.Minutes)
        }
        if ($TimeSpan.Minutes -ge 1) {
            return ('{0}m {1:D2}s' -f $TimeSpan.Minutes, $TimeSpan.Seconds)
        }
        $segundos = [math]::Max($TimeSpan.Seconds, 0)
        return "${segundos}s"
    } catch {
        return 'N/D'
    }
}

function Reset-PerTargetState {
    <#
        Reinicia todo el estado acumulado por-equipo antes de recolectar.
        CRITICO en modo flota: sin este reset, si alguna vez se invoca
        Invoke-SingleTargetReport mas de una vez dentro del mismo proceso (por
        ejemplo, en una sesion interactiva que recorre varios equipos a mano),
        el segundo equipo heredaria las secciones, hallazgos y tiempos del
        primero. En el modo flota real de este script cada equipo corre en un
        proceso hijo separado (ver Invoke-FleetReport / FleetWorker en el
        fragmento 04), pero este reset se hace igual por prolijidad y para
        soportar uso equipo-por-equipo dentro del mismo proceso.
    #>
    $script:ReportSections   = New-Object System.Collections.Generic.List[object]
    $script:Findings         = New-Object System.Collections.Generic.List[object]
    $script:Timings          = New-Object System.Collections.Generic.List[object]
    $script:LogLines         = New-Object System.Collections.Generic.List[string]
    $script:Caps             = @{}
    $script:FindingsSummary  = $null
    $script:ComplianceChecks = @{}
    # SessionFailed queda de una corrida anterior si no se resetea: sin esto,
    # un equipo A inalcanzable dejaria $script:SessionFailed=$true "pegado"
    # para el equipo B en un uso equipo-por-equipo dentro del mismo proceso.
    $script:SessionFailed    = $false
}

function Get-HealthBarTexto {
    <# Barra ASCII de progreso para el puntaje de salud (0-100). #>
    param([int]$Puntaje, [int]$Ancho = 20)
    try {
        $p = $Puntaje
        if ($p -lt 0) { $p = 0 }
        if ($p -gt 100) { $p = 100 }
        $llenos = [int][math]::Round(($p / 100.0) * $Ancho)
        if ($llenos -lt 0) { $llenos = 0 }
        if ($llenos -gt $Ancho) { $llenos = $Ancho }
        $vacios = $Ancho - $llenos
        return ('[' + ('#' * $llenos) + ('-' * $vacios) + "] $p/100")
    } catch {
        return "$Puntaje/100"
    }
}

function Get-VeredictoSalud {
    <# Traduce el puntaje de salud a un veredicto textual en espanol. #>
    param([int]$Puntaje)
    if ($Puntaje -lt 50) { return 'Requiere atencion inmediata' }
    if ($Puntaje -lt 80) { return 'Con observaciones' }
    return 'Saludable'
}
#endregion MAIN-02 - Utilidades locales de orquestacion

#region MAIN-03 - Metadatos para el renderer (Get-ReportMeta)
function Get-ReportMeta {
    <#
        Arma el hashtable $Meta que espera ConvertTo-ReportHtml (fragmento
        05): HostName, GeneradoPor, EquipoGenerador, FechaGeneracion,
        DuracionTotal, FabricanteModelo, UptimeDias, RamTotalGB,
        DiscoMasOcupadoPct, DiscoMasOcupadoNombre, ParchesPendientes,
        UltimoParche.

        Totalmente defensivo: cada extraccion individual va en su propio
        try/catch y cae a $null si la seccion no existe, vino como string
        (mensaje de error o "no instalado") o le falta la propiedad esperada.
        El renderer ya sabe mostrar "N/D" para los valores $null.

        NOTA IMPORTANTE PARA QUIEN ENSAMBLE EL SCRIPT: al escribir este
        fragmento, frag_02_collectors_base.ps1 (que deberia contener los
        colectores de 1.1 Hardware, 1.2.2 Hotfixes, 1.2.3 Actualizaciones
        Faltantes y 1.5.2 Volumenes) todavia no existia en el repositorio, asi
        que los nombres de propiedad EXACTOS de esas secciones no se pudieron
        verificar contra el colector real. Se usa el mismo patron defensivo
        que Get-Prop en el fragmento 04 (varios nombres candidatos por valor,
        en el mismo estilo que las reglas de Storage/Parches de ese
        fragmento) para tolerar variaciones de nombre. Verificar esta funcion
        contra frag_02_collectors_base.ps1 en cuanto exista.
    #>
    param([datetime]$StartTime)

    $meta = @{}

    try { $meta.HostName = $script:TargetName } catch { $meta.HostName = $null }
    try { $meta.GeneradoPor = (Get-UsuarioActual) } catch { $meta.GeneradoPor = $null }
    try { $meta.EquipoGenerador = (Get-EquipoActual) } catch { $meta.EquipoGenerador = $null }
    try { $meta.FechaGeneracion = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss') } catch { $meta.FechaGeneracion = $null }

    try {
        if ($StartTime) {
            $meta.DuracionTotal = Format-DurationTexto -TimeSpan ((Get-Date) - $StartTime)
        } else {
            $meta.DuracionTotal = $null
        }
    } catch { $meta.DuracionTotal = $null }

    # --- 1.1 Hardware: Fabricante + Modelo, RAM total ----------------------
    try {
        $hw = Get-SectionData -Id '1.1'
        if ($hw -is [array]) { $hw = $hw | Select-Object -First 1 }
        if ($hw -and -not ($hw -is [string])) {
            $fabricante = Get-Prop $hw @('Fabricante', 'Manufacturer')
            $modelo     = Get-Prop $hw @('Modelo', 'Model')
            if ($fabricante -or $modelo) {
                $meta.FabricanteModelo = ("$fabricante $modelo").Trim()
            } else {
                $meta.FabricanteModelo = $null
            }
        } else {
            $meta.FabricanteModelo = $null
        }
    } catch { $meta.FabricanteModelo = $null }

    try {
        $hw = Get-SectionData -Id '1.1'
        if ($hw -is [array]) { $hw = $hw | Select-Object -First 1 }
        if ($hw -and -not ($hw -is [string])) {
            $meta.RamTotalGB = Get-Prop $hw @('RamTotalGB', 'RAMTotalGB', 'MemoriaTotalGB', 'MemoriaRamGB', 'TotalRamGB')
        } else {
            $meta.RamTotalGB = $null
        }
    } catch { $meta.RamTotalGB = $null }

    # --- 1.17.1 Uptime y Arranque: dias encendido --------------------------
    try {
        $up = Get-SectionData -Id '1.17.1'
        if ($up -is [array]) { $up = $up | Select-Object -First 1 }
        $dias = $null
        if ($up -and -not ($up -is [string])) {
            $dias = Get-Prop $up @('DiasEncendido', 'UptimeDias', 'DiasUptime')
            if ($null -eq $dias) {
                $arranque = ConvertTo-DateSafe (Get-Prop $up @('UltimoArranque', 'LastBootUpTime'))
                if ($arranque) { $dias = [math]::Round(((Get-Date) - $arranque).TotalDays, 1) }
            }
        }
        $meta.UptimeDias = $dias
    } catch { $meta.UptimeDias = $null }

    # --- 1.5.2 Volumenes: el volumen mas ocupado (100 - LibrePct) ----------
    try {
        $vols = ConvertTo-RowArray (Get-SectionData -Id '1.5.2')
        $peorLibrePct = $null
        $peorNombre   = $null
        foreach ($v in $vols) {
            # Se excluyen los volumenes sin tamano real: una unidad optica o un
            # lector de tarjetas vacio reporta 0 GB totales y 0% libre, y sin este
            # filtro el KPI de portada anunciaba "Disco mas ocupado: 100%, unidad E"
            # sobre un lector de DVD vacio. Mismo criterio que usan las reglas de
            # storage del motor de hallazgos.
            $totalGB = ConvertTo-DoubleSafe (Get-Prop $v @('TotalGB', 'TamanoGB', 'SizeGB'))
            if ($null -eq $totalGB -or $totalGB -lt 1) { continue }
            $fs = "$(Get-Prop $v @('FileSystem', 'SistemaArchivos'))".Trim()
            if ($fs -eq '' -or $fs -match '^(?i:&mdash;|N/D|CDFS|UDF)$') { continue }

            $pctRaw = Get-Prop $v @('LibrePct', 'PctLibre', 'PorcentajeLibre', 'FreePercent')
            if ($null -eq $pctRaw) { continue }
            $pct = ConvertTo-DoubleSafe $pctRaw
            if ($null -eq $peorLibrePct -or $pct -lt $peorLibrePct) {
                $peorLibrePct = $pct
                $peorNombre   = Get-Prop $v @('DriveLetter', 'Unidad', 'Letra') 'N/D'
            }
        }
        if ($null -ne $peorLibrePct) {
            $meta.DiscoMasOcupadoPct    = [math]::Round((100 - $peorLibrePct), 0)
            $meta.DiscoMasOcupadoNombre = $peorNombre
        } else {
            $meta.DiscoMasOcupadoPct    = $null
            $meta.DiscoMasOcupadoNombre = $null
        }
    } catch {
        $meta.DiscoMasOcupadoPct    = $null
        $meta.DiscoMasOcupadoNombre = $null
    }

    # --- 1.2.3 Actualizaciones Faltantes: cantidad de filas ----------------
    try {
        $faltantes = ConvertTo-RowArray (Get-SectionData -Id '1.2.3')
        $meta.ParchesPendientes = $faltantes.Count
    } catch { $meta.ParchesPendientes = $null }

    # --- 1.2.2 Hotfixes: primera fila --------------------------------------
    try {
        $hotfixes = ConvertTo-RowArray (Get-SectionData -Id '1.2.2')
        if ($hotfixes.Count -gt 0) {
            $primero     = $hotfixes[0]
            $idParche    = Get-Prop $primero @('HotFixID', 'KB', 'Id', 'Numero')
            $fechaParche = Get-Prop $primero @('InstalledOn', 'FechaInstalacion', 'Fecha')
            # Formato chileno dd-MM-yyyy. Si el dato viene como texto (sesion
            # remota deserializada) se intenta parsear; si no, se deja tal cual.
            if ($fechaParche) {
                if ($fechaParche -is [datetime]) {
                    $fechaParche = $fechaParche.ToString('dd-MM-yyyy')
                } else {
                    $parsed = [datetime]::MinValue
                    if ([datetime]::TryParse([string]$fechaParche, [ref]$parsed)) {
                        $fechaParche = $parsed.ToString('dd-MM-yyyy')
                    }
                }
            }
            if ($idParche -and $fechaParche) {
                $meta.UltimoParche = "$idParche ($fechaParche)"
            } elseif ($idParche) {
                $meta.UltimoParche = "$idParche"
            } elseif ($fechaParche) {
                $meta.UltimoParche = "$fechaParche"
            } else {
                $meta.UltimoParche = $null
            }
        } else {
            $meta.UltimoParche = $null
        }
    } catch { $meta.UltimoParche = $null }

    return $meta
}
#endregion MAIN-03 - Metadatos para el renderer (Get-ReportMeta)

#region MAIN-04 - Flujo completo por equipo (Invoke-SingleTargetReport)
function Invoke-SingleTargetReport {
    <#
        Ejecuta el flujo completo de recoleccion + analisis + exportacion
        para UN equipo, y devuelve un objeto resumen (el mismo objeto sirve
        tanto para el uso equipo-unico como para el consumo del indice de
        flota si en algun momento se invoca en proceso, aunque el modo flota
        real de este script relanza el .ps1 completo por equipo via
        Invoke-FleetReport / FleetWorker, no llama a esta funcion en proceso).

        La sesion remota (si corresponde) se abre y se cierra integramente
        dentro de esta funcion: Remove-TargetSession corre siempre en un
        bloque finally, incluso si algun paso intermedio lanza una excepcion
        no controlada.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ComputerName)

    $inicioEquipo = Get-Date

    # Paso 1: reset de estado por-equipo. Ver Reset-PerTargetState.
    Reset-PerTargetState

    # Paso 2: contexto del equipo actual.
    $script:TargetName = $ComputerName
    $script:IsRemote     = -not ($ComputerName -eq (Get-EquipoActual) -or $ComputerName -eq 'localhost' -or $ComputerName -eq '.')
    $script:Session      = $null

    $resultado = [PSCustomObject]@{
        Equipo             = $ComputerName
        Estado             = 'Error'
        PuntajeSalud       = $null
        Criticos           = 0
        Advertencias       = 0
        ArchivoHtml        = $null
        ArchivoJson        = $null
        Segundos           = 0
        Error              = $null
        SO                 = $null
        IP                 = $null
        UptimeDias         = $null
        EspacioLibreMinPct = $null
        UltimoParche       = $null
        # Campo adicional (no forma parte del contrato minimo pedido), util
        # para que el resumen de consola liste TODOS los archivos generados
        # (incluye la carpeta CSV y el Markdown, que no tienen campo propio).
        ArchivosGenerados  = @()
    }

    try {
        Write-ReportLog -Message "===== Iniciando recoleccion para '$ComputerName' =====" -Level Info

        # Paso 3: abrir sesion (si es remoto). Si falla, equipo inalcanzable.
        if ($script:IsRemote) {
            $script:Session = New-TargetSession -ComputerName $ComputerName -Credential $Credential
            if (-not $script:Session) {
                $motivo = 'No se pudo abrir sesion remota (WinRM). Verifique: servicio WinRM iniciado en el equipo ' +
                    'de destino, resolucion de nombre/DNS correcta, el firewall permitiendo el puerto TCP 5985 ' +
                    '(HTTP) o 5986 (HTTPS), y que la credencial utilizada (parametro -Credential o el contexto ' +
                    'actual) tenga permisos de administrador local en ese equipo.'
                Write-ReportLog -Message "Equipo '$ComputerName' inalcanzable: $motivo" -Level Error
                $resultado.Estado = 'Inalcanzable'
                $resultado.Error  = $motivo
                return $resultado
            }
        }

        # Paso 4: capacidades del equipo objetivo.
        Measure-Section -Name 'Deteccion de capacidades' -ScriptBlock { Get-TargetCapabilities } | Out-Null

        # Paso 5: todos los colectores (fragmentos 02+03, ver Invoke-AllCollectors).
        Measure-Section -Name 'Recoleccion de todas las secciones' -ScriptBlock { Invoke-AllCollectors } | Out-Null

        # Paso 6: motor de hallazgos + resumen.
        Measure-Section -Name 'Motor de hallazgos' -ScriptBlock {
            Invoke-FindingsEngine | Out-Null
            Update-FindingsSummary | Out-Null
        } | Out-Null

        # Paso 7: comparacion contra baseline, si se pidio.
        if ($BaselinePath) {
            Measure-Section -Name 'Comparacion contra baseline' -ScriptBlock {
                $actual = [PSCustomObject]@{ Sections = $script:ReportSections }
                Compare-ReportBaseline -BaselinePath $BaselinePath -Current $actual | Out-Null
                Update-FindingsSummary | Out-Null
            } | Out-Null
        }

        # Paso 8: metadatos para el renderer.
        $meta = Get-ReportMeta -StartTime $inicioEquipo

        # Paso 9: exportar cada formato pedido.
        $timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
        $nombreBase = "AsBuilt_${ComputerName}_$timestamp"

        foreach ($fmt in $Format) {
            try {
                switch ($fmt) {
                    'HTML' {
                        $rutaHtml = Join-Path $OutputPath "$nombreBase.html"
                        $html = ConvertTo-ReportHtml -Meta $meta
                        [System.IO.File]::WriteAllText($rutaHtml, $html, [System.Text.Encoding]::UTF8)
                        $resultado.ArchivoHtml = $rutaHtml
                        $resultado.ArchivosGenerados += $rutaHtml
                        Write-ReportLog -Message "HTML generado en '$rutaHtml'" -Level Success
                    }
                    'JSON' {
                        $rutaJson = Join-Path $OutputPath "$nombreBase.json"
                        $rutaGenerada = Export-ReportJson -Path $rutaJson
                        if ($rutaGenerada) {
                            $resultado.ArchivoJson = $rutaGenerada
                            $resultado.ArchivosGenerados += $rutaGenerada
                        }
                    }
                    'CSV' {
                        $rutaCsvBase = Join-Path $OutputPath "$nombreBase.csv"
                        $carpetaCsv = Export-ReportCsv -Path $rutaCsvBase
                        if ($carpetaCsv) { $resultado.ArchivosGenerados += $carpetaCsv }
                    }
                    'Markdown' {
                        $rutaMd = Join-Path $OutputPath "$nombreBase.md"
                        $rutaGenerada = Export-ReportMarkdown -Path $rutaMd
                        if ($rutaGenerada) { $resultado.ArchivosGenerados += $rutaGenerada }
                    }
                }
            } catch {
                Write-ReportLog -Message "Error al generar el formato $fmt para '$ComputerName': $($_.Exception.Message)" -Level Error
            }
        }

        # Paso 11 (parcial): completar el objeto de resultado con datos utiles
        # para el indice de flota / resumen de consola.
        $resumen = $script:FindingsSummary
        $resultado.Estado       = 'OK'
        $resultado.PuntajeSalud = if ($resumen) { $resumen.PuntajeSalud } else { $null }
        $resultado.Criticos     = if ($resumen) { $resumen.Crit } else { 0 }
        $resultado.Advertencias = if ($resumen) { $resumen.Warn } else { 0 }
        $resultado.UptimeDias   = $meta.UptimeDias
        $resultado.UltimoParche = $meta.UltimoParche

        try { $resultado.SO = $script:Caps.OSCaption } catch { $resultado.SO = $null }

        try {
            $vols = ConvertTo-RowArray (Get-SectionData -Id '1.5.2')
            $peor = $null
            $volsExcl = Get-VolumenesStorageExcluidos
            foreach ($v in $vols) {
                # Mismo criterio que las reglas de storage: una unidad optica o
                # una ISO montada (0% libre) no es espacio agotado.
                $letraVol = "$(Get-Prop $v @('DriveLetter', 'Unidad', 'Letra'))".Trim().TrimEnd(':').ToUpperInvariant()
                if ($letraVol -and $volsExcl.Contains($letraVol)) { continue }
                $pctRaw = Get-Prop $v @('LibrePct', 'PctLibre', 'PorcentajeLibre', 'FreePercent')
                if ($null -eq $pctRaw) { continue }
                $pct = ConvertTo-DoubleSafe $pctRaw
                if ($null -eq $peor -or $pct -lt $peor) { $peor = $pct }
            }
            $resultado.EspacioLibreMinPct = $peor
        } catch { $resultado.EspacioLibreMinPct = $null }

        try {
            $ipFilas = ConvertTo-RowArray (Get-SectionData -Id '1.4.2')
            if ($ipFilas.Count -gt 0) {
                $resultado.IP = Get-Prop $ipFilas[0] @('IPAddress', 'DireccionIP', 'IP')
            }
        } catch { $resultado.IP = $null }

        # Si alguna seccion termino con error durante la recoleccion, el
        # reporte igual se genero (no se aborta el script), pero se marca el
        # estado para que el resumen final y el codigo de salida lo reflejen.
        $hayErrores = $false
        foreach ($t in $script:Timings) {
            if ($t.Estado -and "$($t.Estado)" -match '^Error') { $hayErrores = $true; break }
        }
        if ($hayErrores) { $resultado.Estado = 'ConErrores' }

    } catch {
        Write-ReportLog -Message "Error inesperado procesando '$ComputerName': $($_.Exception.Message)" -Level Error
        $resultado.Estado = 'Error'
        $resultado.Error  = $_.Exception.Message
    } finally {
        # Paso 10: la sesion se cierra SIEMPRE, incluso si algo exploto arriba.
        Remove-TargetSession -Session $script:Session
        $script:Session = $null
        $resultado.Segundos = [math]::Round(((Get-Date) - $inicioEquipo).TotalSeconds, 1)
        Write-ReportLog -Message "===== Fin de recoleccion para '$ComputerName' en $($resultado.Segundos)s (Estado: $($resultado.Estado)) =====" -Level Info
    }

    return $resultado
}
#endregion MAIN-04 - Flujo completo por equipo (Invoke-SingleTargetReport)

#region MAIN-05 - Resumen final por consola
function Write-SingleTargetSummary {
    <# Imprime (via Write-ReportLog) el resumen final para una corrida de un solo equipo. #>
    param([Parameter(Mandatory)]$Resultado)

    if ($Resultado.Estado -eq 'Inalcanzable') {
        Write-ReportLog -Message "No se genero reporte para '$($Resultado.Equipo)': equipo inalcanzable." -Level Error
        Write-ReportLog -Message "Motivo: $($Resultado.Error)" -Level Error
        return
    }

    $puntaje   = if ($null -ne $Resultado.PuntajeSalud) { [int]$Resultado.PuntajeSalud } else { 0 }
    $barra     = Get-HealthBarTexto -Puntaje $puntaje
    $veredicto = Get-VeredictoSalud -Puntaje $puntaje
    $nivelPuntaje = if ($puntaje -lt 50) { 'Error' } elseif ($puntaje -lt 80) { 'Warn' } else { 'Success' }

    Write-ReportLog -Message '' -Level Info
    Write-ReportLog -Message "=========== RESUMEN: $($Resultado.Equipo) ===========" -Level Info
    Write-ReportLog -Message "Puntaje de salud: $barra - $veredicto" -Level $nivelPuntaje

    $resumen = $script:FindingsSummary
    if ($resumen) {
        Write-ReportLog -Message "Hallazgos por severidad -> CRIT: $($resumen.Crit)  WARN: $($resumen.Warn)  INFO: $($resumen.Info)  OK: $($resumen.Ok)" -Level Info

        $criticos = @($script:Findings | Where-Object { $_.Severity -eq 'CRIT' } | Select-Object -First 5)
        if ($criticos.Count -gt 0) {
            Write-ReportLog -Message 'Hallazgos criticos (hasta 5, ver el HTML para el detalle completo):' -Level Error
            foreach ($c in $criticos) {
                Write-ReportLog -Message "  [CRIT] $($c.Category) - $($c.Item): $($c.Detail)" -Level Error
            }
        }
    }

    if ($Resultado.ArchivosGenerados -and $Resultado.ArchivosGenerados.Count -gt 0) {
        Write-ReportLog -Message 'Archivos generados:' -Level Info
        foreach ($a in $Resultado.ArchivosGenerados) {
            if ($a) { Write-ReportLog -Message "  - $a" -Level Info }
        }
    }

    Write-ReportLog -Message "Duracion total: $(Format-DurationTexto -TimeSpan ([timespan]::FromSeconds([double]$Resultado.Segundos)))" -Level Info

    if ($script:Timings -and $script:Timings.Count -gt 0) {
        $top3 = $script:Timings | Sort-Object -Property Segundos -Descending | Select-Object -First 3
        Write-ReportLog -Message 'Secciones mas lentas (para saber que saltear la proxima vez con -SkipSections):' -Level Info
        foreach ($t in $top3) {
            Write-ReportLog -Message "  - $($t.Seccion): $($t.Segundos)s" -Level Info
        }

        $conError = @($script:Timings | Where-Object { "$($_.Estado)" -match '^Error' })
        if ($conError.Count -gt 0) {
            Write-ReportLog -Message "$($conError.Count) seccion(es) terminaron con error durante la recoleccion. Vea el detalle con -Verbose o en el registro incluido en el reporte." -Level Warn
        }
    }

    if ($Resultado.ArchivoJson) {
        Write-ReportLog -Message 'Para comparar contra este estado en la proxima corrida (deteccion de drift):' -Level Info
        Write-ReportLog -Message "  .\Get-ServerFullReport.ps1 -ComputerName $($Resultado.Equipo) -BaselinePath `"$($Resultado.ArchivoJson)`"" -Level Info
    }

    Write-ReportLog -Message '======================================================' -Level Info
}

function Write-FleetSummary {
    <# Imprime (via Write-ReportLog) el resumen final consolidado del modo flota. #>
    param([Parameter(Mandatory)]$FleetResult)

    Write-ReportLog -Message '' -Level Info
    Write-ReportLog -Message '=========== RESUMEN DE FLOTA ===========' -Level Info
    Write-ReportLog -Message "Equipos OK: $($FleetResult.Ok)   Equipos con error/timeout/inalcanzable: $($FleetResult.ConError)" -Level Info

    if ($FleetResult.Inalcanzables -and $FleetResult.Inalcanzables.Count -gt 0) {
        Write-ReportLog -Message 'Equipos con problemas:' -Level Warn
        foreach ($i in $FleetResult.Inalcanzables) {
            Write-ReportLog -Message "  - $($i.Equipo): $($i.Estado) - $($i.Mensaje)" -Level Warn
        }
    }

    if ($FleetResult.TopCriticos -and $FleetResult.TopCriticos.Count -gt 0) {
        Write-ReportLog -Message 'Top hallazgos criticos mas repetidos en el parque:' -Level Error
        foreach ($t in $FleetResult.TopCriticos) {
            Write-ReportLog -Message ("  - {0} (x{1} equipos)" -f $t.Name, $t.Count) -Level Error
        }
    }

    if ($FleetResult.IndexPath) {
        Write-ReportLog -Message "Indice de flota generado en: $($FleetResult.IndexPath)" -Level Success
    }
    Write-ReportLog -Message '=========================================' -Level Info
}
#endregion MAIN-05 - Resumen final por consola

#region MAIN-06 - Ejecucion (equipo unico vs flota) y codigo de salida
# NOTA sobre -Quiet y Write-Host: por regla dura del contrato compartido, este
# fragmento (como todos los demas) nunca llama a Write-Host directamente;
# todo el texto de este bloque (incluido el resumen final) pasa por
# Write-ReportLog, que ya provisto por el fragmento CORE. Eso significa que,
# tal como esta implementado Write-ReportLog, con -Quiet no se imprime NADA
# por consola (ni siquiera los mensajes de Nivel Error), porque el corte
# temprano de esa funcion no distingue nivel. El log completo (incluyendo
# errores) sigue quedando disponible en $script:LogLines y, por lo tanto,
# dentro del reporte generado (seccion de registro de ejecucion). Si se
# necesita que -Quiet permita ver errores igual por consola, ese cambio debe
# hacerse en Write-ReportLog (fragmento CORE), no aca.

$script:ExitCode = 0

trap {
    # Red de seguridad final: cubre una interrupcion (Ctrl+C) o cualquier
    # excepcion no controlada que logre escapar del try/catch/finally de
    # abajo. Garantiza el cierre de la sesion remota antes de salir.
    try { Write-ReportLog -Message "Ejecucion interrumpida o error fatal no controlado: $($_.Exception.Message)" -Level Error } catch {}
    try { Remove-TargetSession -Session $script:Session } catch {}
    exit 1
}

try {
    if ($ComputerName.Count -gt 1) {
        # --- MODO FLOTA ----------------------------------------------------
        if (-not $PSCommandPath) {
            Write-ReportLog -Message ('El modo flota (-ComputerName con mas de un equipo) necesita relanzarse a ' +
                'si mismo una vez por equipo como proceso hijo, y para eso requiere que este script este guardado ' +
                'como archivo .ps1 en disco (no funciona si el codigo se pego directamente en la consola, porque ' +
                'en ese caso no existe una ruta de archivo para volver a invocar). Guarde este script, por ejemplo ' +
                'como Get-ServerFullReport.ps1, y vuelva a ejecutarlo desde ahi. Alternativa mientras tanto: ' +
                'corralo equipo por equipo, una invocacion por cada uno, por ejemplo: ' +
                '.\Get-ServerFullReport.ps1 -ComputerName SRV01') -Level Error
            $script:ExitCode = 1
        } else {
            $extraParams = @{}
            if ($PSBoundParameters.ContainsKey('Format'))            { $extraParams['Format'] = $Format }
            if ($Sections -and $Sections.Count -gt 0)                { $extraParams['Sections'] = $Sections }
            if ($SkipSections -and $SkipSections.Count -gt 0)        { $extraParams['SkipSections'] = $SkipSections }
            if ($BaselinePath)                                       { $extraParams['BaselinePath'] = $BaselinePath }
            if ($Redact)                                             { $extraParams['Redact'] = $true }
            if ($PamBrokerEndpoint)                                  { $extraParams['PamBrokerEndpoint'] = $PamBrokerEndpoint }
            if ($IncludeMissingUpdates)                              { $extraParams['IncludeMissingUpdates'] = $true }
            if ($PSBoundParameters.ContainsKey('EventLogDays'))      { $extraParams['EventLogDays'] = $EventLogDays }
            if ($PSBoundParameters.ContainsKey('PerfSampleSeconds')) { $extraParams['PerfSampleSeconds'] = $PerfSampleSeconds }
            if ($ScanGitRepos)                                       { $extraParams['ScanGitRepos'] = $true }
            if ($GitScanPaths -and $GitScanPaths.Count -gt 0)        { $extraParams['GitScanPaths'] = $GitScanPaths }

            if ($Credential) {
                Write-ReportLog -Message ('Se especifico -Credential junto con el modo flota. La credencial NO se ' +
                    'propaga a los procesos hijo por linea de comandos (pasar contrasenas en texto plano como ' +
                    'argumento no es seguro). Cada proceso hijo del modo flota va a intentar conectar con el ' +
                    'contexto de seguridad con el que corre el proceso padre. Si necesita una credencial especifica ' +
                    'para uno o mas equipos, ejecute el script para esos equipos por separado con -ComputerName ' +
                    '<equipo> -Credential.') -Level Warn
            }

            $fleetResult = Invoke-FleetReport -ComputerName $ComputerName -ScriptPath $PSCommandPath `
                -OutputPath $OutputPath -ThrottleLimit $ThrottleLimit -ExtraParams $extraParams

            if ($fleetResult) {
                Write-FleetSummary -FleetResult $fleetResult
                $huboInalcanzables = @($fleetResult.Inalcanzables | Where-Object { $_.Estado -eq 'Inalcanzable' }).Count -gt 0
                if ($huboInalcanzables) {
                    $script:ExitCode = 2
                } elseif ($fleetResult.ConError -gt 0) {
                    $script:ExitCode = 1
                } else {
                    $script:ExitCode = 0
                }
            } else {
                Write-ReportLog -Message 'El modo flota no devolvio resultados (ver mensajes de error anteriores).' -Level Error
                $script:ExitCode = 1
            }
        }
    } else {
        # --- MODO EQUIPO UNICO ----------------------------------------------
        $equipoUnico   = $ComputerName[0]
        $resultadoUnico = Invoke-SingleTargetReport -ComputerName $equipoUnico

        Write-SingleTargetSummary -Resultado $resultadoUnico

        switch ($resultadoUnico.Estado) {
            'Inalcanzable' { $script:ExitCode = 2 }
            'ConErrores'   { $script:ExitCode = 1 }
            'Error'        { $script:ExitCode = 1 }
            default        { $script:ExitCode = 0 }
        }
    }
} catch {
    Write-ReportLog -Message "Error fatal no controlado en el bloque principal: $($_.Exception.Message)" -Level Error
    $script:ExitCode = 1
} finally {
    # Red de seguridad adicional: si por algun motivo Invoke-SingleTargetReport
    # no llego a su propio finally (por ejemplo, una excepcion durante el
    # reset de estado, antes de entrar al try de esa funcion), esto asegura
    # que ninguna sesion remota quede abierta al terminar el script.
    try { Remove-TargetSession -Session $script:Session } catch {}
}

exit $script:ExitCode
#endregion MAIN-06 - Ejecucion (equipo unico vs flota) y codigo de salida
