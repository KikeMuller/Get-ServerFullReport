# Inventario real de secciones (`Add-ReportSection`)

Fuente de verdad generada leyendo `frag_02_collectors_base.ps1` y
`frag_03_collectors_new.ps1` línea por línea. Para cada sección se listan las
propiedades exactas de los `[PSCustomObject]` (o columnas de `Select-Object`)
que devuelve su colector. Cuando la sección puede devolver un **string** en
vez de objetos (rol no instalado, error, "no aplica"), se indica.

> Nota: varias secciones solo se registran si el rol correspondiente está
> presente (`$script:Caps.XXX`); en ese caso, si el rol falta, `-Data` es un
> string y NO tiene las propiedades de abajo.

| Id | Titulo | Propiedades | Fragmento |
|---|---|---|---|
| 1.1 | Hardware del Host | Hostname, Fabricante, Modelo, NumeroSerie, CPU, NucleosFisicos, NucleosLogicos, MemoriaTotalGB, TipoSistema, MaquinaVirtual | 02 |
| 1.2.1 | Configuracion del Sistema Operativo | SistemaOperativo, Version, Build, Arquitectura, FechaInstalacion, UltimoArranque, UptimeDias, DirectorioWindows, ZonaHoraria, NombreDominio, RolDominio | 02 |
| 1.2.2 | Hotfixes Instalados | HotFixID, Description, InstalledBy, InstalledOn | 02 |
| 1.2.3 | Actualizaciones de Windows Faltantes | Titulo, KB, Severidad, FechaPublic (string si -IncludeMissingUpdates no se especifico) | 02 |
| 1.2.4 | Drivers Instalados | DeviceName, DriverVersion, Manufacturer, DriverDate, DeviceClass, IsSigned | 02 |
| 1.2.5 | Roles y Caracteristicas Instaladas | DisplayName, Name, FeatureType, Path | 02 |
| 1.2.6 | Aplicaciones Instaladas | DisplayName, DisplayVersion, Publisher, InstallDate, InstallLocation, Arquitectura, EsComponenteDelSistemaOActualizacion | 02 |
| 1.2.7 | Servicios | Name, DisplayName, State, StartMode, StartName, PathName, RutaSinComillas | 02 |
| 1.3.1 | Perfiles de Firewall | Name, Enabled, DefaultInboundAction, DefaultOutboundAction, LogFileName, LogAllowed, LogBlocked | 02 |
| 1.3.2 | Reglas de Firewall - Resumen por Perfil | Perfil, Direccion, Accion, Cantidad | 02 |
| 1.3.3 | Reglas de Firewall Entrantes - Detalle | DisplayName, Direccion, Accion, Perfil, Protocolo, PuertoLocal, PuertoRemoto, DireccionLocal, DireccionRemota | 02 |
| 1.4.1 | Adaptadores de Red | Name, InterfaceDescription, Status, LinkSpeed, MacAddress, MediaType, InterfaceIndex | 02 |
| 1.4.2 | Direcciones IP (IPv4) | InterfaceAlias, IPAddress, PrefixLength, PrefixOrigin, SuffixOrigin, AddressState | 02 |
| 1.4.3 | Configuracion del Cliente DNS | InterfaceAlias, ConnectionSpecificSuffix, RegisterThisConnectionsAddress, UseSuffixWhenRegistering | 02 |
| 1.4.4 | Servidores DNS Configurados por Interfaz | InterfaceAlias, ServerAddresses | 02 |
| 1.4.5 | MTU de Adaptadores de Red | InterfaceAlias, NlMtu, Dhcp, ConnectionState | 02 |
| 1.5.1 | Discos Locales | DeviceID, Model, InterfaceType, SerialNumber, MediaType, TamanoGB, Partitions | 02 |
| 1.5.2 | Volumenes del Host | DriveLetter, FileSystemLabel, FileSystem, HealthStatus, TotalGB, LibreGB, LibrePct | 02 |
| 1.6 | Servidor IIS (rol ausente) | string | 02 |
| 1.6.1 | Application Pools de IIS | Name, EstadoActual, managedRuntimeVersion, managedPipelineMode, IdentityType, IdentityUser, IdleTimeoutMin, PeriodicRestartMin | 02 |
| 1.6.2 | Sitios de IIS - Resumen | Name, ID, State, PhysicalPath, Bindings, ApplicationPool | 02 |
| 1.6.2.1 | Sitios de IIS - Detalle y Aplicaciones Virtuales | Sitio, Estado, RutaFisica, Bindings, Aplicaciones | 02 |
| 1.7 | Servidor de Archivos (rol/shares ausentes) | string | 02 |
| 1.7.1 | Configuracion SMB del Servidor | EnableSMB1Protocol, EnableSMB2Protocol, ServerHidden, AnnounceServer, EncryptData, RejectUnencryptedAccess, AuditSmb1Access (**NO** tiene RequireSecuritySignature; solo existe si `Caps.FileServer` = true) | 02 |
| 1.7.2 | Recursos Compartidos (File Shares) | Name, Path, Description, ShareType, CurrentUsers, FolderEnumerationMode, CachingMode | 02 |
| 1.7.2.N | Permisos del Share: <nombre> | AccountName, AccessControlType, AccessRight | 02 |
| 1.8 / 1.8.1-1.8.4.N | DHCP Server (varias) | ScopeId, Name, StartRange, EndRange, SubnetMask, State, LeaseDuration; TotalScopes, TotalAddresses, AddressesInUse, AddressesAvailable, PercentageInUse; Free, InUse, Reserved; Name, PartnerServer, Mode, State; InterfaceAlias, IPAddress, State; OptionId, Name, Value, VendorClass; DynamicUpdates, DeleteDnsRRonLeaseExpiry, UpdateDnsRRForOlderClients | 02 |
| 1.9 / 1.9.1-1.9.4.4 | DNS Server (varias) | ListenAddresses, BootMethod, EnableDnsSec; ScavengingState, ScavengingInterval, LastScavengeTime, RefreshInterval, NoRefreshInterval; IPAddress, Timeout, UseRootHint; ZoneName, ZoneType, IsAutoCreated, IsDsIntegrated, IsReverseLookupZone, DynamicUpdate; Zona, TransferPolicy; ZoneName, MasterServers, IsDsIntegrated; Zona, AgingEnabled, ScavengeServers | 02 |
| 1.10 | Terminal Services / RD Licensing (rol ausente) | string | 02 |
| 1.10.1 | Configuracion de Licenciamiento RDS | LicenseServer, Mode, ConnectionBroker **o** (fallback CIM) ModoLicenciamiento, ServidoresLicencia | 02 |
| 1.10.2 | Paquetes de Licencias (CALs) - Uso por Tipo | Description, ProductVersion, TypeAndModel, KeyPackType, TotalLicenses, IssuedLicenses, LicenciasDisponibles, ExpirationDate | 02 |
| 1.10.3 | Licencias Emitidas (Detalle) | sIssuedToUser, sIssuedToComputer, FechaEmision, FechaExpiracion, sKeyPackId | 02 |
| 1.10.4 | Historial de Logon RDS | Usuario, UltimoLogon, UltimaIPOrigen, ConexionesEnVentana | 02 |
| 1.11.1 | Certificados SSL/TLS Instalados | CN, Sujeto, Emisor, HuellaDigital, FechaExpiracion, DiasParaExpirar, Expirado, AutoFirmado, UsoMejorado, TieneClavePrivada | 02 |
| 1.11.2 | Protocolos, Cifrados y Hashes SChannel | Tipo (Protocolo/Cifrado/Hash), **Nombre**, Lado, Estado | 02 |
| 1.11.3 | Estado PAM / BeyondTrust | AgentePAMInstalado, ServiciosPAMDetectados, ConectividadResourceBroker, ConectividadOK | 02 |
| 1.11.4 | Configuracion NTP | Estado, Fuentes | 02 |
| 1.11.5 | Tareas Programadas Activas | Nombre, Ruta, Estado, EjecutarComo, NivelPrivilegio, UltimaEjecucion, ProximaEjecucion, UltimoResultado | 02 |
| 1.11.6 | Antivirus / EDR Detallado | WindowsDefender (string volcado, NO expone RealTimeProtectionEnabled/AntivirusSignatureLastUpdated como propiedades sueltas), AgentesEDRTerceros | 02 |
| 1.11.7 | Reinicio Pendiente | PendingFileRename, ComponentBasedServicing, WindowsUpdateReboot, ReinicioPendiente | 02 |
| 1.11.8 | Estado de Backups (Agentes Locales) | DisplayName, Status, StartType (o string si no hay agente) | 02 |
| 1.12 / 1.12.1-1.12.2 | WSUS | Servidor, Puerto, SincronizacionMode, UpstreamServer, ProductosAprobados; StartTime, EndTime, Result, Error | 02 |
| 1.13 | Group Policy Aplicadas (RSOP) | NombreGPO, Habilitada, Version | 02 |
| 1.14.1 | Cuentas de Servicio con Privilegios Elevados | Name, DisplayName, StartName, State, TipoCuenta | 02 |
| 1.15.x | Python/Git | Version, Ruta, Origen; Interprete, TotalPaquetes, Paquetes; Ruta, VersionPython, UltimaModificacion; Entorno, TotalPaquetes; Ruta, RemoteOrigin, RamaActual, UltimaModificacion | 02 |
| 1.14.2 | Administradores Locales | Nombre, TipoObjeto, Origen, SID | 03 |
| 1.14.3 | Usuarios Locales | Nombre, **Habilitado**, PasswordNeverExpires, PasswordRequired, UltimoLogon, PasswordLastSet, DiasDesdeCambioPassword, Descripcion | 03 |
| 1.14.4 | Grupo Remote Desktop Users | Nombre, TipoObjeto, Origen, SID | 03 |
| 1.14.5 | Politica de Contrasenas y Bloqueo | LongitudMinima, VigenciaMaxima, VigenciaMinima, HistorialPasswords, UmbralBloqueo, DuracionBloqueo | 03 |
| 1.14.6 | Politica de Auditoria | Categoria, Subcategoria, Configuracion | 03 |
| 1.14.7 | Derechos de Usuario Sensibles | Derecho, AsignadoA | 03 |
| 1.16.1 | Dominio y Ubicacion en AD | Dominio, Forest, SitioAD, DCAutenticante, DCDelDominio, DNCompleto, CanalSeguroOK | 03 |
| 1.16.2 | Replicacion de Active Directory | SocioReplicacion, NamingContext, UltimoIntento, UltimoResultado, NumeroDeFallos (solo si `Caps.IsDC`) | 03 |
| 1.16.3 | Capa de Virtualizacion | Hipervisor, FabricanteHW, ModeloHW (+ VMwareTools*/IntegrationServices segun hipervisor) | 03 |
| 1.16.3.1 | Maquinas Virtuales Alojadas (Host Hyper-V) | Nombre, Estado, CPUs, MemoriaGB, Uptime, Version | 03 |
| 1.16.4 | Instancias de SQL Server | NombreInstancia, IdInstancia, Version, Edicion, ServicioMotor, ServicioAgente, PuertoTCP | 03 |
| 1.16.5.1 | DFS Namespaces | Path, Type, State, TimeToLiveSec | 03 |
| 1.16.5.2 | Impresoras Compartidas | Name, ShareName, DriverName, PortName, PublishedToAD | 03 |
| 1.16.5.3 | Failover Cluster / Nodos | Name, State, Type | 03 |
| 1.16.5.4 | Failover Cluster - Recursos | Name, State, OwnerGroup, ResourceType | 03 |
| 1.16.5.5 | Contenedores Docker | Nombre, Imagen, Estado, Puertos | 03 |
| 1.16.6 | NIC Teaming / LBFO | NombreTeam, ModoTeaming, ModoBalanceo, Estado, Miembros | 03 |
| 1.16.7 | iSCSI y MPIO | TargetsISCSI, ConexionesISCSI, ConfiguracionMPIO, DiscosMPIO | 03 |
| 1.16.8 | Configuracion de WinRM y PowerShell Remoting | AllowUnencrypted, AutenticacionBasic, AutenticacionKerberos, AutenticacionNegotiate, AutenticacionCert, TrustedHosts, Listeners, PoliticasEjecucion | 03 |
| 1.17.1 | Uptime y Arranque | UltimoArranque, **DiasEncendido**, FastStartupHabilitado, PageFileGestionAuto, ArchivosPaginacion | 03 |
| 1.17.2 | Apagados Inesperados y Reinicios (90 dias) | Fecha, **EventoID**, Tipo, Origen, Mensaje | 03 |
| 1.17.3 | Errores Criticos del Event Log | LogName, EventoID, Proveedor, Nivel, **Cantidad**, PrimeraOcurrencia, UltimaOcurrencia, MensajeEjemplo | 03 |
| 1.17.4 | Baseline de Rendimiento (opt-in `-PerfSampleSeconds`) | Contador (nombre del contador, no columnas fijas), Promedio, Maximo, Minimo, Muestras — **una fila por contador**, no un objeto unico con MemoriaDisponiblePct/CPUPromedio | 03 |
| 1.17.5 | Top 15 Procesos por Memoria | Nombre, PID, MemoriaMB, CPUSegundos, Usuario, RutaEjecutable, Inicio | 03 |
| 1.17.6 | Salud de Discos Fisicos | FriendlyName, MediaType, OperationalStatus, **HealthStatus**, DesgastePct, TamanoGB | 03 |
| 1.4.6 | Puertos TCP/UDP en Escucha | Protocolo, DireccionLocal, PuertoLocal, PID, Proceso, RutaProceso, Servicio | 03 |
| 1.4.7 | Rutas Persistentes y Puerta de Enlace por Defecto | Destino, MascaraPrefijo, SiguienteSalto, InterfaceAlias, Metrica, Persistente | 03 |
| 1.4.8 | Archivo Hosts | IP, NombreHost, LineaCompleta | 03 |
| 1.4.9 | Proxy Configurado | ProxyWinHTTP, ProxyUsuarioHabilitado, ProxyUsuarioServidor, ProxyUsuarioExcepciones | 03 |
| 1.5.3 | Estado de BitLocker | MountPoint, VolumeStatus, ProtectionStatus, EncryptionPercentage, EncryptionMethod, KeyProtectorTypes | 03 |
| 1.5.4 | Deduplicacion de Datos | Volume, Enabled, Capacity, AhorroGB, SavingsPercent, OptimizedFilesCount | 03 |
| 1.5.4.1 | Cuotas FSRM | Path, Size, Usage, Description, PorcentajeUsado | 03 |
| 1.7.3 | Permisos NTFS Efectivos de los Shares | Recurso, RutaLocal, Identidad, Permisos, TipoAcceso, Heredado, EsRiesgoso | 03 |
| 1.11.9 | Estado de Hardening del Sistema Operativo | **Filas** Parametro/ValorActual/ValorRecomendado/Cumple (SMBv1 habilitado, Firma SMB requerida (servidor), LLMNR, NetBIOS sobre TCP/IP, UAC habilitado (EnableLUA), UAC ConsentPromptBehaviorAdmin, RDP habilitado, Puerto RDP, RDP - Network Level Authentication, RDP - Nivel minimo de cifrado) — NO son propiedades sueltas del objeto, son FILAS que hay que buscar por `Parametro` | 03 |
| 1.11.10 | TPM y Secure Boot | **TPM_Presente**, TPM_Listo, TPM_Habilitado, TPM_VersionFabricante, SecureBootHabilitado, VBS_Estado, CredentialGuard_Configurado, CredentialGuard_EnEjecucion, HVCI_Configurado, HVCI_EnEjecucion | 03 |
| 1.11.11 | LAPS (Local Administrator Password Solution) | LAPSDetectado, Habilitado, PoliticaComplejidad, EdadMaximaPassword (o AdmPwdDllPresente en variante legacy) — NO existe una propiedad `LAPSConfigurado`/`Configurado`/`Instalado` | 03 |
| 1.11.12 | Activacion y Licenciamiento de Windows | Nombre, Descripcion, **EstadoLicencia** (string mapeado: 'Licenciado', 'Sin licencia', etc. — NO es el codigo numerico), CanalActivacion, ServidorKMS, DiasRestantes | 03 |
| 1.11.13 | Almacenes de Certificados Adicionales | Almacen, Sujeto, Emisor, HuellaDigital, FechaExpiracion, **Sospechoso** (NO `Sospechosa`) | 03 |

**Total: 110 llamadas a `Add-ReportSection`** en `frag_02_collectors_base.ps1`
(70) + `frag_03_collectors_new.ps1` (40), agrupadas en **82 filas** de esta
tabla (un Id puede tener rama "rol ausente" (string) y rama "rol presente"
con datos estructurados, o generarse dentro de un loop con sufijo numerico
como `1.7.2.$i` / `1.8.4.$j`; esas variantes se agrupan bajo un mismo Id
cuando corresponde).
