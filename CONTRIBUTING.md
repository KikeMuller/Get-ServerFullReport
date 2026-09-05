# Cómo contribuir

Gracias por el interés. Este proyecto es un solo archivo `.ps1`, así que contribuir es más simple de lo habitual — pero tiene algunas reglas duras que conviene conocer antes de escribir código.

## Reportar un bug

**Adjunta el JSON de la ejecución.** Es lo más útil que puedes mandar: contiene todo lo recolectado, los hallazgos que el motor concluyó y el log de ejecución embebido. Con eso se puede reproducir el problema sin acceso a tu equipo.

```powershell
.\Get-ServerFullReport.ps1 -Format JSON -Redact
```

`-Redact` enmascara IPs, cuentas, rutas, seriales y thumbprints. Revísalo igual antes de subirlo a un issue público.

Incluye también:

- Versión de Windows y de PowerShell (`$PSVersionTable.PSVersion`)
- El comando exacto que ejecutaste
- Si corrías como Administrador
- Local o remoto (y si fue remoto, contra qué versión de Windows)

## Reglas duras del código

Estas no son preferencias de estilo. Romperlas rompe el script en producción.

**1. PowerShell 5.1.** El objetivo es que corra en un Windows Server 2016 recién instalado, sin instalar nada. Nada de `??`, `?.`, `ForEach-Object -Parallel`, `ConvertFrom-Json -AsHashtable`. Si necesitas algo de PS 7, va detrás de `if ($script:IsPS7) { ... } else { ...fallback... }`.

**2. Cuidado con el desenrollado de arreglos.** En PS 5.1, `return $arreglo` desenrolla un arreglo de un solo elemento y el llamador recibe el objeto suelto; después `$objeto[0]` devuelve `$null` (en PS 7 devuelve el objeto). Este bug dejó 17 secciones en blanco. Usa siempre el operador coma unario: `return ,$arreglo`.

**3. Nada de caracteres no-ASCII en el código.** Ni en los comentarios ni en las cadenas. Escribe "Configuracion", "Numero", "Ultimo". El HTML generado sí puede usar entidades (`&oacute;`).

**4. Solo lectura.** El script nunca modifica el equipo auditado. Nada de `Set-`, `New-`, `Remove-`, `Start-`, `Stop-` sobre el sistema; solo archivos temporales propios que se limpian.

**5. Cada colector va en su propio try/catch.** Un rol ausente, un permiso denegado o un cmdlet inexistente nunca deben abortar el reporte. En el `catch`, devolvé un string explicativo en español: el renderer lo muestra como mensaje en vez de tabla.

**6. Nunca afirmar sobre un dato que no se pudo leer.** Es el principio central del motor de hallazgos, y el que más veces se violó por accidente. Si el valor viene `N/D`, `$null`, vacío o "No se pudo determinar", la regla no emite nada — ni un hallazgo negativo, ni un OK. Un OK falso es peor que el silencio, porque alguien lo va a creer. Usa `Test-ValorUtilizable` y `Get-NumeroSeguro`, y llama a `Set-TopicoNoEvaluable` cuando no puedas verificar.

**7. Todo el texto visible al usuario, en español.**

## Agregar una sección nueva

```powershell
Measure-Section '1.X.Y Nombre corto' {
    $data = Invoke-Remote -TimeoutSec 60 -ScriptBlock {
        try {
            Get-LoQueSea -ErrorAction Stop | Select-Object Propiedad1, Propiedad2
        } catch {
            "No se pudo consultar X: $($_.Exception.Message)"
        }
    }
    Add-ReportSection -Id '1.X.Y' -Title 'Nombre Legible' -Data $data `
        -Note 'Una frase explicando que es y por que le importa a un sysadmin.'
}
```

La `-Note` es valor real, no relleno. "Aqui se muestran los servicios" no sirve; "Servicios en modo automatico que no estan corriendo suelen indicar una falla de arranque" sí.

Si la sección depende de un rol, usa `$script:Caps` (`IIS`, `DHCP`, `DNS`, `FileServer`, `RDSH`, `WSUS`, `SQL`, `HyperV`, `Cluster`, `DomainMember`, `IsDC`, `IsServerOS`, `IsVM`). Cuando el rol no está, registra la sección igual con un string explicando la ausencia: así el TOC queda completo y el as-built documenta explícitamente que ese rol no existe.

Después de agregarla, actualiza [`docs/INVENTARIO_SECCIONES.md`](docs/INVENTARIO_SECCIONES.md) con el Id, el título y los nombres exactos de las columnas. Ese archivo es la fuente de verdad que usan las reglas del motor de hallazgos; desincronizarlo es cómo se producen los falsos positivos.

## Agregar una regla de hallazgo

```powershell
Invoke-CheckBlock -Categoria 'Seguridad' -Topico 'MiTopico' `
    -OkMessage 'Descripcion afirmativa de lo que se verifico y esta bien.' `
    -RequiereSecciones @('1.X.Y') -Body {
    Invoke-Rule {
        $datos = ConvertTo-RowArray (Get-SectionData -Id '1.X.Y')
        if ($datos.Count -eq 0) { Set-TopicoNoEvaluable; return }

        $valor = Get-NumeroSeguro (Get-Prop $datos[0] @('MiPropiedad'))
        if ($null -eq $valor) { Set-TopicoNoEvaluable; return }

        if ($valor -gt 100) {
            Add-Finding -Severity 'WARN' -Category 'Seguridad' `
                -Item 'Titulo corto del hallazgo' `
                -Detail "Evidencia concreta: valor $valor" `
                -Recommendation 'Que hacer al respecto.' `
                -Control 'CIS X.Y / ISO 27001 A.Z' `
                -SectionRef '1.X.Y'
        }
    }
}
```

`-RequiereSecciones` es obligatorio en reglas nuevas: sin eso, el tópico emite su OK aunque la sección no exista.

## Probar los cambios

No hay framework de tests; hay dos scripts de regresión y una verificación de sintaxis.

```powershell
# Sintaxis
$e = $null
[System.Management.Automation.Language.Parser]::ParseFile(
    "$PWD\Get-ServerFullReport.ps1", [ref]$null, [ref]$e)
if ($e) { $e | ForEach-Object { "L$($_.Extent.StartLineNumber): $($_.Message)" } } else { 'OK' }

# Sin caracteres no-ASCII
Select-String -Path .\Get-ServerFullReport.ps1 -Pattern '[^\x00-\x7F]'

# La prueba que importa: ejecutarlo de verdad, entero
.\Get-ServerFullReport.ps1 -OutputPath $env:TEMP\test -Format HTML,JSON,CSV,Markdown
```

Ese último punto tiene su historia: durante el desarrollo se probaron las funciones por separado (recolección, motor, renderer, exportadores) y todo pasaba, pero el bloque principal nunca se ejecutaba — y ahí estaba el bug que abortaba el script al final. **Ejecuta el script completo antes de mandar un PR**, y si tocaste algo del motor de hallazgos, ejecútalo también contra el JSON de un equipo real y revisa que no aparezcan hallazgos inventados.

## Pull requests

- Una cosa por PR.
- Explica el *por qué*, no el *qué* — el diff ya dice qué cambió.
- Si corriges un bug, cuenta cómo lo reprodujiste.
- Actualiza el `CHANGELOG.md` en la sección correspondiente.
- Sube `$script:ScriptVersion` solo si el PR cierra una versión.
