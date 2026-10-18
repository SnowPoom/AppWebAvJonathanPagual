# =====================================================================
#  PoliMusic - Despliegue en IIS
#  Ejecutar en PowerShell COMO ADMINISTRADOR:
#     powershell -ExecutionPolicy Bypass -File .\deploy_iis.ps1
# =====================================================================
param(
    [string]$Source = (Join-Path $PSScriptRoot "PoliMusic"),
    [string]$Target = "C:\inetpub\wwwroot\PoliMusic",
    [string]$AppCmd = "$env:windir\system32\inetsrv\appcmd.exe"
)

$ErrorActionPreference = "Stop"

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object Security.Principal.WindowsPrincipal($id)
    return $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

Write-Output "== PoliMusic :: despliegue en IIS =="

if (-not (Test-Admin)) {
    Write-Output ""
    Write-Output "ERROR: este script necesita privilegios de administrador."
    Write-Output "Cierre esta ventana y abra PowerShell con 'Ejecutar como administrador'."
    exit 1
}

if (-not (Test-Path $Source)) {
    Write-Output "ERROR: no se encontro la carpeta de origen: $Source"
    exit 1
}

# 0) Archivo de conexion local (no versionado)
$localCfg = Join-Path $Source "includes\config.local.asp"
$exampleCfg = Join-Path $Source "includes\config.example.asp"
if (-not (Test-Path $localCfg)) {
    if (Test-Path $exampleCfg) {
        Copy-Item $exampleCfg $localCfg
        Write-Output "AVISO: se creo includes\config.local.asp a partir de la plantilla."
        Write-Output "       Edite ese archivo con los datos de su SQL Server y vuelva a ejecutar el script."
        exit 1
    }
    Write-Output "ERROR: falta includes\config.local.asp y no hay plantilla config.example.asp."
    exit 1
}

# 1) Caracteristica ASP clasico
$asp = Get-WindowsOptionalFeature -Online -FeatureName IIS-ASP
if ($asp.State -ne "Enabled") {
    Write-Output "Habilitando IIS-ASP..."
    Enable-WindowsOptionalFeature -Online -FeatureName IIS-ASP -All -NoRestart | Out-Null
    Write-Output "IIS-ASP habilitado."
} else {
    Write-Output "IIS-ASP ya estaba habilitado."
}

# 2) Copia de archivos
if (-not (Test-Path $Target)) { New-Item -ItemType Directory -Path $Target | Out-Null }
robocopy $Source $Target /MIR /NFL /NDL /NJH /NJS /NP /R:2 /W:1 | Out-Null
if ($LASTEXITCODE -ge 8) { throw "robocopy fallo con codigo $LASTEXITCODE" }
Write-Output "Archivos sincronizados en $Target"

# 3) Permisos NTFS (lectura para IIS, modificacion para el usuario actual)
icacls $Target /grant "IIS_IUSRS:(OI)(CI)(RX)" /T /Q | Out-Null
icacls $Target /grant "$($env:USERDOMAIN)\$($env:USERNAME):(OI)(CI)(M)" /T /Q | Out-Null
Write-Output "Permisos NTFS aplicados."

# 4) Configuracion de ASP clasico
& $AppCmd set config /section:asp /session.timeout:00:01:00 /commit:apphost | Out-Null
& $AppCmd set config /section:asp /enableParentPaths:true /commit:apphost | Out-Null
& $AppCmd set config /section:asp /scriptErrorSentToBrowser:true /commit:apphost | Out-Null
Write-Output "ASP configurado: sesion 60 min, rutas primarias y errores detallados."

# 5) App pool y sitio
$poolState = (& $AppCmd list apppool "DefaultAppPool" /text:state) -join ""
if ($poolState.Trim() -eq "") {
    & $AppCmd add apppool /name:"DefaultAppPool" | Out-Null
    $poolState = "Started"
}
if ($poolState.Trim() -ne "Started") { & $AppCmd start apppool "DefaultAppPool" | Out-Null }

$siteState = (& $AppCmd list site "Default Web Site" /text:state) -join ""
if ($siteState.Trim() -ne "Started") { & $AppCmd start site "Default Web Site" | Out-Null }
Write-Output "App pool y sitio activos."

# 6) Verificacion HTTP
Start-Sleep -Seconds 2
try {
    $r = Invoke-WebRequest -Uri "http://localhost/PoliMusic/login.asp" -UseBasicParsing -TimeoutSec 25
    Write-Output ("OK  login.asp -> HTTP " + [int]$r.StatusCode)
} catch {
    Write-Output ("AVISO login.asp -> " + $_.Exception.Message)
}
try {
    $r2 = Invoke-WebRequest -Uri "http://localhost/PoliMusic/assets/css/polimusic.css" -UseBasicParsing -TimeoutSec 25
    Write-Output ("OK  polimusic.css -> HTTP " + [int]$r2.StatusCode + " (" + $r2.RawContentLength + " bytes)")
} catch {
    Write-Output ("AVISO polimusic.css -> " + $_.Exception.Message)
}

Write-Output ""
Write-Output "Listo. Abra: http://localhost/PoliMusic/"
Write-Output "Credenciales: userAdmin / 12345678  |  userNormal / 12345678"
Write-Output "== Fin =="
