<#
.SYNOPSIS
    Installeert de CargoWise WiseCloud Client (WiseTech Global) stil, voor AVD-images.
.DESCRIPTION
    Werkwijze:
    1. WiseCloudClientSetup.exe ophalen
       - uit Azure Blob Storage met een managed identity (-ClientId), of
       - van een directe URL zonder login (laat -ClientId leeg), of
       - van een lokaal pad (-InstallerPath, voor testen)
    2. Stil installeren: WiseCloudClientSetup.exe /quiet /norestart /log <bestand>

    De installer (versie 6.5.7517) is een WiX Burn-bundle ('WiX Toolset Bootstrapper Engine').
    De /extract-werkwijze uit de WiseTech-handleiding (nov. 2022) geldt voor een oudere installer;
    Burn kent die optie niet en toont dan de installatie-UI. Burn installeert de onderdelen
    (WiseCloud Client, CargoWise Remote Desktop Services, CargoWise Application Manager) zelf.

    Updates: de CargoWise Application Manager-service werkt de client daarna zelf bij,
    zonder adminrechten. Het image levert alleen de basisinstallatie.

    De installer staat NIET in deze (publieke) repo: WiseTech biedt hem alleen achter een
    login aan. Zet hem in een prive storage account of gebruik een directe link van WiseTech.

    In Azure Image Builder (scriptUri kan geen parameters meegeven), dus:
      - File-stap:       sourceUri = raw-URL van dit script, destination = C:\AVDImage\WiseCloudClient.ps1
      - PowerShell-stap: inline = C:\AVDImage\WiseCloudClient.ps1 -InstallerUrl '<url>' -ClientId '<id>'
.PARAMETER InstallerUrl
    URL van WiseCloudClientSetup.exe (blob-URL of directe downloadlink).
.PARAMETER ClientId
    Client ID van de user-assigned managed identity op de build-VM, met Storage Blob Data
    Reader op de container. Leeg = downloaden zonder authenticatie.
.PARAMETER InstallerPath
    Lokaal pad naar WiseCloudClientSetup.exe. Slaat de download over (voor testen).
.EXAMPLE
    # In Image Builder
    C:\AVDImage\WiseCloudClient.ps1 -InstallerUrl 'https://<account>.blob.core.windows.net/cargowise/WiseCloudClientSetup.exe' -ClientId '<client-id>'
.EXAMPLE
    # Lokaal testen, als administrator
    .\WiseCloudClient.ps1 -InstallerPath 'C:\Temp\WiseCloudClientSetup.exe'
.NOTES
    Versie 1.1 - 23-11-2026 - Burn-bundle: /quiet in plaats van /extract (installer 6.5.7517)
    Versie 1.0 - 23-11-2026 - eerste versie (Slade, ticket 86157695)
#>
[CmdletBinding(DefaultParameterSetName = 'Url')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Url')]
    [string]$InstallerUrl,

    [Parameter(ParameterSetName = 'Url')]
    [string]$ClientId,

    [Parameter(Mandatory, ParameterSetName = 'Local')]
    [string]$InstallerPath
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$workDir = 'C:\AVDImage\WiseCloud'
$logDir  = 'C:\Windows\Temp\WiseCloud'

function Write-Log {
    param([string]$Message)
    Write-Host "WiseCloud | $(Get-Date -Format 'HH:mm:ss') | $Message"
}

function Get-CargoWiseProduct {
    <#
    .SYNOPSIS
        Geeft geinstalleerde producten van WiseTech/CargoWise/WiseCloud terug (uit het register).
    #>
    $keys = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    Get-ItemProperty -Path $keys -ErrorAction SilentlyContinue |
        Where-Object { $_.PSObject.Properties['DisplayName'] -and $_.DisplayName -match 'CargoWise|WiseCloud|WiseTech' } |
        Select-Object DisplayName, DisplayVersion, InstallLocation
}

try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) { throw 'Dit script moet als administrator draaien.' }

    # Schone werkmap, zodat een eerdere run de controles niet beinvloedt
    if (Test-Path -Path $workDir) { Remove-Item -Path $workDir -Recurse -Force }
    foreach ($d in $workDir, $logDir) { $null = New-Item -ItemType Directory -Path $d -Force }
    $setup = Join-Path -Path $workDir -ChildPath 'WiseCloudClientSetup.exe'

    # 1. Installer ophalen
    if ($PSCmdlet.ParameterSetName -eq 'Local') {
        if (-not (Test-Path -Path $InstallerPath)) { throw "Installer niet gevonden: $InstallerPath" }
        Write-Log "Lokale installer: $InstallerPath"
        Copy-Item -Path $InstallerPath -Destination $setup
    }
    else {
        $headers = @{}
        if ($ClientId) {
            # Token via IMDS met de user-assigned identity van de (build-)VM
            Write-Log 'Token ophalen voor Azure Storage via managed identity'
            $imds = "http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=https%3A%2F%2Fstorage.azure.com%2F&client_id=$ClientId"
            $token = $null
            for ($i = 1; $i -le 10 -and -not $token; $i++) {
                try { $token = (Invoke-RestMethod -Uri $imds -Headers @{ Metadata = 'true' } -TimeoutSec 30).access_token }
                catch { Write-Log "Token nog niet beschikbaar (poging $i): $($_.Exception.Message)"; Start-Sleep -Seconds 15 }
            }
            if (-not $token) { throw 'Geen token van IMDS ontvangen.' }
            $headers = @{ Authorization = "Bearer $token"; 'x-ms-version' = '2021-08-06' }
        }
        Write-Log "Downloaden: $InstallerUrl"
        Invoke-WebRequest -Uri $InstallerUrl -OutFile $setup -UseBasicParsing -Headers $headers
    }

    $sig = Get-AuthenticodeSignature -FilePath $setup
    $version = (Get-Item -Path $setup).VersionInfo.ProductVersion
    Write-Log "Installer: versie $version | handtekening $($sig.Status) - $($sig.SignerCertificate.Subject)"
    # Gemanipuleerd bestand -> stoppen; andere status alleen melden
    if ($sig.Status -eq 'HashMismatch') { throw 'Digitale handtekening van WiseCloudClientSetup.exe klopt niet (HashMismatch).' }
    if ($sig.Status -ne 'Valid') { Write-Log "LET OP: handtekening is niet 'Valid' maar '$($sig.Status)'" }

    # 2. Stil installeren via de Burn-engine (de bundle installeert alle onderdelen zelf)
    $log = Join-Path -Path $logDir -ChildPath 'WiseCloudClientSetup.log'
    Write-Log "Installeren: WiseCloudClientSetup.exe /quiet /norestart (log: $log)"
    $p = Start-Process -FilePath $setup -ArgumentList @('/quiet', '/norestart', '/log', "`"$log`"") -Wait -PassThru
    # 0 = OK, 3010 = OK maar herstart nodig, 1641 = herstart gestart
    if ($p.ExitCode -notin 0, 3010, 1641) {
        throw "Installatie mislukt met exitcode $($p.ExitCode). Zie $log en de package-logs in $logDir."
    }
    Write-Log "Installatie OK (exitcode $($p.ExitCode))"

    # Controle: welke onderdelen staan er nu
    $products = @(Get-CargoWiseProduct)
    if ($products.Count -eq 0) { throw 'Installatie meldde succes, maar er is geen CargoWise/WiseCloud-product gevonden in het register.' }
    foreach ($pr in $products) { Write-Log "Product: $($pr.DisplayName) $($pr.DisplayVersion) $($pr.InstallLocation)" }

    # Controle: App Manager-service aanwezig (zorgt voor automatische updates zonder adminrechten)
    $svc = @(Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like '*CargoWise*' -or $_.Name -like '*CargoWise*' })
    if ($svc.Count -eq 0) { Write-Log 'LET OP: geen CargoWise-service gevonden' }
    foreach ($s in $svc) { Write-Log "Service: $($s.Name) [$($s.StartType), $($s.Status)]" }

    Remove-Item -Path $workDir -Recurse -Force
    Write-Log 'Klaar'
}
catch {
    Write-Log "FOUT: $($_.Exception.Message)"
    throw
}
