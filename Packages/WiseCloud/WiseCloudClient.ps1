<#
.SYNOPSIS
    Installeert de CargoWise Cloud Accessor Client en CargoWise Remote Desktop Services per machine, voor AVD-images.
.DESCRIPTION
    Bedoeld voor multi-session hosts (AVD/RDS) waar gewone gebruikers geen software in hun profiel
    mogen installeren. Werkwijze volgens de WiseTech-advisory "WCA and RDS migration" (aug. 2026):

    1. Registerwaarde HKLM\SOFTWARE\WiseTech Global\WiseCloud\PerMachineMigrate = 1 zetten.
       Zonder deze waarde vervangt een automatische upgrade de per-machine installatie door een
       installatie per gebruiker (in %LOCALAPPDATA%), wat op een gedeelde host voor niet-admins faalt.
    2. Beide MSI's ophalen
       - uit Azure Blob Storage met een managed identity (-ClientId), of
       - van een directe URL zonder login (laat -ClientId leeg), of
       - van een lokaal pad (-WcaMsiPath/-RdsMsiPath, voor testen)
    3. Stil per machine installeren:
       msiexec /i CargoWiseCloudAccessorClient.msi INSTALLSCOPE=PerMachine /qn
       msiexec /i CargoWise.RemoteDesktopServices.msi ALLUSERS=1 MSIINSTALLPERUSER="" /quiet /norestart
    4. Snelkoppeling op het openbare bureaublad als die ontbreekt.

    Updates: een per-machine installatie krijgt volgens WiseTech geen automatische updates meer.
    Nieuwe versie = nieuwe MSI's in de container zetten en een nieuw image bouwen.

    De MSI's staan NIET in deze (publieke) repo: WiseTech biedt ze alleen achter een login aan
    (My Account > Client Downloads). Zet ze in een prive storage account.

    In Azure Image Builder (scriptUri kan geen parameters meegeven), dus:
      - File-stap:       sourceUri = raw-URL van dit script, destination = C:\AVDImage\WiseCloudClient.ps1
      - PowerShell-stap: inline = C:\AVDImage\WiseCloudClient.ps1 -WcaMsiUrl '<url>' -RdsMsiUrl '<url>' -ClientId '<id>'
.PARAMETER WcaMsiUrl
    URL van CargoWiseCloudAccessorClient.msi (blob-URL of directe downloadlink).
.PARAMETER RdsMsiUrl
    URL van CargoWise.RemoteDesktopServices.msi.
.PARAMETER ClientId
    Client ID van de user-assigned managed identity op de build-VM, met Storage Blob Data
    Reader op de container. Leeg = downloaden zonder authenticatie.
.PARAMETER WcaMsiPath
    Lokaal pad naar CargoWiseCloudAccessorClient.msi (voor testen).
.PARAMETER RdsMsiPath
    Lokaal pad naar CargoWise.RemoteDesktopServices.msi (voor testen).
.EXAMPLE
    # In Image Builder
    C:\AVDImage\WiseCloudClient.ps1 -WcaMsiUrl 'https://<account>.blob.core.windows.net/cargowise/CargoWiseCloudAccessorClient.msi' -RdsMsiUrl 'https://<account>.blob.core.windows.net/cargowise/CargoWise.RemoteDesktopServices.msi' -ClientId '<client-id>'
.EXAMPLE
    # Lokaal testen, als administrator
    .\WiseCloudClient.ps1 -WcaMsiPath 'C:\Temp\CargoWiseCloudAccessorClient.msi' -RdsMsiPath 'C:\Temp\CargoWise.RemoteDesktopServices.msi'
.NOTES
    Versie 2.0 - 25-11-2026 - per machine via MSI's + PerMachineMigrate (advisory WCA and RDS migration);
                              vervangt WiseCloudClientSetup.exe, die bij de eerste start naar per gebruiker migreerde
    Versie 1.1 - 23-11-2026 - Burn-bundle: /quiet in plaats van /extract (installer 6.5.7517)
    Versie 1.0 - 23-11-2026 - eerste versie (Slade, ticket 86157695)
#>
[CmdletBinding(DefaultParameterSetName = 'Url')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Url')]
    [string]$WcaMsiUrl,

    [Parameter(Mandatory, ParameterSetName = 'Url')]
    [string]$RdsMsiUrl,

    [Parameter(ParameterSetName = 'Url')]
    [string]$ClientId,

    [Parameter(Mandatory, ParameterSetName = 'Local')]
    [string]$WcaMsiPath,

    [Parameter(Mandatory, ParameterSetName = 'Local')]
    [string]$RdsMsiPath
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$workDir = 'C:\AVDImage\WiseCloud'
$logDir  = 'C:\Windows\Temp\WiseCloud'
$WISECLOUD_REG_PATH = 'HKLM:\SOFTWARE\WiseTech Global\WiseCloud'

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

function Get-StorageHeader {
    <#
    .SYNOPSIS
        Haalt via IMDS een storage-token op met de user-assigned identity van de (build-)VM.
    #>
    param([string]$Id)
    $imds = "http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=https%3A%2F%2Fstorage.azure.com%2F&client_id=$Id"
    $token = $null
    for ($i = 1; $i -le 10 -and -not $token; $i++) {
        try { $token = (Invoke-RestMethod -Uri $imds -Headers @{ Metadata = 'true' } -TimeoutSec 30).access_token }
        catch { Write-Log "Token nog niet beschikbaar (poging $i): $($_.Exception.Message)"; Start-Sleep -Seconds 15 }
    }
    if (-not $token) { throw 'Geen token van IMDS ontvangen.' }
    @{ Authorization = "Bearer $token"; 'x-ms-version' = '2021-08-06' }
}

function Test-InstallerSignature {
    <#
    .SYNOPSIS
        Stopt bij een gemanipuleerd bestand; een andere status dan Valid wordt alleen gemeld.
    #>
    param([string]$Path)
    $sig = Get-AuthenticodeSignature -FilePath $Path
    Write-Log "$(Split-Path -Path $Path -Leaf): handtekening $($sig.Status) - $($sig.SignerCertificate.Subject)"
    if ($sig.Status -eq 'HashMismatch') { throw "Digitale handtekening van $Path klopt niet (HashMismatch)." }
    if ($sig.Status -ne 'Valid') { Write-Log "LET OP: handtekening is niet 'Valid' maar '$($sig.Status)'" }
}

function Invoke-MsiInstall {
    <#
    .SYNOPSIS
        Voert msiexec /i uit met extra argumenten en een uitgebreide log; stopt bij een foutcode.
    #>
    param(
        [string]$MsiPath,
        [string[]]$Arguments,
        [string]$LogName
    )
    $log = Join-Path -Path $logDir -ChildPath $LogName
    $argList = @('/i', "`"$MsiPath`"") + $Arguments + @('/l*v', "`"$log`"")
    Write-Log "msiexec $($argList -join ' ')"
    $p = Start-Process -FilePath 'msiexec.exe' -ArgumentList $argList -Wait -PassThru
    # 0 = OK, 3010 = OK maar herstart nodig, 1641 = herstart gestart
    if ($p.ExitCode -notin 0, 3010, 1641) {
        throw "Installatie van $(Split-Path -Path $MsiPath -Leaf) mislukt met exitcode $($p.ExitCode). Zie $log."
    }
    Write-Log "OK (exitcode $($p.ExitCode))"
}

try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) { throw 'Dit script moet als administrator draaien.' }

    # Schone werkmap, zodat een eerdere run de controles niet beinvloedt
    if (Test-Path -Path $workDir) { Remove-Item -Path $workDir -Recurse -Force }
    foreach ($d in $workDir, $logDir) { $null = New-Item -ItemType Directory -Path $d -Force }
    $wcaMsi = Join-Path -Path $workDir -ChildPath 'CargoWiseCloudAccessorClient.msi'
    $rdsMsi = Join-Path -Path $workDir -ChildPath 'CargoWise.RemoteDesktopServices.msi'

    # 1. Per machine vastzetten VOOR de installatie, zodat de eerste upgrade niet naar per gebruiker migreert
    Write-Log "Registerwaarde PerMachineMigrate = 1 in $WISECLOUD_REG_PATH"
    $null = New-Item -Path $WISECLOUD_REG_PATH -Force
    $null = New-ItemProperty -Path $WISECLOUD_REG_PATH -Name 'PerMachineMigrate' -Value 1 -PropertyType DWord -Force

    # 2. MSI's ophalen
    if ($PSCmdlet.ParameterSetName -eq 'Local') {
        foreach ($pair in @(@($WcaMsiPath, $wcaMsi), @($RdsMsiPath, $rdsMsi))) {
            if (-not (Test-Path -Path $pair[0])) { throw "MSI niet gevonden: $($pair[0])" }
            Write-Log "Lokale MSI: $($pair[0])"
            Copy-Item -Path $pair[0] -Destination $pair[1]
        }
    }
    else {
        $headers = @{}
        if ($ClientId) {
            Write-Log 'Token ophalen voor Azure Storage via managed identity'
            $headers = Get-StorageHeader -Id $ClientId
        }
        foreach ($pair in @(@($WcaMsiUrl, $wcaMsi), @($RdsMsiUrl, $rdsMsi))) {
            # Alleen het adres loggen, niet een eventuele SAS-token in de query string
            Write-Log "Downloaden: $(($pair[0] -split '\?')[0])"
            Invoke-WebRequest -Uri $pair[0] -OutFile $pair[1] -UseBasicParsing -Headers $headers
        }
    }
    Test-InstallerSignature -Path $wcaMsi
    Test-InstallerSignature -Path $rdsMsi

    # 3. Per machine installeren (commando's volgens de WiseTech-advisory)
    Invoke-MsiInstall -MsiPath $wcaMsi -LogName 'CloudAccessorClient-install.log' -Arguments @('INSTALLSCOPE=PerMachine', '/qn', '/norestart')
    Invoke-MsiInstall -MsiPath $rdsMsi -LogName 'RemoteDesktopServices-install.log' -Arguments @('ALLUSERS=1', 'MSIINSTALLPERUSER=""', '/quiet', '/norestart')

    # Controle: volgens WiseTech is de installatiemap de enige betrouwbare indicator (niet het register)
    $programDirs = @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { $_ } |
        ForEach-Object { Join-Path -Path $_ -ChildPath 'WiseTech Global' } | Where-Object { Test-Path -Path $_ }
    $clientExe = @($programDirs | ForEach-Object { Get-ChildItem -Path $_ -Filter 'WiseCloudClient.exe' -Recurse -File -ErrorAction SilentlyContinue }) |
        Select-Object -First 1
    if (-not $clientExe) {
        throw "Installatie meldde succes, maar WiseCloudClient.exe staat niet onder Program Files\WiseTech Global. Zie de logs in $logDir."
    }
    Write-Log "Client per machine: $($clientExe.FullName) (versie $($clientExe.VersionInfo.ProductVersion))"
    foreach ($pr in @(Get-CargoWiseProduct)) { Write-Log "Product: $($pr.DisplayName) $($pr.DisplayVersion) $($pr.InstallLocation)" }

    # 4. Snelkoppeling voor alle gebruikers, als de MSI er zelf geen op het openbare bureaublad zet
    $publicDesktop = Join-Path -Path $env:PUBLIC -ChildPath 'Desktop'
    $commonStart   = Join-Path -Path $env:ProgramData -ChildPath 'Microsoft\Windows\Start Menu\Programs'
    $shell = New-Object -ComObject WScript.Shell
    $existing = @(foreach ($dir in $publicDesktop, $commonStart) {
        Get-ChildItem -Path $dir -Filter '*.lnk' -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $shell.CreateShortcut($_.FullName).TargetPath -eq $clientExe.FullName }
    })
    foreach ($lnk in $existing) { Write-Log "Snelkoppeling aanwezig: $($lnk.FullName)" }
    if (-not ($existing | Where-Object { $_.FullName -like "$publicDesktop*" })) {
        $lnkPath = Join-Path -Path $publicDesktop -ChildPath 'CargoWise.lnk'
        $lnk = $shell.CreateShortcut($lnkPath)
        $lnk.TargetPath = $clientExe.FullName
        $lnk.WorkingDirectory = $clientExe.DirectoryName
        $lnk.IconLocation = "$($clientExe.FullName),0"
        $lnk.Save()
        Write-Log "Snelkoppeling aangemaakt: $lnkPath"
    }

    Remove-Item -Path $workDir -Recurse -Force
    Write-Log 'Klaar'
}
catch {
    Write-Log "FOUT: $($_.Exception.Message)"
    throw
}
