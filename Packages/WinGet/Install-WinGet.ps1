.\winget install -s msstore --id 9NBLGGH4NNS1 --silent --accept-package-agreements --accept-source-agreements --force
shutdown -r -t 0

<#https://learn.microsoft.com/en-us/windows/package-manager/winget/
$progressPreference = 'silentlyContinue'
Write-Information "Downloading WinGet and its dependencies..."
Invoke-WebRequest -Uri https://aka.ms/getwinget -OutFile Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle
Invoke-WebRequest -Uri https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx -OutFile Microsoft.VCLibs.x64.14.00.Desktop.appx
Invoke-WebRequest -Uri https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx -OutFile Microsoft.UI.Xaml.2.8.x64.appx
Add-AppxProvisionedPackage -Online Microsoft.VCLibs.x64.14.00.Desktop.appx -SkipLicense 
Add-AppxProvisionedPackage -Online Microsoft.UI.Xaml.2.8.x64.appx -SkipLicense 
Add-AppxProvisionedPackage -Online Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle -SkipLicense 

Add-AppxPackage Microsoft.VCLibs.x64.14.00.Desktop.appx
Add-AppxPackage Microsoft.UI.Xaml.2.8.x64.appx
Add-AppxPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle

&([ScriptBlock]::Create((irm asheroto.com/winget))) -Force
sleep 10
Invoke-WebRequest -Uri https://cdn.winget.microsoft.com/cache/source.msix -OutFile $env:TEMP\source.msix
Add-AppProvisionedPackage -online -packagepath $env:TEMP\source.msix -skiplicense
#>
