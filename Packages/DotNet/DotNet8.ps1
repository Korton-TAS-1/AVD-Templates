# Download WindowsApp from Microsoft
$myDownloadUrl="https://go.microsoft.com/fwlink/?linkid=2262633"

mkdir c:\temp -f
Invoke-WebRequest $myDownloadUrl -OutFile c:\temp\WindowsApp_x64.msix

# Install WindowsApp for all users
Add-AppProvisionedPackage -online -packagepath "c:\temp\WindowsApp_x64.msix" -skiplicense
