# https://nl.visma.com/accountview/support/technisch/meldingen/meldingen-in-accountview-door-netwerkstoringen/
Write-Host "Starting AccountView Optimisation: "
# reg key Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters\CacheFileTimeout
# Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters' 
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters'  -Name CacheFileTimeout -Value 0 -PropertyType DWORD -Force 

# set SmbClientConfiguration 
Set-SmbClientConfiguration -DirectoryCacheLifetime 0 -Confirm:$false
Set-SmbClientConfiguration -FileInfoCacheLifetime 0 -Confirm:$false
Set-SmbClientConfiguration -FileNotFoundCacheLifetime 0 -Confirm:$false
