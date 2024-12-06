Stop-Service wuauserv -Force;
Remove-Item "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate" -Recurse -ErrorAction SilentlyContinue -Force
Start-Service wuauserv;
