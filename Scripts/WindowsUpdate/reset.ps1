Stop-Service wuauserv -Force;
Remove-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Recurse -Force;
Start-Service wuauserv;
