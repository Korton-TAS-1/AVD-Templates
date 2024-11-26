# Get available time zones by name
Get-TimeZone -ListAvailable | Where-Object ({$_.ID -like "*Europe*"})
# Copy the ID of the time zone you get by running Get-TimeZone -ListAvailable | Where-Object ({$_.ID -like "*Europe*"})
# for instance "W. Europe Standard Time" and use the follwing line to set this time zone for the Azure VM 
Set-TimeZone -Id "W. Europe Standard Time"
