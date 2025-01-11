# download office deployment tool 
#$ODTDownloadUrl = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'
$ODTDownloadUrl = 'https://go.microsoft.com/fwlink/p/?LinkID=626065'
$ODTDownloadLinkRegex = '/officedeploymenttool[a-z0-9_-]*\.exe$'
$guid = [guid]::NewGuid().Guid
$tempFolder = (Join-Path -Path "C:\temp\" -ChildPath $guid)
$templateFilePathFolder = "C:\AVDImage"

if (!(Test-Path -Path $tempFolder)) {
   New-Item -Path $tempFolder -ItemType Directory
   }
   Write-Host "AVD AIB Customization Office Apps : Created temp folder $tempFolder"

        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        Write-Host "Starting AVD AIB Customization : Office Apps : $((Get-Date).ToUniversalTime())"

        try {
			# Download XML configuration file
			$XMLUri = "https://raw.githubusercontent.com/Korton-TAS-1/AVD-Templates/refs/heads/main/Packages/Office-C2R/XML/AVD-Configuration.xml"
			$XMLContent = Invoke-WebRequest -Uri $XMLUri -UseBasicParsing
			
			if ($XMLContent.StatusCode -ne 200) {
				throw "Error: Unable to access Configuration XML file. Status Code: $($XMLContent.StatusCode), Description: $($XMLContent.StatusDescription)"
			}

			# Download Office Deployment Tool
			$HttpContent = Invoke-WebRequest -Uri $ODTDownloadUrl -UseBasicParsing
			
			if ($HttpContent.StatusCode -ne 200) {
				throw "Error: Failed to retrieve Office Deployment Tool link. Status Code: $($HttpContent.StatusCode), Description: $($HttpContent.StatusDescription)"
			}

			$ODTDownloadLinks = $HttpContent.Links | Where-Object { $_.href -match $ODTDownloadLinkRegex }
			$ODTToolLink = $ODTDownloadLinks[0].href
			Write-Host "Office Deployment Tool link retrieved: $ODTToolLink"

			$ODTexePath = Join-Path -Path $tempFolder -ChildPath "officedeploymenttool.exe"

			# Download ODT tool
			Write-Host "Downloading Office Deployment Tool to $ODTexePath"
			$ODTResponse = Invoke-WebRequest -Uri $ODTToolLink -UseBasicParsing -OutFile $ODTexePath -PassThru

			if ($ODTResponse.StatusCode -ne 200) {
				throw "Error: Failed to download Office Deployment Tool. Status Code: $($ODTResponse.StatusCode), Description: $($ODTResponse.StatusDescription)"
			}

			# Extract setup.exe
			Write-Host "Extracting Office Deployment Tool to $tempFolder"
			Start-Process -FilePath $ODTexePath -ArgumentList "/extract:`"$($tempFolder)`" /quiet" -Wait -NoNewWindow -ErrorAction Stop

			$setupExePath = Join-Path -Path $tempFolder -ChildPath 'setup.exe'
			$xmlFilePath = Join-Path -Path $tempFolder -ChildPath 'installOffice.xml'

			# Save XML content
			Write-Host "Saving XML configuration to $xmlFilePath"
			$XMLContent.Content | Out-File -FilePath $xmlFilePath -Force -Encoding Ascii

			# Download Office using setup.exe
			Write-Host "Downloading Office components using $setupExePath"
			$ODTRunSetupExe = Start-Process -FilePath $setupExePath -ArgumentList "/download $(Split-Path -Path $xmlFilePath -Leaf)" -Wait -PassThru -WorkingDirectory $tempFolder -WindowStyle Hidden

			if ($ODTRunSetupExe.ExitCode -ne 0) {
				throw "Error: Setup.exe returned exit code $($ODTRunSetupExe.ExitCode) while downloading Office components."
			}

			# Install Office
			Write-Host "Installing Office using $setupExePath"
			$InstallOffice = Start-Process -FilePath $setupExePath -ArgumentList "/configure $(Split-Path -Path $xmlFilePath -Leaf)" -Wait -PassThru -WorkingDirectory $tempFolder -WindowStyle Hidden

			if ($InstallOffice.ExitCode -ne 0) {
				throw "Error: Setup.exe returned exit code $($InstallOffice.ExitCode) while installing Office."
			}

			Write-Host "Office installation completed successfully."

		} catch {
			Write-Error "An error occurred: $($_.Exception.Message)"
			$PSCmdlet.ThrowTerminatingError($_)
		}
        #Cleanup
        if ((Test-Path -Path $tempFolder -ErrorAction SilentlyContinue)) {
            #Remove-Item -Path $tempFolder -Force -Recurse -ErrorAction Continue
        }

        if ((Test-Path -Path $templateFilePathFolder -ErrorAction SilentlyContinue)) {
            #Remove-Item -Path $templateFilePathFolder -Force -Recurse -ErrorAction Continue
        }

        $stopwatch.Stop()
        $elapsedTime = $stopwatch.Elapsed
        Write-Host "Ending AVD AIB Customization : Office Apps - Time taken: $elapsedTime"

