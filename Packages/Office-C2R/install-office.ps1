# download office deployment tool 
$ODTDownloadUrl = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'
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
            
            $XMLContent = Invoke-WebRequest -Uri https://raw.githubusercontent.com/Korton-TAS-1/AVD-Templates/refs/heads/main/Packages/Office-C2R/XML/AVD-Configuration.xml -UseBasicParsing

            if ($XMLContent.StatusCode -ne 200) { 
                throw "Could not access the Configuration XML file  $($XMLContent.StatusCode) ($($XMLContent.StatusDescription))"
            }

         
            $HttpContent = Invoke-WebRequest -Uri $ODTDownloadUrl -UseBasicParsing
            
            if ($HttpContent.StatusCode -ne 200) { 
                throw "Office Installation script failed to find Office deployment tool link -- Response $($HttpContent.StatusCode) ($($HttpContent.StatusDescription))"
            }

            $ODTDownloadLinks = $HttpContent.Links | Where-Object { $_.href -Match $ODTDownloadLinkRegex }

            #pick the first link in case there are multiple
            $ODTToolLink = $ODTDownloadLinks[0].href
            Write-Host "AVD AIB Customization Office Apps : Office deployment tool link is $ODTToolLink"

            $ODTexePath = Join-Path -Path $tempFolder -ChildPath "officedeploymenttool.exe"

            #download office deployment tool

            Write-Host "AVD AIB Customization Office Apps : Downloading ODT tool into folder $ODTexePath"
            $ODTResponse = Invoke-WebRequest -Uri "$ODTToolLink" -UseBasicParsing -UseDefaultCredentials -OutFile $ODTexePath -PassThru

            if ($ODTResponse.StatusCode -ne 200) { 
                throw "Office Installation script failed to download Office deployment tool -- Response $($ODTResponse.StatusCode) ($($ODTResponse.StatusDescription))"
            }

            Write-Host "AVD AIB Customization Office Apps : Extracting setup.exe into $tempFolder"
            # extract setup.exe
            Start-Process -FilePath $ODTexePath -ArgumentList "/extract:`"$($tempFolder)`" /quiet" -PassThru -Wait -NoNewWindow

            $setupExePath = Join-Path -Path $tempFolder -ChildPath 'setup.exe'
            
            # Construct XML config file for Office Deployment Kit setup.exe
            $xmlFilePath = Join-Path -Path $tempFolder -ChildPath 'installOffice.xml'

            Write-Host "AVD AIB Customization Office Apps : Saving xml content into xml file : $xmlFilePath"
            
            $XMLContent.Content | Out-File -FilePath $xmlFilePath -Force -Encoding ascii
            
            [XML]$file = Get-Content $xmlFilePath
            
            Write-Host "AVD AIB Customization Office Apps : Running setup.exe to download Office"
            $ODTRunSetupExe = Start-Process -FilePath $setupExePath -ArgumentList "/download $(Split-Path -Path $xmlFilePath -Leaf)" -PassThru -Wait -WorkingDirectory $tempFolder -WindowStyle Hidden

            if (!$ODTRunSetupExe) {
                Throw "AVD AIB Customization Office Apps : Failed to run `"$setupExePath`" to download Office"
            }

            if ( $ODTRunSetupExe.ExitCode) {
                Throw "AVD AIB Customization Office Apps : Exit code $($ODTRunSetupExe.ExitCode) returned from `"$setupExePath`" to download Office"
            }

            Write-Host "AVD AIB Customization Office Apps : Running setup.exe to Install Office"
            $InstallOffice = Start-Process -FilePath $setupExePath -ArgumentList "/configure $(Split-Path -Path $xmlFilePath -Leaf)" -PassThru -Wait -WorkingDirectory $tempFolder -WindowStyle Hidden

            if (!$InstallOffice) {
                Throw "AVD AIB Customization Office Apps : Failed to run `"$setupExePath`" to install Office"
            }

            if ( $ODTRunSetupExe.ExitCode ) {
                Throw "AVD AIB Customization Office Apps : Exit code $($ODTRunSetupExe.ExitCode) returned from `"$setupExePath`" to download Office"
            }
            
        }
        catch {
            $PSCmdlet.ThrowTerminatingError($PSitem)
        }



        #Cleanup
        if ((Test-Path -Path $tempFolder -ErrorAction SilentlyContinue)) {
            Remove-Item -Path $tempFolder -Force -Recurse -ErrorAction Continue
        }

        if ((Test-Path -Path $templateFilePathFolder -ErrorAction SilentlyContinue)) {
            Remove-Item -Path $templateFilePathFolder -Force -Recurse -ErrorAction Continue
        }

        $stopwatch.Stop()
        $elapsedTime = $stopwatch.Elapsed
        Write-Host "Ending AVD AIB Customization : Office Apps - Time taken: $elapsedTime"



