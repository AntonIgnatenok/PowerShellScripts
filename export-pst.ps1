Function Get-UnifiedExportTool {
    
    # Check if the export tool is installed for the user, and download if not.
    While(-Not ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter microsoft.office.client.discovery.unifiedexporttool.exe -Recurse).FullName | Where-Object{ $_ -notmatch "_none_" } | Select-Object -First 1)){
        Write-Host "Downloading Unified Export Tool."
        Write-Host "This is installed per-user by the Click-Once installer."
        
        $Manifest = "https://complianceclientsdf.blob.core.windows.net/v16/Microsoft.Office.Client.Discovery.UnifiedExportTool.application"
        $ElevatePermissions = $true
        Try {
            Add-Type -AssemblyName System.Deployment
            Write-Host "Starting installation of ClickOnce Application $Manifest "
            $RemoteURI = [URI]::New( $Manifest , [UriKind]::Absolute)
            if (-not  $Manifest){
                throw "Invalid ConnectionUri parameter '$ConnectionUri'"
            }
            $HostingManager = New-Object System.Deployment.Application.InPlaceHostingManager -ArgumentList $RemoteURI , $False
            Register-ObjectEvent -InputObject $HostingManager -EventName GetManifestCompleted -Action { 
                new-event -SourceIdentifier "ManifestDownloadComplete"
            } | Out-Null
            Register-ObjectEvent -InputObject $HostingManager -EventName DownloadApplicationCompleted -Action { 
                new-event -SourceIdentifier "DownloadApplicationCompleted"
            } | Out-Null
            $HostingManager.GetManifestAsync()
            $event = Wait-Event -SourceIdentifier "ManifestDownloadComplete" -Timeout 15
            if($event) {
                $event | Remove-Event
                Write-Host "ClickOnce Manifest Download Completed"
                $HostingManager.AssertApplicationRequirements($ElevatePermissions)
                $HostingManager.DownloadApplicationAsync()
                $event = Wait-Event -SourceIdentifier "DownloadApplicationCompleted" -Timeout 60
                if($event) {
                    $event | Remove-Event
                    Write-Host "ClickOnce Application Download Completed"
                }
                else {
                    Write-error "ClickOnce Application Download did not complete in time (60s)"
                }
            }
            else {
                Write-error "ClickOnce Manifest Download did not complete in time (15s)"
            }
        }
        Finally {
            Get-EventSubscriber | ?{$_.SourceObject.ToString() -eq 'System.Deployment.Application.InPlaceHostingManager'} | Unregister-Event
        }
    }
}


Function Export-TeamsData ($user, $msCred) {
    
    $exportexe = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter microsoft.office.client.discovery.unifiedexporttool.exe -Recurse).FullName | Where-Object{ $_ -notmatch "_none_" } | Select-Object -First 1)
    $dateString = get-date -Format dd-MM-yyyy_HH-mm
    $username = ($user.Split("@"))[0]
    $SearchName = $username + "_TEAMS_" + $dateString
   
    if (!(Test-Path $exportexe)) {
        write-host "ERROR! Downloader Not found!" -ForegroundColor Red
    }

    # Connect to security & Compliance
    if(!$(get-psSession).Name -like "ExchangeOnline*"){
        Import-Module ExchangeOnlineManagement
        write-host "Connect to Exchange Online" -ForegroundColor green
        Connect-IPPSSession -Credential $msCred -WarningAction SilentlyContinue
    }

    #$Session2search = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $msCred -Authentication Basic -AllowRedirection
    #Import-PSSession $Session2search -AllowClobber -DisableNameChecking

    if ($(get-complianceSearch $searchname -errorAction silentlyContinue)) {
    }
    else {
        write-host "Create new Compliance Search in $user mailbox"
        $complianceSearch = New-ComplianceSearch -ExchangeLocation $user -Name $SearchName -ContentMatchQuery "kind:microsoftteams AND kind:im" -Description "Teams data export"
        Start-Sleep -s 3
        Start-ComplianceSearch $SearchName
        Do {
            # You'll see red on the screen if you don't wait a bit for O365 to actually create the search.
            Start-Sleep -s 10
            
            $complianceSearch = Get-ComplianceSearch $SearchName
            write-host "Compliance Search status: $($complianceSearch.Status)"
        # Added the -or to check if it actually EXISTS yet.
        }
        While (($complianceSearch.Status -ne 'Completed') -or (!(get-complianceSearch $searchName)))
    }
    # Microsoft automatically adds the _Export suffix to all exports, so we use that name to run our query.
    $JobName = $SearchName+"_Export"
    if (get-complianceSearchAction -identity $jobname -erroraction SilentlyContinue) {
        
    }
    else {
        # Create Compliance Search in exportable format. GIVE it the SEARCH name not the JOB name.
        write-host "Create Compliance Search in exportable format."
        New-ComplianceSearchAction -SearchName $SearchName -EnableDedupe $true -Export -Format FxStream -ArchiveFormat PerUserPST
    }
     Do {
        # Check every 5 seconds that the search has been CREATED.
        Start-Sleep -s 5
        #$index = Get-ComplianceSearchAction -Identity "kozinsky.d_TEAMS_20-12-2021_03-59_Export" -includeCredential
        $index = Get-ComplianceSearchAction -Identity $jobname -includeCredential
        $y=$index.Results.split(";")
        $url = $y[0].trimStart("Container url: ")
        $sasKey = $y[1].trimStart(" SAS token: ")
        $estSize = $y[18]
        $transferreditems = $y[21]
        $progress = $y[22]
        # These dont appear to be populated yet by the time we try to read them.
        if($estSize) {
            write-host "$estSize"
            write-host "$transferreditems"
        }
        else {
            write-host "$progress"
        }
    } 
    Until($index.Status -eq 'Completed')
    write-host "Compliance search action status - Completed."
    write-host "Download URL: $url"
    write-host "Download Key: $sasKey"

    # Download the exported files from Office 365
    $traceFileName = $exportlocation+"\"+$SearchName+"\"+$user+".log"
    $errorFilename = $exportlocation+"\"+$SearchName+"\Errorlog_"+$user+".txt"
  
    $arguments = "-name `"$jobName`" -source `"$url`" -key `"$sasKey`" -dest `"$exportlocation`" -trace $traceFileName"
    $downLoadProcess = Start-Process -FilePath "$exportexe" -ArgumentList $arguments -Windowstyle Normal -RedirectStandardError $errorfilename -PassThru
    write-host "Start downloading process."
    While(Get-Process -Name "microsoft.office.client.discovery.unifiedexporttool" -ErrorAction SilentlyContinue){
        
        # Get export details.
        $SearchActionStatusDetails = Get-ComplianceSearchAction -Identity $jobname -IncludeCredential -Details;
        $SearchActionStatusDetails = $SearchActionStatusDetails.Results.split(";");
        #$ExportEstSize = [double]::Parse(((($SearchActionStatusDetails[18].TrimStart(" Total estimated bytes: ")))), [cultureinfo] 'en-US');
        $ExportProgress = $SearchActionStatusDetails[22].TrimStart(" Progress: ").TrimEnd("%");
        $ExportStatus = $SearchActionStatusDetails[25].TrimStart(" Export status: ");
        $ExportProgress
        # Get download content.
        $Downloaded = Get-ChildItem -Path ("{0}" -f $exportlocation) -Recurse | Measure-Object -Property Length -Sum | Select-Object -ExpandProperty Sum;

        # Get procent downloaded.
        #$ProcentDownloaded = ($Downloaded/$ExportEstSize*100);

        # Start sleep.
        Start-Sleep -Seconds 5;
    }    
}

### Begin Main Code ######################################################################

$exportlocation = "C:\Temp" # NO Trailing Slash!
$user = "user@csitltd.ru"
$msCred = Get-Credential

Get-UnifiedExportTool

Export-TeamsData $user $msCred