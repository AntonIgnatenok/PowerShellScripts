<#
.Synopsis
    Teams chat Export Tool

.DESCRIPTION
   PS script for export Teams chat messages.
#>

### Enter paths for export and User List ###

$exportlocation = "C:\temp\TeamsUserData" # NO Trailing Slash!
$UsersTXT = "C:\temp\TeamsUserData\Users.txt"
$AdminAccount = "Jiri.Jares@eu.lukoil.com"

###########################################

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

Clear-Host

Function Export-TeamsData ($user, $SearchName) {
    $exportexe = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter microsoft.office.client.discovery.unifiedexporttool.exe -Recurse).FullName | Where-Object{ $_ -notmatch "_none_" } | Select-Object -First 1)
    write-host "Start time: $(get-date -Format HH:mm:ss)" -ForegroundColor Green

    if (!(Test-Path $exportexe)) {
        write-host "ERROR! Unified Export Tool not found!" -ForegroundColor Red
        break
    }
    # Connect to security & Compliance
    if(!((get-psSession).ComputerName -like "*ps.compliance.protection.outlook.com")){
        if(!(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement)){
            Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
        }
        Import-Module ExchangeOnlineManagement
        write-host "Enter tenant admin credentials:"
        write-host "Please specify Office 365 Global Admin Credentials"
        write-host
        write-host -ForegroundColor Red "If you have never run this tool before please verify if your account has the following Ediscovery Manager and Compliance Administrator permissions." 
        Write-host
        Write-host -ForegroundColor Red "If your account does not have these privileges, then you can add them by following the following steps:"
        Write-host -ForegroundColor Red "1. Go to https://protection.office.com/permissions"
        Write-host -ForegroundColor Red "2. Add your admin account into the Ediscovery Manager and Compliance Administrator Permissions"
        write-host -ForegroundColor Red "3. After adding your admin into those permissions wait 30 to 40 minutes to be effective"
        write-host
        write-host "Connecting to Exchange Online..."
        Connect-IPPSSession -UserPrincipalName $AdminAccount -WarningAction SilentlyContinue -InformationAction SilentlyContinue
    }

    if(!(get-complianceSearch $searchname -errorAction silentlyContinue)){
        write-host "Create new Compliance Search in " -NoNewline
        write-host $user -ForegroundColor Yellow -NoNewline
        write-host " mailbox."
        $complianceSearch = New-ComplianceSearch -ExchangeLocation $user -Name $SearchName -ContentMatchQuery "kind:microsoftteams AND kind:im" -Description "Teams data export"
        Start-Sleep -s 3
        Start-ComplianceSearch $SearchName
        Start-Sleep -s 5
        $complianceSearch = Get-ComplianceSearch $SearchName
        
        write-host "Compliance Search ($SearchName) status: $($complianceSearch.Status)" -NoNewline
        Do {
            Start-Sleep -s 5
            write-host "." -NoNewline
            $complianceSearch = Get-ComplianceSearch $SearchName
        }
        While(($complianceSearch.Status -ne 'Completed') -or (!(get-complianceSearch $searchName)))
        write-host
    }
    # Microsoft automatically adds the _Export suffix to all exports, so we use that name to run our query.
    $JobName = $SearchName+"_Export"
    if(get-complianceSearchAction -identity $jobname -erroraction SilentlyContinue){  
    }
    else {
        # Create Compliance Search in exportable format. GIVE it the SEARCH name not the JOB name.
        write-host "Create Compliance Search in exportable format."
        New-ComplianceSearchAction -SearchName $SearchName -EnableDedupe $true -Export -Format FxStream -ArchiveFormat PerUserPST | Out-Null
    }
    write-host "Waiting for export to complete." -NoNewline
    Do {
        # Check every 5 seconds that the search has been CREATED.
        Start-Sleep -s 5
        $index = Get-ComplianceSearchAction -Identity $jobname -includeCredential
        $y = $index.Results.split(";")
        $url = $y[0].trimStart("Container url: ")
        $sasKey = $y[1].trimStart(" SAS token: ")
        $estSize = $y[18]
        $transferreditems = $y[21]
        $progress = $y[22]
        write-host "." -NoNewline
    } 
    Until($index.Status -eq 'Completed')
    write-host
    if(($progress -ne "") -and ($estSize -ne "") -and ($transferreditems -ne "")){
        write-host $progress
        write-host $estSize
        write-host $transferreditems
    }
    write-host "Compliance search export status - Completed."
    write-host "Download URL: $url"
    write-host "Download Key: $sasKey"

    # Download the exported PST file
    $traceFileName = $exportlocation+"\"+$JobName+"\Log.txt"
    $errorFileName = $exportlocation+"\"+$JobName+"\Errorlog.txt"
  
    $arguments = "-name `"$jobName`" -source `"$url`" -key `"$sasKey`" -dest `"$exportlocation`" -trace $traceFileName"
    $downLoadProcess = Start-Process -FilePath "$exportexe" -ArgumentList $arguments -Windowstyle Normal -RedirectStandardError $errorFileName -PassThru
    write-host "Start downloading process."
 
    $proc = Get-Process | where -Property name -EQ "microsoft.office.client.discovery.unifiedexporttool"
    Start-Sleep -s 1
    if($proc){
        write-host "Downloading." -NoNewline
        while(Get-Process -Name "microsoft.office.client.discovery.unifiedexporttool" -ErrorAction SilentlyContinue){
            write-host "." -NoNewline
            Start-Sleep -s 5
        }
    }
    write-host
    write-host "Done."    
    write-host "End time: $(get-date -Format HH:mm:ss)" -ForegroundColor Green
    write-host
}

Get-UnifiedExportTool

$Users = Get-Content -Path $UsersTXT

if($Users.count -gt 1){
    $dateString = get-date -Format dd-MM-yyyy_HH-mm
    $SearchName = "Export_TeamsData_for_"+$($Users.count) + "_users_" + $dateString
    $User = $Users
    Export-TeamsData $User $SearchName
}

if($Users.count -eq 1){
    $dateString = get-date -Format dd-MM-yyyy_HH-mm
    $user = $users | Select -First 1
    $username = ($user.Split("@"))[0]
    $SearchName = $username + "_" + $dateString
    Export-TeamsData $User $SearchName
}

if($Users.count -lt 1){
    write-host
    write-host "ERROR! Users not present in file $UsersTXT!" -ForegroundColor Red
    write-host
    break
}