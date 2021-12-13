function Get-MigrationBatchStatus {
  param (
    $BatchID
  )
  Check-LoadedModule "ExchangeOnlineManagement"
  get-migrationuser -BatchId $BatchID | sort DataConsistencyScore | select Identity,Status,ErrorSummary,DataConsistencyScore,HasUnapprovedSkippedItems
}

function Get-SkippedItems {
  param (
    $BatchID
  )
  Check-LoadedModule "ExchangeOnlineManagement"
  get-migrationuser -BatchId $BatchID | 
  ?{ $_.HasUnapprovedSkippedItems -eq $True } | 
  Get-MigrationUserStatistics -IncludeSkippedItems | 
  select -Expand SkippedItems @{label="UserIdentity";expression={$_.Identity}} | 
  ? {$_.Kind -ne "CorruptFolderACL" } | 
  select @{label="Identity";expression={$_.UserIdentity}},Kind,FolderName,Subject,DateReceived,@{label="MessageSizeMB";expression={$_.MessageSize/1024/1024}}
}


####### Skipped items
$userStats = Get-MigrationUserStatistics -Identity user@fabrikaminc.net -IncludeSkippedItems
$userStats.SkippedItems | ft -a Subject, Sender, DateSent, ScoringClassifications

##############################

$mu = get-migrationuser -resultsize unlimited | select *
$count = ($mu | measure).count
$res = @()
$c = 0
foreach ($u in $mu){
    Write-Progress -Activity "Processing users" -Status "processing user $($u.identity)" -PercentComplete ($c/$count*100)

    
    $ms = $u | get-migrationuserstatistics -includeskippeditems | select *
    foreach ($si in $ms.skippeditems)    {
    $obj = new-object psobject
    $obj | add-member -MemberType NoteProperty -Name Identity -Value $U.identity
    $obj | add-member -MemberType NoteProperty -Name FolderName -Value $si.foldername
    $obj | add-member -MemberType NoteProperty -Name Kind -Value $si.kind
    $obj | add-member -MemberType NoteProperty -Name Failure -Value $si.failure
    $res += $obj
    }
    $c +=1 
}
$res | export-csv -path "C:\Temp\skippeditems.csv" -NoTypeInformation

################

# Remove existing ADAL identities if Outlook doesn't prompt
Get-ChildItem -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Identity\Identities" | Remove-Item cmdkey /list | ForEach-Object{if($_ -like "*Target:*"){cmdkey /del:($_ -replace " ","" -replace "Target:","")}}

#Install Azure AD Directory Broker Plugin
if (-not (Get-AppxPackage Microsoft.AAD.BrokerPlugin)) { Add-AppxPackage -Register "$env:windir\SystemApps\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\Appxmanifest.xml" -DisableDevelopmentMode -ForceApplicationShutdown } Get-AppxPackage Microsoft.AAD.BrokerPlugin
