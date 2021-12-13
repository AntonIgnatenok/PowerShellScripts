$batchid = "GmailBatch1"

$migrationmbs = Get-MigrationUser -resultsize unlimited -batchid "$batchid" | Get-MigrationUserStatistics

$transferred = [math]::Round(($migrationmbs | Measure-Object BytesTransferred -sum).sum /1gb,2)
$totalmbs = $migrationmbs.count
$movedmbs = $migrationmbs | ?{$_.status -eq "completed" -or $_.status -eq "synced"}
$failedmbs = $migrationmbs | ?{$_.status -eq "failed"}
$syncedmbs = $migrationmbs | ?{$_.status -eq "synced"}
$movingmbs = $migrationmbs | ?{$_.status -ne "completed" -and $_.status -ne "synced" -and $_.status -ne "failed"}
$movedmbscount = $movedmbs.count
$failedmbscount = $failedmbs.count
$syncedmbscount = $syncedmbs.count
$movingmbscount = $movingmbs.count 
$movingtotaltransfer = [math]::Round(($movingmbs | Measure-Object EstimatedTotalTransferSize -sum).sum /1gb,2)
$movingtransferred = [math]::Round(($movingmbs | Measure-Object BytesTransferred -sum).sum /1gb,2)
$lefttotransfer = [math]::Round($movingtotaltransfer - $movingtransferred,2)
$totaltotransfer = $transferred + $lefttotransfer
$percentcomplete = [math]::Round($transferred/$totaltotransfer *100)

If ($movingmbscount -gt 0)
{
Write-output "Migration batch ""$batchid"" contains a total of $totalmbs mailboxes being migrated."
Write-output "The total data size of the migration batch is estimated to be at least $totaltotransfer GB."
Write-output "$transferred GB have been processed thusfar, accounting for $percentcomplete percent of the data to be transferred."
Write-output "$movedmbscount Mailboxes have been successully Synced or Completed"
Write-output "$movingmbscount mailboxes are still being moved, with an estimated $lefttotransfer GB left to process"
}
else 
{
Write-output "Migration batch ""$batchid"" has completed or fully synced."
Write-output "The total data size of the migration batch was $transferred GB"
Write-output "$failedmbscount mailboxes failed to successfully be migrated."
Write-output "$syncedmbscount mailboxes are in a Synced state and need to be finalized."
}
