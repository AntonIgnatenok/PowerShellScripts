Import-Module ActiveDirectory
$Users=Get-ADUser -Filter *

function guidtobase64
{
    param($str);
    $g = new-object -TypeName System.Guid -ArgumentList $str;
    $b64 = [System.Convert]::ToBase64String($g.ToByteArray());
    return $b64;
}

$Users | Select UserPrincipalName,@{Expression={(guidtobase64($_.ObjectGUID))}; Label="ImmutableID"} #| Export-Csv -Path "C:\temp\UPNs-and-ImmutableIDs.csv" -Delimiter ';' -NoTypeInformation

Import-Module MSOnline
Connect-MsolService
foreach($User in $Users){
    try{
      $UPN = $User.UserPrincipalName
      $ImmutableID = $User.ImmutableID
      Get-MsolUser -UserPrincipalName $UPN | Set-MsolUser -ImmutableId $ImmutableId
    }
    catch{
      Write-Output $_.Exception.Message
    }
}
