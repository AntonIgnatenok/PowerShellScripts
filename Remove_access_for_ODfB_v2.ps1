Clear-Host

#Set Runtime Parameters
$AdminSiteURL  = "https://contoso-admin.sharepoint.com"
$SiteCollAdmin = "SiteCollAdmin@contoso.com"

#Check module SharePoint Online
$spm = Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable
if(!$spm){
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell
}
 
#Get Credentials to connect to the SharePoint Admin Center
$Cred = Get-Credential
 
#Connect to SharePoint Online Admin Center
Connect-SPOService -Url $AdminSiteURL -credential $Cred
 
#Get all OneDrive for Business Site collections
$OneDriveSites = Get-SPOSite -Template "SPSPERS" -Limit ALL -IncludePersonalSite $True
Write-Host -f Yellow "Total Number of OneDrive Sites Found: "$OneDriveSites.count

#remove Site Collection Admin from each OneDrive
$i = 0
Foreach($Site in $OneDriveSites)
{
    try{
        $SAUser = Get-SPOUser -Site $Site.Url -LoginName $SiteCollAdmin | select IsSiteAdmin
        $IsSiteAdmin = $($SAUser.IsSiteAdmin).ToString()
    }
    catch{
        $IsSiteAdmin = "False"
    }
    $SiteUrl = $Site.Url
    if(($IsSiteAdmin -ne "False") -and ($SiteUrl -notlike "*SiteCollAdmin*") -and ($SiteUrl -notlike "*admin*") -and ($SiteUrl -notlike "*ai_*") -and ($SiteUrl -notlike "*sync_*")){
        Write-Host -f Yellow "Removing Site Collection Admin from: " -NoNewline
        Write-Host $SiteUrl 
        try{
            Set-SPOUser -Site $Site.Url -LoginName $SiteCollAdmin -IsSiteCollectionAdmin $False | Out-Null
            Write-Host -f Yellow "Removed from: "$Site.URL 
            $i++
        }
        catch{
            Write-Host -f Red "Error! Can not removing from: "$Site.URL 
        }
    }
}
Write-Host "Site Collection Admin ($SiteCollAdmin) removed from $i OneDrive sites successfully!" -f Green