# You need to import the SharePoint Online PowerShell module into your session and connect to SharePoint with an Admin account before 
# running this code. Something like the command below will do fine, substituting your tenant name...
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportOneDriveStorageUsage.PS1
# if(-not(Get-Module -name Microsoft.Online.Sharepoint.PowerShell)) {Import-Module Microsoft.Online.Sharepoint.PowerShell} 
# Connect-SPOService -url https://tenant-admin.sharepoint.com -Credential $O365Cred
# Now that we're connected, we can run this code
# Get a list of OneDrive for Business sites in the tenant sorted by the biggest consumer of quota
$ODFBSites = Get-SPOSite -IncludePersonalSite $True -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'" | Select Owner, Title, URL, StorageQuota, StorageUsageCurrent | Sort StorageUsageCurrent -Desc
$TotalODFBGBUsed = [Math]::Round(($ODFBSites.StorageUsageCurrent | Measure-Object -Sum).Sum /1024,2)
$Report = @()
ForEach ($Site in $ODFBSites) {
      $ReportLine = [PSCustomObject][Ordered]@{
             Owner    = $Site.Title
             Email    = $Site.Owner
             URL      = $Site.URL
             QuotaGB  = [Math]::Round($Site.StorageQuota/1024,2) 
             UsedGB   = [Math]::Round($Site.StorageUsageCurrent/1024,4) }
      $Report += $ReportLine }
$Report | Export-CSV -NoTypeInformation c:\temp\OneDriveSiteConsumption.CSV
Write-Host "Current OneDrive for Business storage consumption:" $TotalODFBGBUsed " Report is in C:\temp\OneDriveSiteConsumption.CSV" 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
