# FindOrphanOneDriveSites.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindOrphanOneDriveSites.PS1
# A script to find orphan OneDrive for Business Accounts and add an admin user to the accounts so that they can be accessed
# Needs connections to the SharePoint Online and Microsoft Graph PowerShell SDK modules
# V2.0 13 May 2025

Connect-MgGraph  -Scopes "User.Read.All", "Organization.Read.All" -NoWelcome

# Define the account to add to each orphan site
$NewSiteAdmin = "Administrator@office365itpros.com"

[array]$Domains = (Get-MgOrganization).verifiedDomains
$DefaultDomain = $Domains | Where-Object {$_.IsDefault -eq $true}
$SPOAdminRoot = ("https://{0}-admin.sharepoint.com" -f $DefaultDomain.Name.split('.')[0])
Write-Host "Connecting to SharePoint Online..."
Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell
Connect-SPOService -Url $SPOAdminRoot
If (Get-SPOTenant) {
    Write-Host ("Connected to SharePoint Online at {0}" -f $SPOAdminRoot)
} Else {
    Write-Host "Failed to connect to SharePoint Online"
    Break
}

# Create list for output report
$Report = [System.Collections.Generic.List[Object]]::new()

# Find OneDrive for Business accounts
Write-Host "Finding OneDrive for Business accounts..."
[array]$ODSites = Get-SPOSite -IncludePersonalSite $True -Limit All -Filter "url -like '-my.sharepoint.com/personal/'"
# Find Entra ID acounts and create hash table for lookup
Write-Host "Finding Entra ID user accounts..."
[array]$Users = Get-MgUser -All -Filter "Usertype eq 'Member'" -Property Id, DisplayName, UserPrincipalName

$UserAccounts = @{}
$Users.ForEach( {
       $UserAccounts.Add([String]$_.UserPrincipalName, $_.DisplayName) } )
# Process the sites
[int]$i = 0
ForEach ($Site in $ODSites) {
      If (!($UserAccounts.Item($Site.Owner))) { #Allocate a new owner to the OneDrive site
      Write-Host "Adding administator to" $Site.URL
      $Status = $null
      Try {
         $Status = Set-SPOUser -Site $Site.URL -LoginName $NewSiteAdmin -IsSiteCollectionAdmin $True -ErrorAction Stop
      }
      Catch {
         Write-Host "Couldn't add" $NewSiteAdmin "to" $Site.URL -ForegroundColor Red
      }
      If ($Status) { #Update output report file
         $i++
         $ReportLine = [PSCustomObject]@{  #Update with details of what we have done
           Site             = $Site.URL
           "Previous Owner" = $Site.Title
           OwnerUPN         = $Site.Owner
           "New Owner"      = $NewSiteAdmin
           LastModified     = Get-Date($Site.LastContentModifiedDate) -format g
           StorageUsage     = $Site.StorageUsageCurrent } 
         $Report.Add($ReportLine) } # End If
      } #End If
} # End ForEach

If ($i -gt 0) {
   Write-Host $NewSiteAdmin "added to" $i "OneDrive for Business accounts - details in c:\temp\OrphanOneDrive.csv"
   $Report | Export-CSV -NoTypeInformation c:\temp\OrphanOneDrive.csv }
Else {
   Write-Host "No orphan OneDrive for Business accounts found" 
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.

