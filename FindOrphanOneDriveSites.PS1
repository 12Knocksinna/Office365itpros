# FindOrphanOneDriveSites.PS1
# A script to find orphan OneDrive for Business Accounts and add an admin user to the accounts so that they can be accessed
# Needs connections to the SharePoint Online and Azure AD modules

# Connect to AzureAD and SPO
$AzureADCheck = Get-Module -Name AzureADPreview
If ($AzureADCheck -eq $Null) {
     Write-Host "Your PowerShell session is not connected to Azure Active Directory."
     Write-Host "Please connect to Azure Active Directory using an administrative account and retry."; Break }

Connect-AzureAD -Credential $O365Cred

If(-not(Get-Module -name Microsoft.Online.Sharepoint.PowerShell)) {Import-Module Microsoft.Online.Sharepoint.PowerShell} 
Connect-SPOService -URL $TenantAdminUrl -Credential $O365Cred

# Define the SharePoint admin endpoint for the tenant
$TenantAdminUrl = "https://redmondassociates-admin.sharepoint.com"
# Define the account to add to each orphan site
$NewSiteAdmin = "Administrator@Redmondassociates.org"

# Create list for output report
$Report = [System.Collections.Generic.List[Object]]::new()

# Find Azure AD accounts
$AADUsers = Get-AzureADUser -All $True -Filter "Usertype eq 'Member'"
# Find OneDrive for Business accounts
$ODSites = Get-SPOSite -IncludePersonalSite $True -Limit All -Filter "url -like '-my.sharepoint.com/personal/'"
$i = 0
# Do the checks...
ForEach ($Site in $ODSites) {
   If ($Site.Owner -notin $AADUsers.UserPrincipalName) { #Allocate a new owner to the OneDrive site
      Write-Host "Adding user to" $Site.URL
      $Status = $Null
      Try {
         $Status = Set-SPOUser -Site $Site.URL -LoginName $NewSiteAdmin -IsSiteCollectionAdmin $True }
      Catch {
         Write-Host "Couldn't add" $NewSiteAdmin "to" $Site.URL }
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
   Write-Host "No orphan OneDrive for Business accounts found" }


