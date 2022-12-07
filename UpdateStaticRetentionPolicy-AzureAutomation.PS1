# UpdateStaticRetentionPolicy-AzureAutomation.PS1
# Obtain AccessToken for Microsoft Graph via the managed identity
$ResourceURL = "https://graph.microsoft.com/" 
$Response = [System.Text.Encoding]::Default.GetString((Invoke-WebRequest -UseBasicParsing -Uri "$($env:IDENTITY_ENDPOINT)?Resource=$ResourceURL" -Method 'GET' -Headers @{'X-IDENTITY-HEADER' = "$env:IDENTITY_HEADER"; 'Metadata' = 'True'}).RawContentStream.ToArray()) | ConvertFrom-Json 
$AccessToken = $Response.access_token 

# Connect to AzAccount to access Key Vault to fetch variables used by the script
$AzConnection = Connect-AzAccount -Identity | Out-Null
# Get username and password from Key Vault
$UserName = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -Name "ExoAccountName" -AsPlainText
$UserPassword = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -name "ExoAccountPassword" -AsPlainText
# Create credentials object from the username and password
[securestring]$SecurePassword = ConvertTo-SecureString $UserPassword -AsPlainText -Force
[pscredential]$UserCredentials = New-Object System.Management.Automation.PSCredential ($UserName, $SecurePassword)

#Connect to the Microsoft Graph using the aquired AccessToken
Connect-Graph -AccessToken $accessToken
#Define the desired graph endpoint
Select-MgProfile -Name 'Beta'

# Get organization details so that we can figure out the SharePoint admin endpoint to connect to
$Tenant = Get-MgOrganization
$TenantName = $Tenant.VerifiedDomains | Where-Object {$_.IsInitial -eq $True} | Select -ExpandProperty Name
$TenantRoot = "https://" + $TenantName.Split(".")[0] + "-admin.sharepoint.com"

# Make sure that you update the organization name here
Connect-ExchangeOnline -ManagedIdentity -organization xxxxx.onmicrosoft.com 
Connect-IPPSSession -Credential $UserCredentials

# Define retention policy to update
$RetentionPolicy = "Office 365 for IT Pros Static Retention Policy"
# Find Mailboxes with the custom attribute set
[array]$Mbx = Get-ExoMailbox -Filter {CustomAttribute8 -eq "Office365itpros.com"} -Properties CustomAttribute8 
# Populate an array with the primary email addresses of the mailboxes
#[array]$Locations = $M.PrimarySmtpAddress

$ExoLocations = [System.Collections.Generic.List[Object]]::new()
  ForEach ($M in $Mbx) {
     $ExoLocation = [PSCustomObject] @{    
         Id = $M.ExternalDirectoryObjectId
         Mail = $M.PrimarySmtpAddress
         DisplayName = $M.DisplayName
     }     
     $ExoLocations.Add($ExoLocation)
}

# Connect to SharePoint Online
Connect-SPOService -url $TenantRoot -Credential $UserCredentials

# Find all OneDrive personal sites
[Array]$ODSites = Get-SPOSite -IncludePersonalSite $True -Limit All -Filter "Url -like 'my.sharepoint.com/personal/"
# Put the site owners and URLs into a hash table for easy lookup
$OneDriveSites = @{}    
ForEach ($OD in $ODSites) { 
   Try { $OneDriveSites.Add([string]$OD.Owner, [string]$OD.Url) }
   Catch {}
}

$OneDriveURls = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {
  $MbxOD = $OneDriveSites[$M.UserPrincipalName]
  If ($MbxOD) {
    $ODLine = [PSCustomObject] @{    
     "OneDrive Site" =  $MbxOD }
  $OneDriveUrls.Add($ODLine) }
}
[array]$ODSites = $OneDriveUrls | Select-Object -ExpandProperty "OneDrive Site" | Sort-Object  -Unique

# Find the existing locations covered by the policy
Write-Output ("Fetching existing Exchange and OneDrive for Business locations from the {0} retention policy" -f $RetentionPolicy)
$PolicyExoLocations = Get-RetentionCompliancePolicy -Identity $RetentionPolicy -DistributionDetail | Select-Object -ExpandProperty ExchangeLocation
$PolicyODLocations = Get-RetentionCompliancePolicy -Identity $RetentionPolicy -DistributionDetail | Select-Object -ExpandProperty OneDriveLocation

# Check what Exchange mailboxes in the set are not currently covered by the retention policies
$ExoAddToPolicy = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $ExoLocations) {
  If ($M.Id -notin $PolicyExoLocations.ImmutableIdentity) {
     $ExoLine = [PSCustomObject] @{    
       Id = $M.Id }
     $ExoAddToPolicy.Add($ExoLine) }
}
# Check what OneDrive sites are not currently covered by the policy
$ODAddToPolicy = [System.Collections.Generic.List[Object]]::new()
ForEach ($O in $ODSites) {
   If ($O -notin $PolicyODLocations.Name) {
      $ODLine = [PSCustomObject] @{    
       Id = $O }
     $ODAddToPolicy.Add($ODLine)}
}

# Now figure out if any Exchange and OneDrive locations are currently covered by the policy but not in the set found by the custom attribute
$ExoRemoveFromPolicy = [System.Collections.Generic.List[Object]]::new()
ForEach ($L in $PolicyExoLocations) {
  If ($L.ImmutableIdentity -notin $ExoLocations.Id) {
     $ExoLine = [PSCustomObject] @{    
       Id = $L.ImmutableIdentity }
     $ExoRemoveFromPolicy.Add($ExoLine) }
}

$ODRemoveFromPolicy = [System.Collections.Generic.List[Object]]::new()
ForEach ($S in $PolicyODLocations) {
   If ($S.Name -notin $ODSites) {
      $ODLine = [PSCustomObject] @{    
       Id = $S.Name }
     $ODRemoveFromPolicy.Add($ODLine)}
}

Write-Output "Results after check"
Write-Output "-------------------"
If ($ExoAddToPolicy) {
Write-Output ""
Write-Output "Mailboxes to add"
Write-Output ""
$ExoAddToPolicy.Id }
If ($ExoRemoveFromPolicy) {
Write-Output ""
Write-Output "Mailboxes to remove"
Write-Output ""
$ExoRemoveFromPolicy.Id }
If ($ODAddToPolicy) {
Write-Output ""
Write-Output "OneDrive sites to add"
Write-Output ""
$ODAddToPolicy.Id }
If ($ODRemoveFromPolicy) {
Write-Output ""
Write-Output "OneDrive sites to remove"
Write-Output ""
$ODRemoveFromPolicy.Id }

# Update the retention policy with new Exchange mailboxes and OneDrive for Business personal sites
Try {
Set-RetentionCompliancePolicy -Identity $RetentionPolicy -AddExchangeLocation $ExoAddToPolicy.Id -AddOneDriveLocation $ODAddtoPolicy.Id -RemoveExchangeLocation $ExoRemoveFromPolicy.Id -RemoveOneDriveLocation $ODRemoveFromPolicy.Id }
Catch {
  Write-Output ("An error occurred updating the {0} compliance policy" -f $RetentionPolicy)
}

# Hopefully all done
$ExistingLocations = Get-RetentionCompliancePolicy -Identity $RetentionPolicy -DistributionDetail | Select-Object -ExpandProperty ExchangeLocation
Write-Output ""
Write-Output "The following mailboxes are now covered by the policy"
Write-Output "-----------------------------------------------------"
$ExistingLocations.Name

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
