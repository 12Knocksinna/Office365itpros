# Remove-Microsoft365UserAccount.PS1
# Prototype script to show how to remove a user account from Microsoft 365 and set up retention policies for the user's mailbox and OneDrive for Business site
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Remove-Microsoft365UserAccount.PS1

$UserToRemove = "Grady.Mike@Contoso.com"
$DefaultSecondarySiteAdmin = "Jayne.Sixsmith@contoso.com"
$RetentionPolicyForExEmployees = "Ex-Employees Retention Policy"

Connect-MgGraph -NoWelcome -Scopes Directory.Read.All, User.ReadWrite.All, Directory.AccessAsUser.All
# Connect to Exchange Online to add the user to the special retention policy for ex-employees
$Modules = Get-Module | Select-Object Name
if ('ExchangeOnlineManagement' -notin $Modules) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
    Connect-IPPSSession
}

# Do we have a valid account?
$User = Get-MgUser -UserId $UserToRemove -ErrorAction SilentlyContinue
If (!($User)) {
    Write-Host ("User {0} not found" -f $UserToRemove)
    Break
}

Write-Host "Disabling user account and revoking refresh tokens..."
# First thing to do is to block the account and set a new password
$NewPassword = @{}
$NewPassword["Password"]= "!NewYorkCity2022?"
$NewPassword["ForceChangePasswordNextSignIn"] = $True
Update-MgUser -UserId $User.Id -AccountEnabled:$False -PasswordProfile $NewPassword

# Revoke the user's refresh tokens
$RevokeStatus = Revoke-MgUserSignInSession -UserId $User.Id
If ($RevokeStatus.Value -eq $true) {
    Write-Host (“User access revoked for {0}!” -f $User.displayName)
}

# Disable registered devices
[array]$UserDevices = Get-MgUserRegisteredDevice -UserId $User.Id
If ($UserDevices) {
ForEach ($Device in $UserDevices) {
  Update-MgDevice -DeviceId $Device.Id -AccountEnabled:$False}
}

# Find the user's manager and set them as a secondary administrator for the user's OneDrive for Business site
$ManagerData = Get-MgUser -UserId $UserToRemove -ExpandProperty Manager
If ($ManagerData.Manager.Id) {
    $SecondarySiteAdmin = $ManagerData.Manager.AdditionalProperties['userPrincipalName']
} Else {
    $SecondarySiteAdmin = $DefaultSecondarySiteAdmin
}

# Connect to SharePoint Online to set a secondary administrator for the ex-employee's OneDrive for Business site
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

# Find the URI for the OneDrive for Business site
$OneDriveSiteURI = ("https://{0}-my.sharepoint.com/personal/{1}" -f ($DefaultDomain.Name.split('.')[0]), ($UserToRemove -replace '\.|\@', '_'))

# Assign the secondary administrator to the OneDrive for Business site
$Status = Set-SPOUser -Site $OneDriveSiteURI -LoginName $SecondarySiteAdmin -IsSiteCollectionAdmin $true
If ($Status) {
    Write-Host ("{0} is now a secondary administrator for {1}" -f $SecondarySiteAdmin, $OneDriveSiteURI)
} Else {
    Write-Host "Failed to assign secondary administrator"
}

# Add the user's mailbox to the retention policy for ex-employees
Write-Host "Adding mailbox and OneDrive for Business account to retention policy..."
Set-RetentionCompliancePolicy -Identity $RetentionPolicyForExEmployees -AddExchangeLocation $UserToRemove -AddOneDriveLocation $OneDriveSiteURI
# Also add a litigation hold to the mailbox
Set-Mailbox -Identity $UserToRemove -LitigationHoldEnabled $True
[array]$RetentionLocations = Get-RetentionCompliancePolicy -Identity $RetentionPolicyforExEmployees -DistributionDetail `
  | Select-Object -ExpandProperty ExchangeLocation
If ($UserToRemove -notin $RetentionLocations.Name) {
    Write-Host "Failed to add mailbox to retention policy"
    break
} Else {
    Write-Host ("Mailbox and OneDrive for Business account for {0} added to retention policy {1}" -f $UserToRemove, $RetentionPolicyForExEmployees)
}

Write-Host "Pausing to allow the retention hold to take effect"
# Delay to allow the retention hold to take effect
Start-Sleep -Seconds 30

Write-Host "Deleting user..."
# Now go ahead and remove the user's account
Remove-MgUser -UserId $User.Id -Confirm:$false
If (Get-MgUser -UserId $User.Id -ErrorAction SilentlyContinue) {
    Write-Host ("Failed to remove user account {0}" -f $UserToRemove)
} Else {
    Write-Host ("User account {0} removed" -f $UserToRemove)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.