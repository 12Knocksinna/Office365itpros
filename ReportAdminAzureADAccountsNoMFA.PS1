# ReportAdminAzureADAccountsNoMFA.PS1
# A script to find Azure AD accounts with privileged roles that aren't protected by MFA
# 
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportAdminAzureADAccountsNoMFA.PS1
#
# Uses both AzureAD and MSOnline modules

$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "AzureAD")) {Write-Host "Please connect to the Azure AD module and then restart the script" ; break}
If (!($ModulesLoaded -match "MSOnline")) {Write-Host "Please connect to the Microsoft Online Services module and then restart the script"; break}

# Define GUIDs for the Privileged Roles (from Get-AzureADDirectoryRole)

$UserAccountAdmin = Get-AzureADDirectoryRole | Where-Object {$_.DisplayName -eq ‘User Account Administrator’} | Select ObjectId
$TenantAdmin = Get-AzureADDirectoryRole | Where-Object {$_.DisplayName -eq ‘Company Administrator’} | Select ObjectId
$TeamsAdmin = Get-AzureADDirectoryRole | Where-Object {$_.DisplayName -eq ‘Teams Service Administrator’} | Select ObjectId
$ExchangeAdmin = Get-AzureADDirectoryRole | Where-Object {$_.DisplayName -eq ‘Exchange Service Administrator’} | Select ObjectId
$SharePointAdmin = Get-AzureADDirectoryRole | Where-Object {$_.DisplayName -eq ‘Sharepoint Service Administrator’} | Select ObjectId

# Find out the set of accounts that hold these admin roles in the tenant
$UserAccountAdmins = Get-AzureADDirectoryRoleMember -ObjectId $UserAccountAdmin.ObjectID | Select ObjectId, UserPrincipalName
$TenantAdmins = Get-AzureADDirectoryRoleMember -ObjectId $TenantAdmin.ObjectID | Select ObjectId, UserPrincipalName
$TeamsAdmins = Get-AzureADDirectoryRoleMember -ObjectId $TeamsAdmin.ObjectID | Select ObjectId, UserPrincipalName
$ExchangeAdmins = Get-AzureADDirectoryRoleMember -ObjectId $ExchangeAdmin.ObjectID | Select ObjectId, UserPrincipalName
$SharePointAdmins = Get-AzureADDirectoryRoleMember -ObjectId $SharePointAdmin.ObjectID | Select ObjectId, UserPrincipalName

$MFAReport = [System.Collections.Generic.List[Object]]::new() # Create output file 
$i = 0
Write-Host "Checking MFA status for accounts..."
$Users = (Get-MsolUser -All | ? {$_.UserType -eq "Member" -and $_.Islicensed -eq $True} | Sort DisplayName)
ForEach ($User in $Users) {
   $MFAMethods = $User.StrongAuthenticationMethods.MethodType
   $MFAEnforced = $User.StrongAuthenticationRequirements.State
   $DefaultMFAMethod = ($User.StrongAuthenticationMethods | ? {$_.IsDefault -eq "True"}).MethodType
   If (($MFAEnforced -eq "Enforced") -or ($MFAEnforced -eq "Enabled")) {
      Switch ($DefaultMFAMethod) {
        "OneWaySMS"             { $MethodUsed = "One-way SMS" }
        "TwoWayVoiceMobile"     { $MethodUsed = "Phone call verification" }
        "PhoneAppOTP"           { $MethodUsed = "Hardware token or authenticator app" }
        "PhoneAppNotification"  { $MethodUsed = "Authenticator app" }
      } #End Switch
    }
    Else {
          $MFAEnforced= "Not Enabled"
          $MethodUsed = "MFA Not Used" }
  
   $MFAReportLine = [PSCustomObject] @{
           UserPrincipalName  = $User.UserPrincipalName
           DisplayName        = $User.DisplayName
           MFAUsed            = $MFAEnforced
           MFAMethod          = $MethodUsed
           ObjectId           = $User.ObjectId }
                 
    $MFAReport.Add($MFAReportLine) 
} # End For

# Extract users whose accounts don't have MFA
$MFAUsers = $MFAReport | ? {$_.MFAUsed -ne "Enforced"}
If (!($MFAUsers)) { Write-Host "No privileged accounts found without MFA protection" ; break}

Write-Host "Checking holders of admin roles..."
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
# Check Admin Roles if MFA not enabled
ForEach ($User in $MFAUsers) {
  $Roles = $Null
  If ($UserAccountAdmins.ObjectId -Contains $User.ObjectId) {
         Write-Host $User.DisplayName "Account holds the User Account Admin role" -ForegroundColor Red 
         $Roles = "Account Admin" }
  If ($TenantAccountAdmins.ObjectId -Contains $User.ObjectId) {
         Write-Host $User.DisplayName "Account holds the Tenant Admin role" -ForegroundColor Red 
         If ($Roles -eq $Null) { $Roles = "Tenant Admin" } Else { $Roles = $Roles + "; Tenant Admin" } }
  If ($TeamsAdmins.ObjectId -Contains $User.ObjectId) {
         Write-Host $User.DisplayName "Account holds the Teams Admin role" -ForegroundColor Red 
         If ($Roles -eq $Null) { $Roles = "Teams Admin" } Else { $Roles = $Roles + "; Teams Admin" } }
  If ($ExchangeAdmins.ObjectId -Contains $User.ObjectId) {
         Write-Host $User.DisplayName "Account holds the Exchange Admin role" -ForegroundColor Red
         If ($Roles -eq $Null) { $Roles = "Exchange Admin" } Else { $Roles = $Roles + "; Exchange Admin" } }
  If ($SharePointAdmins.ObjectId -Contains $User.ObjectId) {
         Write-Host $User.DisplayName "Account holds the SharePoint Admin role" -ForegroundColor Red 
         If ($Roles -eq $Null) { $Roles = "SharePoint Admin" } Else { $Roles = $Roles + "; SharePoint Admin" } }      
 If ($Roles -ne $Null) {Write-Host "User" $User.DisplayName "is assigned the following roles:" $Roles -ForeGroundColor Yellow;  $i++ 
    $ReportLine = [PSCustomObject]@{
       User      = $User.DisplayName
       UPN       = $User.UserPrincipalName
       Roles     = $Roles
       MFA       = $User.MFAUsed }   
   $Report.Add($ReportLine) } #End if
}

Write-Host "All done." $i "privileged accounts found which aren't protected by MFA - see C:\temp\MFAReport.CSV for details"
$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation C:\temp\MFAReport.CSV 
