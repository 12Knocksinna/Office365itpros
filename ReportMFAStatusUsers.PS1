# ReportMFAStatusUsers.
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportMFAStatusUsers.PS1
# A script to report the authentication methods used by Azure AD user accounts
 
# Check that we have loaded the Azure AD module
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "AzureAD")) {Write-Host "Please connect to the Azure Ad module and then restart the script"; break}
$TenantInfo = (Get-AzureADTenantDetail)
$TenantId = $TenantInfo.ObjectId
$TenantName = $TenantInfo.DisplayName

# Now connect to the Microsoft Graph SDK for PowerShell
Connect-MgGraph -TenantId $TenantId -Scope "UserAuthenticationMethod.Read.All, Directory.Read.All, User.Read.All, Auditlog.Read.All"
Select-MgProfile Beta
$Details = Get-MgContext
$Scopes = $Details | Select -ExpandProperty Scopes
$Scopes = $Scopes -Join ", "
$ProfileName = (Get-MgProfile).Name
CLS
Write-Host "Microsoft Graph Connection Information"
Write-Host "--------------------------------------"
Write-Host " "
Write-Host ("Connected to Tenant {0} ({1}) as account {2}" -f $Details.TenantId, $TenantName, $Details.Account)
Write-Host "+-------------------------------------------------------------------------------------------------------------------+"
Write-Host ("Profile set as {0}. The following permission scope is defined: {1}" -f $ProfileName, $Scopes)
Write-Host ""

# Get user accounts (exclude guests)
Write-Host "Looking for Azure AD user accounts to check"
[array]$Users = Get-MgUser -All -Filter "UserType eq 'Member'"

If (!($Users)) { Write-Host "No accounts found for some reason... exiting" ; break}
   Else { Write-Host ("{0} Azure AD user accounts found" -f $Users.count ) }
$CheckedUsers = 0
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    # Try and find a sign in record for the user - this eliminates unused accounts 
   # Write-Host "Checking" $User.DisplayName
    [array]$LastSignIn = Get-MgAuditLogSignIn -Filter "UserId eq '$($User.Id)'" -Top 1
    If ($LastSignIn) {
    $CheckedUsers++
    Write-Host "Sign in found - checking authentication methods for" $User.DisplayName
    [array]$MfaData = Get-MgUserAuthenticationMethod -UserId $User.Id 
    # Process each of the authentication methods found for an account
      ForEach ($MfaMethod in $MfaData) {   
        Switch ($MfaMethod.AdditionalProperties["@odata.type"]) {
         "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod"  { # Microsoft Authenticator App
               $AuthType     = 'AuthenticatorApp'
               $AuthTypeDetails = $MfaMethod.AdditionalProperties["displayName"] } 
         "#microsoft.graph.phoneAuthenticationMethod"                  { # Phone authentication
               $AuthType     = 'PhoneAuthentication'
               $AuthTypeDetails = $MfaMethod.AdditionalProperties["phoneType", "phoneNumber"] -join ' '  } 
         "#microsoft.graph.fido2AuthenticationMethod"                   { # FIDO2 key
               $AuthType     = 'Fido2'
               $AuthTypeDetails = $MfaMethod.AdditionalProperties["model"] }  
         "#microsoft.graph.passwordAuthenticationMethod"                { # Password
               $AuthType     = 'PasswordAuthentication'
               $AuthTypeDetails = $MfaMethod.AdditionalProperties["displayName"] } 
         "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" { # Windows Hello
               $AuthType     = 'WindowsHelloForBusiness'
               $AuthTypeDetails = $MfaMethod.AdditionalProperties["displayName"] }                        
         "#microsoft.graph.emailAuthenticationMethod"                   { # Email Authentication
               $AuthType     = 'EmailAuthentication'
               $AuthTypeDetails = $MfaMethod.AdditionalProperties["emailAddress"] }               
         "microsoft.graph.temporaryAccessPassAuthenticationMethod"    { # Temporary Access pass
               $AuthType     = 'TemporaryAccessPass'
               $AuthTypeDetails = 'Access pass lifetime (minutes): ' + $MfaMethod.AdditionalProperties["lifetimeInMinutes"] }
         "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" { # Passwordless
               $AuthType     = 'Passwordless'
               $AuthTypeDetails = $MfaMethod.AdditionalProperties["displayName"] }                      
           } # End switch 
       # Note what we found
        $ReportLine  = [PSCustomObject][Ordered]@{
            User            = $User.DisplayName
            UPN             = $User.UserPrincipalName 
            Method          = $AuthType
            Details         = $AuthTypeDetails }
       $Report.Add($ReportLine) 
    } #End Foreach MfaMethod
  } # End if
} # End ForEach Users

# Take the report file and check each user to see if they use a strong authentication method 
$OutputFile = [System.Collections.Generic.List[Object]]::new()
[array]$AuthUsers = $Report | Sort UPN -Unique | Select UPN, User
ForEach ($AuthUser in $AuthUsers) {
    $MFAStatus = $Null
    $Records = $Report | ? {$_.UPN -eq $AuthUser.UPN}
    $Methods = $Records.Method | Sort -Unique
    Switch ($Methods) {
      "Fido2"               { $MFAStatus = "Good" }
      "PhoneAuthentication" { $MFAStatus = "Good" }
      "AuthenticatorApp"    { $MFAStatus = "Good" }
      "Passwordless"        { $MFAStatus = "Good" }
       Default              { $MFAStatus = "Check!" }
    } # End Switch
    $ReportLine  = [PSCustomObject][Ordered]@{
         User            = $AuthUser.User
         UPN             = $AuthUser.UPN
         Methods         = $Methods -Join ", "
         MFAStatus       = $MFAStatus }
    $OutputFile.Add($ReportLine) 
} 
   
$OutputFile | Out-GridView
$OutputFile | Export-CSV -NoTypeInformation c:\Temp\MFAStatus.CSV

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
