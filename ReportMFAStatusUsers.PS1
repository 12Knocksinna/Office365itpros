# ReportMFAStatusUsers.
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportMFAStatusUsers.PS1
# A script to report the authentication methods used by Azure AD user accounts
# Updated for SDK V2 5-Jan-2024
 
# Connect to the Microsoft Graph SDK for PowerShell
Connect-MgGraph -Scope UserAuthenticationMethod.Read.All, Directory.Read.All, User.Read.All, Auditlog.Read.All, UserAuthenticationMethod.Read.All -NoWelcome

# Get user accounts (exclude guests)
Write-Host "Looking for Entra ID user accounts to check"
[array]$Users = Get-MgUser -All -Filter "UserType eq 'Member'"

If (!($Users)) { Write-Host "No accounts found for some reason... exiting" ; break}
    Else { Write-Host ("{0} Entra ID member accounts found (not all are user accounts which authenticate)" -f $Users.count ) }
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
                  "#microsoft.graph.softwareOathAuthenticationMethod" { # Software Authenticator App
                  $AuthType = 'Third-party Authenticator App'
                  $AuthTypeDetails = $MfaMethod.AdditionalProperties["displayName"] }
                  } # End switch 
            # Note what we found
            $ReportLine  = [PSCustomObject][Ordered]@{
                  User            = $User.DisplayName
                  UPN             = $User.UserPrincipalName 
                  Method          = $AuthType
                  Details         = $AuthTypeDetails
                  LastSignIn      = $LastSignIn.CreatedDateTime
                  LastSignInApp   = $LastSignIn.AppDisplayName}
            $Report.Add($ReportLine) 
            } #End Foreach MfaMethod
      } # End if
} # End ForEach Users

# Take the report file and check each user to see if they use a strong authentication method 
$OutputFile = [System.Collections.Generic.List[Object]]::new()
[array]$AuthUsers = $Report | Sort-Object UPN -Unique | Select-Object UPN, User, LastSignIn, LastSignInApp
ForEach ($AuthUser in $AuthUsers) {
    $MFAStatus = $Null
    $Records = $Report | Where-Object {$_.UPN -eq $AuthUser.UPN}
    $Methods = $Records.Method | Sort-Object -Unique
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
         MFAStatus       = $MFAStatus
          LastSignIn      = $AuthUser.LastSignIn
         LastSignInApp   = $AuthUser.LastSignInApp }
    $OutputFile.Add($ReportLine) 
} 
   
$OutputFile | Out-GridView
$OutputFile | Export-CSV -NoTypeInformation c:\Temp\MFAStatus.CSV

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
