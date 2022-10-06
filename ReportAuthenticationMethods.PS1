# ReportAuthenticationMethods.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportAuthenticationMethods.PS1
# A report of the authentication methods for Azure AD licensed accounts

Connect-MgGraph -Scopes UserAuthenticationMethod.Read.All, Directory.Read.All, User.Read.All
Select-MgProfile Beta

Write-Host "Finding licensed Azure AD accounts"
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All
If (!($Users)) { Write-Host "No licensed users found in Azure AD... exiting!"; break }

$i = 0
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
 $i++
 Write-Host ("Processing user {0} {1}/{2}." -f $User.DisplayName, $i, $Users.Count)
 $AuthMethods = Get-MgUserAuthenticationMethod -UserId $User.Id
 ForEach ($AuthMethod in $AuthMethods) {
  $P1 = $Null; $P2 = $Null
  $Method = $AuthMethod.AdditionalProperties['@odata.type']
  Switch ($Method) {
     "#microsoft.graph.passwordAuthenticationMethod" {
       $DisplayMethod = "Password"
       $P1 = "Traditional password"
     }
     "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
       $DisplayMethod = "Authenticator" 
       $P1 = $AuthMethod.AdditionalProperties['displayName']
       $P2 = $AuthMethod.AdditionalProperties['deviceTag'] + ": " + $AuthMethod.AdditionalProperties['clientAppName'] 
     }
     "#microsoft.graph.fido2AuthenticationMethod" {
       $DisplayMethod = "Fido 2 Key"
       $P1 = $AuthMethod.AdditionalProperties['displayName']
       $P2 = Get-Date($AuthMethod.AdditionalProperties['creationDateTime']) -format g
     }
     "#microsoft.graph.phoneAuthenticationMethod" {
       $DisplayMethod = "Phone" 
       $P1 = "Number: " + $AuthMethod.AdditionalProperties['phoneNumber']
       $P2 = "Type: " + $AuthMethod.AdditionalProperties['phoneType']
     }
    "#microsoft.graph.emailAuthenticationMethod" {
      $DisplayMethod = "Email"
      $P1 = "Address: " + $AuthMethod.AdditionalProperties['emailAddress']
     }
    "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" {
      $DisplayMethod = "Passwordless"
      $P1 = $AuthMethod.AdditionalProperties['displayName']
      $P2 = Get-Date($AuthMethod.AdditionalProperties['creationDateTime']) -format g
    }
     "Default" {
      $DisplayMethod = $Method
    }
  }
  
  $ReportLine   = [PSCustomObject] @{ 
     User   = $User.DisplayName
     Method = $DisplayMethod
     Id     = $AuthMethod.Id
     P1     = $P1
     P2     = $P2 
     UserId = $User.Id }
  $Report.Add($ReportLine)
 } #End ForEach Authentication Method
} #End ForEach User
   
$Report = $Report | Sort-Object User 
Write-Host ""
Write-Host "Authentication Methods found"
Write-Host "----------------------------"
Write-Host ""
$Report | Group-Object Method | Sort-Object Count -Descending | Select Name, Count

$Report | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
