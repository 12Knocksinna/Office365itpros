# Report-UserPasswordChanges
# A script to show how to report details of user password settings including dates for last password changes and 
# information about account MFA enablement and uses the sign in information for user accounts to figure out if 
# an account is active.
# V1.0 5-Jan-2024
# V1.1 17-Mar-2024
# V1.2 12-June-2024 Add check for per-user MFA state
# V1.3 20-June-2024 Add better output for authentication methods registered for a user

# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-UserPasswordChanges.PS1

Param (
    [Parameter(Mandatory = $false)]
    [switch]$PrivacyFlag  # Must be true to output minimal MFA data. This is the default. If you want to see full MFA data in the report, set to $true
)
function Get-AuthMethods {
    # Function to return details of the authentication methods registered for a user
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$UserId
    )
    [array]$AllMethods = @()
    [array]$AuthMethods = Get-MgUserAuthenticationMethod -UserId $UserId
    ForEach ($AuthMethod in $AuthMethods) {
       $P1 = $null; $P2 = $null
       $Method = $AuthMethod.AdditionalProperties['@odata.type']
       Switch ($Method) {
           "#microsoft.graph.passwordAuthenticationMethod" {
               $DisplayMethod = "Password"
               $P1 = "Traditional password"
               $MethodSummary = "Password"
           }
           "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
               $DisplayMethod = "Authenticator app" 
               $P1 = $AuthMethod.AdditionalProperties['displayName']
               $P2 = $AuthMethod.AdditionalProperties['deviceTag'] + " " + $AuthMethod.AdditionalProperties['phoneAppVersion'] 
               $MethodSummary = ("{0} on {1} ({2})" -f $DisplayMethod, $P1, $P2)
           }
           "#microsoft.graph.fido2AuthenticationMethod" {
               If ($AuthMethod.AdditionalProperties['aaGuid'] -eq "90a3ccdf-635c-4729-a248-9b709135078f") {
                   $DisplayMethod = ("Passkey: {0}" -f $AuthMethod.AdditionalProperties['model'])
               } Else {
                   $DisplayMethod = "Fido 2 Key"
               }
               $P1 = $AuthMethod.AdditionalProperties['displayName']
               $P2 = Get-Date($AuthMethod.AdditionalProperties['createdDateTime']) -format "dd-MMM-yyyy HH:mm"
               $MethodSummary = ("Passkey on {0}" -f $AuthMethod.AdditionalProperties['model'])
           }
           "#microsoft.graph.phoneAuthenticationMethod" {
               $DisplayMethod = "SMS" 
               $P1 = "Number: " + $AuthMethod.AdditionalProperties['phoneNumber']
               $P2 = "Type: " + $AuthMethod.AdditionalProperties['phoneType']
               $MethodSummary = ("{0} to {1} ({2})" -f $DisplayMethod, $AuthMethod.AdditionalProperties['phoneNumber'], $AuthMethod.AdditionalProperties['phoneType'])
           }
           "#microsoft.graph.emailAuthenticationMethod" {
               $DisplayMethod = "Email (SSPR)"
               $P1 = "Address: " + $AuthMethod.AdditionalProperties['emailAddress']
               $MethodSummary = ("{0} to {1}" -f $DisplayMethod, $AuthMethod.AdditionalProperties['emailAddress'] )
           }
           "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" {
               $DisplayMethod = "Passwordless"
               $P1 = $AuthMethod.AdditionalProperties['displayName']
               $P2 = Get-Date($AuthMethod.AdditionalProperties['createdDateTime']) -format "dd-MMM-yyyy HH:mm"
               $MethodSummary = ("{0} enabled on {1}" -f $DisplayMethod, $P2)
           }
           "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {    
               $DisplayMethod = "Windows Hello for Business"
               $P1 = $AuthMethod.AdditionalProperties['displayName']
               $P2 = Get-Date($AuthMethod.AdditionalProperties['createdDateTime']) -format "dd-MMM-yyyy HH:mm"
               $MethodSummary = ("{0} enabled for {1} on {2}" -f $DisplayMethod, $P1, $P2)
           }
           Default {
               $DisplayMethod = "Unknown authentication method"
               $MethodSummary = ("Unknown method: {0}" -f $AuthMethod.AdditionalProperties['@odata.type'])
           }
        }
    $AllMethods += $MethodSummary 
    }
    $AllMethods = $AllMethods -join ", "
    Return $AllMethods
}

Connect-MgGraph -NoWelcome -Scopes AuditLog.Read.All, Directory.Read.All, UserAuthenticationMethod.Read.All, Policy.ReadWrite.AuthenticationMethod

[string]$RunDate = Get-Date -format "dd-MMM-yyyy HH:mm:ss"
$Version = "1.3"
$CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\UserAuthenticationReport.CSV"
$ReportFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\UserAuthenticationReport.html"

# Define the optional MFA user sign-in file downloaded from the Entra ID admin center
$UserSignInDataFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InteractiveSignIns.csv"
# Flag indicating if MFA user sign-in data is available
$MFADataAvailable = $false

# If the file is available, import it (the “incoming token type” column must be removed first)
If (Test-Path $UserSignInDataFile) {
    [array]$MFASignIns = Import-CSV $UserSignInDataFile -ErrorAction SilentlyContinue
    # If some data was imported, reduce it down to unique entries for user accounts that have used MFA
    If ($MFASignIns) {
        [array]$MFAUserData = $MFASignIns | Where-Object {$_.'user type' -eq 'Member' -and $_.Status -eq 'Success'} |`
          Sort-Object {$_.'Date (UTC)' -as [datetime]} | Sort-Object 'User ID' -Unique | `
          Select-Object 'User ID', 'User', 'UserName', 'Date (UTC)','Multifactor authentication result', 'Multifactor authentication auth method'
        $MFADataAvailable = $true
        Write-Host ("MFA User Sign In Data available: using file dated {0}" -f  (Get-Item $UserSignInDataFile).LastWriteTime)
    } Else {
        Write-Host ("MFA User Sign In Data file ({0}) unavailable" -f $UserSignInDataFile)
    }
}

Write-Host "Retrieving user details"
# We use a Graph API request here instead of an SDK cmdlet because the Last successful interactive sign in data is
# unavailable in SDK V2.19

$ContentType = "application/json"
$Headers = @{ConsistencyLevel="Eventual"}  
$Uri = "https://graph.microsoft.com/beta/users?`$count=true&`$filter=(assignedLicenses/`$count ne 0 and userType eq 'Member')&$`top=999&`$select=id, displayName, userprincipalname, usertype, signInActivity, SignInSessionsValidFromDateTime, LastPasswordChangeDateTime, passwordPolicies"
[array]$Data = Invoke-MgGraphRequest -Uri $Uri -Headers $Headers -ContentType $ContentType -Method Get
[array]$Users = $Data.Value

If (!($Users)) {
    Write-Host "Can't find any users... exiting!" ; break
}

# Paginate until we have all the user accounts
$Nextlink = $Data.'@odata.nextLink'
While ($null -ne $Nextlink) {
    Write-Host ("Fetching more user accounts - currently at {0}" -f $Users.count)
    [array]$Data = Invoke-MgGraphRequest -Uri $Nextlink -Headers $Headers -ContentType $ContentType -Method Get
    $Users = $Users += $Data.Value
    $Nextlink = $Data.'@odata.nextLink'
 }
 Write-Host ("All available user accounts fetched ({0}) - now processing report" -f $Users.count) -ForegroundColor Yellow

 $Users = $Users | Sort-Object displayName
 # Get MFA data
 [array]$MFAData = Get-MgBetaReportAuthenticationMethodUserRegistrationDetail -All -PageSize 999
 $MFAData = $MFAData | Where-Object {$_.userType -eq 'Member'}

 # Report what we've found
 $Report = [System.Collections.Generic.List[Object]]::new()
 [int]$i = 0
 ForEach ($User in $Users) {
    $i++
    Write-Host ("Processing {0} ({1}/{2})..." -f $User.displayname, $i, $Users.count)
    $DaysSinceLastSignIn = $null; $DaysSinceLastSuccessfulSignIn = $null
    $DaysSincePasswordChange = $null; $PasswordPoliciesOutput = $null
    $DaysSinceLastSignIn = "N/A"; $DaysSinceLastSuccessfulSignIn = "N/A"
    
    If (!([string]::IsNullOrWhiteSpace($User.signInActivity.lastSuccessfulSignInDateTime))) {
        [datetime]$LastSuccessfulSignIn = $User.signInActivity.lastSuccessfulSignInDateTime
         $DaysSinceLastSuccessfulSignIn = (New-TimeSpan $LastSuccessfulSignIn).Days 
    }
    If (!([string]::IsNullOrWhiteSpace($User.signInActivity.lastSignInDateTime))) {
        [datetime]$LastSignIn = $User.signInActivity.lastSignInDateTime
        $DaysSinceLastSignIn = (New-TimeSpan $LastSignIn).Days
    }    
    If (!([string]::IsNullOrWhiteSpace($User.LastPasswordChangeDateTime))) {
        $LastPasswordChange = $User.LastPasswordChangeDateTime
        $DaysSincePasswordChange = (New-TimeSpan $LastPasswordChange).Days 
    }

    $SessionTokensValidFrom = $User.SignInSessionsValidFromDateTime
    $LastPasswordChange = $User.LastPasswordChangeDateTime
    [array]$PasswordPolicies = $User.passwordPolicies

    If ($PasswordPolicies) {
        $PasswordPoliciesOutput = $PasswordPolicies -join ", "
    }

    # Get MFA status for the user. If the privacy flag is set, we get the basic data reported for the user
    # If not, we get the full set of authentication methods registered for the user, including the details of each method
    $UserMFAStatus  = $MFAData | Where-Object {$_.Id -eq $User.Id}
    If ($PrivacyFlag -eq $true) {   
        $AuthenticationTypesOutput = $UserMFAStatus.MethodsRegistered -join ", "
    } Else { 
        $AuthenticationTypesOutput = Get-AuthMethods -UserId $User.Id
    }

    # Get per-user MFA data if possible
    $Uri = ("https://graph.microsoft.com/beta/users/{0}/authentication/requirements" -f $User.id)
    $Data = Invoke-MgGraphRequest -Uri $Uri -Method Get

    # Make sure dates are all in a common format
    If ($LastSignIn) {
        $LastSignInOutput = (Get-Date $LastSignIn -format 'dd-MMM-yyyy HH:mm')
    }
    If ($LastPasswordChange) {
        $LastPasswordChangeOutput = (Get-Date $LastPasswordChange -format 'dd-MMM-yyyy HH:mm')
    }
    If ($LastSuccessfulSignIn) {
        $LastSuccessfulSignInOutput = (Get-Date $LastSuccessfulSignIn -format 'dd-MMM-yyyy HH:mm')
    }
    If ($SessionTokensValidFrom) {
        $SessionTokensValidFromOutput = (Get-Date $SessionTokensValidFrom -format 'dd-MMM-yyyy HH:mm')
    }

    If ($MFADataAvailable) {
        $MFAVerifiedDate = $null
        $MFAVerifiedDate = $MFAUserData | Where-Object {$_.'User Id' -eq $User.Id} | Select-Object -ExpandProperty 'Date (UTC)'
        If ($MFAVerifiedDate) {
            $MFAVerifiedDate = (Get-Date $MFAVerifiedDate -format 'dd-MMM-yyyy HH:mm')
        }

        $DataLine = [PSCustomObject][Ordered]@{
            User                               = $User.displayName
            UserId                             = $User.Id
            UPN                                = $User.userPrincipalName
            UserType                           = $User.userType
            'Last password change'             = $LastPasswordChangeOutput
            'Days since password change'       = $DaysSincePasswordChange
            'Last successful sign in'          = $LastSuccessfulSignInOutput
            'Last sign in'                     = $LastSignInOutput
            'Days since successful sign in'    = $DaysSinceLastSuccessfulSignIn
            'Days since sign in'               = $DaysSinceLastSignIn
            'Session tokens valid from'        = $SessionTokensValidFromOutput
            'Password policies applied'        = $PasswordPoliciesOutput
            'Authentication methods'           = $AuthenticationTypesOutput
            'Admin flag'                       = $UserMFAStatus.isAdmin
            'MFA capable'                      = $UserMFAStatus.IsMfaCapable
            'MFA registered'                   = $UserMFAStatus.IsMfaRegistered
            'MFA last used'                    = $MFAVerifiedDate
            'MFA default method'               = $UserMFAStatus.DefaultMfaMethod
            'Secondary auth. method'           = $UserMFAStatus.UserPreferredMethodForSecondaryAuthentication
            'Per user MFA state'               = $Data.perUserMfaState
        }
    } Else {
        $DataLine = [PSCustomObject][Ordered]@{
            User                               = $User.displayName
            UserId                             = $User.Id
            UPN                                = $User.userPrincipalName
            UserType                           = $User.userType
            'Last password change'             = $LastPasswordChangeOutput
            'Days since password change'       = $DaysSincePasswordChange
            'Last successful sign in'          = $LastSuccessfulSignInOutput
            'Last sign in'                     = $LastSignInOutput
            'Days since successful sign in'    = $DaysSinceLastSuccessfulSignIn
            'Days since sign in'               = $DaysSinceLastSignIn
            'Session tokens valid from'        = $SessionTokensValidFromOutput
            'Password policies applied'        = $PasswordPoliciesOutput
            'Authentication methods'           = $AuthenticationTypesOutput
            'Admin flag'                       = $UserMFAStatus.isAdmin
            'MFA capable'                      = $UserMFAStatus.IsMfaCapable
            'MFA registered'                   = $UserMFAStatus.IsMfaRegistered
            'MFA default method'               = $UserMFAStatus.DefaultMfaMethod
            'Secondary auth. method'           = $UserMFAStatus.UserPreferredMethodForSecondaryAuthentication
            'Per user MFA state'               = $Data.perUserMfaState
        }
    }
    # Output report line
    $Report.Add($DataLine)
 }
 
# Now to generate a HTML report
Write-Host "Generating HTML report..."
[array]$PerUserMFAStates = "enabled", "enforced"
$OrgName  = (Get-MgOrganization).DisplayName
#  First, define the header.
$HTMLHead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.admin{background: #B7EB83;}
	   td.mfacapable{background: #E3242B;}
       td.mfaperuserenabled{background: #FFFF00;}
       td.mfaregistered{background: #FF474C;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>User Passwords and Authentication Report</h1></p>
           <p><h2><b>For the " + $Orgname + " tenant</b></h2></p>
           <p><h3>Generated: " + $RunDate + "</h3></p></div>"

# This section highlights whether a conditional access policy is enabled or disabled in the summary.
# Idea from https://stackoverflow.com/questions/37662940/convertto-html-highlight-the-cells-with-special-values
# First, convert the CA Policies report to HTML and then import it into an XML structure

$HTMLTable = $Report | Select-Object User, UPN, 'Last password change', 'Days since password change', 'Last successful sign in', 'Days since sign in', 'Password policies applied', `
    'Authentication methods', 'Admin flag', 'MFA capable', 'MFA registered', 'MFA last used', 'MFA default method', 'Secondary auth. method', 'Per user MFA state' | ConvertTo-Html -Fragment
[xml]$XML = $HTMLTable
# Create an attribute class to use, name it, and append to the XML table attributes
$TableClass = $XML.CreateAttribute("class")
$TableClass.Value = "State"
$XML.table.Attributes.Append($TableClass) | Out-Null
# Conditional formatting for the table rows.  
ForEach ($TableRow in $XML.table.SelectNodes("tr")) {
    # each TR becomes a member of class "tablerow"
    $TableRow.SetAttribute("class","tablerow")
    # If row has the admin flag set to true
    If (($TableRow.td) -and ([string]$TableRow.td[8] -eq 'True'))  {
        ## tag the TD with the color for admin in the heading
        $TableRow.SelectNodes("td")[8].SetAttribute("class","admin")
    }
    # If MFA capable
    If (($TableRow.td) -and ([string]$TableRow.td[9] -eq 'True')) {
        $TableRow.SelectNodes("td")[9].SetAttribute("class","mfacapable")
    }
    # If MFA registered
    If (($TableRow.td) -and ([string]$TableRow.td[10] -eq 'True')) {
        $TableRow.SelectNodes("td")[10].SetAttribute("class","mfaregistered")
    }
    # Per-user MFA status
    If (($TableRow.td) -and ([string]$TableRow.td[14] -in $PerUserMFAStates)) {
        $TableRow.SelectNodes("td")[14].SetAttribute("class","mfaperuserenabled")
    }
}
# Wrap the output table with a div tag
$HTMLBody = [string]::Format('<div class="tablediv">{0}</div>',$XML.OuterXml)

[array]$MFAUsers = $Report | Where-Object {$_.'MFA Registered' -eq $True}
[array]$AdminUsers = $Report | Where-Object {$_.'Admin Flag' -eq $True}
[array]$AdminNoMfA = $AdminUsers | Where-Object {$_.'MFA Registered' -eq $False}
[string]$AdminNoMFANames = $AdminNoMFA.User -Join ", "
[int]$NumberAdminNoMFA = $AdminNoMFA.Count
[int]$NumberUsersNoMFA =  ($Users.Count - $MFAUsers.count)
$PercentMFAUsers = ($NumberUsersNoMFA/$Users.Count).ToString("P")
If ($NumberofAdmins -eq 0) {
    $PercentMFAAdmins = "N/A"
} Else {
    $PercentMFAAdmins = ($NumberAdminNoMFA/$AdminUsers.Count).ToString("P")
}
[array]$PerUserMFAEnabled = $Report | Where-Object {$_.'Per user MFA state' -eq 'enabled'}
[array]$PerUserMFAEnforced = $Report | Where-Object {$_.'Per user MFA state' -eq 'enforced'}


# If MFA data is available, find out how many MFA-capable users are actually connecting with MFA
If ($MFADataAvailable) {
    [array]$MFAActiveUsers = $Report | Where-Object {$_.'MFA last used' -ne $null}
    $PercentMFAActive = ($MFAActiveUsers.Count/$MFAUsers.Count).toString("P")
}

 # End stuff to output
$HTMLTail = "<p>Report created for the " + $OrgName + " tenant on " + $RunDate + "<p>" +
"<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
"<p>Number of user accouts analyzed:         " + $Users.Count + "</p>" +
"<p>Number of admin accounts found:          " + $AdminUsers.Count + "</p>" +
"<p>Number of accounts registered for MFA:   " + $MFAUsers.Count + "</p>" 

IF ($PerUserMFAEnabled.Count -gt 0 -or $PerUserMFAEnforced.Count -gt 0) {
    $HTMLTail = $HTMLTail +
    "<p>Number of accounts with per-user MFA:    " + $PerUserMFAEnabled.Count + "</p>" +
    "<p>Names of accounts enabled for per-user MFA: " + ($PerUserMFAEnabled.User -join ", ") + "</p>" +
    "<p>Number of accounts with per-user MFA enforced: " + $PerUserMFAEnforced.Count + "</p>" +
    "<p>Names of accounts with per-user MFA enforced: " + ($PerUserMFAEnforced.User -join ", ") + "</p>"
}

If ($MFADataAvailable) {
    $HTMLTail = $HTMLTail +
    "<p>Number of MFA accounts active:          " + $MFAActiveUsers.count + "</p>" +
    "<p>Percentage of MFA accounts active:      " + $PercentMFAActive + "</p>"
}

$HTMLTail = $HTMLTail + 
"<p>User accounts not registered for MFA:    " + $NumberUsersNoMFA + " (" + $PercentMFAUsers + ")</p>" +
"<p>Admin accounts not registered for MFA:   " + $NumberAdminNoMFA + " (" + $PercentMFAAdmins + ")</p>" +
"<p>Names of admin accounts not registrered: " + $AdminNoMFANames + "</p>" +
"<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
"<p>Entra ID User Passwords and Authentication Report<b> " + $Version + "</b>"	

$HTMLReport = $HTMLHead + $HTMLBody + $HTMLtail
$HTMLReport | Out-File $ReportFile  -Encoding UTF8

$Report | Export-Csv -NoTypeInformation $CSVOutputFile -Encoding utf8
Write-Host ("HTML format report is available in {0} and CSV file in {1}" -f $ReportFile, $CSVOutputFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
