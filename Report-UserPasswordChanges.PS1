# Report-UserPasswordChanges

Connect-MgGraph -NoWelcome -Scopes AuditLog.Read.All, Directory.Read.All

[string]$RunDate = Get-Date -format "dd-MMM-yyyy HH:mm:ss"
$Version = "1.0"
$CSVOutputFile = "c:\temp\UserPasswordReport.CSV"
$ReportFile = "c:\temp\UserPasswordReport.html"

Write-Host "Retrieving user details"
# We use a Graph API request here instead of an SDK cmdlet because the Last successful interactive sign in data is
# unavailable in SDK V2.11

$Headers = @{ConsistencyLevel="Eventual"}  
$Uri = "https://graph.microsoft.com/beta/users?`$count=true&`$filter=(assignedLicenses/`$count ne 0 and userType eq 'Member')&$`top=999&`$select=id, displayName, usertype, signInActivity, SignInSessionsValidFromDateTime, LastPasswordChangeDateTime, passwordPolicies"
[array]$Data = Invoke-MgGraphRequest -Uri $Uri -Headers $Headers
[array]$Users = $Data.Value

If (!($Users)) {
    Write-Host "Can't find any users... exiting!" ; break
}

# Paginate until we have all the user accounts
While ($Null -ne $Data.'@odata.nextLink') {
    Write-Host ("Fetching more user accounts - currently at {0}" -f $Users.count)
    $Uri = $Data.'@odata.nextLink'
    [array]$Data = Invoke-MgGraphRequest -Uri $Uri -Headers $Headers
    $Users = $Users + $Data.Value
 }
 Write-Host ("All available user accounts fetched ({0}) - now processing report" -f $Users.count)

 $Users = $Users | Sort-Object displayName
 # Get MFA data
 [array]$MFAData = Get-MgBetaReportAuthenticationMethodUserRegistrationDetail
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
    If (!([string]::IsNullOrWhiteSpace($LastPasswordChange))) {
        $DaysSincePasswordChange = (New-TimeSpan $LastPasswordChange).Days 
    }

    $SessionTokensValidFrom = $User.SignInSessionsValidFromDateTime
    $LastPasswordChange = $User.LastPasswordChangeDateTime
    [array]$PasswordPolicies = $User.passwordPolicies

    If ($PasswordPolicies) {
        $PasswordPoliciesOutput = $PasswordPolicies -join ", "
    }

    # Get MFA status for the user
  
    $UserMFAStatus  = $MFAData | Where-Object {$_.Id -eq $User.Id}
    $AuthenticationTypesOutput = $UserMFAStatus.MethodsRegistered -join ", "

     $DataLine = [PSCustomObject][Ordered]@{
         User = $User.displayName
         UserId = $User.Id
         'Last password change'             = $LastPasswordChange
         'Days since password change'       = $DaysSincePasswordChange
         'Last successful sign in'          = $LastSuccessfulSignIn
         'Last sign in'                     = $LastSignIn
         'Days since successful sign in'    = $DaysSinceLastSuccessfulSignIn
         'Days since sign in'               = $DaysSinceLastSignIn
         'Session tokens valid from'        = $SessionTokensValidFrom 
         'Password policies applied'        = $PasswordPoliciesOutput
         'Authentication types'             = $AuthenticationTypesOutput
         'Admin flag'                       = $UserMFAStatus.isAdmin
         'MFA capable'                      = $UserMFAStatus.IsMfaCapable
         'MFA registered'                   = $UserMFAStatus.IsMfaRegistered
         'MFA default method'               = $UserMFAStatus.DefaultMfaMethod
         'Secondary auth. method'           = $UserMFAStatus.UserPreferredMethodForSecondaryAuthentication
     }
     $Report.Add($DataLine)
 }
 
# Now to generate a HTML report
Write-Host "Generating HTML report..."
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
       td.mfaregistered{background: #FFFF00;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>User Password Report</h1></p>
           <p><h2><b>For the " + $Orgname + " tenant</b></h2></p>
           <p><h3>Generated: " + $RunDate + "</h3></p></div>"


# This section highlights whether a conditional access policy is enabled or disabled in the summary.
# Idea from https://stackoverflow.com/questions/37662940/convertto-html-highlight-the-cells-with-special-values
# First, convert the CA Policies report to HTML and then import it into an XML structure
$HTMLTable = $Report | ConvertTo-Html -Fragment
[xml]$XML = $HTMLTable
# Create an attribute class to use, name it, and append to the XML table attributes
$TableClass = $XML.CreateAttribute("class")
$TableClass.Value = "State"
$XML.table.Attributes.Append($TableClass) | Out-Null
# Conditional formatting for the table rows. The number of available units is in table row 6, so we update td[5]
ForEach ($TableRow in $XML.table.SelectNodes("tr")) {
    # each TR becomes a member of class "tablerow"
    $TableRow.SetAttribute("class","tablerow")
    # If row has the admin flag set to true
    If (($TableRow.td) -and ([string]$TableRow.td[11] -eq 'True'))  {
        ## tag the TD with either the color for admin in the heading
        $TableRow.SelectNodes("td")[11].SetAttribute("class","admin")
    }
    # If MFA capable
    If (($TableRow.td) -and ([string]$TableRow.td[12] -eq 'True')) {
        $TableRow.SelectNodes("td")[12].SetAttribute("class","mfacapable")
    }
    # If MFA registered
    If (($TableRow.td) -and ([string]$TableRow.td[13] -eq 'True')) {
        $TableRow.SelectNodes("td")[13].SetAttribute("class","mfaregistered")
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
$PercentMFAAdmins = ($NumberAdminNoMFA/$AdminUsers.Count).ToString("P")


 # End stuff to output
$HTMLTail = "<p>Report created for the " + $OrgName + " tenant on " + $RunDate + "<p>" +
"<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
"<p>Number of user accouts analyzed:         " + $Users.Count + "</p>" +
"<p>Number of admin accounts:                " + $AdminUsers.Count + "</p>" +
"<p>User accounts not registered for MFA:    " + $NumberUsersNoMFA + " (" + $PercentMFAUsers + ")</p>" +
"<p>Admin accounts not registered for MFA:   " + $NumberAdminNoMFA + " (" + $PercentMFAAdmins + ")</p>" +
"<p>Names of admin accounts not registrered: " + $AdminNoMFANames + "</p>" +
"<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
"<p>Entra ID User Account Password Report<b> " + $Version + "</b>"	

$HTMLReport = $HTMLHead + $HTMLBody + $HTMLtail
$HTMLReport | Out-File $ReportFile  -Encoding UTF8

$Report | Export-Csv -NoTypeInformation $CSVOutputFile -Encoding utf8
Write-Host ("HTML format report is available in {0}  and CSV file in {1}" -f $ReportFile, $CSVOutputFile)

 # An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.