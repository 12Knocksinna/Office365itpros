# Get-UsersWithPhoneMFAMethod.PS1
# Find user accounts with a phone-based MFA method so that we can send them a nagging email
# V1.0 10-Jan-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Get-UsersWithPhoneMFAMethod.PS1

# Requires Users.Read.All and UserAuthenticationMethod.Read.All and Mail.Send
$Thumbprint = '32C9529B1FFD08BCD483A5D98807E47A472C5318'
$AppId = '9e289bb8-05bd-486a-aaec-55ecac4aa6ea'
$TenantId = 'b662313f-14fc-43a2-9a7a-d2e27f4f3478'
Connect-MgGraph -NoWelcome -AppId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint

# Find licensed user accounts
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
    -ConsistencyLevel eventual -CountVariable UsersFound -Sort "displayName" `
    -Property Id, DisplayName, GivenName , UserPrincipalName, UserType, Mail
If (!$Users) {
    Write-Host "No user accounts found"
    Exit
}
Write-Output ("Checking MFA methods for {0} user accounts" -f $UsersFound)
$Report = [System.Collections.Generic.List[Object]]::new()
    
ForEach ($User in $Users) {
    Write-Host ("Checking MFA methods for {0}" -f $User.UserPrincipalName)
    $Method = $null; $SystemPreferred = $false
    $Methods =  Get-MgBetaUserAuthenticationSignInPreference -UserId $User.Id
    If ($Methods.IsSystemPreferredAuthenticationMethodEnabled) {
        Write-Host ("User {0} has system preferred method enabled" -f $User.UserPrincipalName)
        $SystemPreferred = $true
    } 
    # Figure out what secondary method the account uses
    Switch ($Methods.UserPreferredMethodForSecondaryAuthentication) {
        'push' {
            Write-Host ("User {0} uses the Authenticator app for their MFA method" -f $User.UserPrincipalName)
            $Method = "App"
        }
        'oauth' {
            Write-Host ("User {0} uses OAuth Token MFA method" -f $User.UserPrincipalName)
            $Method = "Oauth Token"
        }
        'voiceMobile' {
            Write-Host ("User {0} uses Voice to Mobile for their MFA method" -f $User.UserPrincipalName)
            $Method = "Voice Mobile"
        }
        'voiceAlternateMobile' {
            Write-Host ("User {0} uses Voice to Alternate Mobile their MFA method" -f $User.UserPrincipalName)
            $Method = "Voice Mobile Alternate"
        }
        'voiceOffice' {
            Write-Host ("User {0} uses Voice to Office Phone for their MFA method" -f $User.UserPrincipalName)
            $Method = "Voice Office"
        }
        'sms' {
            Write-Host ("User {0} uses a phone for their MFA method" -f $User.UserPrincipalName)
            $Method = "SMS"
        }
    }
    If ($null -eq $Methods.UserPreferredMethodForSecondaryAuthentication) {
        Write-Host ("User {0} has no MFA method" -f $User.UserPrincipalName)
        $Method = "None"
    }
    $DataLine = [PSCustomObject][Ordered]@{
        UserPrincipalName   = $User.UserPrincipalName
        DisplayName         = $User.DisplayName
        GivenName           = $User.GivenName
        Method              = $Method
        Mail                = $User.Mail
        'System Preferred'  = $SystemPreferred
    }
    $Report.Add($DataLine)
}

# Define domains and specific users that we don't want to send nagging emails to
[array]$UserNoNag = 'Azure.Management.Account@office365itpros.com'
[array]$DomainsNoNag = 'redmondassociates.org'

[array]$UsersToNag = $Report | Where-Object {($_.Method -eq 'SMS') -and ($_.UserPrincipalName -notin $UserNoNag) -and ($_.UserPrincipalName -notlike "*@$($DomainsNoNag)")}
[array]$UsersNoMFA = $Report | Where-Object {($_.Method -eq 'None') -and ($_.UserPrincipalName -notin $UserNoNag) -and ($_.UserPrincipalName -notlike "*@$($DomainsNoNag)")}

Write-Host ("{0} users selected for nagging because they use the SMS MFA method" -f $UsersToNag.Count)
Write-Host ("{0} users selected for nagging because their account has no MFA method" -f $UsersNoMFA.Count)

# Send Email to the users with SMS MFA method - first define the message sender, which can be the SMTP address for a user or shared mailbox
$MsgFrom = 'Customer.Services@office365itpros.com'
# Define some variables used to construct the HTML content in the message body
#HTML header with styles
$HtmlHead="<html>
    <style>
    BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    H4{font-size: 8px; padding-top: 4px;}
</style>"

$MsgSubject = "Important: Please upgrade your MFA method"

ForEach ($User in $UsersToNag) {
    Write-Host ("Sending email to {0} to ask them to change their MFA method" -f $User.UserPrincipalName)
    $ToRecipients = @{}
    $ToRecipients.Add("emailAddress",@{'address'=$User.Mail})
    [array]$MsgTo = $ToRecipients
    # Customize the message 
    $HtmlHeaderUser = "<h2>"+ $User.DisplayName + ": It's time to say goodbye to SMS authentications</h2>"   
    #Content for the message - obviously, this is very customizable and should reflect what you want to say to new users
    $HtmlBody = "<body>
    <h1>SMS responses to Multifactor Challenges are no longer secure enough!</h1>
    <h2><u>We need you to increase the strength of your multifactor authentication</u></h2>
    <p><b>Dear $($User.GivenName),</b></p>
    <p>Our records show that you use SMS to respond to multifactor authentication challenges. In other words, you receive a code via SMS and enter that number when prompted by Entra ID.</p>
    <p>SMS was a great method some years ago, but now it's prone to compromise. We need you to upgrade to a more secure authentication method, such as the Microsoft Authenticator app, which is free and easy to use.</p>
    <p>To learn more about the Authenicator app, please <a href=https://support.microsoft.com/en-us/account-billing/about-microsoft-authenticator-9783c865-0308-42fb-a519-8cf666fe0acc>click here</a> </p>
    <p>Have a great time and be sure to call the help desk if you need assistance. And please read all the great articles about Microsoft 365 published on <a href=https://Practical365.com>Practical365.com</a>.</p>
    <p><p>
    <p><h4>Generated:</strong> $(Get-Date -Format g)</h4></p>"
    $HtmlMsg = "</body></html>" + $HtmlHead + $Htmlheaderuser + $HtmlBody + "<p>"
    # Construct the message body
    $MsgBody = @{}
    $MsgBody.Add('Content', "$($HtmlMsg)")
    $MsgBody.Add('ContentType','html')

    $Message = @{}
    $Message.Add('subject', $MsgSubject)
    $Message.Add('toRecipients', $MsgTo)
    $Message.Add('body', $MsgBody)
    $Params = @{}
    $Params.Add('message', $Message)
    $Params.Add('saveToSentItems', $true)
    $Params.Add('isDeliveryReceiptRequested', $true)    

    Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params

}

Write-Host ""
Write-Host ("All done - nagging emails sent to {0} users" -f $UsersToNag.Count)
$UsersToNag | Format-Table DisplayName, Mail -AutoSize

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.