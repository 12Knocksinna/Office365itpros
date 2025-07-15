# Check-SignInsUtilityAccounts.PS1
# Use the Microsoft Graph to check sign-ins for accounts that shoudn't be signing in, such as the accounts
# used for break glass access, service accounts, and other utility accounts.

# V1.01 12-Jul-2025
# GitHub link:  https://github.com/12Knocksinna/Office365itpros/blob/master/Check-SignInsUtilityAccounts.PS1

If ([Environment]::UserInteractive) { 
    # We're running interactively...
    Write-Host "Running interactively..."
    Connect-MgGraph -Scopes AuditLog.Read.All, Mail.Send, User.Read.All -NoWelcome
} Else { 
    # We're not, so likely in Azure Automation
    Write-Output "Running the Check Utility account script in Azure Automation..."
    Connect-MgGraph -Identity
}

# Find marked utility accounts
[array]$UtilityAccounts = Get-MgUser -Filter "(onPremisesExtensionAttributes/extensionAttribute1 eq 'BG' or onPremisesExtensionAttributes/extensionAttribute1 eq 'Utility') and userType eq 'Member'" -All -PageSize 250 -ConsistencyLevel eventual -CountVariable Records

If ($UtilityAccounts) {
    Write-Host "Found $($Users.Count) utility accounts to process..." -ForegroundColor Yellow
} Else {
    Write-Host "No utility accounts found." -ForegroundColor Yellow
    Exit
}

Write-Host "Processing utility accounts..." -ForegroundColor Cyan
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($User in $UtilityAccounts) {
    Write-Host "Processing account: $($User.displayName)" -ForegroundColor Cyan
    Try {
        $UserId = $User.Id
        [array]$SignIn = Get-MgAuditLogSignIn -Filter "userid eq '$UserId'" -Top 1 -ErrorAction Stop
    } Catch {
        Write-Host "Failed to retrieve sign-ins for user $($User.displayName): $($_.Exception.Message)" -ForegroundColor Red
        Continue
    }

    If ($SignIn) {
        Write-Host "Found sign-ins for user $($User.displayName)" -ForegroundColor Red
        If ($SignIn.Status.ErrorCode -eq 0) {
            $SignInStatus = "Success"
        } Else {
            $SignInStatus = "Failure: $($SignIn.Status.ErrorCode)"
        }
        $ReportLine = [PSCustomObject]@{
            User          = $User.displayName
            UPN           = $User.userPrincipalName
            SignInTime    = Get-Date $SignIn.createdDateTime -format 'dd-MMM-yyyy HH:mm:ss'
            Status        = $SignInStatus
            SignInIP      = $SignIn.ipAddress
            SignInLocation= $SignIn.location.city + ", " + $SignIn.location.countryOrRegion
        }
        $Report.Add($ReportLine)
    }
}

# Define HTML style 
$HtmlStyle = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; background: #f4f6f8; color: #222; }
h1 { background: #0078d4; color: #fff; padding: 16px; border-radius: 6px 6px 0 0; margin-bottom: 0; }
table { border-collapse: collapse; width: 100%; background: #fff; border-radius: 0 0 6px 6px; overflow: hidden; }
th, td { padding: 10px 12px; text-align: left; }
th { background: #e5eaf1; color: #222; }
tr { background: #fff; color: #222; }
tr:nth-child(even) { background: #f0f4fa; color: #222; }
tr:hover { background: #d0e7fa; color: #222; }
.caption { font-size: 14px; color: #555; margin-bottom: 12px; }
</style>
"@

# Create the HTML content
$HtmlTable = $Report | ConvertTo-HTML -Fragment -As Table

$HtmlReport = @"
<html><head>
$HtmlStyle
<title>Sign-ins Detected for Utility Accounts</title>
</head><body><p>Report generated: <b>$((Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss 'UTC'"))</b></p>
<h1>Sign-In Records Observed for Utility Accounts</h1>
<p><p>Please check these sign-in events to validate that the accounts are not being used incorrectly.</p>
$HtmlTable
</body>
</html>
"@

If ($Report) {
    # Create and send a message to the administrator
    # Change this email address to be the sender of the message
    $MsgFrom = 'Background.Monitoring.Service@office365itpros.com'
    # Build the array of a single TO recipient detailed in a hash table - change this address to be the desired recipient of the mail
    $ToRecipient = @{}
    $ToRecipient.Add("emailAddress",@{'address'="AdminsDL@office365itpros.com"})
    [array]$MsgTo = $ToRecipient
    # Define the message subject
    $MsgSubject = "Important: Sign-ins Detected for Utility Accounts"

    # Construct the message body 	
    $MsgBody = @{}
    $MsgBody.Add('Content', "$($HtmlReport)")
    $MsgBody.Add('ContentType','html')

    # Build the parameters to submit the message
    $Message = @{}
    $Message.Add('subject', $MsgSubject)
    $Message.Add('toRecipients', $MsgTo)
    $Message.Add('body', $MsgBody)
    $EmailParameters = @{}
    $EmailParameters.Add('message', $Message)
    $EmailParameters.Add('saveToSentItems', $true)
    $EmailParameters.Add('isDeliveryReceiptRequested', $true)

    # Send the message
    Try {
        Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters -ErrorAction Stop
        Write-Host "Email sent successfully to $($MsgTo[0].emailAddress.address)" -ForegroundColor Green
    } Catch {
        Write-Host "Failed to send email: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.