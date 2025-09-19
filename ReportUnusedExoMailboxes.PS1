# ReportUnusedExoMailboxes.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportUnusedExoMailboxes.PS1
# Find and report unused Exchange Online mailboxes
# Needs to connect to Exchange Online and the Microsoft Graph PowerShell SDK
# V1.1 Updated for Graph SDK V2
# V1.2 19-Sep-2025 Make sure that the Graph connection happens before Exchange Online to ensure that a mismatch doesn't happen with
# the MSAL assemblies. Also removed the post tom Teams channel via a webhook and replaced it with an email to a specified address. The email
# could be sent to a Teams channel, if you add the email address of the channel

# Connect to the Microsoft Graph PowerShell SDK so that we can read sign in data
Connect-MgGraph -Scopes User.Read.All, AuditLog.Read.All -NoWelcome

# Check for Exchange Online
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {
  Write-Host "Loading Exchange Online PowerShell module" -ForegroundColor Yellow
  Connect-ExchangeOnline -ShowBanner:$False
}

# Find mailboxes and check if they are unused
$Now = Get-Date -format s
[int]$i = 0
Write-Host "Looking for User Mailboxes..."
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | `
    Select-Object DisplayName, DistinguishedName, UserPrincipalName, ExternalDirectoryObjectId | Sort-Object DisplayName
  Write-Host ("Reporting {0} mailboxes..." -f $Mbx.Count)
  $Report = [System.Collections.Generic.List[Object]]::new() 
  ForEach ($M in $Mbx) {
    $i++  
    Write-Host ("Processing {0} {1}/{2}" -f $M.DisplayName, $i, $Mbx.count) 
    $LastActive = $Null
    $Log = Export-MailboxDiagnosticLogs -Identity $M.DistinguishedName -ExtendedProperties 
    $xml = [xml]($Log.MailboxLog) 
    $LastEMail = $Null; $LastCalendar = $Null; $LastContacts = $Null; $LastFile = $Null
    $LastEmail = ($xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastEmailTimeCurrentValue"}).Value
    $LastCalendar = ($xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastCalendarTimeCurrentValue"}).Value
    $LastContacts = ($xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastContactsTimeCurrentValue"}).Value
    $LastFile = ($xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastFileTimeCurrentValue"}).Value
    $LastLogonTime = ($xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastLogonTime"}).Value 
    $LastActive = ($xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastUserActionWorkloadAggregateTime"}).Value 
    
    # This massaging of dates is to accommodate the different U.S. date format returned by Export-MailboxDiagnosticsData
    [datetime]$LastActiveDateTime = Get-Date
    If ([string]::IsNullOrEmpty($LastActive)) {
        $DaysSinceActive = "N/A"
    }
    If (($LastActive.IndexOf("M") -gt -0)) { # U.S. format date with AM or PM in it
        $LastActiveDateTime = [datetime]$LastActive
    } Else {
        $LastActiveDateTime = Get-Date ($LastActive) 
    }
    If ($LastActiveDateTime) {
        $DaysSinceActive = (New-TimeSpan -Start $LastActiveDateTime -End $Now).Days 
    }
  
    # Get Mailbox statistics
    $Stats = (Get-ExoMailboxStatistics -Identity $M.DistinguishedName)
    $MbxSize = ($Stats.TotalItemSize.Value.ToString()).Split("(")[0] 
    # Get last Sign in from Entra ID sign in logs
    $LastUserSignIn = $null
    $LastUserSignIn = (Get-MgAuditLogSignIn -Filter "UserId eq '$($M.ExternalDirectoryObjectId)'" -Top 1).CreatedDateTime
    If ($LastUserSignIn) {
       $LastUserSignInDate = Get-Date($LastUserSignIn) -format g 
    } Else {
       $LastUserSignInDate = "No sign in records found in last 30 days" 
    }
    # Get account enabled status
    $AccountEnabled = (Get-MgUser -UserId $M.ExternalDirectoryObjectId -Property AccountEnabled).AccountEnabled
    $ReportLine = [PSCustomObject][Ordered]@{ 
        Mailbox         = $M.DisplayName 
        UPN             = $M.UserPrincipalName
        Enabled         = $AccountEnabled
        Items           = $Stats.ItemCount 
        Size            = $MbxSize 
        LastLogonExo    = $LastLogonTime
        LastLogonAD     = $LastUserSignInDate
        DaysSinceActive = $DaysSinceActive
        LastActive      = $LastActive
        LastEmail       = $LastEmail
        LastCalendar    = $LastCalendar
        LastContacts    = $LastContacts
        LastFile        = $LastFile } 
    $Report.Add($ReportLine) 
  } 
$Report | Sort-Object DaysSinceActive -Descending | Out-GridView

# Extract the mailboxes that are inactive for more than 60 days but only take 25 because that's how much we can post in Teams
[array]$UnusedMailboxes = $Report | Where-Object {$_.DaysSinceActive -ge 60 } | Sort-Object DaysSinceActive -Descending | Select-Object -First 25
If ($UnusedMailboxes.Count -eq 0) { 
  Write-Host "No unused mailboxes found!" ; break 
}

# The original script posted to a Teams channel. Changed in this version to email to an address
$MsgFrom = (Get-MgContext).Account
$MsgSubject = "Report of possibly unused Exchange Online mailboxes"
$ToRecipient = @{}
# Update the target email address here
$ToRecipient.Add("emailAddress",@{'address'="Help.Desk@office365itpros.com"})
[array]$MsgTo = $ToRecipient
$HtmlMsg = $UnusedMailboxes | Select-Object Mailbox, DaysSinceActive, LastEmail, LastActive, LastLogonExo, LastLogonAD | ConvertTo-Html -Fragment
# Construct the message body 	
$MsgBody = @{}
$MsgBody.Add('Content', "$($HtmlMsg)")
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
  Write-Host "Sending email to $($MsgTo.emailAddress.address)" -ForegroundColor Yellow
  Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters
} Catch {
  Write-Host "Failed to send email to $($MsgTo.emailAddress.address)" -ForegroundColor Red
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
