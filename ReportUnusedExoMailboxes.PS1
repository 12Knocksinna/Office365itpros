# ReportUnusedExoMailboxes.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportUnusedExoMailboxes.PS1
# Find and report unused Exchange Online mailboxes
# Needs to connect to Exchange Online and the Microsoft Graph PowerShell SDK
# Updated for Graph SDK V2

$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}

# Connect to the Microsoft Graph PowerShell SDK so that we can read sign in data
Connect-MgGraph -Scopes User.Read.All, AuditLog.Read.All -NoWelcome

# Find mailboxes and check if they are unused
$Now = Get-Date -format s; [int]$i = 0
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
    [datetime]$LastActiveDateTime | Out-Null
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
    # Get last Sign in from Azure AD sign in logs
    $LastUserSignIn = $Null
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
If ($UnusedMailboxes.Count -eq 0) { Write-Host "No unused mailboxes found!" ; break }

# Critical - make sure that you change this line to point to the webhook URI for the Teams channel you want to use
$WebHookUri = "https://office365itpros.webhook.office.com/webhookb2/672b3d70-31db-4c8c-b2d9-6b0892cd5b5b@b662313f-14fc-43a2-9a7a-d2e27f4f3478/IncomingWebhook/d5e79ee65089445c8d0b28ba3a15b123/eff4cd58-1bb8-4899-94de-795f656b4a18"

$InactiveTable = [System.Collections.Generic.List[Object]]::new() 
$UserImage = "https://i0.wp.com/office365itpros.com/wp-content/uploads/2022/04/pexels-photo-5605061.jpeg?ssl=1&resize=438%2C438"
ForEach ($R in $UnusedMailboxes) {
     $UserData = @{
     ActivityTitle = "$($R.Mailbox)"
     ActivitySubTitle = "-----------------------------------------------"
     ActivityText = "$($R.Mailbox)'s last email activity was on $($R.LastEmail)"
     ActivityImage = $UserImage
     Facts = @(
        @{    
      name  = 'Mailbox:'
      value = $R.Mailbox     },
    @{
      name  = 'Days since last active:'
      value = $R.DaysSinceActive   },
    @{
      name  = 'Last Email activity:'
      value = $R.LastEmail     },
    @{
      name  = 'Days since last activity:'
      value = $R.LastActive  },
    @{
      name  = 'Last logon:'
      value = $R.LastLogon    } )
    }
   $InactiveTable.Add($UserData)
  # Write-Host $R.Mailbox $R.LastEmail $DaysSince 
} # End ForEach 

$Body = ConvertTo-Json -Depth 8 @{
    Title = "Possibly Inactive Exchange Online Mailboxes"
    Text  = "Here are the top $($InactiveTable.Count) mailboxes with no detected email activity"
    Sections = $InactiveTable }
# Post to the channel
Write-Host "Posting unused mailbox information to Teams" -ForegroundColor Yellow
$Status = Invoke-RestMethod -uri $WebHookUri -Method Post -body $body -ContentType 'application/json' 
If ($Status) {
  Write-Host "Message posted to channel successfully"
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
