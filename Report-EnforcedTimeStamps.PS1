# Report-EnforcedTimeStamps.PS1
# GitHub Link:

# Report the compliance time stamps written into Exchange Online mailboxes to guide the Managed Folder Assistant in how
# it applies retention policies and holds to mailboxes.
# V1.0 11-Dec-2025

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {   
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
}

# Define the event types
$EventTypeMap = @{
    1  = 'RetentionHoldApplied'
    2  = 'RetentionHoldRemoved'
    3  = 'LitigationHoldApplied'
    4  = 'LitigationHoldRemoved'
    5  = 'DelayHoldApplied'
    6  = 'DelayHoldRemoved'
    7  = 'RetentionPolicyApplied'
    8  = 'RetentionPolicyRemoved'
    9  = 'RetentionPolicyTagApplied'
    10 = 'RetentionPolicyTagRemoved'
    11 = 'CalendarLoggingDisabled'
    12 = 'CalendarLoggingEnabled'
    13 = 'AuditEnabled'
    14 = 'AuditDisabled'
    15 = 'SingleItemRecoveryDisabled'
    16 = 'SingleItemRecoveryEnabled'
    17 = 'DiscoveryHoldApplied'
    18 = 'DiscoveryHoldRemoved'
    26 = 'DiscoveryHoldRemoved (Legacy)'
}

# Find mailboxes
[array]$Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Sort-Object DisplayName
Write-Host ("Found {0} mailboxes to process" -f $Mailboxes.Count)
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Mailbox in $Mailboxes) {
    [array]$Events = $Mailbox.EnforcedTimestamps
    If ($Events.Count -eq 0) {
        Write-Host ("No enforced time stamps found for mailbox {0} ({1})" -f $Mailbox.DisplayName, $Mailbox.PrimarySmtpAddress)
        Continue
    }   
    $EventData = $Events | ConvertFrom-Json
    ForEach ($Event in $EventData) {
        $EventType = $null
        [int]$EventType = $Event.EventType
        [string]$EventTypeString = $EventTypeMap[$EventType]
        $EventTimeStamp = Get-Date $Event.EventTimeStamp -Format 'dd-MMM-yyyy HH:mm'
        $EnforcedUntilTimestamp = If ($Event.EnforcedUntilTimestamp) { Get-Date $Event.EnforcedUntilTimestamp -Format 'dd-MMM-yyyy HH:mm' } Else { 'N/A' }
        $ReportLine = [PSCustomObject][Ordered]@{
            DisplayName             = $Mailbox.DisplayName
            PrimarySmtpAddress      = $Mailbox.PrimarySmtpAddress
            EventType               = $EventType
            'Event Name'            = $EventTypeString
            EventTimeStamp          = $EventTimeStamp
            EnforcedUntilTimestamp  = $EnforcedUntilTimestamp
        }
        $Report.Add($ReportLine)
    }
}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.