# Transfer-MeetingOwnership.PS1
# This script transfers ownership of Microsoft Teams meetings from one user to another based on specified criteria.

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Transfer-MeetingOwnership.PS1
# V1.0 05-Aug-2025

# Exchange Online PowerShell 3.10.1 has the bug fix to support certificate-based authentication.
#Requires -Version 3.0
#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.10.1" }

# The app used for the script must be assigned the Exchange<ManageAsApp permisssion and the Exchange administrator role to be able to run Exchangae Online cmdlets.
# See https://office365itpros.com/2022/10/13/exchange-online-powershell-app/ for background.

# The script also uses the Microsoft Graph PowerShell SDK and needs the Calendar.Read and User.ReadBasic.All permissions to read the calendar and user information. 
# The same app can be used to authenticate with both Exchange Online and the Graph.

[CmdletBinding()]
param(
    [string]$SourceMailbox,
    [string]$TargetMailbox,
    [datetime]$StartDate
)

If (-not $SourceMailbox -or -not $TargetMailbox -or -not $StartDate) {
    Write-Host "Please provide all required parameters: SourceMailbox, TargetMailbox, and StartDate."
    return
}

[datetime]$CheckDate = (Get-Date).AddHours(3)
# The start date for the transfer must be more than 3 hours in the future to allow for processing time and to avoid conflicts with ongoing meetings.
If ($StartDate -lt $CheckDate) {
    Write-Host ("The value of the StartDate parameter ({0}) must be more than 3 hours in the future to allow for ongoing meetings to conclude. Please provide a valid date." -f (Get-Date $StartDate -format 'dd-MMM-yyyy HH:mm'))
    return
}

$AppId = '03ef5d75-e8dd-4f9e-a71c-fa81939d5ca2' # Replace with your Entra ID app ID
$Thumbprint = "0CF6CE3F3548FD73E7AC8CF20226ED447E125C71" # Replace with your certificate thumbprint
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3478" # Replace with your tenant ID

# Connect to the Microsoft Graph in app-only mode
Write-Host "Connecting to Microsoft Graph..."
Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome
# Connect to Exchange Online in app-only mode
Write-Host "Connecting to Exchange Online..."
Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $Thumbprint -Organization "redmondassociates.onmicrosoft.com" -ShowBanner:$false

# Check that the source and target mailboxes exist and that they are user mailboxes
Write-Host "Validating source and target mailboxes..."
$SourceMailboxDetails = Get-Mailbox -Identity $SourceMailbox -ErrorAction SilentlyContinue
If (-not $SourceMailboxDetails -or $SourceMailboxDetails.RecipientTypeDetails -ne 'UserMailbox') {
    Write-Host ("The source mailbox '{0}' does not exist or is not a user mailbox. Please provide a valid source mailbox." -f $SourceMailbox)
    return
}
$TargetMailboxDetails = Get-Mailbox -Identity $TargetMailbox -ErrorAction SilentlyContinue
If (-not $TargetMailboxDetails -or $TargetMailboxDetails.RecipientTypeDetails -ne 'UserMailbox') {
    Write-Host ("The target mailbox '{0}' does not exist or is not a user mailbox. Please provide a valid target mailbox." -f $TargetMailbox)
    return
}

$SourceMailboxId = $SourceMailboxDetails.ExternalDirectoryObjectId
$TargetMailboxId = $TargetMailboxDetails.ExternalDirectoryObjectId

Write-Host ("Checking the calendar of {0} to find meetings that can be transferred to {1} starting from {2}" -f $SourceMailbox, $TargetMailbox, (Get-Date $StartDate -format 'dd-MMM-yyyy HH:mm'))
# Filter for events past the start date that are not cancelled.
Try {
    [array]$Events = Get-MgUserDefaultCalendarEvent -UserId $SourceMailboxId -All `
         -Filter "start/datetime gt '$($StartDate.ToString('yyyy-MM-ddTHH:mm:ss'))' and IsCancelled eq false" -ErrorAction Stop
    Write-Host ("Found $($Events.Count) meetings in the calendar of {0}. Now checking transfer eligibility for each event..." -f $SourceMailbox)
} Catch {
    Write-Host ("Error retrieving events from the source mailbox: {0}" -f $_.Exception.Message)
    return
}

$Report = [System.Collections.Generic.List[object]]::new()
[int]$EventsNotProcessed = 0

ForEach ($CalendarEvent in $Events) {
    $EventStartTime =  (Get-Date $CalendarEvent.Start.DateTime).toUniversalTime().toString('dd-MMM-yyyy HH:mm' + " UTC")
    $EventEndTime =  (Get-Date $CalendarEvent.End.DateTime).toUniversalTime().toString('dd-MMM-yyyy HH:mm' + " UTC")
    $TeamsMeeting = $false
    $TransferStatus = "Transferred"
    
    # Check if the meeting is a Teams meeting and if the source user is the organizer
    If ($CalendarEvent.Attendees.count -eq 0) {
        $EventsNotProcessed++
        Write-Host ("Skipping meeting '{0}' scheduled for {1} because it is a personal event with no attendees." -f $CalendarEvent.Subject, $EventStartTime) -ForegroundColor Yellow
        continue
    }
    # Check if the event is a meeting with participants and if the source user is the organizer
    If ($CalendarEvent.Organizer.EmailAddress.Address -eq $SourceMailboxDetails.PrimarySmtpAddress){
        Write-Host ("Transferring ownership of meeting '{0}' scheduled for {1} to {2}..." -f $CalendarEvent.Subject, $EventStartTime, $TargetMailbox)
         # Is it a Teams Meeting?
        If ($CalendarEvent.IsOnlineMeeting -and $CalendarEvent.OnlineMeetingProvider -eq 'teamsForBusiness') {
            Write-Host ("Meeting '{0}' scheduled for {1} is a Teams online meeting." -f $CalendarEvent.Subject, $EventStartTime) -ForegroundColor Red
            $TeamsMeeting = $true
        }
        If ($null -ne $CalendarEvent.recurrence.pattern.type) {
            Write-Host ("Meeting '{0}' scheduled for {1} is a recurring meeting. " -f $CalendarEvent.Subject, $EventStartTime) -ForegroundColor Yellow}
            Try {
                Invoke-ChangeMeetingOrganizer -Identity $SourceMailboxId -EventId $CalendarEvent.Id -NewOrganizer $TargetMailboxId `
                    -TransferSeriesStartDate $StartDate
                $EventType = "Recurring"
            } Catch {
                Write-Host ("Error transferring recurring meeting '{0}': {1}" -f $CalendarEvent.Subject, $_.Exception.Message) -ForegroundColor Red
                $TransferStatus = "Failed: " + $_.Exception.Message
                Continue
            }
        } Else {
            Try {
                Invoke-ChangeMeetingOrganizer -Identity $SourceMailboxId -EventId $CalendarEvent.Id -NewOrganizer $TargetMailboxId 
                $EventType = "Single"
            } Catch {
                Write-Host ("Error transferring meeting '{0}': {1}" -f $CalendarEvent.Subject, $_.Exception.Message) -ForegroundColor Red
                $TransferStatus = "Failed: " + $_.Exception.Message
                Continue
            }
        }
        Start-Sleep -Milliseconds 500 # Pause to avoid throttling
       
        $ReportLine = [PSCustomObject]@{
            Subject         = $CalendarEvent.Subject
            StartTime       = $EventStartTime
            EndTime         = $EventEndTime
            IsTeamsMeeting  = $TeamsMeeting
            EventType       = $EventType
            TransferStatus  = $TransferStatus
            Organizer       = $CalendarEvent.Organizer.EmailAddress.Address
            NewOrganizer    = $TargetMailboxDetails.PrimarySmtpAddress
            Attendees       = ($CalendarEvent.Attendees | ForEach-Object { $_.EmailAddress.Address }) -join ', '
            StartTimeUTC    = $CalendarEvent.Start.DateTime
            EndTimeUTC      = $CalendarEvent.End.DateTime
        }
        $Report.Add($ReportLine)
} 

$TeamsMeetings = $Report | Where-Object { $_.IsTeamsMeeting -eq $true} 
$OutputCSVFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Transferred Teams Meetings.csv"
Remove-Item -Path $OutputCSVFile -ErrorAction SilentlyContinue
$TeamsMeetings | Export-Csv -Path $OutputCSVFile -NoTypeInformation -Encoding Utf8

Write-Host ""
Write-Host ("Transfer process completed. Summary of {0} meetings transferred:" -f $Report.Count)
$Report | Format-Table Subject, StartTime, EndTime, IsTeamsMeeting -AutoSize
Write-Host ""
Write-Host ("Total personal events not processed: {0}" -f $EventsNotProcessed)
Write-Host ("Total Teams Meetings in transferred set: {0}" -f ($TeamsMeetings.Count) )
Write-Host "The Teams meetings must be created by the new organizer to ensure that the meeting link is valid and that the new organizer has control over the meeting."
Write-Host ("A CSV report of the transferred Teams meetings has been saved to: {0}" -f $OutputCSVFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.