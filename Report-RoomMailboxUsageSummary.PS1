# Report-RoomMailboxUsageSummary.PS1
# A very simple script to show how to detect whether room mailboxes are being used
# Application needs Calendar.ReadBasic permissions in the Graph to run successfully
# See https://practical365.com/room-mailbox-statistics-new/ for an article explaining how to generate a more comprehensive
# analysis of room mailbox usage using the Graph API and PowerShell. 

# V1.0 23-May-2026
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-RoomMailboxUsageSummary.PS1

# This is the command to use the Places API to find room mailboxes
#[array]$RoomMailboxes = Get-MgPlaceAsRoom -All -Sort "displayName"

# Modify these values for your own tenant and application registration
$AppId = 'f99054ee-efb6-4c69-95cc-500801071fa8'
$CertThumbprint = '0CF6CE3F3548FD73E7AC8CF20226ED447E125C71'
$TenantId = 'a662313f-14fc-43a2-9a7a-d2e27f4f3478'

# Connect to Exchange Online and then to the Graph using app-only mode
Connect-ExchangeOnline
Connect-MgGraph -AppId $AppId -TenantId $TenantId -CertificateThumbprint $CertThumbprint 

# Use Exchange Online to Find Mailboxes
[array]$RoomMailboxes = Get-ExoMailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited
# Set date range to check for
$StartDate = (Get-Date).AddMonths(-6)
$EndDate = (Get-Date).AddMonths(6)

ForEach ($RoomMailbox in $RoomMailboxes) {
    Try {
        Write-Host ("Getting calendar events for room mailbox {0}..." -f $RoomMailbox.DisplayName)
        [array]$Events = Get-MgUserCalendarView -UserId $RoomMailbox.PrimarySmtpAddress -All -StartDateTime $StartDate -EndDateTime $EndDate -ErrorAction Stop
    } Catch {
        Write-Host ("Couldn't get calendar events for room mailbox {0} - check the error message" -f $RoomMailbox.DisplayName)
    } 
    If ($Events) {
        Write-Host ("{0} events found for {1}" -f $Events.count, $RoomMailbox.DisplayName)
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
