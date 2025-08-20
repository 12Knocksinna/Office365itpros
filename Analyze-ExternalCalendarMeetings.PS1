# Analyze-ExternalCalendarMeetings.PS1
# Find out what external domains users who attend meetings belong to

# V1.01 19-Aug-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Analyze-ExternalCalendarMeetings.PS1

# Define the credentials to be used to allow the SDK to access user calendars using application permissions
$Thumbprint = '6FBD36F588E12DE291100371E16219863E399818'
$TenantId = 'a662313f-14fc-43a2-9a7a-d2e27f4f3478'
$AppId = 'a4cfce82-2b2f-4c87-b32f-6cb32b5fdc5b'

# Permissions used
# User.Read.All - Read user account information, including licenses.
# Group.Read.All - Read test group info.
# GroupMember.Read.All - Read members of the test group
# Calendars.Read - Read user calendars
# Domain.Read.All - Read accepted domains defined for the tenant

Connect-MgGraph -AppId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome
Clear-Host

# Define the GUID for the Teams service plan
$TeamsServicePlanId = '57ff2da0-773e-42df-b2af-ffb7a2317929'
# Define the name of an Entra ID group whose members are used to test the script (or run the script for a subset of users)
$GroupName = "Test Users for Graph Scripts"
# See if the test group exists containing the user accounts for testing. If not, find all user accounts licensed with a Teams service plan
Write-Host "Finding user accounts to analyze..."
Try {
    $Group = Get-MgGroup -Filter "displayName eq '$GroupName'" -ErrorAction Stop
    $GroupMembers = Get-MgGroupMember -Group $Group.Id
    [array]$Users = $GroupMembers | Select-Object -ExpandProperty Id
} Catch {
    [array]$Users = Get-MgUser -Filter "assignedPlans/any(c:c/servicePlanId eq $TeamsServicePlanId and capabilityStatus eq 'Enabled')" `
        -ConsistencyLevel eventual -CountVariable Test -All -PageSize 500 | Select-Object -ExpandProperty Id
}

If ($Users.Count -eq 0) {
    Write-Host "No users found in the group or with the Teams service plan. Exiting script."
    Break
} Else {
    Write-Host ("Found {0} users to analyze" -f $Users.Count)
}

# Get the domains defined for the tenant
Write-Host ("Finding domains defined for tenant {0}" -f $TenantId)
[array]$Domains = Get-MgDomain -All | Select-Object -ExpandProperty Id | Sort-Object

# Start and end dates for the reporting period
$StartDate = (Get-Date).AddDays(-365)
$EndDate = (Get-Date).AddDays(1)
  
$CalendarInfo = [System.Collections.Generic.List[Object]]::new() 
[array]$DomainsFound = $null

ForEach ($User in $Users) {
    Write-Host "Processing calendar for $User"
 
    [array]$CalendarData = Get-MgUserCalendarView -UserId $User -StartDateTime $StartDate -EndDateTime $EndDate -All -PageSize 250 
    # Drop cancelled events and filter to get online meetings - if you want to exclude private events from the set, use Where-Object {$_.sensitivity -ne "private"}
    $CalendarData = $CalendarData | Where-Object {$_.isCancelled -eq $False -and $_.isOnlineMeeting -eq $True}
    
    If ($CalendarData) {
        Write-Host ("Found {0} online meetings for {1}" -f $CalendarData.Count, $User)
    } Else {
        Continue
    }   

    [array]$UserDomainsFound = $null
    # Extract attendees and find what domains they come from
    ForEach ($CalendarEvent in $CalendarData) {

        # Check if the meeting organizer is from our tenant. If not, the event is organized elsewhere and we ignore it
        $Organizer = $CalendarEvent.organizer.emailaddress
        If ($Organizer.address.Split("@")[1] -notin $Domains) {
            Continue
        }
        # Get the set of attendees and extract the domains for attendees not in our tenant
        [array]$EventAttendees = $CalendarEvent.attendees.emailaddress
        If ($EventAttendees) {
            ForEach ($Attendee in $EventAttendees) {
                $AttendeeDomain = $Attendee.Address.Split('@')[1]
                If ($AttendeeDomain -notin $Domains) {
                    $UserDomainsFound += $AttendeeDomain
                }
            }
        }
    }

    $DomainsFound = $DomainsFound + $UserDomainsFound  

    $ReportLine = [PSCustomObject]@{
        User            = $User
        MeetingCount    = $CalendarData.Count
        ExternalDomains = ($UserDomainsFound | Sort-Object -Unique) -join "; "
    }
    $CalendarInfo.Add($ReportLine)

}

[array]$UniqueDomains = $DomainsFound | Sort-Object -Unique
Write-Host ""
Write-Host "All done... Unique domains found" ($UniqueDomains -join "; ")

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.