# ReportRoomMailboxUsage.PS1
# A script to report how busy room mailboxes are

# Requires the Graph Calendar.Read.All and Place.Read.All application permissions

# Some functions to get going

function Get-GraphData {
# Based on https://danielchronlund.com/2018/11/19/fetch-data-from-microsoft-graph-with-powershell-paging-support/
# GET data from Microsoft Graph.
    param (
        [parameter(Mandatory = $true)]
        $AccessToken,

        [parameter(Mandatory = $true)]
        $Uri
    )

    # Check if authentication was successful.
    if ($AccessToken) {
    $Headers = @{
         'Content-Type'  = "application\json"
         'Authorization' = "Bearer $AccessToken" 
         'ConsistencyLevel' = "eventual"  }

        # Create an empty array to store the result.
        $QueryResults = @()

        # Invoke REST method and fetch data until there are no pages left.
        do {
            $Results = ""
            $StatusCode = ""

            do {
                try {
                    $Results = Invoke-RestMethod -Headers $Headers -Uri $Uri -UseBasicParsing -Method "GET" -ContentType "application/json"

                    $StatusCode = $Results.StatusCode
                } catch {
                    $StatusCode = $_.Exception.Response.StatusCode.value__

                    if ($StatusCode -eq 429) {
                        Write-Warning "Got throttled by Microsoft. Sleeping for 45 seconds..."
                        Start-Sleep -Seconds 45
                    }
                    else {
                        Write-Error $_.Exception
                    }
                }
            } while ($StatusCode -eq 429)

            if ($Results.value) {
                $QueryResults += $Results.value
            }
            else {
                $QueryResults += $Results
            }

            $uri = $Results.'@odata.nextlink'
        } until (!($uri))

        # Return the result.
        $QueryResults
    }
    else {
        Write-Error "No Access Token"
    }
}

function GetAccessToken {
# function to return an Oauth access token

# Define the values applicable for the application used to connect to the Graph (you need to update these values for your tenant)
$AppId = "a6a7d55c-a847-443d-b7b9-f24b67ec4709"
$TenantId = "xxx-xxx-xxx-xxx"
$AppSecret = '_O28Q~.QZNE5dd4pOCaxqTSx13dbkDIadadfNs5blI'

# Construct URI and body needed for authentication
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials"
}

# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
# Unpack Access Token
$Global:Token = ($tokenRequest.Content | ConvertFrom-Json).access_token

Return $Token
}

# End functions - start doing some real work
#

$Token = GetAccessToken
If (!($Token)) {Write-Host "Can't get a valid Azure AD access token - exiting" ; break }

$Headers = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $Token" 
            'ConsistencyLevel' = "eventual" }

$StartDate = (Get-Date).AddDays(-60)
$EndDate = (Get-Date).AddDays(1)

$Start = Get-Date($StartDate) -format s
$End = Get-Date($EndDate) -format s
$ReportingPeriodDays = (($EndDate - $StartDate).Days)-1

# Find room mailboxes - this ignores room mailboxes marked as workspaces
$Uri = "https://graph.microsoft.com/beta/places/microsoft.graph.room"
[Array]$RoomMailboxes = Get-GraphData -Uri $Uri -AccessToken $Token
If (!($RoomMailboxes)) {Write-Host "No room mailboxes found - exiting" ; break}

# Find workspaces
$Uri = "https://graph.microsoft.com/beta/places/microsoft.graph.workspace"
[array]$WorkSpaces = Get-GraphData -Uri $Uri -AccessToken $Token

# Combine workspaces with room mailboxes if any are found
If ($WorkSpaces) { $RoomMailboxes = $RoomMailboxes + $WorkSpaces }

Write-Host ("Scanning room mailboxes for calendar events from {0} to {1}" -f $StartDate, $EndDate)
$CalendarInfo = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Room in $RoomMailboxes) {
    $Uri = "https://graph.microsoft.com/V1.0/users/" + $Room.emailAddress + "/calendar/calendarView?startDateTime=" + $Start + "&endDateTime=" + $End
    [array]$CalendarData = Get-GraphData -Uri $Uri -AccessToken $Token
    # Drop cancelled events
    $CalendarData = $CalendarData | Where-Object {$_.isCancelled -eq $False}
    
   If ($CalendarData) { # Found some events  
  
     Write-Host ("Found {0} calendar events for the {1} room" -f $CalendarData.Count, $Room.DisplayName)
     ForEach ($Event in $CalendarData) {
        [datetime]$MeetingStart =  Get-Date($Event.start.datetime) 
        [datetime]$MeetingEnd   = Get-Date($Event.end.datetime)

        # Calculate meeting duration in minutes. If it's an all-day event, use 480 minutes
        If ($Event.IsAllDay -eq $False) {
            $Duration =  ($MeetingEnd - $MeetingStart).TotalMinutes 
         }  Else { 
            $Duration = 480 
         }
     
        [array]$AllAttendees = ($Event.Attendees | Where-Object {$_.Type -ne "resource"} )
        [array]$RequiredAttendees = ($Event.Attendees | Where-Object {$_.Type -eq "required"}) 
        [array]$OptionalAttendees = ($Event.Attendees | Where-Object {$_.Type -eq "optional"})
        # Create output line - add one to the total attendees to account for the organizer
        $DataLine = [PSCustomObject] @{
          Room              = $Room.displayName
          Mail              = $Room.emailAddress
          Type              = $Event.type
          Organizer         = $Event.organizer.emailaddress.name
          OrganizerEmail    = $Event.organizer.emailaddress.address
          Created           = Get-Date($Event.createdDateTime) -format g
          Modified          = Get-Date($Event.lastModifiedDateTime) -format g
          TimeZone          = $Event.originalStartTimeZone
          Subject           = $Event.Subject
          AllDay            = $Event.IsAllDay
          Online            = $Event.isOnlineMeeting
          OnlineProvider    = $Event.onlineMeetingProvider
          Start             = Get-Date($MeetingStart) -format g
          End               = Get-Date($MeetingEnd) -format g
          Duration          = $Duration
          Location          = $event.location.displayname
          RequiredAttendees = $RequiredAttendees.emailaddress.name -join ", "
          OptionalAttendees = $OptionalAttendees.emailaddress.name -join ", "
          TotalAttendees    = $AllAttendees.Count
          Required          = $RequiredAttendees.Count
          Optional          = $OptionalAttendees.Count
          TotalAtEvent      = $AllAttendees.Count + 1
          EventId           = $Event.Id }
       $CalendarInfo.Add($DataLine)

     } #End ForEach Event
    } #End if
} #End ForEach Room

$TotalEvents = $CalendarInfo.Count
[array]$TopRooms = $CalendarInfo | Group-Object Room -NoElement | Sort-Object Count -Descending | Select-Object Name, Count
[array]$TopOrganizers = $CalendarInfo | Group-Object Organizer -NoElement | Sort-Object Count -Descending | Select-Object Name, Count
[array]$OnlineMeetings = $CalendarInfo | Where-Object {$_.Online -eq $True}
[array]$Rooms = $CalendarInfo | Sort-Object Room -Unique | Select-Object -ExpandProperty Room
$PercentOnline = ($OnlineMeetings.Count/$TotalEvents).toString("P")

# Calculate per-room summary data
$RoomSummary = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Room in $Rooms) {
   [array]$RoomData = $CalendarInfo | Where-Object {$_.Room -eq $Room} 
   [array]$RoomOnlineEvents = $RoomData | Where-Object {$_.Online -eq $True}
   [array]$RoomAllDayEvents = $RoomData | Where-Object {$_.IsAllDay -eq $True}
   $TotalMinutes = ($RoomData.Duration | measure-object -sum).Sum
   $TotalRoomAttendees = ($RoomData.TotalAtEvent | Measure-Object -Sum).Sum
   $AverageDuration = $TotalMinutes/$RoomData.Count
   $AverageAttendees = $TotalRoomAttendees/$RoomData.Count
   $AverageEventsPerDay = $RoomData.Count/$ReportingPeriodDays
   $RoomDataLine = [PSCustomObject] @{   
       Room                 = $Room
       Events               = $RoomData.Count
       "Avg Events/day"     = $AverageEventsPerDay.ToString("#.##")
       "Total Minutes"      = $TotalMinutes
       "Avg Event Duration" = $AverageDuration.Tostring("#.#")
       "Online Events"      = $RoomOnlineEvents.Count
       "All-day Events"     = $RoomAllDayEvents.Count
       "Total attendees"    = $TotalRoomAttendees
       "Average attendees"  = $AverageAttendees.Tostring("#.#")
   }
   $RoomSummary.Add($RoomDataLine)
}

Write-Host ""
Write-Host ("Meeting Room Statistics from {0} to {1}" -f $StartDate, $EndDate)
Write-Host "-----------------------------------------------------------------------"
Write-Host ""
Write-Host "Total events found: " $TotalEvents
Write-Host "Online events:      " $OnlineMeetings.Count "" $PercentOnline
Write-Host ""
Write-Host "Most popular rooms"
Write-Host "------------------"
$TopRooms | Format-Table Name, Count -AutoSize
Write-Host "Most active meeting organizers"
Write-Host "------------------------------"
$TopOrganizers | Format-Table Name, Count -AutoSize
Write-Host "Individual Room Statistics"
Write-Host "--------------------------"
$RoomSummary | Format-Table Room, Events, "Avg events/day", "Total minutes", "Avg Event Duration", "Total Attendees", "Average Attendees" -AutoSize

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the nternet without first validating the code in a non-production environment. 
