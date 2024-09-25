# ReportRoomMailboxUsage.PS1
# A script to report how busy room mailboxes are
# V1.1 8-Feb-2023 - Updated to include daily usage pattern statistics
# A version that uses the Graph SDK is available at https://github.com/12Knocksinna/Office365itpros/blob/master/Report-RoomMailboxUsage.PS1

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

# Define the values applicable for the application used to connect to the Graph
$AppId = "a6a7d55c-a847-443d-b7b9-f24b67ec4709"
$TenantId = "c662313f-14fc-43a2-9a7a-d2e27f4f3478"
$AppSecret = '_O28Q~.QZNE5QJ4pOCaxqTSx13dbkDI2_2Ns5blI'

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

Function Generate-DayString {
  [cmdletbinding()]
    Param(
        [string]$InputDay,
        [int]$DayEvents,  
        [int]$TotalEvents)

   $Balls = ($DayEvents/$TotalEvents) * 100
   $PercentEvents = ($DayEvents/$TotalEvents).toString("P")
   [int]$P = ($Balls/2)
   If ($Balls -eq 0) { $G = $Null 
    } Else {     
    [int]$i = 0; [string]$G = $Null
    Do {
      
      $G = $G + "o"; $i++
 
    } While ($i -lt $P)
   }
     
   $OutputString = ("{0} events: {1} ({2}) `t{3}>" -f $InputDay, $DayEvents, $PercentEvents, $G)

Return $OutputString
}

# End functions - start doing some real work
#

$Token = GetAccessToken
If (!($Token)) {
    Write-Host "Can't get a valid Entra ID access token - exiting" ; break 
}

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
$Uri = "https://graph.microsoft.com/V1.0/places/microsoft.graph.room"
[Array]$RoomMailboxes = Get-GraphData -Uri $Uri -AccessToken $Token
If (!($RoomMailboxes)) {Write-Host "No room mailboxes found - exiting" ; break}

# Find workspaces
$Uri = "https://graph.microsoft.com/beta/places/microsoft.graph.workspace"
[array]$WorkSpaces = Get-GraphData -Uri $Uri -AccessToken $Token

# Combine workspaces with room mailboxes if any are found
If ($WorkSpaces) { 
    $RoomMailboxes = $RoomMailboxes + $WorkSpaces 
}

$RoomMailboxes = $RoomMailboxes | Where-Object {$_.EmailAddress -ne $Null}

Write-Host ("Scanning room mailboxes for calendar events from {0} to {1}" -f $StartDate, $EndDate)
$CalendarInfo = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Room in $RoomMailboxes) {
    $Data = $false # Assume no data in the targeted range
    $Uri = ("https://graph.microsoft.com/V1.0/users/{0}/calendar/calendarView?startDateTime={1}&endDateTime={2}" -f $Room.emailAddress, $Start, $End)
    [array]$CalendarData = Get-GraphData -Uri $Uri -AccessToken $Token
    # Drop cancelled events - if you want to exclude private events from the set, use Where-Object {$_.isCancelled -eq $False -and $_.sensitivity -ne "private"}
    $CalendarData = $CalendarData | Where-Object {$_.isCancelled -eq $False}
    # This code makes sure that we only attempg to report data when the Graph returns some calendar data for the room
    If (!($CalendarData) -or $CalendarData[0].'@odata.context') { 
        $Data = $false 
    } Else {
        $Data = $true
    }
    If ($Data) {
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
          Day               = (Get-Date($MeetingStart)).DayOfWeek
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
   [array]$RoomAllDayEvents = $RoomData | Where-Object {$_.AllDay -eq $True}
   $TotalMinutes = ($RoomData.Duration | measure-object -sum).Sum
   $TotalRoomAttendees = ($RoomData.TotalAtEvent | Measure-Object -Sum).Sum
   $AverageDuration = $TotalMinutes/$RoomData.Count
   $AverageAttendees = $TotalRoomAttendees/$RoomData.Count
   $AverageEventsPerDay = $RoomData.Count/$ReportingPeriodDays
   $OverAllUsage = ($RoomData.Count/$CalendarInfo.Count).toString("P")
   # Extract meetings for each day of the week
   [array]$Monday = $RoomData | Where-Object {$_.Day -eq "Monday"}
   [array]$Tuesday = $RoomData | Where-Object {$_.Day -eq "Tuesday"}
   [array]$Wednesday = $RoomData | Where-Object {$_.Day -eq "Wednesday"}
   [array]$Thursday = $RoomData | Where-Object {$_.Day -eq "Thursday"}
   [array]$Friday = $RoomData | Where-Object {$_.Day -eq "Friday"}
   [array]$Saturday = $RoomData | Where-Object {$_.Day -eq "Saturday"}
   [array]$Sunday = $RoomData | Where-Object {$_.Day -eq "Sunday"}
   # Generate a basic graph for the room usage per day
   $MondayOutput = Generate-DayString -InputDay "Monday" -DayEvents $Monday.Count -TotalEvents $RoomData.Count
   $TuesdayOutput = Generate-DayString -InputDay "Tuesday" -DayEvents $Tuesday.Count -TotalEvents $RoomData.Count
   $WednesdayOutput = Generate-DayString -InputDay "Wednesday" -DayEvents $Wednesday.Count -TotalEvents $RoomData.Count
   $ThursdayOutput = Generate-DayString -InputDay "Thursday" -DayEvents $Thursday.Count -TotalEvents $RoomData.Count
   $FridayOutput = Generate-DayString -InputDay "Friday" -DayEvents $Friday.Count -TotalEvents $RoomData.Count
   $SaturdayOutput = Generate-DayString -InputDay "Saturday" -DayEvents $Saturday.Count -TotalEvents $RoomData.Count
   $SundayOutput = Generate-DayString -InputDay "Sunday" -DayEvents $Sunday.Count -TotalEvents $RoomData.Count
         
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
       "% Overall use"      = $OverAllUsage
       Monday               = $MondayOutput
       Tuesday              = $TuesdayOutput
       Wednesday            = $WednesdayOutput
       Thursday             = $ThursdayOutput
       Friday               = $FridayOutput
       Saturday             = $SaturdayOutput
       Sunday               = $SundayOutput
   }
   $RoomSummary.Add($RoomDataLine)
}

# Generate the overall usage pattern across all rooms
[array]$Monday = $CalendarInfo | Where-Object {$_.Day -eq "Monday"}
[array]$Tuesday = $CalendarInfo | Where-Object {$_.Day -eq "Tuesday"}
[array]$Wednesday = $CalendarInfo | Where-Object {$_.Day -eq "Wednesday"}
[array]$Thursday = $CalendarInfo | Where-Object {$_.Day -eq "Thursday"}
[array]$Friday = $CalendarInfo | Where-Object {$_.Day -eq "Friday"}
[array]$Saturday = $CalendarInfo | Where-Object {$_.Day -eq "Saturday"}
[array]$Sunday = $CalendarInfo | Where-Object {$_.Day -eq "Sunday"}
$MondayOutput = Generate-DayString -InputDay "Monday" -DayEvents $Monday.Count -TotalEvents $TotalEvents
$TuesdayOutput = Generate-DayString -InputDay "Tuesday" -DayEvents $Tuesday.Count -TotalEvents $TotalEvents
$WednesdayOutput = Generate-DayString -InputDay "Wednesday" -DayEvents $Wednesday.Count -TotalEvents $TotalEvents
$ThursdayOutput = Generate-DayString -InputDay "Thursday" -DayEvents $Thursday.Count -TotalEvents $TotalEvents
$FridayOutput = Generate-DayString -InputDay "Friday" -DayEvents $Friday.Count -TotalEvents $TotalEvents
$SaturdayOutput = Generate-DayString -InputDay "Saturday" -DayEvents $Saturday.Count -TotalEvents $TotalEvents
$SundayOutput = Generate-DayString -InputDay "Sunday" -DayEvents $Sunday.Count -TotalEvents $TotalEvents

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

Write-Host ""
Write-Host "Daily usage pattern across all room mailboxes"
$MondayOutput
$TuesdayOutput
$WednesdayOutput
$ThursdayOutput
$FridayOutput
$SaturdayOutput
$SundayOutput

Write-Host ""
Write-Host "Individual Room Statistics"
Write-Host "--------------------------"
$RoomSummary | Format-Table Room, Events, "Avg events/day", "Total minutes", "Avg Event Duration", "Total Attendees", "Average Attendees" -AutoSize

ForEach ($Room in $Rooms) {
   Write-Host ("Daily usage pattern for {0}" -f $Room)
   $RoomSummary | Where-Object {$_.Room -eq $Room} | Select-Object -ExpandProperty Monday
   $RoomSummary | Where-Object {$_.Room -eq $Room} | Select-Object -ExpandProperty Tuesday
   $RoomSummary | Where-Object {$_.Room -eq $Room} | Select-Object -ExpandProperty Wednesday
   $RoomSummary | Where-Object {$_.Room -eq $Room} | Select-Object -ExpandProperty Thursday
   $RoomSummary | Where-Object {$_.Room -eq $Room} | Select-Object -ExpandProperty Friday
   $RoomSummary | Where-Object {$_.Room -eq $Room} | Select-Object -ExpandProperty Saturday
   $RoomSummary | Where-Object {$_.Room -eq $Room} | Select-Object -ExpandProperty Sunday
   Write-Host ""
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
