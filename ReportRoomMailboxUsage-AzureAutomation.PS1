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

# End functions - start doing some real work
#
# Get an access token for the managed identity
$ResourceURL = "https://graph.microsoft.com/" 
$Response = [System.Text.Encoding]::Default.GetString((Invoke-WebRequest -UseBasicParsing -Uri "$($env:IDENTITY_ENDPOINT)?Resource=$ResourceURL" -Method 'GET' -Headers @{'X-IDENTITY-HEADER' = "$env:IDENTITY_HEADER"; 'Metadata' = 'True'}).RawContentStream.ToArray()) | ConvertFrom-Json 
$AccessToken = $Response.access_token 

# Make sure that you change this line to point to your tenant
Connect-ExchangeOnline -ManagedIdentity -Organization office365itpros.onmicrosoft.com

$StartDate = (Get-Date).AddDays(-60)
$EndDate = (Get-Date).AddDays(1)

$Start = Get-Date($StartDate) -format s
$End = Get-Date($EndDate) -format s
$ReportingPeriodDays = (($EndDate - $StartDate).Days)-1

# Find room mailboxes - this ignores room mailboxes marked as workspaces
$Uri = "https://graph.microsoft.com/beta/places/microsoft.graph.room"
[Array]$RoomMailboxes = Get-GraphData -Uri $Uri -AccessToken $AccessToken
If (!($RoomMailboxes)) {Write-Output "No room mailboxes found - exiting" ; break}

# Find workspaces
$Uri = "https://graph.microsoft.com/beta/places/microsoft.graph.workspace"
[array]$WorkSpaces = Get-GraphData -Uri $Uri -AccessToken $AccessToken

# Combine workspaces with room mailboxes if any are found
If ($WorkSpaces) { $RoomMailboxes = $RoomMailboxes + $WorkSpaces }

Write-Output ("Scanning room mailboxes for calendar events from {0} to {1}" -f $StartDate, $EndDate)
$CalendarInfo = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Room in $RoomMailboxes) {
    $Uri = "https://graph.microsoft.com/V1.0/users/" + $Room.emailAddress + "/calendar/calendarView?startDateTime=" + $Start + "&endDateTime=" + $End
    [array]$CalendarData = Get-GraphData -Uri $Uri -AccessToken $AccessToken
    # Drop cancelled events - if you want to exclude private events from the set, use Where-Object {$_.isCancelled -eq $False -and $_.sensitivity -ne "private"}
    $CalendarData = $CalendarData | Where-Object {$_.isCancelled -eq $False}
    
   If ($CalendarData) { # Found some events  
  
     Write-Output ("Found {0} calendar events for the {1} room" -f $CalendarData.Count, $Room.DisplayName)
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

Write-Output ""
Write-Output ("Meeting Room Statistics from {0} to {1}" -f $StartDate, $EndDate)
Write-Output "-----------------------------------------------------------------------"
Write-Output ""
Write-Output "Total events found: " $TotalEvents
Write-Output "Online events:      " $OnlineMeetings.Count "" $PercentOnline
Write-Output ""
Write-Output "Most popular rooms"
Write-Output "------------------"
$TopRooms | Format-Table Name, Count -AutoSize
Write-Output "Most active meeting organizers"
Write-Output "------------------------------"
$TopOrganizers | Format-Table Name, Count -AutoSize
Write-Output "Individual Room Statistics"
Write-Output "--------------------------"
$RoomSummary | Format-Table Room, Events, "Avg events/day", "Total minutes", "Avg Event Duration", "Total Attendees", "Average Attendees" -AutoSize
Write-Output ""
Write-Output "Access to room mailbox calendars granted by RBAC for Applications"

