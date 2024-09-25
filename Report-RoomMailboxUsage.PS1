# Report-RoomMailboxUsage.PS1
# A version of the original ReportRoomMailboxUsage script converted to replace Graph API requests 
# with cmdlets from the Microsoft Graph PowerShell SDK

# V1.0 19-Sep-2024  Conversion from Graph API to Graph SDK cmdlets

# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-RoomMailboxUsage.PS1

Function New-DayString {
    [cmdletbinding()]
      Param(
        [string]$InputDay,
        [int]$DayEvents,  
        [int]$TotalEvents)
  
    $Balls = ($DayEvents/$TotalEvents) * 100
    $PercentEvents = ($DayEvents/$TotalEvents).toString("P")
    [int]$P = ($Balls/2)
    If ($Balls -eq 0) { 
        $G = $Null 
    } Else {     
        [int]$i = 0; [string]$G = $Null
        Do {    
            $G = $G + "o"; $i++ 
      } While ($i -lt $P)
     }
       
    $OutputString = ("{0} events: {1} ({2}) `t{3}>" -f $InputDay, $DayEvents, $PercentEvents, $G)
  
  Return $OutputString
}
  
# Values to connect to the Graph in app-only mode. The Entra ID app must have 
# the Graph Calendar.Read.All and Place.Read.All application permissions
$AppId = 'x66a7d55c-a847-443d-b7b9-f24b67ec4709'
$TenantId = 'a662313f-14fc-43a2-9a7a-d2e27f4f3478'
$Thumbprint = 'F79286DB88C21491110109A0222348FACF694CBD'
Connect-MgGraph -NoWelcome -AppId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint

# Start and end dates for the reporting period
$StartDate = (Get-date).AddDays(-60).toString('dd-MMM-yyyy')
$EndDate = (Get-Date).AddDays(1).toString('dd-MMM-yyyy')
  
[datetime]$Start = Get-Date($StartDate) -format s
[datetime]$End = Get-Date($EndDate) -format s
$ReportingPeriodDays = (($End - $Start).Days)-1

Write-Host "Looking for room mailboxes and workspaces..."
[Array]$RoomMailboxes = Get-MgPlaceAsRoom -All -PageSize 500
If (!($RoomMailboxes)) {
    Write-Host "No room mailboxes found - exiting" ; break
}
# No SDK cmdlet available for workspaces, so we need to run a Graph API request
$Uri = "https://graph.microsoft.com/beta/places/microsoft.graph.workspace"
[array]$WorkSpaces = Invoke-MgGraphRequest -Uri $Uri -Method GET

$WorkSpaces = $WorkSpaces.Value
# Combine workspaces with room mailboxes if any are found
If ($WorkSpaces) { 
    $RoomMailboxes = $RoomMailboxes + $WorkSpaces
}
# Eliminate any room mailboxes or workspaces that don't have an email address
$RoomMailboxes = $RoomMailboxes | Where-Object {$_.EmailAddress -ne $Null} | Sort-Object DisplayName

Write-Host ("Scanning room mailboxes for calendar events from {0} to {1}" -f $StartDate, $EndDate)
$CalendarInfo = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Room in $RoomMailboxes) {
    $Data = $false # Assume no data in the targeted range
    [array]$CalendarData = Get-MgUserCalendarView -UserId $Room.emailAddress -StartDateTime $Start -EndDateTime $End -All -PageSize 250
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
        } Else { 
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
   $MondayOutput = New-DayString -InputDay "Monday" -DayEvents $Monday.Count -TotalEvents $RoomData.Count
   $TuesdayOutput = New-DayString -InputDay "Tuesday" -DayEvents $Tuesday.Count -TotalEvents $RoomData.Count
   $WednesdayOutput = New-DayString -InputDay "Wednesday" -DayEvents $Wednesday.Count -TotalEvents $RoomData.Count
   $ThursdayOutput = New-DayString -InputDay "Thursday" -DayEvents $Thursday.Count -TotalEvents $RoomData.Count
   $FridayOutput = New-DayString -InputDay "Friday" -DayEvents $Friday.Count -TotalEvents $RoomData.Count
   $SaturdayOutput = New-DayString -InputDay "Saturday" -DayEvents $Saturday.Count -TotalEvents $RoomData.Count
   $SundayOutput = New-DayString -InputDay "Sunday" -DayEvents $Sunday.Count -TotalEvents $RoomData.Count
         
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
$MondayOutput = New-DayString -InputDay "Monday" -DayEvents $Monday.Count -TotalEvents $TotalEvents
$TuesdayOutput = New-DayString -InputDay "Tuesday" -DayEvents $Tuesday.Count -TotalEvents $TotalEvents
$WednesdayOutput = New-DayString -InputDay "Wednesday" -DayEvents $Wednesday.Count -TotalEvents $TotalEvents
$ThursdayOutput = New-DayString -InputDay "Thursday" -DayEvents $Thursday.Count -TotalEvents $TotalEvents
$FridayOutput = New-DayString -InputDay "Friday" -DayEvents $Friday.Count -TotalEvents $TotalEvents
$SaturdayOutput = New-DayString -InputDay "Saturday" -DayEvents $Saturday.Count -TotalEvents $TotalEvents
$SundayOutput = New-DayString -InputDay "Sunday" -DayEvents $Sunday.Count -TotalEvents $TotalEvents

Write-Host ""
Write-Host ("Meeting Room Statistics from {0} to {1}" -f (Get-Date($Start) -format 'dd-MMM-yyyy'), (Get-Date($End) -format 'dd-MMM-yyyy'))
Write-Host "-------------------------------------------------------"
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

Write-Host ""
Write-Host "Analysis complete. Generating output files..."
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $RoomSummaryOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Room Mailboxes Usage Report.xlsx"
    $RoomSummary | Export-Excel -Path $RoomSummaryOutputFile -WorksheetName "Room Mailboxes Usage" `
        -Title ("Room Mailboxes Usage Report from {0} to {1}" -f (Get-Date($Start) -format 'dd-MMM-yyyy'), (Get-Date($End) -format 'dd-MMM-yyyy')) -TitleBold -TableName "RoomMailboxesUsage"  
    $TopOrganizersOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Top Room Organizers.xlsx"
    $TopOrganizers | Export-Excel -Path $TopOrganizersOutputFile -WorksheetName "Top Room Organizers" `
        -Title ("Top Room Organizers Report from {0} to {1}" -f (Get-Date($Start) -format 'dd-MMM-yyyy'), (Get-Date($End) -format 'dd-MMM-yyyy')) -TitleBold -TableName "TopRoomOrganizers" 
} Else {
    $RoomSummaryCSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Room Mailboxes Usage Report.CSV"
    $RoomSummary | Export-Csv -Path $RoomSummaryCSVOutputFile -NoTypeInformation -Encoding Utf8
    $TopOrganizersCSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Top Room Organizers Report.CSV"
    $TopOrganizers | Export-Csv -Path $TopOrganizersCSVOutputFile -NoTypeInformation -Encoding Utf8
}

If ($ExcelGenerated) {
    Write-Host ("Excel worksheets generated in your Downloads folder: {0}, {1}" -f $RoomSummaryOutputFile, $TopOrganizersOutputFile)
} Else {
    Write-Host ("CSV files generated in your Downloads folder: {0}, {1}" -f $RoomSummaryCSVOutputFile, $TopOrganizersCSVOutputFile)
} 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 