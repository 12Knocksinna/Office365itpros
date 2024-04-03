# Report-WeeklyCalendarForRoomMailboxes.PS1
# Example script showing how to generate a weekly events report for room mailboxes
# App requires Calendar.Read and Place.read.all permissions

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
    $AppId = "66a7d55c-a847-443d-b7b9-f24b67ec4709"
    $TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f3478"
    $AppSecret = 'B7h8Q~pzvflr5pPb.4mgzOWS_hLuK5l4qbkRMcqv'
    
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
    
$Token = GetAccessToken
If (!($Token)) {
    Write-Host "Can't get a valid Entra ID access token - exiting" ; break
}
$Headers = @{
    'Content-Type'  = "application\json"
    'Authorization' = "Bearer $Token" 
    'ConsistencyLevel' = "eventual" }

$HtmlReportFile = "C:\Temp\RoomMailboxReport.html"
$CSVOutputFile = "C:\Temp\RoomMailboxReport.CSV"
[datetime]$Today = (Get-Date).Date
[datetime]$NextMonday =  Get-Date($Today.AddDays(2 + $Today.DayOfWeek.value__)) -format 'dd-MMM-yyyy'
[datetime]$NextSunday = Get-Date($NextMonday).AddDays(7) -format 'dd-MMM-yyyy'
# Sortable dates that can be used with the Graph request to find calendar data
$StartDate = Get-Date($NextMonday) -format s
$EndDate = Get-Date($NextSunday) -format s

Write-Host ("Searching for events scheduled in room mailboxes between {0} and {1}" -f $NextMonday, $NextSunday) -ForegroundColor Red
# Find room mailboxes - this ignores room mailboxes marked as workspaces
$Uri = "https://graph.microsoft.com/V1.0/places/microsoft.graph.room"
[Array]$RoomMailboxes = Get-GraphData -Uri $Uri -AccessToken $Token
If (!($RoomMailboxes)) {
    Write-Host "No room mailboxes found - exiting" ; break
} Else {
    $RoomMailboxes = $RoomMailboxes | Where-Object {$null -ne $_.EmailAddress}
    $RoomMailboxes = $RoomMailboxes | Sort-Object DisplayName
}
Write-Host ("{0} room mailboxes found" -f $RoomMailboxes.Count)
$CalendarInfo = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Room in $RoomMailboxes) {
    $Data = $false # Assume no data in the targeted range
    $Uri = ("https://graph.microsoft.com/V1.0/users/{0}/calendar/calendarView?startDateTime={1}&endDateTime={2}" -f $Room.emailAddress, $StartDate, $EndDate)
    [array]$CalendarData = Get-GraphData -Uri $Uri -AccessToken $Token
    # Drop cancelled events - if you want to exclude private events from the set, use Where-Object {$_.isCancelled -eq $False -and $_.sensitivity -ne "private"}
    $CalendarData = $CalendarData | Where-Object {$_.isCancelled -eq $False}
    # This code makes sure that we only attempg to report data when the Graph returns some calendar data for the room
    If (!($CalendarData) -or $CalendarData[0].'@odata.context') { 
        $Data = $false 
    } Else {
        $Data = $true
    }
    If ($Data) { # Found some events  
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
                Created           = Get-Date($Event.createdDateTime) -format 'dd-MMM-yyyy HH:mm'
                Modified          = Get-Date($Event.lastModifiedDateTime) -format 'dd-MMM-yyyy HH:mm'
                TimeZone          = $Event.originalStartTimeZone
                Subject           = $Event.Subject
                AllDay            = $Event.IsAllDay
                Online            = $Event.isOnlineMeeting
                OnlineProvider    = $Event.onlineMeetingProvider
                Start             = Get-Date($MeetingStart) -format 'dd-MMM-yyyy HH:mm'
                End               = Get-Date($MeetingEnd) -format 'dd-MMM-yyyy HH:mm'
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

# Create the HTML report
$HtmlBody = $null
$RunDate = (Get-Date).ToString("dd-MMM-yyyy HH:mm:ss")
$StartDate = Get-Date $StartDate -format 'dd-MMM-yyyy'
$EndDate = Get-Date $EndDate -format 'dd-MMM-yyyy'
$HtmlHead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 10pt;}
	   H1{font-size: 32px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 24px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 20px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.warn{background: #FFF275;}
	   td.fail{background: #FF2626; color: #ffffff;}
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Room Mailboxes Scheduling Report</h1></p>
           <p><h2>From <b>" + $StartDate + "</b> to <b>" + $EndDate + "</b></h2></p>"

# Loop through the records found for room mailboxes and create a section for each room
ForEach ($Room in $RoomMailboxes) {
    [array]$CalendarEvents = $CalendarInfo | Where-Object {$_.Room -eq $Room.DisplayName}
    If ($CalendarEvents) {
        $HtmlRoomSeparator = "<p><h2>Room: <b>" + $Room.DisplayName + "</b></h2></p>"   
        $HtmlRoomData = $CalendarEvents | Select-Object Organizer, Subject, Start, End, Duration, Online, TotalAttendees | `
            ConvertTo-Html -Fragment            
        $HtmlBody += $HtmlRoomSeparator + $HtmlRoomData
    }
}

# Finish up
$HtmlBody = $HtmlBody + "<p>Report created: " + $RunDate + "</p>" 
$HtmlReport = $HtmlHead + $HtmlBody + "</div></body></html>"
$HtmlReport | Out-File $HtmlReportFile -Encoding UTF8
Write-Host ("The output HTML report is available in {0}" -f $HtmlReportFile) -ForegroundColor Red
$CalendarInfo | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding UTF8  
Write-Host ("The data for calendar events for each room is available in {0}" -f $CSVOutputFile) -ForegroundColor Red                          

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 