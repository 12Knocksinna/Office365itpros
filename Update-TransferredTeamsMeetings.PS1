# Update-TransferredTeamsMeetings.PS1
# This script updates the Teams online meeting information for meetings transferred from another user to make the target user
# the organizer of the meeting. It is intended to be run in a PowerShell session with the necessary permissions to access and modify Teams meetings.

# OnlineMeetings.ReadWrite.All
# Calendars.ReadWrite
# Users.ReadBasic.All

# In addition, the registered app used by this script must be permitted by a Teams application access policy to access the
# calendar where the meetings to be updated are stored.

# V1.0 6-Aug-2026
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Update-TransferredTeamsMeetings.PS1

# The input is a CSV file containing the following columns:
# Subject
# StartTime
# Newrganizer

# The CSV file output from the https://github.com/12Knocksinna/Office365itpros/blob/master/Transfer-MeetingOwnership.PS1 script can be used as input.

[CmdletBinding()]
param(
    [string]$TransferInputFile

)

# Helper function to turn dates stored with a date/time and time zone into a UTC time.

function Get-UtcEventTime {
    param(
        $InputDate,
        $InputTimeZone
    )

    [datetime]$EventTime = $InputDate
    [string]$EventTimeZone = $InputTimeZone

    if ($EventTimeZone -eq 'UTC') {
        return $EventTime
    }

    $Zone = [System.TimeZoneInfo]::FindSystemTimeZoneById($EventTimeZone)

    return [System.TimeZoneInfo]::ConvertTimeToUtc($EventTime, $Zone)
}

# Test do we have a valid input file to process?
# $TransferInputFile = "c:\temp\Transferred Teams Meetings.csv"
If (-not (Test-Path $TransferInputFile)) {
    Write-Host ("The input file '{0}' does not exist. Please provide a valid CSV file." -f $TransferInputFile) -ForegroundColor Red
    return
}   

# Input the data and see if we have any meetings to process.
[array]$InputMeetings = Import-Csv -Path $TransferInputFile
If (-not $InputMeetings) {
    Write-Host ("The input file '{0}' does not contain any data. Please provide a valid CSV file." -f $TransferInputFile) -ForegroundColor Red
    return
}   Else {
    Write-Host ("The input file '{0}' contains {1} meetings to be updated." -f $TransferInputFile, $InputMeetings.Count) -ForegroundColor Green
}

$AppId = '03ef5d75-e8dd-4f9e-a71c-fa81939d5ca2' # Replace with your Entra ID app IDcd
$Thumbprint = "0CF6CE3F3548FD73E7AC8CF20226ED447E125C71" # Replace with your certificate thumbprint
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3478" # Replace with your tenant ID

# Connect to the Microsoft Graph in app-only mode
Write-Host "Connecting to Microsoft Graph..."
Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome

# Do we have a valid target mailbox?
$TargetMailbox = $InputMeetings | Select-Object -ExpandProperty NewOrganizer | Select-Object -Unique
Try {
    $TargetMailboxDetails = Get-MgUser -UserId $TargetMailbox -ErrorAction Stop
} Catch {
    Write-Host ("The target mailbox '{0}' does not exist or is not a user mailbox. Please provide a valid target mailbox." -f $TargetMailbox) -ForegroundColor Red
    return
}

ForEach ($Item in $InputMeetings) {
    # Set up the details to search the calendar for the target meeting
    $ItemStartDate = $Item.StartTime.Split(' ')[0] + " " + $Item.StartTime.Split(' ')[1]
    $ItemStartTimeZone = $Item.StartTime.Split(' ')[-1]
    $EventStartDate = Get-UtcEventTime $ItemStartDate $ItemStartTimeZone
    $ItemEndDate = $Item.EndTime.Split(' ')[0] + " " + $Item.EndTime.Split(' ')[1]
    $ItemEndTimeZone = $Item.EndTime.Split(' ')[-1]       
    $EventEndDate = Get-UtcEventTime $ItemEndDate $ItemEndTimeZone
    $SearchDateStart = $EventStartDate.toString("yyyy-MM-dd")
    $SearchDateEnd = (Get-Date $SearchDateStart).AddDays(1).toString("yyyy-MM-dd")
    $EventSubject = $Item.Subject.trim()
    # Attempt to find the target event
    [array]$Events = Get-MgUserDefaultCalendarEvent -UserId $TargetMailboxDetails.Id -Filter "start/datetime ge '$SearchDateStart' and end/datetime le '$SearchDateEnd' and subject eq '$EventSubject'"
    If ($Events.Count -eq 1) {
        $CalendarEvent = $Events[0]
        Write-Host ("Found target event '{0}' on {1} for update." -f $CalendarEvent.Subject, $CalendarEvent.Start.DateTime) -ForegroundColor Green
    } Else {
        Write-Host ("{0} matching events found. We cannot update the meeting." -f $Events.count)
        Continue
    }

    $ZeroEventDetails = "<p><b><h2>Microsoft Teams meeting information removed</h2></b></p>"
    # Remove the current Teams meeting information from the event
    $EventBody = @{}
    $EventBody.Add("contentType", "HTML")
    $EventBody.Add("content", $ZeroEventDetails)

    $EventDetails = @{}
    #$EventDetails.Add("isOnlineMeeting", $false)
    $EventDetails.Add("body", $EventBody)
    #$EventDetails.Add("onlineMeetingProvider", "null")
    Try {
        $Status = Update-MgUserEvent -EventId $CalendarEvent.Id -UserId $TargetMailboxDetails.Id -BodyParameter $EventDetails -ErrorAction Stop
    } Catch {
        Write-Host ("Error removing old Teams join information from event '{0}': {1}" -f $CalendarEvent.Subject, $_.Exception.Message) -ForegroundColor Red
        Continue
    }
    
    # Now attempt to create a new Teams online event.
    $OnlineMeetingBody = @{
        subject       = $CalendarEvent.Subject
        startDateTime = $EventStartDate.ToString("yyyy-MM-ddTHH:mm:ss")
        endDateTime   = $EventEndDate.ToString("yyyy-MM-ddTHH:mm:ss")

        lobbyBypassSettings = @{
            scope = "everyone"                
        }
        allowedPresenters = "everyone"        
    }
    # Can’t use UPN - the cmdlet only takes the user identifier
    Try {
        $onlineMeeting = New-MgUserOnlineMeeting -UserId $TargetMailboxDetails.Id -BodyParameter $OnlineMeetingBody -ErrorAction Stop
    } Catch {
        Write-Host ("Error creating online meeting for event '{0}': {1}" -f $CalendarEvent.Subject, $_.Exception.Message) -ForegroundColor Red
        Continue
    }
   
    # The join information is stored in the JoinInformation.Content property, which is a URL-encoded HTML string. We need to decode it to extract the join URL, meeting ID, and passcode.
    Add-Type -AssemblyName System.Web
    # Decode the meeting content
    $Html = [System.Uri]::UnescapeDataString(($onlineMeeting.JoinInformation.Content -replace '^data:text/html,'))

    # Extract Join URL
    $JoinUrl = [Regex]::Match($Html,'https://teams\.microsoft\.com/meet/[^"&<\s]+').Value

    # Extract Meeting ID
    $MeetingId = ([Regex]::Match($Html,'\d{3}\+\d{3}\+\d{3}\+\d{3}\+\d{3}').Value) -replace '\+',' '

    # Extract Passcode
    $PasscodeSection = $Html.Substring($Html.IndexOf('Passcode'), 200)
    $Passcode = ([Regex]::Matches($PasscodeSection,'[A-Za-z0-9]{8}') | Select-Object -Last 1).Value
  
    # Now that we have an online meeting, we can update the calendar event with the new meeting information.
    $EventLink = "<p><b><h2>Microsoft Teams meeting</h2></b></p>"
    $EventLink += "<p><b><h3>Join:</b> <a href='$JoinUrl'>$JoinUrl</a></h3></b></p>"
    $EventLink += "<p><b><h3>Meeting ID:</b> $MeetingId</h3></p>"
    $EventLink += "<p><b><h3>Passcode:</b> $Passcode</h3></p>"
    $EventLink += "<p><p><b>Note:</b> This meeting was transferred from the previous organizer. The new organizer is $($TargetMailboxDetails.DisplayName).</p>"
    $EventLink += "<p>If you have any issues joining the meeting, please contact the new organizer.</p>"
    $EventBody = @{}
    $EventBody.Add("contentType", "HTML")
    $EventBody.Add("content", $EventLink)

    $EventDetails = @{}
    $EventDetails.Add("isOnlineMeeting", $true)
    $EventDetails.Add("body", $EventBody)
    $EventDetails.Add("onlineMeetingProvider", "teamsForBusiness")

    Try {
        $Status = Update-MgUserEvent -EventId $CalendarEvent.Id -UserId $TargetMailboxDetails.Id -BodyParameter $EventDetails -ErrorAction Stop
        If ($Status) {
            Write-Host ("Successfully updated event '{0}' with new Teams meeting information." -f $CalendarEvent.Subject) -ForegroundColor Green
        } Else {
            Write-Host ("Failed to update event '{0}' with new Teams meeting information." -f $CalendarEvent.Subject) -ForegroundColor Red
        }   
    } Catch {
        Write-Host ("Error updating event '{0}': {1}" -f $CalendarEvent.Subject, $_.Exception.Message) -ForegroundColor Red
        Continue
    }   

    # The code above specifies that the meeting is online and the meeting provider is Teams. This means that the calendar inserts 
    # details of another Teams meeting into the event. We need to remove that information to avoid confusion. The code below is a hack to
    # remove the additional Teams information from the event body by removing the content after the "please contact the new organizer" message.
    # This is not a perfect solution, but it works for now. If Microsoft changes the way they insert Teams information into calendar events,
    # this code may break. If that happens, please update the script.
    [array]$CalendarEvent = Get-MgUserEvent -UserId $TargetMailboxDetails.Id -EventId $Status.Id
    $Position = $CalendarEvent.body.content.indexOf("please contact the new organizer")
    $Position = $Position + 40
    If ($Position -gt 690) {
        $HtmlContent = $CalendarEvent.body.content.substring(0, $Position)
        $CalendarEvent.body.content = $HtmlContent
        $EventBody = @{}
        $EventBody.Add("contentType", "HTML")
        $EventBody.Add("content", $HtmlContent)

        $EventDetails = @{}
        $EventDetails.Add("body", $EventBody)

        Try {
            $Status = Update-MgUserEvent -EventId $CalendarEvent.Id -UserId $TargetMailboxDetails.Id -BodyParameter $EventDetails -ErrorAction Stop
        } Catch {
            Write-Host ("Error performing post processing of Teams information for event '{0}': {1}" -f $CalendarEvent.Subject, $_.Exception.Message) -ForegroundColor Red
            Continue
        }
    }

}

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.