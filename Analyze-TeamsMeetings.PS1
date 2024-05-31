# Analyze-TeamsMeetings.PS1
# a script to show how to use Graph APIs to analyze Teams meetings and attendance reports for the last 60 days
# V1.0 28-May-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Analyze-TeamsMeetings.PS1

# Connect to the Graph using an appId, tenantId, and certificate thumbprint
# The app must have the necessary application permissions to read calendar data and online meeting information
# An application access policy must be in the place to allow the app to access online meeting data for the target accounts
# Insert the correct values for your tenant and app
$AppId = "2817e4bc-69dd-4cfb-9411-c3f661717208"
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f3478"
$CertThumbprint = "A1B2C3D4E5F6A7B8C9D0E1F2A3B4C5D6E7F8A9B0"
#Permissions required:
# Calendars.Read = Read calendar data
# Group.Read.All = Read group membership
# OnlineMeetings.Read.All = Read online meeting data
# OnlineMeetingArtifact.Read.All = Read attendance reports
# CrossTenantInformation.ReadBasic.All = Read tenant information for external (federated) participants
# Organization.Read.All = Read organization information

Connect-MgGraph -AppId $AppId -TenantId $TenantId -CertificateThumbprint $CertThumbprint -NoWelcome
$Organization = Get-MgOrganization
$TenantName = $Organization.DisplayName

$StartDate = (Get-Date).AddDays(-60)
$EndDate = (Get-Date)
$StartDateSearch = Get-Date $StartDate -format s
$EndDateSearch = Get-Date $EndDate -format s
$HtmlReportFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\TeamsOnlineMeetingReport.html"

# Get members of the group that we want to process - change the name of the group to match your tenant
[array]$TeamsOrganizers = Get-MgGroupMember -GroupId (Get-MgGroup -Filter "displayName eq 'Teams Meeting Organizers'").Id

$MeetingReport = [System.Collections.Generic.List[Object]]::new()
$AttendanceData = [System.Collections.Generic.List[Object]]::new()
ForEach ($Organizer in $TeamsOrganizers) {
    $DisplayName = $Organizer.additionalProperties.displayName
    Write-Host ("Checking online Teams meetings for {0}" -f $DisplayName)
    [array]$CalendarItems = Get-MgUserCalendarView -UserId $Organizer.id -Startdatetime $StartDateSearch -Enddatetime $EndDateSearch -All
    $CalendarItems = $CalendarItems | Where-Object {$_.isCancelled -eq $False -and $_.OnlineMeetingProvider -eq "teamsForBusiness" `
         -and $_.IsOrganizer -eq $true}
    If ($CalendarItems) {
        Write-Host ('Found {0} Teams meetings for {1}' -f $CalendarItems.Count, $DisplayName) -ForegroundColor Yellow
    } Else {
        Write-Host ('No Teams meetings found for {0}' -f $DisplayName)
        continue
    }
    Write-Host ("Analyzing Teams meetings for {0}..." -f $DisplayName)

    ForEach ($Item in $CalendarItems) {
        $MeetingDuration = $null; [array]$MeetingData = $null
        # Get the meeting URL
        $MeetingURL =  $Item.onlinemeeting.joinUrl.trim()
        $DecodedURL = [System.Web.HttpUtility]::UrlDecode($MeetingURL)
        $MeetingIdStart = $DecodedURL.IndexOf("19:") 
        $MeetingIdEnd = $DecodedURL.IndexOf("thread")
        $MeetingId = $DecodedURL.Substring($MeetingIdStart, $MeetingIdEnd - $MeetingIdStart +9)
        $MeetingIdLookup = ("1*{0}*0**{1}" -f $Organizer.id, $MeetingId)
        $Base64MeetingId = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($MeetingIdLookup))
        $Uri = ("https://graph.microsoft.com/v1.0/users/{0}/onlineMeetings/{1}" -f $Organizer.Id, $Base64MeetingId)
        Try {
            [array]$MeetingData = Invoke-MgGraphRequest -Method Get -Uri $Uri
        } Catch {
            Write-Host ("Ran into an issue retieving information for meeting {0} (created {1})" -f $Item.subject, $Item.CreatedDateTime, $($PSItem.ToString())) -ForegroundColor Red
            Continue
        }
        If (!($MeetingData)) {
            Write-Host ("No data found for meeting {0}" -f $MeetingId)
            continue    
        }
        $MeetingDuration = $MeetingData.endDateTime -$MeetingData.startDateTime
        $ReportLine = [PSCustomObject]@{
            MeetingId               = $MeetingId
            Organizer               = $DisplayName
            'Creation timestamp'    = (Get-Date $MeetingData.creationDateTime -format 'dd-MMM-yyyy HH:mm:ss')
            'Meeting start'         = (Get-Date $MeetingData.startDatetime -format 'dd-MMM-yyyy HH:mm')
            'Meeting snd'           = (Get-Date $MeetingData.endDateTime -format 'dd-MMM-yyyy HH:mm')
            'Time zone'             = $Item.originalStartTimeZone
            'Meeting duration'      = ("{0:hh\:mm\:ss}" -f $MeetingDuration)
            Subject                 = $MeetingData.subject
            'Allow presenters'      = $MeetingData.allowedpresenters
            'Lobby bypass'          = $MeetingData.lobbyBypassSettings.scope
            'Auto admit'            = $MeetingData.autoAdmittedUsers
            'Allow recording'       = $MeetingData.allowRecording
            'Record automatically'  = $MeetingData.recordAutomatically
            'Allow chat'            = $MeetingData.allowMeetingChat
            'Allow transcripton'    = $MeetingData.allowTranscription
            'Allow reactions'       = $MeetingData.allowTeamWorkReactions
            MeetingURL              = $MeetingURL
        }
        $MeetingReport.Add($ReportLine)
    }
    # attempt at throttling control
    Start-Sleep -Milliseconds 500

    # Now get the attendance reports for the meeting
    Write-Host "Fetching attendance reports for the meetings..."
    $Uri = $Uri = ("https://graph.microsoft.com/v1.0/users/{0}/onlineMeetings/{1}/attendanceReports" -f $Organizer.Id, $Base64MeetingId)
    [array]$AttendanceReports = Invoke-MgGraphRequest -Method get -Uri $Uri
    $AttendanceReports = $AttendanceReports.value
 
    If ($AttendanceReports) {
        ForEach ($AR in $AttendanceReports) {
          
            $AttendanceRecords = $null; $Participant = $null
            $Uri = $Uri = ("https://graph.microsoft.com/v1.0/users/{0}/onlineMeetings/{1}/attendanceReports/{2}?`$expand=attendanceRecords" -f `
                $Organizer.Id, $Base64MeetingId, $AR.Id)
            [array]$AttendanceRecords = Invoke-MgGraphRequest -Method get -Uri $Uri
            $AttendanceRecords = $AttendanceRecords.AttendanceRecords
            ForEach ($Participant in $AttendanceRecords) {  
                If ($Participant.identity.tenantId -eq $TenantId) {
                    $ParticipantTenantName = $TenantName
                } Else {
                    $LookUpTenantId = $Participant.identity.tenantId
                    $Uri = ("https://graph.microsoft.com/V1.0/tenantRelationships/findTenantInformationByTenantId(tenantId='{0}')" -f $LookUpTenantId)
                    $ExternalTenantData = Invoke-MgGraphRequest -Uri $Uri -Method Get
                    $ParticipantTenantName = $ExternalTenantData.displayName
                }
                $TimeInMeeting = [timespan]::fromseconds($Participant.totalAttendanceInSeconds)

                $ReportLine = [PSCustomObject]@{
                    MeetingId               = $AR.Id
                    'Number Participants'   = $AR.totalParticipantCount
                    Email                   = $Participant.emailAddress
                    DisplayName             = $Participant.identity.displayName
                    'Meeting start time'    = (Get-Date $AR.MeetingStartDateTime -format 'dd-MMM-yyyy HH:mm:ss')
                    'Meeting end time'      = (Get-Date $AR.MeetingEndDateTime -format 'dd-MMM-yyyy HH:mm:ss')
                    ParticipantId           = $Participant.Id
                    Role                    = $Participant.role   
                    'Join time'             = $Participant.attendanceIntervals.joinDateTime
                    'Leave time'            = $Participant.attendanceIntervals.leaveDateTime
                    'Attendance in seconds' = $Participant.totalAttendanceInSeconds
                    'Time in meeting'       = ("{0:hh\:mm\:ss}" -f $TimeInMeeting)
                    TenantId                = $Participant.identity.tenantId
                    'Tenant name'           = $ParticipantTenantName
                    MatchMeetingId          = $MeetingId
                }
                $AttendanceData.Add($ReportLine)              
            }   
             # Add the attendee data to the overall set
           #  $OverallMeetingAttendees += $AttendanceData   
       
        }
    }
}

# Now let's generate a HTML report
$HtmlBody = $null
$RunDate = (Get-Date).ToString("dd-MMM-yyyy HH:mm:ss")
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
           <div align=center>"
$HtmlHead1 = ("<p><h1>Teams Online Meeting Report for the {0} tenant</h1></p>" -f $TenantName)
$HtmlHead2 = ("<p><h2>Details extracted for period <b>{0}</b> to <b>{1}</b></h2></p>" -f (Get-Date $StartDate -format 'dd-MMM-yyyy HH:mm'), (Get-Date $EndDate  -format 'dd-MMM-yyyy HH:mm'))     
$HtmlHead = $HtmlHead + $HtmlHead1 + $HtmlHead2        

# Output meeting details

ForEach ($M in $MeetingReport) {
    $HtmlBody = $HtmlBody + "<h3>Meeting <u>{0}</u> organized by {1} for {2}</h3>" -f $M.Subject, $M.Organizer, $M.'Meeting start'
    $MeetingHtml = $M | Select-Object Organizer, 'Meeting start', 'Meeting end', 'Time zone', 'Meeting duration', 'Allow presenters',`
     'Lobby bypass', 'Auto admit', 'Allow recording', 'Record automatically', 'Allow chat', 'Allow transcripton', 'Allow reactions' | ConvertTo-HTML -Fragment
    $HtmlBody = $HtmlBody + $MeetingHtml

    # Output attendance details for the meeting
    [array]$Attendees = $AttendanceData | Where-Object {$_.MatchMeetingId -eq $M.MeetingId} | Select-Object `
       Email, DisplayName, Role, 'Join time', 'Leave time', 'Time in meeting', 'Tenant name' 
    If ($Attendees.count -eq 0) {
        $HtmlBody = $HtmlBody + "<p>No attendance data available for this meeting</p>"
    } Else {
        $AttendeesHTML = $Attendees | ConvertTo-HTML -Fragment 
        $HtmlBody = $HtmlBody + ("<h4>Attendance details for meeting <u>{0}</u></h4><p></p>" -f $M.Subject)
        $HtmlBody = $HtmlBody + $AttendeesHTML + "<p></p>"
    }
}

$HtmlBody = $HtmlBody + "<p>Report created: " + $RunDate + "</p>" 
$HtmlReport = $HtmlHead + $HtmlBody + "</div></body></html>"
$HtmlReport | Out-File $HtmlReportFile -Encoding UTF8

Write-Host ("Output HTML file is available in {0}" -f $HtmlReportFile) -ForegroundColor Green

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
