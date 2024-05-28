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
# Calendar.Read = Read calendar data
# Group.Read.All = Read group membership
# OnlineMeeting.Read.All = Read online meeting data
# OnlineMeetingArtifact.Read.All = Read attendance reports
# User.Read.All = Read user data
# CrossTenantInformation.ReadBasic.All = Read tenant information for external participants

Connect-MgGraph -AppId $AppId -TenantId $TenantId -CertificateThumbprint $CertThumbprint -NoWelcome
$Organization = Get-MgOrganization
$TenantName = $Organization.DisplayName
[array]$OverMeetingAttendees = $null

$StartDate = (Get-Date).AddDays(-62)
$EndDate = (Get-Date)
$StartDateSearch = Get-Date $StartDate -format s
$EndDateSearch = Get-Date $EndDate -format s
# Get members of the group that we want to process - change the name of the group to match your tenant
[array]$TeamsOrganizers = Get-MgGroupMember -GroupId (Get-MgGroup -Filter "displayName eq 'Teams Meeting Organizers'").Id

$MeetingReport = [System.Collections.Generic.List[Object]]::new()

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
    Write-Host ("Proceeding to analyze Teams meetings for {0}" -f $DisplayName)

    ForEach ($Item in $CalendarItems) {
        $MeetingDuration = $null; [array]$MeetingData = $null
        # Get the meeting URL
        $MeetingURL =  $Item.onlinemeeting.joinUrl.trim()
        $DecodedURL = [System.Web.HttpUtility]::UrlDecode($MeetingURL)
        $MeetingIdStart = $DecodedURL.IndexOf("19:") 
        $MeetingIdEnd = $DecodedURL.IndexOf("thread")
        $MeetingId = $DecodedURL.Substring($MeetingIdStart, $MeetingIdEnd - $MeetingIdStart +9)
        #$OrganizerIdStart = $DecodedURL.IndexOf("Oid")
        #$OrganizerId = $DecodedURL.Substring($OrganizerIdStart + 6, 36)
        $MeetingIdLookup = ("1*{0}*0**{1}" -f $Organizer.id, $MeetingId)
        $Base64MeetingId = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($MeetingIdLookup))
        $Uri = ("https://graph.microsoft.com/v1.0/users/{0}/onlineMeetings/{1}" -f $Organizer.Id, $Base64MeetingId)
        Try {
            [array]$MeetingData = Invoke-MgGraphRequest -Method Get -Uri $Uri
        } Catch {
            Write-Host ("Ran into an issue retieving information for meeting {0} (created {1}): {1}" -f $Item.subject, $Item.CreatedDateTime, $($PSItem.ToString()))
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
            'Creation date'         = (Get-Date $MeetingData.creationDateTime -format 'dd-MMM-yyyy HH:mm:ss')
            'Meeting Start'         = (Get-Date $MeetingData.startDatetime -format 'dd-MMM-yyyy HH:mm:ss')
            'Meeting End'           = (Get-Date $MeetingData.endDateTime -format 'dd-MMM-yyyy HH:mm:ss')
            'Time zone'             = $Item.originalStartTimeZone
            'Meeting Duration'      = ("{0:hh\:mm\:ss}" -f $MeetingDuration)
            Subject                 = $MeetingData.subject
            'Allow Presenters'      = $MeetingData.allowedpresenters
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
        $AttendanceData = [System.Collections.Generic.List[Object]]::new()
        ForEach ($AR in $AttendanceReports) {
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
                }
                $AttendanceData.Add($ReportLine)
            }
        }
    }
    # Add the data to the overall list
    $OverMeetingAttendees += $AttendanceData 

}






# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
