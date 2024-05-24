# Analyze-TeamsMeetings.Ps1

Connect-ExchangeOnline -SkipLoadingCmdletHelp

Connect-MgGraph -NoWelcome -Scopes OnlineMeetings.Read
$Organization = Get-MgOrganization
$TenantId = $Organization.id
$TenantName = $Organization.DisplayName
[array]$OverMeetingAttendees = $null

# Find meeting details
$StartDate = (Get-Date).AddDays(-90)
$EndDate =(Get-Date).AddDAys(1)
Write-Host "Searching for audit events for Teams Meetings..."
[array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations MeetingDetail -Formatted -SessionCommand ReturnLargeSet -ResultSize 5000
If (!($Records)) {
    Write-Host "No meeting details found"
    Break
}
# remove duplicates
$Records = $Records | Sort-Object Identity -Unique | Sort-Object { $_.CreationDate -as [datetime]} -Descending
Write-Host ("{0} meeting details found" -f $Records.Count)  
$MeetingData = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
    $AuditData = ConvertFrom-Json $Rec.AuditData
    [datetime]$Start = $AuditData.StartTime
    [datetime]$End = $AuditData.EndTime
    $Duration = $End - $Start
    $MeetingDuration = ("{0:hh\:mm\:ss}" -f $Duration)

    $ReportLine = [PSCustomObject]@{
        CreationDate    = (Get-Date $Rec.CreationDate -format "dd-MMM-yyyy HH:mm:ss")
        MeetingId       = $AuditData.Id
        UserId          = $AuditData.UserId
        MeetingType     = $AuditData.CommunicationSubType
        StartTime       = (Get-Date $AuditData.StartTime -format "dd-MMM-yyyy HH:mm:ss")
        EndTime         = (Get-Date $AuditData.EndTime -format "dd-MMM-yyyy HH:mm:ss")
        Duration        = $MeetingDuration
        Organizer       = $AuditData.Organizer.UserObjectId
        Modalities      = $AuditData.Modalities
        MeetingURL      = $AuditData.MeetingURL

    }
    $MeetingData.Add($ReportLine)
}







$MeetingURL = "https://" +  "teams.microsoft.com/l/meetup-join/19%3ameeting_ZDdlYmFiZGMtMDY5NC00NTQ5LTgyZTAtYjQwMGFiNDUwMTM4%40thread.v2/0?context=%7b%22Tid%22%3a%22b662313f-14fc-43a2-9a7a-d2
e27f4f3478%22%2c%22Oid%22%3a%22eff4cd58-1bb8-4899-94de-795f656b4a18%22%7d"
$DecodedURL = [System.Web.HttpUtility]::UrlDecode($MeetingURL)
$MeetingIdStart = $DecodedURL.IndexOf("19:") 
$MeetingIdEnd = $DecodedURL.IndexOf("thread")
$MeetingId = $DecodedURL.Substring($MeetingIdStart, $MeetingIdEnd - $MeetingIdStart +9)
$OrganizerIdStart = $DecodedURL.IndexOf("Oid")
$OrganizerId = $DecodedURL.Substring($OrganizerIdStart + 6, 36)
$MeetingIdLookup = ("1*{0}*0**{1}" -f $OrganizerId, $MeetingId)
$Base64MeetingId = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($MeetingIdLookup))
$Uri = ("https://graph.microsoft.com/v1.0/users/{0}/onlineMeetings/{1}" -f $OrganizerId, $Base64MeetingId)
[array]$Data = Invoke-MgGraphRequest -Method get -Uri $Uri

$MeetingSubject = $Data.subject

# Get the set of attendance reports for the meeting
$Uri = $Uri = ("https://graph.microsoft.com/v1.0/users/{0}/onlineMeetings/{1}/attendanceReports" -f $OrganizerId, $Base64MeetingId)
[array]$AttendanceReports = Invoke-MgGraphRequest -Method get -Uri $Uri
$AttendanceReports = $AttendanceReports.value
# Extract the individual attendee records from each report
$AttendanceData = [System.Collections.Generic.List[Object]]::new()
If ($AttendanceReports) {
    ForEach ($AR in $AttendanceReports) {
        $Uri = $Uri = ("https://graph.microsoft.com/v1.0/users/{0}/onlineMeetings/{1}/attendanceReports/{2}?`$expand=attendanceRecords" -f `
            $OrganizerId, $Base64MeetingId, $AR.Id)
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
                ReportId                = $AR.Id
                'Number Participants'   = $AR.totalParticipantCount
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
                Email                   = $Participant.emailAddress
                DisplayName             = $Participant.identity.displayName
            }
            $AttendanceData.Add($ReportLine)
        }
    }
}

$OverMeetingAttendees += $AttendanceData 



# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
