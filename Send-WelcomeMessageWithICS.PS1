# Send-WelcomeMessageWithICS.PS1
# a script to send a welcome message to new user accounts, cokmplete with an attachment to welcome the new account
# and one ore more ICS files for upcoming corporate events.
# Github link: https://github.com/12Knocksinna/Office365itpros/blob/master/Send-WelcomeMessageWithICS.PS1

# See article at https://office365itpros.com/2026/05/25/online-meeting-ics-files/

# V1.0 12-May-2026
# Requires the Microsoft Graph PowerShell SDK running in app-only mode with the following permissions:
# Mail.Send, Organization.Read.All, Calendars.ReadWrite, User.Read.All, and OnlineMeetings.Read.All

# In addition, a Teams application access policy must be created and assigned to the app to allow it to access online meeting information.
# This script is intended to be run as an Azure Automation runbook or in app-only mode in an interactive PowerShell session. 

Function Update-MessageRecipients {
    [cmdletbinding()]
    Param(
        [array]$ListOfAddresses )
    ForEach ($SMTPAddress in $ListOfAddresses) {
        @{
            emailAddress = @{address = $SMTPAddress}
        }    
    }    
}

Function Update-MessageAttachments {
    [cmdletbinding()]
    Param(
        [array]$ListOfAttachments 
    )

    [array]$MsgAttachments = $null
    ForEach ($File in $ListOfAttachments) {
        $ConvertedContent = [Convert]::ToBase64String([IO.File]::ReadAllBytes($File))
        $FileExtension = [System.IO.Path]::GetExtension($File) 
        Switch ($FileExtension) {
            ".pdf" {
                $ContentType = "application/pdf"
            }
            ".docx" {
                $ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            }
            ".xlsx" {
                $ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }   
            ".pptx" {
                $ContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            }   
            ".jpg" {
                $ContentType = "image/jpeg"
            }   
            ".png" {
                $ContentType = "image/png"
            }   
            default {
                $ContentType = "application/octet-stream"
            }
        }
        $AttachmentDetails = @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            Name = $File
            ContentType =  $ContentType
            ContentBytes = $ConvertedContent
        }
        $MsgAttachments += $AttachmentDetails
    }
    Return $MsgAttachments  
}

function New-IcsEvent {
    param (
        [Parameter(Mandatory)] $EventDetails
    )

    # Make sure that meeting dates are converted to UTC and formatted as needed for the ICS file
    $dts =[datetime]$EventDetails.Start
    $dte =[datetime]$EventDetails.End
    $tz  = [timeZoneInfo]::FindSystemTimeZoneById($EventDetails.Timezone)
    $dtso = [DateTimeOffset]::new($dts, $tz.GetUtcOffset($dts))
    $dtse = [DateTimeOffset]::new($dte, $tz.GetUtcOffset($dte))

    $uid            = [guid]::NewGuid().ToString()
    $dtStart        = ([DateTime]$dtso.utcDateTime).ToUniversalTime().ToString("yyyyMMddTHHmmssZ")
    $dtEnd          = ([DateTime]$dtse.UtcDateTime).ToUniversalTime().ToString("yyyyMMddTHHmmssZ")
    $dtStamp        = (Get-Date).ToUniversalTime().ToString("yyyyMMddTHHmmssZ")
    $teamsUrl       = $EventDetails.TeamsUrl
    $subject        = $EventDetails.Subject
    $Organizer      = $EventDetails.OrganizerName
    $OrganizerEmail = $EventDetails.OrganizerEmail
    $meetingId       = $EventDetails.MeetingId
    $MeetingPasscode = $EventDetails.MeetingPasscode

    # Build the Teams description blob exactly as Outlook expects it
    $description  = "Microsoft Teams meeting\n"
    $description += "Join on your computer or mobile app\n"
    $description += "Click here to join the meeting: $teamsUrl\n"
    If ($meetingId) {
        $description += "Meeting ID: $meetingId\n"
        $description += "Passcode: $MeetingPasscode\n"
    }

    # Audio conferencing details if available
    If ($EventDetails.audioConferencing) {
        $dialIn     = $EventDetails.audioConferencing.tollNumber
        $confId     = $EventDetails.audioConferencing.conferenceId
        if ($dialIn) { $description += "Dial in (audio only): $dialIn\n" }
        if ($confId) { $description += "Phone Conference ID: $confId\n" }
    }
  
# Build the ICS using the information extracted from the calendar event and online meeting  
$ICSLines = @(
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//RA//Meeting Scheduler//EN",
    "CALSCALE:GREGORIAN",
    "BEGIN:VEVENT",
    "UID:$uid",
    "DTSTAMP:$dtStamp",
    "DTSTART:$dtStart",
    "DTEND:$dtEnd",
    (Set-IcsLine "SUMMARY:$subject"),
    (Set-IcsLine "DESCRIPTION:$description"),
    "LOCATION:Microsoft Teams Meeting",
    (Set-IcsLine "X-MICROSOFT-SKYPETEAMSMEETINGURL:$teamsUrl"),
    (Set-IcsLine "X-MICROSOFT-ONLINEMEETINGCONFLINK:$teamsUrl"),
    (Set-IcsLine "ORGANIZER;CN=$($Organizer):mailto:$($OrganizerEmail)"),
    "BEGIN:VALARM",
    "TRIGGER:-PT15M",
    "ACTION:DISPLAY",
    "DESCRIPTION:Reminder",
    "END:VALARM",
    "END:VEVENT",
    "END:VCALENDAR"
)
    # Make sure that all lines in the ICS are terminated with a line feed
    $ICS = [string]::Join("`r`n", $icsLines)
    $ICS += "`r`n"

    # Create the ICS file
    $Path = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\$Uid.ics"
    [System.IO.File]::WriteAllText($Path, $ICS, [System.Text.UTF8Encoding]::new($false))
    Get-Item $Path
}

function Set-IcsLine {
    param([string]$line)
    # Line folding function to make sure that the lines in the ICS file do not exceed 75 characters, as per the iCalendar specification. Lines longer than 75 characters are folded by inserting a CRLF followed by a space.
    if ($line.Length -le 75) { return $line }
    $result    = $line.Substring(0, 75)
    $remaining = $line.Substring(75)
    while ($remaining.Length -gt 0) {
        $chunk     = if ($remaining.Length -gt 74) { $remaining.Substring(0, 74) } else { $remaining }
        $result   += "`r`n " + $chunk
        $remaining = if ($remaining.Length -gt 74) { $remaining.Substring(74) } else { "" }
    }
    return $result
}

# Start of processing

$Thumbprint = '0CF6CE3F3548FD73E7AC8CF20226ED447E125C71'
$TenantId = 'a662313f-14fc-43a2-9a7a-d2e27f4f3478'
$AppId = 'dfe6637d-8d2a-4d8f-90eb-c2cbb7479b55'

# Connect in app-only mode
Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint

# Runbook - connect using a managed identity. Make sure that the automation account has the 
# necessary permissions and that the required modules like Microsoft.Graph.Authentication, Microsoft.Graph.Calendar,  
# Microsoft.Graph.Users, and Microsoft.Graph.Users.Actions are added to the runbook environment
Connect-MgGraph -Identity

# We're going to look for user accounts created in the last week, but you can change this as needed
$WeekAgo = (Get-Date).AddDays(-7).toString("yyyy-MM-ddTHH:mm:ssZ")    
# and the user accounts must be licensed for Exchange Online -
$ExoServicePlan1 = "9aaf7827-d63c-4b61-89c3-182f06f82e5c"
$ExoServicePlan2 = "efb87545-963c-4e0d-99df-69c6916d9eb0"
$NewICS = $null

# Find the set of user accounts licensed for Exchange Online created in the last seven days
[array]$Users = Get-MgUser -Filter "(createddateTime ge $WeekAgo and userType eq 'Member') `
    and (assignedPlans/any(c:c/servicePlanId eq $ExoServicePlan1 and capabilityStatus eq 'Enabled') `
     or assignedPlans/any(c:c/servicePlanId eq $ExoServicePlan2 and capabilityStatus eq 'Enabled'))" `
    -ConsistencyLevel eventual -CountVariable Test -All -PageSize 500 -Sort ('displayname') `
    -Property Id, displayName, userprincipalName, assignedPlans, CreatedDateTime, mail

If (!$Users) { 
    Write-Host "No new users found. Exiting!" 
    break 
} Else {
    Write-Host "Found" $Test "new users created in the last week. Processing..." 
}

# Message sender can be any user in the tenant.
$MsgFrom = 'Azure.Management.Account@office365itpros.com'
# Get Tenant name
$TenantName = (Get-MgOrganization).displayName
$MsgSubject = "A warm welcome to $($TenantName)"

# Define some variables used to construct the HTML content in the message body
#HTML header with styles
$HtmlHead="<html><style>BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    </style>"

# Define attachments  we're only using one here, which we fetch from a web site
# if you want to add more files, add the file names to the $AttachmentsList array.
# Obviously, change the file to something that makes sense for your organization.
$WebAttachmentFile = "https://office365itpros.com/wp-content/uploads/2022/02/WelcomeToOffice365ITPros.docx"
# fetch the content of the web file and store it in the downloads folder
$Attachment = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) +"\WelcomeNewEmployeeToOffice365itpros.docx"
Invoke-WebRequest -Uri $WebAttachmentFile -OutFile $Attachment
[array]$Attachments = $Attachment

#Content for the message - obviously, this is very customizable and should reflect what you want to say to new users
$HtmlBody = "<body>
     <h1>Welcome to $($TenantName)</h1>
     <p><strong>Generated:</strong> $(Get-Date -Format 'dd-MMM-yyyy')</p>  
     <h2><u>We're Pleased to Have You Here</u></h2>
     <p><b>Welcome to your new Microsoft 365 account</b></p>
     <p>You can open your account to access your email and documents by clicking <a href=http://www.portal.office.com>here</a> </p>
     <p>Have a great time and be sure to call the help desk if you need assistance. And be sure to read all the great articles about Microsoft 365 published by our <a href=https://office365itpros.com>Microsoft 365 experts</a>.</p>"

# Determine if any calendar events for corporate events should be added to the message. We look for events in the next six weeks.
# If any events are found, create an .ics file and add it to the attachments list
$CalendarUser = Get-MgUser -UserId "Corporate.Events.Meetings@office365itpros.com"
Try {
    $NoCorporateEvents = $False
    $Calendar = Get-MgUserDefaultCalendar -UserId $CalendarUser.Id -ErrorAction Stop
} Catch {
    Write-Host "Can't find calendar for corporate events" $CalendarUser " - skipping addition of calendar events to message attachments"
    $NoCorporateEvents = $True
}
$StartDateTime = (Get-Date).ToString("yyyy-MM-dd")
$EndDateTime = (Get-Date).AddDays(42).ToString("yyyy-MM-dd")

Try {
    [array]$Events = Get-MgUserCalendarView -UserId $CalendarUser.Id -CalendarId $Calendar.Id -all `
        -EndDateTime $EndDateTime -StartDateTime $StartDateTime -ErrorAction Stop    
} Catch {
    Write-Host "Error fetching events from corporate events calendar - skipping addition of calendar events to message attachments"
    $NoCorporateEvents = $True
}

# If we have some events, go and process them
If ($Events -and $NoCorporateEvents -eq $False) {
    Write-Host "Found" $Events.Count "upcoming events in the corporate calendar. Adding to message attachments."
    ForEach ($ICSEvent in $Events) {
        $EventJoinURL = $ICSEvent.OnlineMeeting.JoinUrl
        If (!$EventJoinURL) {
            # Only interested in online events...
            Write-Host "Event" $ICSEvent.Subject "does not have an online meeting URL - skipping"
            continue
        }
        # Fetch detail of the online meeting
        $OnlineMeetingInfo = Get-MgUserOnlineMeeting -UserId $CalendarUser.Id -Filter "JoinWebURL eq '$EventJoinURL'" -ErrorAction Stop
        $EventDetails = @{
            Subject = $ICSEvent.Subject
            Start = $ICSEvent.Start.DateTime
            End = $ICSEvent.End.DateTime
            Location = $ICSEvent.Location.DisplayName
            Body = $ICSEvent.Body.Content
            Description = 'Corporate Event'
            TeamsUrl = $OnlineMeetingInfo.JoinWebUrl
            OrganizerEmail = $ICSEvent.Organizer.EmailAddress.Address
            OrganizerName = $ICSEvent.Organizer.EmailAddress.Name
            MeetingId  = $OnlineMeetingInfo.JoinMeetingIdSettings.JoinMeetingId
            MeetingPasscode = $OnlineMeetingInfo.JoinMeetingIdSettings.Passcode
            Timezone = $ICSEvent.Start.TimeZone
            AudioConferencing = @{
                tollNumber = $OnlineMeetingInfo.AudioConferencing.TollNumber
                conferenceId = $OnlineMeetingInfo.AudioConferencing.ConferenceId
            }
        }
        # Create .ics file for the event and add to attachments list
        $NewICS = New-IcsEvent -EventDetails $EventDetails
        $Attachments += $NewICS.VersionInfo.filename
    }   
}
# Create the array of attachments for the message
[array]$MsgAttachments = Update-MessageAttachments -ListOfAttachments $Attachments

# Loop through each user to generate and send the welcome message, customizing the content for each user.
ForEach ($user in $users) {
    $ToRecipientList   = @( $User.Mail )
    [array]$MsgToRecipients = Update-MessageRecipients -ListOfAddresses $ToRecipientList
    Write-Host "Sending welcome email to" $User.DisplayName
    # Customize the message 
    $htmlHeaderUser = "<h2>A Huge Welcome to Our New User " + $User.DisplayName + "</h2>"    
    $HtmlMsg = "</body></html>" + $HtmlHead + $htmlHeaderUser + $HtmlBody + "<p>"
    If ($NewICS) {
        $HtmlMsg += "<p>Also, we've included an .ics file for any upcoming corporate events that you might want to attend. Just open each file to add the event to your calendar.</p>"
    }
    # Construct the message body
    $MsgBody = @{
        Content = "$($HtmlMsg)"
        ContentType = 'html'   
    }

    $Message =  @{subject           = $MsgSubject}
    $Message += @{toRecipients      = $MsgToRecipients}  
    $Message += @{ccRecipients      = $MsgCcRecipients}  
    $Message += @{attachments       = $MsgAttachments}
    $Message += @{body              = $MsgBody}
    $Params   = @{'message'         = $Message}
    $Params  += @{'saveToSentItems' = $True}
    $Params  += @{'isDeliveryReceiptRequested' = $True}

    Try {
        Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params -ErrorAction Stop
    } Catch {
        Write-Host "Error sending message to" $User.DisplayName ":" $_.Exception.Message
    }
  
}

Write-Host "Processing complete. Sent welcome messages to" $Users.Count "new users created in the last week."

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 