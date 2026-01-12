<# 
.SYNOPSIS
Creates a channel-bound Teams meeting, a calendar event in the Team's Group calendar,
and optionally posts a channel announcement with the join link.

.PARAMETER TeamName
Display name of the Team (M365 Group) backing your channel.

.PARAMETER ChannelName
Display name of the channel (e.g., "General").

.PARAMETER Subject
Meeting/event subject.

.PARAMETER Start
Start date/time in local time, e.g. "2026-01-15 09:30".

.PARAMETER End
End date/time in local time, e.g. "2026-01-15 10:15".

.PARAMETER TimeZone
IANA or Windows time zone to apply to the calendar event (e.g., "Europe/Dublin" or "GMT Standard Time").

.PARAMETER Agenda
file name to read agenda into the meeting body (optional).

.PARAMETER Attendees
Comma-separated email addresses to invite (optional). Channel members can still join from the channel.

.PARAMETER PostAnnouncement
Switch. If present, posts a message with the join link into the channel.

.EXAMPLE
.\New-TeamsChannelMeeting.ps1 -TeamName "Microsoft Graph Gurus" -ChannelName "General" `
  -Subject "WAM and Graph SDK V2.34" -Start "2026-02-15 0930" -End "2026-02-15 1015" `
  -TimeZone "UTC" -Attendees "Sean.Landy@office365itpros.com" -PostAnnouncement -Agenda "c:\temp\Agenda.txt"

.NOTES
Requires Microsoft.Graph SDK and delegated scopes:
OnlineMeetings.ReadWrite, Group.ReadWrite.All, ChannelMessage.Send, Channel.ReadBasic.All.

Version 1.0 12-Jan-2026
GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/New-TeamsChannelMeeting.PS1

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $TeamName,

    [Parameter(Mandatory)]
    [string] $ChannelName,

    [Parameter(Mandatory)]
    [string] $Subject,

    [Parameter(Mandatory)]
    [string] $Start,

    [Parameter(Mandatory)]
    [string] $End,

    [string] $TimeZone = "Europe/Dublin",

    [string] $Agenda = "",

    [string] $Attendees,

    [switch] $PostAnnouncement
)
function Connect-GraphSession {
    $requiredScopes = @(
        "OnlineMeetings.ReadWrite",
        "Group.ReadWrite.All",
        "ChannelMessage.Send",
        "Channel.ReadBasic.All"
    )

    if (-not (Get-MgContext)) {
        Write-Host "Connecting to Microsoft Graph and requesting delegated scopes..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
        
        return
    }

    # Ensure we have all scopes
    [array]$Scopes = Get-MgContext | Select-Object -ExpandProperty Scopes
    $MissingScopes = $RequiredScopes | Where-Object { $_ -notin $Scopes }
    if ($MissingScopes) {
        Write-Host "Disconnecting from Graph to allow the addition of missing scopes: $($MissingScopes -join ', ')" -ForegroundColor Yellow
        Disconnect-MgGraph
        Break
        
    }
}

function New-ChannelMeeting {
    # Function to create a meeting in a group calendar
    param(
        [string] $Subject,
        [string] $StartLocal,
        [string] $EndLocal,
        [string] $TimeZone,
        [string] $AgendaFile,
        [string[]] $AttendeeEmails
    )
    Write-Host "Creating online meeting..." -ForegroundColor Cyan

    If (Get-Content $AgendaFile -ErrorAction SilentlyContinue) {
        $Agenda = Get-Content $AgendaFile -Raw
    } Else {
        $Agenda = "No agenda provided."
    }   

    # Build attendees array
    $attendeesArray = @()
    foreach ($EmailAddress in $AttendeeEmails) {
        $trim = $EmailAddress.Trim()
        if ($trim) {
            $attendeesArray += @{
                emailAddress = @{ address = $trim; name = $trim }
                type = "required"
            }
        }
    }

    $EventParams = @{
        Subject = $Subject
        Body    = @{
            ContentType = "HTML"
            Content     = $Agenda
        }
        Start   = @{ DateTime = ([DateTime]::Parse($StartLocal).ToString("yyyy-MM-ddTHH:mm:ss")); TimeZone = $TimeZone }
        End     = @{ DateTime = ([DateTime]::Parse($EndLocal).ToString("yyyy-MM-ddTHH:mm:ss")); TimeZone = $TimeZone }
        Location = @{ DisplayName = "Teams Channel meeting" }
        Attendees = $attendeesArray
        IsOnlineMeeting       = $true
        OnlineMeetingProvider = "teamsForBusiness"
        AllowNewTimeProposals = $true
        TransactionId         = [Guid]::NewGuid().ToString()  # must be a *string*
    }

# Create the meeting in the group calendar
Try {
    $Global:NewMeeting = New-MgGroupEvent -GroupId $TeamId -BodyParameter $EventParams
    $Global:StartTime = Get-Date $NewMeeting.Start.DateTime -format "dd-MMM-yyyy HH:mm"
    $Global:EndTime = Get-Date $NewMeeting.End.DateTime -format "dd-MMM-yyyy HH:mm"
    $Global:HtmlBodyContent = $NewMeeting.Body.content
} Catch {
    Write-Output "Failed to create event in group calendar: $TeamName"
}
    
return $NewMeeting

}

function New-ChannelAnnouncement {
    param(
        [string] $TeamId,
        [string] $ChannelId,
        [string] $Subject
    )

    # Extract Teams meeting URL using regex
    $Pattern = 'https://teams\.microsoft\.com/meet/[^\s"<]+'
    If ($HtmlBodyContent -match $Pattern) {
        $TeamsJoinUrl = $Matches[0]
    } Else {
        $TeamsJoinURL = "Unknown Teams meeting link – contact organizer for details"
    }

    # Create HTML body for channel message
$htmltext = @"
<div><b><u>All</u></b> channel members are expected to attend this meeting.<br/>
<a href="$TeamsJoinURL">Join the online Teams meeting</a></div>
<p>Starting at <b>$StartTime</b> until <b>$EndTime</b></p>
"@

# Create rest of body with an @ mention to the channel
$Body = @{
	body = @{
        contentType = "html"
	    content = ('<div><div>Hello <at id="0">' + $ChannelName + '</at>. This meeting about <b> ' + $Subject + '</b> is important for operational reasons. </div></div>') + $htmltext
	}
	subject = "Critical Meeting for Channel Members to Attend"
    importance = "High"
	mentions = @(
	    @{
	        id = 0
	        mentionText = $ChannelName
	        mentioned = @{
	        conversation = @{
	            id = $ChannelId
	            displayName = $ChannelName
	            conversationIdentityType = "channel"
	          }
	        }
        }
    )
}

    # Post channel message
    Try {
        $Message = New-MgTeamChannelMessage -TeamId $TeamId -ChannelId $ChannelId -Bodyparameter $Body -ErrorAction Stop
    } Catch {
        throw "Failed to post channel announcement: $_"
    }
    If ($Message) {
        Write-Host "Posted channel announcement with meeting link." -ForegroundColor Green
    }
}

# --- MAIN ---

# Check our permissions
Connect-GraphSession
# Creating a meeting in a group calendar only works with delegated permissions

$Global:UserId = (Get-MgContext).Account
# Make sure that we can find the target team and channel
Try {
    [array]$Teams = Get-MgUserJoinedTeam -UserId $UserId  -ErrorAction Stop
    $Team = $Teams | Where-Object { $_.DisplayName -eq $TeamName }
} catch {
    Write-Output "Failed to retrieve list of joined Teams to validate membership of signed-in user."
    Break
}

If ($Team) {
    $Global:TeamId = $Team.Id
} Else {
    Write-Output "You are not a member of a Team named '$TeamName'. Please check the name and try again."
    Break
}
Try {
    $Channel = Get-MgTeamChannel -TeamId $TeamId -Filter "displayName eq '$ChannelName'" -ErrorAction Stop
} catch {
    Write-Output "Failed to retrieve channels for Team '$TeamName'. Ensure you are a member of the Team."
    Break
}

If ($Channel) {
    $Global:ChannelId = $Channel.Id
} Else {
    Write-Output "The Team '$TeamName' does not contain a channel named '$ChannelName'. Please check the name and try again."
    Break
}

Write-Host "Attempting to create meeting in Team '$TeamName' Channel '$ChannelName'..." -ForegroundColor Cyan

# Create the meeting in the group calendar
[array]$Attendees = $Attendees -split ","

$NewMeeting = New-ChannelMeeting -Subject $Subject -StartLocal $Start -EndLocal $End -TimeZone $TimeZone `
    -AgendaFile $Agenda -AttendeeEmails $Attendees

# Post announcement if requested
If ($PostAnnouncement) {
    $Global:Subject = $Subject
    New-ChannelAnnouncement -NewMeeting $NewMeeting -TeamId $TeamId -ChannelId $ChannelId -Subject $Subject
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.