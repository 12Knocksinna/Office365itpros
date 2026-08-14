
# Download-TeamsMeetingTranscripts.PS1
# An example of how to download Teams transscript files for a meeting using the Microsoft Graph API. The script uses the Microsoft Graph PowerShell SDK to fetch the transcripts for a specific Teams meeting based on the user and meeting subject provided by the user.
# V1.0 14-Aug-2026

# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Download-TeamsMeetingTranscripts.PS1

# See message center notification MC1393806 for details of what needs to be done to allow Graph API access to Teams meeting transcripts. T
# The notification is titled "Microsoft Graph API support for Teams meeting transcripts" and was published on 16-June-2026.
# In addition, for an app to be able to access transcripts, a Teams application access policy must be configured appropriately.

# These permissions are sufficient if you fetch transcripts using delegated permissions (your own meetings)
Connect-MgGraph -Scopes "OnlineMeetings.Read.All", "OnlineMeetings.Transcripts.Read.All", "User.ReadBasic.All", "Calendars.Read" -NoWelcome

# To fetch transcripts for meetings of other users, you need to use application permissions. The following permissions are required for application permissions:
# - OnlineMeetings.Read.All: Read all online meetings in the organization 
# - OnlineMeetings.Transcripts.Read.All: Read all online meeting transcripts in the organization
# - User.ReadBasic.All: Read all users' basic information in the organization
# - Calendars.Read: Read all users' calendars in the organization

# What calendar do we need to search to find the meeting?
$SearchUser = Read-Host "What is the meeting organizer's email address"

Try {
    $User = Get-MgUser -UserId $SearchUser -ErrorAction Stop
    $UserId = $User.Id
} Catch {
    Write-Error ("Failed to retrieve user identifier for {0}: {1}" -f $SearchUser, $_.Exception.Message)
    Return
}

# Now that we have the user ID, we can search for the meeting by subject. Note that this search is case-sensitive and must match the subject exactly. If there are multiple meetings with the same subject, the script will use the first one found.    
$SearchSubject = Read-Host "What is the meeting subject"
Try {
    [array]$Events = Get-MgUserEvent -Filter "subject eq '$SearchSubject'" -UserId $UserId -ErrorAction Stop
} Catch {
    Write-Error ("Failed to retrieve events for user {0} with subject '{1}': {2}" -f $SearchUser, $SearchSubject, $_.Exception.Message)
    Return
}
If ($Events.count -eq 0) {
    Write-Warning ("No events found for user {0} with subject '{1}'" -f $SearchUser, $SearchSubject)
    Return
}

# This code only deals with a single meeting...
If ($Events.count -gt 1) {
    Write-Warning ("Multiple events found with subject '{0}', using the first one" -f $SearchSubject)
    $Event = $Events[0]
} Else {        
    $Event = $Events[0]
}

# And the meeting must be a Teams meeting, otherwise there will be no transcripts to download.
If ($Event.onlineMeetingProvider -ne 'teamsForBusiness') {
    Write-Warning ("The event with subject '{0}' is not a Teams meeting" -f $SearchSubject)
    Return
}

# Fetch the JoinUrl for the event and endcode it for use in the Graph API call to fetch the meeting details. The JoinUrl is used to uniquely identify the meeting in the Graph API.
$JoinURL = $Event.onlineMeeting.joinUrl
$EncodedJoinUrl = [System.Web.HttpUtility]::UrlEncode($JoinUrl)

# Can we find the online meeting? It's required, because transcripts are only available for online meetings. The meeting ID is required to fetch the transcripts.
Try {
    $Meeting = Get-MgUserOnlineMeeting -Filter "JoinWebURL eq '$EncodedJoinUrl'" -UserId $UserId -ErrorAction Stop
    $MeetingId = $Meeting.Id
    $MeetingSubject = $Meeting.Subject
} Catch {
    Write-Error ("Failed to retrieve meeting details for {0}: {1}" -f $JoinURL, $_.Exception.Message)
}
[array]$Transcripts = Get-MgUserOnlineMeetingTranscript -OnlineMeetingId $MeetingId -UserId $UserId

# This header is used to request the transcript content without speaker attribution
# $GraphHeaders = @{ "Accept" = "application/vnd.microsoft.graph.transcript+text" }
# Theis header fetches transcripts with speaker attribution included
$GraphHeaders = @{ "Accept" = "text/vtt" }

# Loop through the transcripts and download each one to the user's Downloads folder. 
# The transcript files are named using the meeting subject, the transcript creation date, and a counter to ensure unique filenames.
[int]$TranscriptCounter = 0
[array]$FetchedFiles = $null
ForEach ($Transcript in $Transcripts) {
    $TranscriptCounter++
    Write-Output ("Transcript ID: {0}, Created: {1}, End: {2}" -f $Transcript.Id, $Transcript.CreatedDateTime, $Transcript.EndDateTime)
    $TranscriptId = $Transcript.Id
    $TranscriptFileName = ("Transcript_{0}_{1}_{2}.txt" -f $MeetingSubject, ($Transcript.CreatedDateTime.ToUniversalTime().ToString("yyyy-MM-dd-HH-mm-ss")), $TranscriptCounter)
    $Uri = ("https://graph.microsoft.com/v1.0/users/{0}/onlineMeetings/{1}/transcripts/{2}/content" -f $UserId, $MeetingId, $TranscriptId)
    # $Uri can also be fetched from $Transcript.TranscriptContentUrl, but it is not guaranteed to be present in all cases
    $OutputTranscriptFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) `
         + "\" + $TranscriptFileName
    Try {
        Invoke-MgGraphRequest -Method GET -Uri $Uri -Headers $GraphHeaders -OutputFilePath $OutputTranscriptFile -ErrorAction Stop
        $FetchedFiles += $OutputTranscriptFile
        Write-Output ("Transcript saved to {0}" -f $OutputTranscriptFile)
    } Catch {
        Write-Error ("Failed to download transcript {0}: {1}" -f $TranscriptId, $_.Exception.Message)
    }
}

Write-Host ""
Write-Host ("A total of {0} transcript files were downloaded." -f $FetchedFiles.Count)
Write-Host "The files are downloaded to your Downloads folder. Please check the folder for the transcripts of the meeting."
Write-Host ""
Write-Host $FetchedFiles | Format-Table -AutoSize

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.