# Send-TeamsChatWithLotsofMentions.PS1
# Example of using the Microsoft Graph PowerShell SDK to send a Teams chat with lots of mentions and 
# other bits like emojis and inline images.

Connect-MgGraph -Scopes Chat.ReadWrite, User.Read.All -NoWelcome

# Get details of the signed in user
$SendingUser = (Get-MgContext).Account
$SendingUserId = (Get-MgUser -UserId $SendingUser).Id
Write-Host ("Chats will be sent by {0}" -f $SendingUser)


[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement -notin $Modules") {
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}   

$GroupName = "Emergency Contacts"
# Download the icon we want to use if it's not already available - use your own image if you want
$WebImage = "https://i0.wp.com/office365itpros.com/wp-content/uploads/2024/04/Warning.jpg"
# Download target is in user download folder
$ContentFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Warning.jpg"
If (!(Get-Item -Path $ContentFile -ErrorAction SilentlyContinue)) {
  Invoke-WebRequest $WebImage -OutFile $ContentFile
}

# Target group to receive chat messages
$TargetGroupforChats = "Information Quality and Accuracy"
$Team = Get-MgTeam -Filter "displayName eq '$TargetGroupforChats'"
If ($Team) {
   [array]$TeamMembers = (Get-MgGroupMember -GroupId $Team.Id).Id 
   Write-Host ("Found {0} members in the {1} team" -f $TeamMembers.Count, $GroupName)
} Else {
   Write-Host ("Can't find the target team ({0}) to send chats" -f $GroupName)
   Break
}

# Find the group to use for contact information
$Group = Get-MgGroup -Filter "displayName eq '$GroupName'"
If (!($Group)) {    
    Write-Host ("Can't find the {0} group with emergency contact information" -f $GroupName)
    Break
}
# Group owner is the primary contact
$PrimaryContact = Get-MgGroupOwner -GroupId $Group.Id
$PrimaryContactDetails = $PrimaryContact | Select-Object -ExpandProperty AdditionalProperties

# Get the identifiers for the other contacts
[array]$Contacts = Get-MgGroupMember -GroupId $Group.Id


# Define the content of the chat message, starting with the imline image
$Content = '<img height="200" src="../hostedContents/1/$value" width="200" style="vertical-align:bottom; width:200px; height:200px">'
$Content = $Content + '<p><b>Warning!</b></p>'
$Content = $Content + '<p><b>Immediate Action Required</b><p>Please contact the situation coordinator '

# Create the mention for the primary contact
$MentionedUserDetails = @{}
$MentionedUserDetails.add("userIdentityType", "aadUser")
$MentionedUserDetails.add("Id", $PrimaryContact.id)
# Define a hashtable to point to the hash table holding the user details
$MentionedUser = @{}
$MentionedUser.add("user", $MentionedUserDetails)
$MentionPrimaryContact = @{}
$MentionPrimaryContact.add("Id","0")
$MentionPrimaryContact.add("Mentiontext",$PrimaryContactDetails.displayName)
$MentionPrimaryContact.add("Mentioned", $MentionedUser)
[array]$MentionIds = $MentionPrimaryContact

# add the contact mention to the message text
$PrimaryContactMention = ('<at id="0">{0}</at> for more information. Contact details are available in the profile card.<p>' -f $PrimaryContactDetails.displayName)
$Content = $Content + $PrimaryContactMention
$Content = $Content + "<p><p><u>Location-specific contacts are:</u></p>"


# Loop through the users in the $Contacts array to add them to the message
[int]$i = 0
ForEach ($Contact in $Contacts) {
    $i++
    $ContactName = $Contact.additionalProperties.displayName
    $ContactLocation = $Contact.additionalProperties.officeLocation
    $MentionedUserDetails = @{}
    $MentionedUserDetails.add("userIdentityType", "aadUser")
    $MentionedUserDetails.add("Id", $Contact.Id) 
    $MentionedUser = @{}
    $MentionedUser.add("user", $MentionedUserDetails)
    $MentionContact = @{}
    $MentionContact.add("Id",$i)
    $MentionContact.add("Mentiontext",$ContactName)
    $MentionContact.add("Mentioned", $MentionedUser)
    $MentionIds += $MentionContact
    $ContactLine = ('<p><at id="{0}">{1}</at> ({2})</p>' -f $i, $ContactName, $ContactLocation)
    $Content = $Content + $ContactLine
}

$Content = $Content + "<p>-------------------------------------------------------------------<p></p>"
$Content = $Content + '<p><emoji alt="😎"></emoji>'

# Create a hash table to hold the image content that's used with the HostedContents parameter. Hosted content
# is stored by Teams in its store and includes images and code snippets. File attachments are in OneDrive
# or SharePoint Online
$ContentDataDetails = @{}
$ContentDataDetails.Add("@microsoft.graph.temporaryId", "1")
$ContentDataDetails.Add("contentBytes", [System.IO.File]::ReadAllBytes("$ContentFile"))
$ContentDataDetails.Add("contentType", "image/jpeg")
[array]$ContentData = $ContentDataDetails

$Body = @{}
$Body.add("content", $Content)
$Body.add("contentType", 'html')

# Loop through the set of team members and send a chat to each
[int]$ChatMessagesSent = 0
ForEach ($TeamUser in $TeamMembers) {
    # No need to chat with the sender, so ignore them if they're in the team membership
    If ($TeamUser -eq $SendingUserId) {
        Write-Host "Skipping sending chat to self"
        Continue
    }
    $User = Get-MgUser -UserId $TeamUser -Property id, displayName, userprincipalName, userType
    # Can't handle MTO accounts - communicate should be with their real account
    If ($User.UserPrincipalName -like "*#EXT*" -and $User.userType -eq "Member") {
        Write-Host ("Skipping MTO account {0}" -f $User.DisplayName)
        Continue
    }
    [array]$MemberstoAdd = $SendingUserId, $TeamUser
    [array]$Members = $null
    ForEach ($Member in $MemberstoAdd){
        $MemberId = ("https://graph.microsoft.com/v1.0/users('{0}')" -f $Member)
        $MemberDetails = @{}
        [array]$MemberRole = "owner"
        If ($User.userType -eq "Guest") {
            [array]$MemberRole = "guest"
        }
        $MemberDetails.Add("roles", $MemberRole.trim())
        $MemberDetails.Add("@odata.type", "#microsoft.graph.aadUserConversationMember")
        $MemberDetails.Add("user@odata.bind", $MemberId.trim())
        $Members += $MemberDetails
    }
    # Add the members to the chat body
    $OneOnOneChatBody = @{}
    $OneOnOneChatBody.Add("chattype", "oneOnOne")
    $OneOnOneChatBody.Add("members", $Members)
    # Set up the chat - if one already exists between these two participants, Teams returns the id for that chat
    $NewChat = New-MgChat -BodyParameter $OneOnOneChatBody
    If ($NewChat) {
        Write-Host ("Chat {0} available" -f $NewChat.id)
    } Else {
        Write-Host "Failed to create chat"
    }
    # Send the message to the chat
    Write-Host ("Sending chat to {0}" -f $User.DisplayName)
    $ChatMessage = New-MgChatMessage -ChatId $NewChat.Id -Body $Body -Mentions $MentionIds `
        -HostedContents $ContentData -Importance Urgent
    If ($ChatMessage) {
        Write-Host ("Chat sent to {0}" -f $User.DisplayName) -ForegroundColor Yellow
        $ChatMessagesSent++
    } Else {
        Write-Host ("Failed to send chat message to {0}" -f $User.DisplayName) -ForegroundColor Red
    }
}
Write-Host ("All done. {0} chat messages sent" -f $ChatMessagesSent)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.