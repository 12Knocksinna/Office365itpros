# Send-TeamsChatMessages.PS1
# Demonstration of how to use the Microsoft Graph PowerShell SDK to send a message to a Teams chat
# https://github.com/12Knocksinna/Office365itpros/blob/master/Send-TeamsChatMessage.PS1
# V1.0 17-Apr-2024

Connect-MgGraph -Scopes Chat.ReadWrite, User.Read.All -NoWelcome

# Get details of the signed in user
$SendingUser = (Get-MgContext).Account
$SendingUserId = (Get-MgUser -UserId $SendingUser).Id
Write-Host ("Chats will be sent by {0}" -f $SendingUser)
# Define the group that's used to find the set of users to send chats to (make sure you change this for your tenant)
$GroupName = "Information Quality and Accuracy"

$WebImage = "https://i0.wp.com/office365itpros.com/wp-content/uploads/2024/04/Emergency.jpg"
# Download the icon we want to use if it's not already available - use your own image if you want
$ContentFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Emergency.jpg"
If (!(Get-Item -Path $ContentFile -ErrorAction SilentlyContinue)) {
  Invoke-WebRequest $WebImage -OutFile $ContentFile
}

# Define the content of the chat message, starting with the imline image
$Content = '<img height="200" src="../hostedContents/1/$value" width="200" style="vertical-align:bottom; width:200px; height:200px">'
$Content = $Content + '<p><b>Emergency!</b></p>'
$Content = $Content + '<p><b>Action Required Now</b><p> Please contact situation coordinator <at id="0">James Ryan</at> for more information. Contact details are available in his profile card.</p>'

# Create a hash table to hold the image content that's used with the HostedContents parameter
$ContentDataDetails = @{}
$ContentDataDetails.Add("@microsoft.graph.temporaryId", "1")
$ContentDataDetails.Add("contentBytes", [System.IO.File]::ReadAllBytes("$ContentFile"))
$ContentDataDetails.Add("contentType", "image/jpeg")
[array]$ContentData = $ContentDataDetails

# Define the body of the chat message
$Body = @{}
$Body.add("content", $Content)
$Body.add("contentType", 'html')

# Fetch people to mention in the chat message
# Define the details of the mentioned user in a hash table
$MentionedUserDetails = @{}
$MentionedUserDetails.add("userIdentityType", "aadUser")
$MentionedUserDetails.add("Id", "cad05ccf-a359-4ac7-89e0-1e33bf37579e")

# Define a hashtable to point to the hash table holding the user details
$MentionedUser = @{}
$MentionedUser.add("user", $MentionedUserDetails)

# Construct the hash table to pass the mention in the Mentions parameter for the New-MgChatMessage cmdlet
$MentionIds = @{}
$MentionIds.add("Id","0")
$MentionIds.add("Mentiontext","James Ryan")
$MentionIds.add("Mentioned", $MentionedUser)

# Find the set of users to send chats to - these are members of a team
$Team = Get-MgTeam -Filter "displayName eq '$GroupName'"
If ($Team) {
   [array]$TeamMembers = (Get-MgGroupMember -GroupId $Team.Id).Id 
   Write-Host ("Found {0} members in the {1} team" -f $TeamMembers.Count, $GroupName)
} Else {
   Write-Host ("Can't find the {0} team" -f $GroupName)
   Break
}	

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