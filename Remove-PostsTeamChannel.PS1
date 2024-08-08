# Remove-PostsTeamChannel.PS1
# A script to demonstrate how to clean out all messages from a channel
# V1.0 9-Feb-2024
# https://github.com/12Knocksinna/Office365itpros/blob/master/Remove-PostsTeamChannel.PS1
# See https://practical365.com/remove-channel-messages/ for more information

Function Get-Response ([string]$Prompt,[int]$NumberPossibleAnswers) {
    # Helper function to prompt a question and get a response
       $OKtoProceed = $False
       While ($OKToProceed -eq $False) {
        [int]$Answer = Read-Host $Prompt
        If ($Answer -gt 0 -and $Answer -le $NumberPossibleAnswers) {
            $OKtoProceed = $True
            Return ($Answer) 
        } ElseIf ($Answer -eq 0) { #break out of loop
            $OKtoProceed = $True
            Return ($Answer)
        }
       } #End while
}

Connect-MgGraph -NoWelcome -Scopes ChannelMessage.ReadWrite, Team.ReadBasic.All, ChannelMessage.Read.All, TeamMember.Read.All, ChannelSettings.Read.All

$CheckTeam = Read-Host "What team is the channel to clear out in?"
Write-Host ("Searching for the {0} team" -f $CheckTeam)
$Team = Get-MgTeam -Filter "displayName eq '$CheckTeam'"
If (!$Team) {
    Write-Host ("Can't find the {0} team. Please retry" -f $CheckTeam)
    break
}
[array]$TeamMembers = Get-MgTeamMember -TeamId $Team.Id 
[array]$TeamOwners = $TeamMembers | Where-Object {$_.Roles -ne $null} 
# Extract array to get details of the team owners
$TeamOwners = $TeamOwners.additionalProperties
$ThisUser = (Get-MgContext).Account
If ($ThisUser -notin $TeamOwners.email) {
    Write-Host ("User {0} is not a team owner - exiting" -f $ThisUser)
    break
}
[array]$Channels = Get-MgTeamChannel -TeamId $Team.Id 
Write-Host ""
Write-Host "Please select a channel from the following"
Write-Host ""
[int]$i = 0
ForEach ($Channel in $Channels) {
    $i++
    Write-Host ("{0}: {1}" -f $i, $Channel.displayName)
}
[Int]$Answer = Get-Response -Prompt  "Enter the number of the channel to select" -NumberPossibleAnswers $i
[int]$SelectedChannelNumber = ($Answer - 1)
[array]$SelectedChannel = $Channels[$SelectedChannelNumber]
Write-Host ""
Write-Host ("You selected channel number {0} - {1}" -f $Answer, $SelectedChannel.displayName)

# Check channel settings to see if owner can remove messages.
[array]$Settings = (Get-MgTeam -TeamId $Team.Id).MessagingSettings
If ($Settings.AllowOwnerDeleteMessages -eq $false) {
    Write-Host ("Owners aren't allowed to remove messages from the {0} team" -f $Team.displayName)
    Break
}
# Fetch messages
Write-Host ("Checking messages in the {0} channel..." -f $SelectedChannel.displayName )
[array]$ChannelMessages = Get-MgTeamChannelMessage -TeamId $Team.Id -ChannelId $SelectedChannel.Id -PageSize 50 -All -ErrorAction SilentlyContinue
$ChannelMessages = $ChannelMessages | Where-Object {$null -eq $_.DeletedDateTime}

If ($ChannelMessages) {
    Write-Host ("There {0} messages in the {1} channel." -f $ChannelMessages.count, $SelectedChannel.displayName)
}
$ConfirmationHeader = ("Confirm removal of messages from {0} channel" -f $SelectedChannel.displayName)
$Question = 'OK to go ahead and remove these messages?'
$Choices = '&Yes', '&No'
$Decision = $Host.UI.PromptForChoice($ConfirmationHeader, $Question, $Choices, 1)

if ($Decision -eq 0) {
    [int]$RepliesRemoved = 0; [int]$MessagesRemoved = 0
    Write-Host 'Removing messages... please wait'
    [int]$i = 0
    ForEach ($Message in $ChannelMessages) {
        $i++
        Write-Host ("Processing message {0} of {1}" -f $i, $ChannelMessages.count)
        # Check if any replies exist - if so, remove them first
        [array]$Replies = Get-MgTeamChannelMessageReply -TeamId $team.id -ChannelId $SelectedChannel.Id `
            -ChatMessageId $Message.Id
        If ($Replies) {
            $RepliesRemoved = $Replies.count
            ForEach ($Reply in $Replies) {
            $Status = Invoke-MgSoftTeamChannelMessageReplyDelete -TeamId $Team.Id -ChannelId $SelectedChannel.Id `
                	-ChatMessageId $Message.Id -ChatMessageId1 $Reply.Id -ErrorAction SilentlyContinue
            }
        }
        # Now remove the message
        $Status = Invoke-MgSoftTeamChannelMessageDelete -TeamId $Team.Id -ChannelId $SelectedChannel.Id `
            -ChatMessageId $Message.Id -ErrorAction SilentlyContinue
        $MessagesRemoved++
        # Brief sleep to prevent throttling
        Start-Sleep -Milliseconds 30
    }
    Write-Host ("All done {0} messages and {1} replies removed from channel" -f $i, $RepliesRemoved.count)
} Else {
    Write-Host 'Exiting script - no messages removed'
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
