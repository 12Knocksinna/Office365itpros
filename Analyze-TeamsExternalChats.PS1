# Analyze-TeamsExternalChats.PS1
# Analyze User chats to find chats with people outside the tenant
# Because the chat API uses delegated access in interactive mode, this code can only process
# chats for other users when run as an app or in an Azure Automation runbook

Connect-MgGraph -NoWelcome -Scopes User.Read.All, Chat.Read.Basic, Mail.Send

[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All
Write-Output ("Found {0} user accounts to process..." -f $users.Count)
$ChatReport = [System.Collections.Generic.List[Object]]::new()

ForEach ($User in $Users) {
    Write-Output ("Processing chats for user {0}..." -f $User.DisplayName)
    # Find the onOnOne chats for this user
    [array]$Chats = Get-MgBetaUserChat -Userid $User.Id -All -Filter "ChatType eq 'oneOnOne'"
    If ($Chats) {
        Write-Output ("Found {0} one-on-one chats for {1}" -f $Chats.count, $User.DisplayName)
        ForEach ($Chat in $Chats) {
            # Find participants
            [array]$Participants = Get-MgChatMember -ChatId $Chat.Id
            # Remove the user we're processing
            [array]$OtherParticipant = $Participants | `
               Where-Object {$_.additionalProperties.email -ne $User.userPrincipalName}
            If ($OtherParticipant.displayName) {   
             
               [array]$Chats = Get-MgChatMessage -ChatId $Chat.id 
               $Reportline = [PsCustomObject]@{
                Chat         = $Chat.Id 
                Owner        = $User.userPrincipalName
                Participant  = $OtherParticipant.additionalProperties.email
                Domain       = $OtherParticipant.additionalProperties.email.split('@')[1]
                Role         = $OtherParticipant.Roles
                ChatCreated  = $Chat.CreatedDateTime
                LastUpdated  = $Chat.LastUpdatedDateTime
                ChatMessages = $Chats.count
               }
               $ChatReport.Add($ReportLine) 
            }
        }
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.