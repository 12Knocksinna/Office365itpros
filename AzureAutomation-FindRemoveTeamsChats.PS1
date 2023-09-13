# Runbook to search Teams chat threads in a tenant to find and remove threads with a specific topic.
# An example of using the Teams Delete Chat API
# V1.0 9-Sept-2023

Function Add-MessageRecipients {
# Function to build an addressee list to send email   
 [cmdletbinding()]
     Param(
     [array]$ListOfAddresses )
      ForEach ($SMTPAddress in $ListOfAddresses) {
           @{
              emailAddress = @{address = $SMTPAddress}
           }    
        }
} 
# Start
# Connect to the Graph SDK with the correct permissions
Connect-MgGraph -Identity -Scopes User.Read.All, Chat.Read.Basic, Mail.Send
$Organization = Get-MgOrganization

# Find users with licenses
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All
Write-Output ("Found {0} user accounts to process..." -f $users.Count)
# Define an array of chat thread topics that we want to find and remove
# Other filters are possible, but this is the easiest one to use to illustrate the principal
[array]$Topics = "Loopy Conversation", "Sensitive Stuff", "Project Aurora Stand-up", "Supervision troubleshooting"
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($User in $Users) {
    $UserId = $User.Id
    # See if any matching chat threads exist for this user
    [array]$Chats = Get-MgBetaUserChat -Userid $UserId -All -Filter "chatType eq 'Group' or chatType eq 'Meeting'" | Where-Object {$_.Topic -in $Topics}
    If ($Chats) {
        ForEach ($Chat in $Chats) {
            Write-Output ("Found chat in account {0} with topic {1}" -f $User.displayName, $Chat.topic) 
            # Get members of the chat thread
            [array]$Members = Get-MgBetaUserChatMember -UserId $UserId -ChatId $Chat.Id
            # Extract the member email addresses and remove any blanks (accounts that no longer exist)
            [array]$MemberNames = $Members.additionalProperties.email | Sort-Object -Unique
            $NumberOfParticipants = $MemberNames.Count
            If ($MemberNames.Count -gt 0) {
                $MemberNamesOutput = $MemberNames -Join ", "
            }
            # If the chat thread originates in this tenant, we can delete it
            If ($Chat.TenantId -eq $Organization.Id) {
            # Soft-delete the chat thread
                Write-Output ("Removing the chat thread {0}..." -f $Chat.Id)
                Remove-MgBetaChat -ChatId $Chat.Id
                # Wait a second to avoid throttling
                Start-Sleep -Seconds 1
                $ReportLine = [PSCustomObject][Ordered]@{
                  Timestamp            = (Get-date)
                  Action               = "Chat deleted"
                  User                 = $User.UserPrincipalName
                  UserId               = $User.Id
                  ChatType             = $Chat.ChatType
                  Topic                = $Chat.Topic
                  Participants         = $MemberNamesOutput
                  NumberOfParticipants = $NumberOfParticipants
                  Created              = $Chat.CreatedDateTime
                  LastUpdated          = $Chat.LastUpdatedDateTime
                  TenantId             = $Chat.TenantId
                  Id                   = $Chat.Id }     
            } Else {  
                 Write-Host ("Chat thread {0} originates in tenant {1} and cannot be removed" -f $Chat.topic, $Chat.tenant)
                 $ReportLine = [PSCustomObject][Ordered]@{
                  Timestamp            = (Get-date)
                  Action               = "Chat found in other tenant"
                  User                 = $User.UserPrincipalName
                  UserId               = $User.Id
                  ChatType             = $Chat.ChatType
                  Topic                = $Chat.Topic
                  Participants         = $MemberNamesOutput
                  NumberOfParticipants = $NumberOfParticipants
                  Created              = $Chat.CreatedDateTime
                  LastUpdated          = $Chat.LastUpdatedDateTime
                  TenantId             = $Chat.TenantId
                  Id                   = $Chat.Id }  
            }    
            # Update what we found     
            $Report.Add($ReportLine) 
        } # End Foreach Chats
    } # End if Chats
} #End Foreach Users

$Report | Format-Table Timestamp, Action, User, Topic -AutoSize

$EmailRecipient = "Peter.Bedson@o365maestro.onmicrosoft.com"
# Send a message from the shared mailbox
$MsgFrom = "Azure.Management.Account@office365itpros.com"
# Add your recipient address here
$ToRecipientList   = @( $EmailRecipient )
[array]$MsgToRecipients = Add-MessageRecipients -ListOfAddresses $ToRecipientList
$MsgSubject = "Teams chat thread deletions"
$HtmlHead = "<h2>Teams chat thread deletions</h2><p>The following requests to remove chat threads have been processed.</p>"
$HtmlBody = $Report | ConvertTo-Html -Fragment 
$HtmlMsg = "</body></html><p>" + $HtmlHead + $Htmlbody + "<p>"
# Construct the message body
$MsgBody = @{
  Content = "$($HtmlMsg)"
  ContentType = 'html'  
 }

$Message =  @{subject           = $MsgSubject}
$Message += @{toRecipients      = $MsgToRecipients}  
$Message += @{body              = $MsgBody}
$Params   = @{'message'         = $Message}
$Params  += @{'saveToSentItems' = $True}
$Params  += @{'isDeliveryReceiptRequested' = $True}

# And send the message using the parameters that we've filled in
Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params
Write-Output ("Message containing deleted chat information sent to {0}!" -f $EmailRecipient)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
