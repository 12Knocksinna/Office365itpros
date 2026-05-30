# Runbook to search Teams chat threads in a tenant to find and remove threads with a specific topic.
# An example of using the Teams Delete Chat API
# V1.0 9-Sept-2023
# V1.1 30-May-2026 Updated cmdlets to make sure that everything works with the latest Graph SDK and V1.0 cmdlets

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
# Connect to the Graph SDK with the correct permissions - if run in Azure Automation, make sure that the automation account has these permissions.

# User.Read.All - read user accounts
# Chat.ReadWrite.All - read chat messages
# Chat.ManageDeletion.All - manage chat deletions

# For Azure Automation, use a managed identity 
Connect-MgGraph -Identity 

# Otherwise for app-only access
# $AppId = '8af0f6ae-bb6c-416d-ab6b-45668a3c15ee'
# $TenantId = '72f988bf-86f1-4111-9a1a-0123456789ab'
# $CertificateThumbprint = '1234567890abcdef1234567890abcdef12345678
# Connect-MgGraph -ClientId "app-id" -TenantId "tenant-id" -CertificateThumbprint "thumbprint"
$TenantId = (Get-MgOrganization).Id

# Parameters for search
# Search only goes back a month - change this if you think it should go cover a different period. 
$StartDate = (Get-Date).AddMonths(-1).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$EndDate   = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
# Define an array of chat thread topics that we want to find and remove
# Other filters are possible, but this is the easiest one to use to illustrate the principal
[array]$Topics = "Loopy Conversation", "Sensitive Stuff", "Project Aurora Stand-up", "Supervision troubleshooting", "Inappropriate content"

# Compiled regex pattern to check if any of the topics are in the chat topic or message content. This is used to avoid having to loop through all the topics for each message. The regex will be used with the -match operator, which is case-insensitive by default.
$Pattern = ($Topics | ForEach-Object {[regex]::Escape($_)}) -join "|"
$Regex = [regex]::new($Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

# Find users with licenses
Try {
    [array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All
} Catch {
    Write-Output "Error getting users: $_"
    Break
}

If ($Users) {
    Write-Output ("Found {0} user accounts to process..." -f $Users.Count)
} Else {
    Write-Output "No user accounts found to process"
    Break
}

ForEach ($User in $Users) {
    # See if any matching chat threads exist for this user
    Try {
        # Find group or meeting topics created in the last month about the topics we're loooking for.
        [array]$Chats = Get-MgUserChat -Userid $UserId -All -Filter "(chatType eq 'Group' or chatType eq 'Meeting') and (createdDateTime ge $StartDate and createdDateTime le $EndDate) and (tenantId eq '$TenantId')" 
        Write-Host ("Checking {0} chats for user {1}" -f $Chats.Count, $User.displayName)
    } Catch {
        Write-Output ("Error getting chats for user {0}: $_" -f $User.UserPrincipalName)
        Continue
    }
    If ($Chats) {
        $Report = [System.Collections.Generic.List[Object]]::new()        
        ForEach ($Chat in $Chats) {
            $DeleteThread = $false
            $MemberNamesOutput = $null
            $TopicFound = $null
            # Is the chat topic one of the ones we're looking for? If so, it should be deleted
            If ($Chat.Topic -in $Topics) {
                Write-Output ("Found chat with topic {0} for user {1}" -f $Chat.Topic, $User.displayName)
                $DeleteThread = $true
                $TopicFound = "Topic"
            }
            If ($DeleteThread -eq $false) {
                # If not, we need to check the content of the messages in the thread to see if any of them contain the topic. 
                # The API doesn't allow us to filter messages based on content, so we have to get all messages and check them one by one.
                # This is not ideal, but it's the only way to do it at the moment. 
            # Is the banned topic in the text of any of the messages in the thread?
                [array]$Messages = Get-MgChatMessage -ChatId $Chat.Id -All
                ForEach ($Message in $Messages) {
                    # Strip out HTML tags
                    $Content = ($Message.Body.Content -replace '<[^>]+>', '')
                    # Check messages to see if they contain any of the topics we're looking for. If so, the thread should be deleted. We only need one message to match to delete the thread, so we can stop checking after the first match.
                    If ($Regex.IsMatch($Content) -and $DeleteThread -eq $false) {
                        Write-Output ("Found chat message with matching content in chat {0} for user {1}" -f $Chat.Id, $User.displayName)
                        $DeleteThread = $true
                        $TopicFound = "Content"
                    }
                }
            }
            If ($DeleteThread -eq $true) {
                # Get members of the chat thread
                [array]$Members = Get-MgUserChatMember -UserId $UserId -ChatId $Chat.Id
                # Extract the member email addresses and remove any blanks (accounts that no longer exist)
                [array]$MemberNames = $Members.additionalProperties.email | Sort-Object -Unique
                $NumberOfParticipants = $MemberNames.Count
                If ($MemberNames.Count -gt 0) {
                    $MemberNamesOutput = $MemberNames -Join ", "
                }

                # Soft-delete the chat thread
                Write-Output ("Removing the chat thread {0}..." -f $Chat.Id)
                Try {
                    Remove-MgChat -ChatId $Chat.Id
                    $ReportLine = [PSCustomObject][Ordered]@{
                        Timestamp            = (Get-date)
                        Action               = "Chat deleted"
                        User                 = $User.UserPrincipalName
                        UserId               = $User.Id
                        ChatType             = $Chat.ChatType
                        Topic                = $Chat.Topic
                        "Problem found in"   = $TopicFound
                        Participants         = $MemberNamesOutput
                        NumberOfParticipants = $NumberOfParticipants
                        Created              = $Chat.CreatedDateTime
                        LastUpdated          = $Chat.LastUpdatedDateTime
                        TenantId             = $Chat.TenantId
                        Id                   = $Chat.Id 
                    }                     
                    # Update what we found     
                    $Report.Add($ReportLine) 
                } Catch {
                    Write-Output ("Error removing chat {0}: $_" -f $Chat.Id)
                }
                # Wait a second to avoid throttling
                Start-Sleep -Seconds 1
                
            }
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
