# SendMsgToGroupOwners.PS1
# Script to send a polite informational message to owners of Office 365 Groups to tell them that their groups might be a tad
# obsolete because no one is doing anything inside the group
# https://github.com/12Knocksinna/Office365itpros/blob/master/SendMsgToGroupOwners.ps1
# Send-MailMessage https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage?view=powershell-6
CLS
Write-Host "Working..."
$Date = Get-Date
# Select groups with a fail status from the array populated by the TeamsGroupsActivityReport.ps1 script
# Also exclude groups that don't have an owner as we can't send them email
$FailedGroups = $Report | ? {$_.Status -eq "Fail" -and $_.ManagerSmtp -ne $Null}
#Reinitialize progress bar
$ProgDelta = 100/($FailedGroups.count)
$CheckCount = 0
$GroupNumber = 0
# Check do we have suitable credentials
If (-not $SmtpCred) {
    $SmtpCred = (Get-Credential)}

$MsgFrom = $SmtpCred.UserName
$SmtpServer = "smtp.office365.com"
$SmtpPort = '587'

#HTML header with styles
$htmlhead="<html>
     <style>
      BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    </style>"
		
ForEach ($R in $FailedGroups) {      
     $GroupNumber++
     $CheckCount += $ProgDelta
     $ChatMessage = "Group not enabled for Teams"
     $MsgTo = $R.ManagerSmtp
     $GroupStatus = $MsgTo + " for " + $($R.GroupName) + " ["+ $GroupNumber +"/" + $FailedGroups.Count + "]"
     Write-Progress -Activity "Sending email to" -Status $GroupStatus -PercentComplete $CheckCount
     $MsgSubject = "You need to check the activity for the " + $($R.GroupName) + " group"
     If ($R.TeamEnabled -eq "True") {
        If ($R.LastChat -ne "No chats") {
            $LastChatDate = ([DateTime]$R.LastChat).ToShortDateString()
            $DaysSinceLastChat = (New-TimeSpan -Start $LastChatDate -End $Date).Days
            $ChatMessage = "Last Teams activity on " + $LastChatDate + " (" + $DaysSinceLastChat + " days ago)" }
        Else {$ChatMessage = "Group is Teams-enabled, but no conversations have taken place"}}
     If ($R.NumberConversations -gt 0) {
        $LastGroupConversation = ([DateTime]$R.LastConversation).ToShortDateString()
        $DaysSinceLastConversation = (New-TimeSpan -Start $LastGroupConversation -End $Date).Days
        $GroupMessage = "Last Inbox activity on " + $LastGroupConversation + " (" + $DaysSinceLastConversation + " days ago)" }
     Else {
        $GroupMessage = "No inbox activity for this group" }

     # Build HTML message
     $HtmlBody = "<body>
     <h1>Office 365 Group Non-Activity Notification</h1>
     <p><strong>Generated:</strong> $Date</p>
     Please review the activity in the <u><b>$($R.GroupName)</b></u> group as it doesn't seem to have been used too much recently. Perhaps we can remove it?
     <h2><u>Details</u></h2>
     <p>Member count:            <b>$($R.Members)</b>
     <p>Guests:                  <b>$($R.ExternalGuests)</b>
     <p>Mailbox status:          <b>$($R.MailboxStatus)</b></p>
     <p>Last conversation:       <b>$GroupMessage</b></p>
     <p>Number of conversations: <b>$($R.NumberConversations)</b></p>
     <p>Team-enabled:            <b>$($R.TeamEnabled)</b></p>
     <p>Last chat:               <b>$ChatMessage</b></p>
     <p>Number of messages:      <b>$($R.NumberChats)</b></p>
     <p>SharePoint activity:     <b>$($R.SPOActivity)</b></p>
     <p>SharePoint status:       <b>$($R.SPOStatus)</b></p>
     <p>Overall status:          <b>$($R.Status)</b><p>
     <p>
     <p>If a group has a <b><u>Fail</b></u> overall status, it means that the group is a candidate for removal due to lack of use.</p>
     </body></html>"
     $HtmlMsg = $HtmlHead + $HtmlBody
     # Set TLS version
     [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
     # Construct the message parameters and send it off...
     $MsgParam = @{
         To = $MsgTo
         From = $MsgFrom
         Subject = $MsgSubject
         Body = $HtmlMsg
         SmtpServer = $SmtpServer
         Port = $SmtpPort
         Credential = $SmtpCred
     }
     Send-MailMessage @msgParam -UseSSL -BodyAsHTML 
     Start-Sleep -Seconds 1
}

Write-Host $GroupNumber "Notification Messages sent"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
