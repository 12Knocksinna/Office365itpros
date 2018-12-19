# SendMsgToGroupOwners.PS1
# Script to send a polite informational message to owners of Office 365 Groups to tell them that their groups might be a tad
# obsolete because no one is doing anything inside the group
CLS
Write-Host "Working..."
# Select groups with a fail status from the array populated by the TeamsGroupsActivityReport.ps1 script
# Also exclude groups that don't have an owner as we can't send them email
$FailedGroups = $Report | ? {$_.Status -eq "Fail" -and $_.ManagerSmtp -ne $Null}
#Reinitialize progress bar
$ProgDelta = 100/($FailedGroups.count)
$CheckCount = 0
$GroupNumber = 0
If (-not $SmtpCred) {
    $SmtpCred = (Get-Credential)}

# Credential is in $O365Cred, so we can fetch the SMTP address for the user
#$SmtpFrom = $O365Cred.UserName  
#$SmtpCred = $O365Cred
#$SmtpServer = "office365itpros-com.mail.protection.outlook.com"
#$SmtpPort = "25"
$MsgFrom = $SmtpCred.UserName
$SmtpServer = "smtp.office365.com"
$SmtpPort = '587'

#HTML HEAD with styles
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
     $MsgTo = $R.ManagerSmtp
     $GroupStatus = $ToAddress + " for " + $($R.GroupName) + " ["+ $GroupNumber +"/" + $FailedGroups.Count + "]"
     Write-Progress -Activity "Sending email to" -Status $GroupStatus -PercentComplete $CheckCount
     $MsgSubject = "You need to check the activity for the " + $($R.GroupName) + " group"
     $HtmlBody = "<body>
     <h1>Office 365 Group Non-Activity Notification</h1>
     <p><strong>Generated:</strong> $Date</p>
     Please review the activity in the $GroupName group as it doesn't seem to have been used too much recently. Perhaps we can remove it?
     <h2><u>Details</u></h2>
     <p>Member count:            <b>$($R.Members)</b>
     <p>Guests:                  <b>$($R.ExternalGuests)</b>
     <p>Mailbox status:          <b>$($R.MailboxStatus)</b></p>
     <p>Last conversation:       <b>$($R.LastConversation)</b></p>
     <p>Number of conversations: <b>$($R.Numberconversations)</b></p>
     <p>Team-enabled:            <b>$($R.TeamEnabled)</b></p>
     <p>Last chat:               <b>$($R.LastChat)</b></p>
     <p>Number of messages:      <b>$($R.NumberChats)</b></p>
     <p>SharePoint activity:     <b>$($R.SPOActivity)</b></p>
     <p>SharePoint status:       <b>$($R.SPOStatus)</b></p>
     <p>Overall status:          <b>$($R.Status)</b><p>
     <p>
     <p>If a group has a <b><u>Fail</b></u> overall status, it means that the group is a candidate for removal due to lack of use.</p>
     </body></html>"
     $HtmlMsg = $HtmlHead + $HtmlBody
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

