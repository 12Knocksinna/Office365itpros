# CheckAuditEventsRunbook.PS1
# A script to execute in an Azure Automation runbook to search the Microsoft 365 audit log to find high-priority audit events and
# report them via email to admins.
# Get access token from the Azure Automation account and use it to connect to the Graph

# V1.0 October 2022 for https://practical365.com/azure-automation-managed-identity-exchange/
# V2.0 5-Nov-2026   Updated to use AuditLogQuery API to get audit records instead of the Get-AdminAuditLogSearch cmdlet 
#                   which is being deprecated in October 2027.

# GitHub Link:      https://github.com/12Knocksinna/Office365itpros/blob/master/CheckAuditEventsRunbook.PS1

Connect-MgGraph -Identity

# Get tenant name
$TenantName = (Get-MgOrganization).DisplayName
# Connect to Exchange Online with the managed identity - update your organization name here
Connect-ExchangeOnline -ManagedIdentity -Organization xxxxx.onmicrosoft.com 
# Define variables for the mailbox used to send the message, the recipient, and the message subject
# Change these values to match your own tenant
$MsgFrom = "Azure.Management.Account@mydomain.com"
$ToAddress = "AdminDL@mydomain.com"
$MsgSubject = "High-Priority Audit Events Found for $($TenantName)"

# Define the set of operations we're interested in picking up in the audit log
[array]$AuditOperationFilters = "New-TransportRule", "New-InboundConnector", "Set-TransportRule"
$AuditJobName = ("Transport rules audit job created at {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm'))
$AuditQueryStart = ((Get-Date).AddDays(-30).toUniversalTime().toString('yyyy-MM-ddTHH:mm:ss'))
$AuditQueryEnd = ((Get-Date).AddHours(1).toUniversalTime().toString('yyyy-MM-ddTHH:mm:ss'))

Try {
  $AuditJob = New-MgSecurityAuditLogQuery -DisplayName $AuditJobName -OperationFilters $AuditOperationFilters `
   -FilterStartDateTime $AuditQueryStart -FilterEndDateTime $AuditQueryEnd
  Write-Output "Audit job created successfully with ID: $($AuditJob.Id)"
} Catch {
  Write-Output "Error creating audit job: $_"
  Return
}

# Check the audit query status every 20 seconds until it completes
[int]$i = 1
[int]$SleepSeconds = 20
$SearchFinished = $false; [int]$SecondsElapsed = 20
Write-Host "Checking audit query status..."
# Initial wait to let the audit job spin up
Start-Sleep -Seconds 30
$AuditQueryStatus = Get-MgSecurityAuditLogQuery -AuditLogQueryId $AuditJob.Id

While ($SearchFinished -eq $false) {
   $i++
   Write-Host ("Waiting for audit search to complete. Check {0} after {1} seconds. Current state {2}" -f $i, $SecondsElapsed, $AuditQueryStatus.status)
   If ($AuditQueryStatus.status -eq 'succeeded') {
      $SearchFinished = $true
   } Else {
      Start-Sleep -Seconds $SleepSeconds
      $SecondsElapsed = $SecondsElapsed + $SleepSeconds
      $AuditQueryStatus =  Get-MgSecurityAuditLogQuery -AuditLogQueryId $AuditJob.Id    
   }
}

[array]$AuditRecords = Get-MgSecurityAuditLogQueryRecord -All -AuditLogQueryId $AuditJob.Id
# If no records are found, exit
If (!($AuditRecords)) { 
  Write-Output "No records found - exiting" 
  Return
} Else {
  Write-Output ("Found {0} audit records" -f $AuditRecords.Count)
}

$AuditRecords | Format-Table CreatedDateTime, Operation, Userid -AutoSize
$AuditRecords = $AuditRecords | Sort-Object {$_.CreatedDateTime -as [datetime]} -Descending

# Parse out audit information to make it useful
$Report = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Record in $AuditRecords) {
  $AuditData = $Record.AuditData.additionalProperties 
  $UserDisplayName = Get-Exomailbox -Identity $Record.UserId | Select-Object -ExpandProperty DisplayName
  $ReportLine   = [PSCustomObject] @{
     TimeStamp   = Get-Date($Record.CreatedDateTime) -Format 'dd-MMM-yyyy HH:mm:ss'
     User        = $Record.UserId
     DisplayName = $UserDisplayName
     Operation   = $Record.Operation
     Object      = $AuditData.ObjectId
  }
  $Report.Add($ReportLine) 
}

# Define HTML header with styles
$htmlhead="<style>
	.UserTable {
		border:1px solid #C0C0C0;
		border-collapse:collapse;
		padding:5px;
	}
	.UserTable th {
		border:1px solid #C0C0C0;
		padding:5px;
		background:#F0F0F0;
	}
	.UserTable td {
		border:1px solid #C0C0C0;
		padding:5px;
	}
</style>"

# Build the message including the audit details in a table
$HtmlBody = "<body>
<p><font size='2' face='Segoe UI'>
<p><strong>Generated:</strong> $(Get-Date -Format 'dd-MMM-yyyy HH:mm')</p>  
<h2><u>Please Check Audit Events</u></h2>
<p><b>We've discovered some high-priority events in the Microsoft 365 unified audit log.</b></p>
<p>Please investigate the details of these events.</p><p></p>
<table class='UserTable'>
	<caption><h2><font face='Segoe UI'>High-Priority Audit Events for Review</h2></font></caption>
	<thead>
	<tr>
	    <th>Timestamp</th>
	    <th>User</th>
          <th>DisplayName</th>
          <th>Operation</th>
          <th>Object</th>
          <th>P1</th>
          <th>P2</th>
	</tr>
	</thead>
	<tbody>"

ForEach ($A in $Report) {
      $HtmlBody += "<tr><td><font face='Segoe UI'>$($A.Timestamp)</font></td><td><font face='Segoe UI'>$($A.User)</td></font><td><font face='Segoe UI'>$($A.DisplayName)</td></font><td><font face='Segoe UI'>$($A.Operation)</td></font><td><font face='Segoe UI'>$($A.Object)</td></font><td><font face='Segoe UI'>$($A.Parameter1)</td><td><font face='Segoe UI'>$($A.Parameter2)</td></tr></font>"
    }
$HtmlBody += "</tbody></table><p>" 
$HtmlBody += '</body></html>'

$EmailAddress  = @{address = $ToAddress} 
$EmailRecipient = @{EmailAddress = $EmailAddress}  
    
$HtmlHeaderUser = "<h2>High Priority Audit Events</h2>"    
$HtmlMsg = "</html>" + $HtmlHead + $htmlbody + "<p>"
# Construct the message body
$MessageBody = @{
    content = "$($HtmlBody)"
    ContentType = 'html'  }

# Create a draft message in the mailbox used to send the message
Try {
  $NewMessage = New-MgUserMessage -UserId $MsgFrom -Body $MessageBody -ToRecipients $EmailRecipient -Subject $MsgSubject 
} Catch {
  Write-Output "Error creating message: $_"
  Return
}
  # Send the message
Try {
  Send-MgUserMessage -UserId $MsgFrom -MessageId $NewMessage.Id  
  Write-Output "Email sent successfully to $($ToAddress)"
} Catch {
  Write-Output "Error sending email: $_"
  Return
}
# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 