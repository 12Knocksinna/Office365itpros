# CheckAuditEventsRunbook.PS1
# A script to execute in an Azure Automation runbook to search the Office 365 audit log to find high-priority audit events and
# report them via email to admins.
# Get access token from the Azure Automation account and use it to connect to the Graph
Connect-AzAccount -Identity
$AccessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
Connect-MgGraph -AccessToken $AccessToken.Token

#Define the desired graph endpoint
Select-MgProfile Beta
# Get tenant name
$TenantName = (Get-MgOrganization).DisplayName
# Connect to Exchange Online with the managed identity - update your organization name here
Connect-ExchangeOnline -ManagedIdentity -Organization xxxxx.onmicrosoft.com 
$StartDate = (Get-Date).AddDays(-30)
$EndDate = (Get-Date).AddDays(1)
# Define the set of operations we're interested in picking up in the audit log
[array]$Operations = "New-TransportRule", "New-InboundConnector", "Set-TransportRule"

[array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 -Operations $Operations -Formatted
# If no records are found, exit
If (!($Records)) { Write-Output "No records found - exiting" ; Break }
$records | ft creationdate, operations, userids

# Parse out audit information to make it useful
$Report = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Record in $Records) {
  $AuditData = $Record.AuditData | ConvertFrom-Json

  $P1 = $Null; $P2 = $Null
  Switch ($Record.Operations) { # Process the audit record for each operation to extract important parameters
   "New-InboundConnector" {
      $P1 = $AuditData.Parameters | Where-Object {$_.Name -eq "EFSkipIPs"} | Select-Object -ExpandProperty Value
      $P2 = $AuditData.Parameters | Where-Object {$_.Name -eq "ConnectorType"} | Select-Object -ExpandProperty Value
   }
   "New-TransportRule" {
     $P1 = $AuditData.Parameters| Where-Object {$_.Name -eq "Name"} | Select-Object -ExpandProperty Value
     ForEach ($V in $AuditData.Parameters) {
       If ($V.Name -ne "Name") { $P2 += " " + $V.Name + ": " + $V.Value } }
     $P2 = $P2.SubString(1) # Trim first leading space
    }
   "Set-TransportRule" {
    $P1 = $AuditData.Parameters| Where-Object {$_.Name -eq "Name"} | Select-Object -ExpandProperty Value
    ForEach ($V in $AuditData.Parameters) {
       If ($V.Name -ne "Name") { $P2 += " " + $V.Name + ": " + $V.Value } }
     $P2 = $P2.SubString(1) # Trim first leading space
    }
  }
  $UserDisplayName = Get-Exomailbox -Identity $Record.UserIds | Select-Object -ExpandProperty DisplayName
  $ReportLine   = [PSCustomObject] @{
     TimeStamp  = Get-Date($Record.CreationDate) -Format g
     User       = $UserDisplayName
     Operation  = $Record.Operations
     Object     = $AuditData.Parameters | Where-Object {$_.Name -eq "Name"} | Select-Object -ExpandProperty Value
     Parameter1 = $P1
     Parameter2 = $P2
   }
  $Report.Add($ReportLine) 
}
# Define variables for the mailbox used to send the message, the recipient, and the message subject
# Change these values to match your own tenant
$MsgFrom = "Azure.Management.Account@mydomain.com"
$ToAddress = "AdminDL@mydomain.com"
$MsgSubject = "High-Priority Audit Events Found for $($TenantName)"

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
<p><strong>Generated:</strong> $(Get-Date -Format g)</p>  
<h2><u>Please Check Audit Events</u></h2>
<p><b>We've discovered some high-priority events in the unified audit log.</b></p>
<p>Please investigate the details of these events.</p><p></p>
<table class='UserTable'>
	<caption><h2><font face='Segoe UI'>High-Priority Audit Events for Review</h2></font></caption>
	<thead>
	<tr>
	    <th>Timestamp</th>
	    <th>User</th>
          <th>Operation</th>
          <th>Object</th>
          <th>P1</th>
          <th>P2</th>
	</tr>
	</thead>
	<tbody>"

ForEach ($A in $Report) {
      $HtmlBody += "<tr><td><font face='Segoe UI'>$($A.Timestamp)</font></td><td><font face='Segoe UI'>$($A.User)</td></font><td><font face='Segoe UI'>$($A.Operation)</td></font><td><font face='Segoe UI'>$($A.Object)</td></font><td><font face='Segoe UI'>$($A.Parameter1)</td><td><font face='Segoe UI'>$($A.Parameter2)</td></tr></font>"
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
$NewMessage = New-MgUserMessage -UserId $MsgFrom -Body $MessageBody -ToRecipients $EmailRecipient -Subject $MsgSubject 
# Send the message
Send-MgUserMessage -UserId $MsgFrom -MessageId $NewMessage.Id  
