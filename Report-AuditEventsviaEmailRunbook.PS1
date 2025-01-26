# Report-AuditEventsViaEmailRunbook.PS1
# an Azure Automation runbook to show how to send email about an audit event. You can run the script
# interactively using an app-only session. This means that you'll connect to the Graph SDK with an
# appId, tenantId, and certificate. The app must have consent for the required application permissions.

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-AuditEventsviaEmailRunbook.PS1
# V1.0 27-Jan-2025

# Requires Mail.Send and AuditLogsQuery.Read.All. Alternatively, consider using RBAC for Applications to limit
# access for the managed identity to specific mailboxes. See https://practical365.com/rbac-for-applications-azure-automation/
Connect-MgGraph -Identity

Set-MgRequestContext -MaxRetry 10 -RetryDelay 15
$AuditQueryName = ("Azure Automation Runbook Scan created at {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm'))
$StartDate = (Get-Date).AddDays(-7)
$EndDate = (Get-Date).AddDays(1)
$AuditQueryStart = (Get-Date $StartDate -format s)
$AuditQueryEnd = (Get-Date $EndDate -format s)
[array]$AuditQueryOperations = "Update user."
$AuditQueryParameters = @{}
#$AuditQueryParameters.Add("@odata.type","#microsoft.graph.security.auditLogQuery")
$AuditQueryParameters.Add("displayName", $AuditQueryName)
$AuditQueryParameters.Add("OperationFilters", $AuditQueryOperations)
$AuditQueryParameters.Add("filterStartDateTime", $AuditQueryStart)
$AuditQueryParameters.Add("filterEndDateTime", $AuditQueryEnd)

# Submit the audit query
$AuditJob =  New-MgBetaSecurityAuditLogQuery -BodyParameter $AuditQueryParameters

# Check the audit query status every 20 seconds until it completes
[int]$i = 1
[int]$SleepSeconds = 20
$SearchFinished = $false; [int]$SecondsElapsed = 20
Write-Host "Checking audit query status..."
Start-Sleep -Seconds 30
# This cmdlet is not working...
#$AuditQueryStatus = Get-MgBetaSecurityAuditLogQuery -AuditLogQueryId $AuditJob.Id
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}" -f $AuditJob.id)
$AuditQueryStatus = Invoke-MgGraphRequest -Uri $Uri -Method Get

While ($SearchFinished -eq $false) {
    $i++
    Write-Host ("Waiting for audit search to complete. Check {0} after {1} seconds. Current state {2}" -f $i, $SecondsElapsed, $AuditQueryStatus.status)
    If ($AuditQueryStatus.status -eq 'succeeded') {
        $SearchFinished = $true
    } Else {
        Start-Sleep -Seconds $SleepSeconds
        $SecondsElapsed = $SecondsElapsed + $SleepSeconds
        # $AuditQueryStatus = Get-MgBetaSecurityAuditLogQuery -AuditLogQueryId $AuditJob.Id
        $AuditQueryStatus = Invoke-MgGraphRequest -Uri $Uri -Method Get
    }
}

# Fetch the audit records returned by the query
# This cmdlet isn't working either
# [array]$AuditRecords = Get-MgBetaSecurityAuditLogQueryRecord -AuditLogQueryId $AuditJob.Id -All -PageSize 999
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}/records" -f $AuditJob.Id)
[array]$SearchRecords = Invoke-MgGraphRequest -Uri $Uri -Method GET
[array]$AuditRecords = $SearchRecords.value

$NextLink = $SearchRecords.'@Odata.NextLink'
While ($null -ne $NextLink) {
    $SearchRecords = $null
    [array]$SearchRecords = Invoke-MgGraphRequest -Uri $NextLink -Method GET 
    $AuditRecords += $SearchRecords.value
    Write-Host ("{0} audit records fetched so far..." -f $AuditRecords.count)
    $NextLink = $SearchRecords.'@odata.NextLink' 
}

Write-Host ("Audit query {0} returned {1} records" -f $AuditQueryName, $AuditRecords.Count)

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $AuditRecords) {
    $AuditData = $Rec.AuditData
    If ('UserPrincipalName' -in $rec.AuditData.ModifiedProperties.Name) {
        # Extract the values for the old and new UPNs from the Auditdata payload
        $OldUPN = $AuditData.ModifiedProperties | Where-Object {$_.Name -eq 'UserPrincipalName'} | Select-Object -ExpandProperty OldValue
        $EmailAddress = $OldUPN -match '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b' | Out-Null
        $OldUPN = $Matches[0]
        $NewUPN = $AuditData.ModifiedProperties | Where-Object {$_.Name -eq 'UserPrincipalName'} | Select-Object -ExpandProperty NewValue
        $EmailAddress = $NewUPN -match '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b' | Out-Null
        $NewUPN = $Matches[0]
        If ($OldUPN -ne $NewUPN) {
            $ReportLine = [PSCustomObject][Ordered]@{
                Timestamp        = $Rec.CreatedDateTime 
                Actor            = $Rec.userPrincipalName
                Operation        = $Rec.operation
                OldValue         = $OldUPN
                NewValue         = $NewUPN
            } 
            $Report.Add($ReportLine)
        }
    }
}

$Report | Sort-Object {$_.Timestamp -as [datetime]} | Format-Table -AutoSize

$MsgFrom = 'Customer.Services@office365itpros.com'
# Define some variables used to construct the HTML content in the message body
#HTML header with styles
$HtmlHead="<html>
    <style>
    BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    H4{font-size: 8px; padding-top: 4px;}
</style>"

$MsgSubject = "Audit Events for Your Review"

$ToRecipients = @{}
$ToRecipients.Add("emailAddress",@{'address'='tony.redmond@office365itpros.com'})
[array]$MsgTo = $ToRecipients
# Customize the message 
$HtmlHeaderUser = "<h2>Administrative alert: User Principal Name Updates in the last week</h2>"   
# Add some content for the message - obviously, this is very customizable and should reflect what you want to say about the data being reported
$HtmlBody = "<body> <h1>Please Check these audit events</h1><p></p>"
$HtmlBody = $HtmlBody + ($Report | ConvertTo-Html -Fragment)
$HtmlBody = $HtmlBody + "<p>These audit records are highlighted because of the impact a change to a user principal name can have on systems.</p>"
$HtmlBody = $HtmlBody + "<p><h4>Generated:</strong> $(Get-Date -Format g)</h4></p>"
$HtmlMsg = "</body></html>" + $HtmlHead + $Htmlheaderuser + $HtmlBody + "<p>"
# Construct the message body
$MsgBody = @{}
$MsgBody.Add('Content', "$($HtmlMsg)")
$MsgBody.Add('ContentType','html')

$Message = @{}
$Message.Add('subject', $MsgSubject)
$Message.Add('toRecipients', $MsgTo)    
$Message.Add('body', $MsgBody)
$Params = @{}
$Params.Add('message', $Message)
$Params.Add('saveToSentItems', $true)
$Params.Add('isDeliveryReceiptRequested', $true)    

Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.