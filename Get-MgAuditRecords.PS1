# Get-MgAuditRecords.PS1
# Example script to show how to run an audit query using the Microsoft Graph PowerShell SDK.
# V1.0 13-Aug-2024 (Tested with SDK V2.21)
# V1.1 10-Nov-2026 Updated with production cmdlets
# See https://github.com/12Knocksinna/Office365itpros/blob/master/Search-AuditLogGraph.PS1 for an example using Graph API requests
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Get-MgAuditRecords.PS1

Connect-MgGraph -NoWelcome -Scopes AuditLogsQuery.Read.All

Set-MgRequestContext -MaxRetry 10 -RetryDelay 15 
$AuditQueryName = ("Audit Job created at {0}" -f (Get-Date))
$StartDate = (Get-Date).AddDays(-180)
$EndDate = (Get-Date).AddDays(1)
$AuditQueryStart = (Get-Date $StartDate -format s)
$AuditQueryEnd = (Get-Date $EndDate -format s)
[array]$AuditQueryOperations = "FileModified", "FileUploaded"
$AuditQueryParameters = @{}
#$AuditQueryParameters.Add("@odata.type","#microsoft.graph.security.auditLogQuery")
$AuditQueryParameters.Add("displayName", $AuditQueryName)
$AuditQueryParameters.Add("OperationFilters", $AuditQueryOperations)
$AuditQueryParameters.Add("filterStartDateTime", $AuditQueryStart)
$AuditQueryParameters.Add("filterEndDateTime", $AuditQueryEnd)

# Submit the audit query
$AuditJob =  New-MgSecurityAuditLogQuery -BodyParameter $AuditQueryParameters
If ($AuditJob) {
    Write-Host ("Audit query {0} created successfully with ID {1}" -f $AuditQueryName, $AuditJob.Id)
} Else {
    Write-Host "Audit query failed to start - exiting"
    Return
}

# Check the audit query status every 20 seconds until it completes
[int]$i = 1
[int]$SleepSeconds = 20
$SearchFinished = $false; [int]$SecondsElapsed = 30
Write-Host "Checking audit query status..."
# Wait 30 seconds before checking the status of the audit query to give it time to start processing
Start-Sleep -Seconds 30
$AuditQueryStatus = Get-MgSecurityAuditLogQuery -AuditLogQueryId $AuditJob.Id

While ($SearchFinished -eq $false) {
   $i++
   Write-Host ("Waiting for audit search to complete. Check {0} after {1} seconds. Current state {2}" -f $i, $SecondsElapsed, $AuditQueryStatus.status)
   Switch ($AuditQueryStatus.status) {
        'succeeded' { 
            $SearchFinished = $true 
        }
        'failed'    { 
            Write-Host "Audit query failed - exiting"
            Return
        }
        default     { 
            Start-Sleep -Seconds $SleepSeconds
            $SecondsElapsed = $SecondsElapsed + $SleepSeconds
            $AuditQueryStatus = Get-MgSecurityAuditLogQuery -AuditLogQueryId $AuditJob.Id 
        }
   }
}

# Fetch the audit records returned by the query
[array]$AuditRecords = Get-MgSecurityAuditLogQueryRecord -AuditLogQueryId $AuditJob.Id -All -PageSize 999
Write-Host ("Audit query {0} returned {1} records" -f $AuditQueryName, $AuditRecords.Count)

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Record in $AuditRecords) {
    $ReportLine = [PSCustomObject][Ordered]@{
        Service          = $Record.Service
        Timestamp        = $Record.CreatedDateTime 
        UPN              = $Record.userPrincipalName
        Operation        = $Record.operation
    } 
    $Report.Add($ReportLine)
}

$Report | Sort-Object {$_.Timestamp -as [datetime]} | Out-GridView -Title ("Audit Records fetched by query {0}" -f $AuditQueryName)

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.

