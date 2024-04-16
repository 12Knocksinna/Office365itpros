# Search-AuditLogGraph.PS1
# Example of how to search the Microsoft 365 unified audit log using the Graph API
# In this case, we use the Microsoft Graph PowerShell SDK to do the magic. Tested with SDK V2.15.
# See this article for more information: https://practical365.com/audit-log-query-api/
# V1.0 29-Feb-2024
# https://github.com/12Knocksinna/Office365itpros/blob/master/Search-AuditLogGraph.PS1

Connect-MgGraph -Scopes AuditLogsQuery.Read.All -NoWelcome
$SearchId = $null

# Construct basic search parameters
# For multiple operations, use "operationFilters" = @("fileaccessed","filedeleted")
# For record type filters, use "recordTypeFilters" = @("sharePointFileOperation","threatIntelligence")
$StartDate = $null; $EndDate = $null; $Operations = $null
$StartDate = Read-Host "Start Date for audit search"
$EndDate = Read-Host "End Date for audit search"
[array]$UserOperations = Read-Host "Enter the audit operations to search for (separated by commas)"
[array]$Operations = $UserOperations.split(",").trim(" ").tolower()
Try {
    $StartDateSearch = (Get-Date $StartDate -format s) + "Z"
    }
Catch {
    Write-Host ("{0} is not a valid date" -f $StartDate)
    Break
}
Try {
    $EndDateSearch = (Get-Date $EndDate -format s) + "Z"
    }
Catch {
    Write-Host ("{0} is not a valid date" -f $EndDate)
    Break
}

If (!($Operations)) {
    Write-Host "No audit operations specified - exiting"
    Break
}

$SearchName = ("Audit Search {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm'))

$SearchParameters = @{
    "displayName"           = $SearchName
    "filterStartDateTime"   = $StartDateSearch
    "filterEndDateTime"     = $EndDateSearch
    "operationFilters"      = $Operations
}

Write-Host "Creating an audit search query..."
$Uri = "https://graph.microsoft.com/beta/security/auditLog/queries"
$SearchQuery = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $SearchParameters
$SearchId = $SearchQuery.Id
If ($null -eq $SearchId) {
    Write-Host "Search not created"
    Break
} Else {
    $SearchId = $SearchQuery.Id
    Write-Host ("Audit log search created with id: {0} and name {1}" -f $SearchId, $SearchQuery.displayname)
}

[int]$i = 1
[int]$SleepSeconds = 20
$SearchFinished = $false; [int]$SecondsElapsed = 20
Write-Host "Checking audit query status..."
Start-Sleep -Seconds 20
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}" -f $SearchId)
$SearchStatus = Invoke-MgGraphRequest -Uri $Uri -Method GET
While ($SearchFinished -eq $false) {
    $i++
    Write-Host ("Waiting for audit search to complete. Check {0} after {1} seconds. Current state {2}" -f $i, $SecondsElapsed, $SearchStatus.status)
    If ($SearchStatus.status -eq 'succeeded') {
        $SearchFinished = $true
    } Else {
        Start-Sleep -Seconds $SleepSeconds
        $SecondsElapsed = $SecondsElapsed + $SleepSeconds
        $SearchStatus = Invoke-MgGraphRequest -Uri $Uri -Method GET
    }
}

Write-Host "Fetching audit records found by the search..."
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}/records?`$Top=999" -f $SearchId)
[array]$SearchRecords = Invoke-MgGraphRequest -Uri $Uri -Method GET

[array]$AuditRecords = $SearchRecords.value
# Paginate to fetch all available audit records
$NextLink = $SearchRecords.'@odata.NextLink'
While ($null -ne $NextLink) {
    $SearchRecords = $null
    [array]$SearchRecords = Invoke-MgGraphRequest -Uri $NextLink -Method GET 
    $AuditRecords += $SearchRecords.value
    Write-Host ("{0} audit records fetched so far..." -f $AuditRecords.count)
    $NextLink = $SearchRecords.'@odata.NextLink' 
}

Write-Host ("Total of {0} audit records found" -f $AuditRecords.count) -ForegroundColor Red
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

$Report | Sort-Object {$_.Timestamp -as [datetime]} | Out-GridView -Title 'Audit Records'

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.