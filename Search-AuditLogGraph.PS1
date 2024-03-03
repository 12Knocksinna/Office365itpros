# Search-AuditLogGraph.PS1
# Example of how to search the Microsoft 365 unified audit log using the Graph API
# In this case, we use the Microsoft Graph PowerShell SDK to do the magic. Tested with SDK V2.15
# V1.0 29-Feb-2024
#

Connect-MgGraph -Scopes AuditLogsQuery.Read.All -NoWelcome
$SearchId = $null

# Construct basic search parameters
# For multiple operations, use "operationFilters" = @("fileaccessed","filedeleted")
# For record type filters, use "recordTypeFilters" = @("sharePointFileOperation","threatIntelligence")
$StartDate = $null; $EndDate = $null; $Operations = $null
$StartDate = Read-Host "Start Date for audit search"
$EndDate = Read-Host "End Date for audit search"
[array]$UserOperations = Read-Host "Enter the audit operations to search for (separated by commas)"
[array]$Operations = $UserOperations.split(",").trim(" ")
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

$Uri = "https://graph.microsoft.com/beta/security/auditLog/queries"
$SearchName = ("Audit Search {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm'))
$SearchParameters = @{}
$SearchParameters.Add("operationsFilters",$Operations)
$SearchParameters.Add("displayname", $SearchName)
$SearchParameters.Add("filterStartDateTime", $StartDateSearch)
$SearchParameters.Add("filterEndDateTime", $EndDateSearch)

$SearchQuery = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $SearchParameters
$SearchId = $SearchQuery.Id
If ($null -eq $SearchId) {
    Write-Host "Search not created"
    Break
} Else {
    $SearchId = $SearchQuery.Id
    Write-Host ("Audit log search created with id: {0} and name {1}" -f $SearchId, $SearchQuery.displayname)
}

$int = 0
$SearchFinished = $false; [int]$SecondsElapsed = 0
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}" -f $SearchId)
While ($SearchFinished -eq $false) {
   $int++
   Write-Host ("Waiting for search to complete check {0} after {1} seconds" -f $int, $SecondsElapsed)
   $SearchStatus = Invoke-MgGraphRequest -Uri $Uri -Method GET
   If ($SearchStatus.status -eq 'succeeded') {
      $SearchFinished = $true
   }
   Start-Sleep -Seconds 20
   $SecondsElapsed = $SecondsElapsed + 15
}

Write-Host "Fetching audit records found by the search..."
$SearchContinues = $false
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}/records" -f $SearchId)
[array]$SearchRecords = Invoke-MgGraphRequest -Uri $Uri -Method GET

[array]$AuditRecords = $SearchRecords.value
If ($null -ne $SearchRecords.'@odata.nextLink') {
    $SearchContinues = $true
    While ($SearchContinues -eq $true) {
        $Uri = $SearchRecords.'@odata.nextLink'
        Write-Host "Fetching next page..."
        $SearchRecords = $null
        [array]$SearchRecords = Invoke-MgGraphRequest -Uri $Uri -Method GET
        $AuditRecords += $SearchRecords.value
        If ($null -eq $SearchRecords.'@odata.nextLink') {
            $SearchContinues = $false
        }
    }
}


Write-Host ("Found {0} audit records" -f $SearchRecords.value.count)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.