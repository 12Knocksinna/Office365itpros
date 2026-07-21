# Update-UserAccountsBatch.PS1
# This script is used to update user accounts in batch. It connects to Microsoft Graph, retrieves a set of user accounts, and updates
# two properties for each account using JSON batching.

# V1.0 3-July-2026
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Update-UserAccountsBatch.PS1
# Article at: https://office365itpros.com/2026/07/21/json-batching-primer/

Connect-MgGraph -Scopes "User.ReadWrite.All"

# Get a set of users to process. We only need the identifier and displayname 
[array]$Users = Get-MgUser -All -Filter "userType eq 'member'" -Property Id, DisplayName, CompanyName, OfficeLocation | Select-Object Id, DisplayName, CompanyName, OfficeLocation

Write-Host "Found $($Users.Count) users to update"

# Define the values for the properties to update. In this example, we update the company name and office location.
$newCompanyName = "Office 365 Maestro (UK) Limited"
$newOfficeLocation = "London"

$Users = $Users | Where-Object {$_.CompanyName -ne $newCompanyName -or $_.OfficeLocation -ne $newOfficeLocation}
If (!$Users) {
   Write-Host "No users need to be updated"
   Exit
}

# The Graph batch endpoint allows max 20 requests per batch. Figure out how
# many batches are needed
$BatchSize = 20
[int]$Batches = [Math]::Ceiling($Users.Count/$BatchSize)

For ($i = 0; $i -lt $Batches; $i++) {

    $Start = $i * $BatchSize
    $End = [Math]::Min($Start + $BatchSize - 1, $Users.Count - 1)
    $CurrentBatch = $Users[$Start..$End] | Select-Object Id, displayName

    # Build the batch request body - each entry needs a unique "id",
    # the target method/url, and the body of the PATCH
    $batchRequests = [System.Collections.Generic.List[object]]::new()
    $Counter = 0

    ForEach ($User in $CurrentBatch) {
        $Counter++
        $BatchRequest = [PSCustomObject][Ordered]@{
            id      = "$Counter"
            method  = "PATCH"
            url     = "/users/$($User.Id)"
            headers = @{ "Content-Type" = "application/json" }
            body    = @{ companyName = $newCompanyName 
                         officeLocation = $newOfficeLocation }
        }
        $BatchRequests.Add($BatchRequest)
    }

    $BatchBody = @{ Requests = $batchRequests }

    Try {
        $Response = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/`$batch" `
            -Body ($batchBody | ConvertTo-Json -Depth 10) -ErrorAction Stop
        # Check each individual response inside the batch for errors
        ForEach ($Result in $Response.responses) {
            If ($Result.status -ge 400) {
                $FailedUser = $CurrentBatch[[int]$Result.id - 1]
                Write-Warning "Failed to update $($FailedUser.DisplayName): Status $($Result.status) - $($Result.body.error.message)"
            }
        }
        Write-Host "Batch $($i + 1) of $Batches completed"
    }
    Catch {
        Write-Error "Batch $($i + 1) failed entirely: $_"
    }

    # Small pause to avoid hammering the throttling limits between batches
    Start-Sleep -Milliseconds 200
}

Write-Host "Company name update complete."

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
