# Purge-ComplianceSearchItems.PS1
# An example script to show how to purge items found by a compliance search

Clear-Host
$SearchName = 'Compliance Search Test October 2024'
Remove-ComplianceSearch -Identity $SearchName -Confirm:$False -ErrorAction SilentlyContinue
Write-Host ("Creating and running new compliance search ({0})" -f $SearchName) -ForegroundColor Yellow 
New-ComplianceSearch -Name $SearchName -ExchangeLocation All -ContentMatchQuery "A new comment on the post" -Description 'Compliance Search Test' | Out-Null
Start-ComplianceSearch -Identity $SearchName

Write-Host "Waiting for compliance search to finish..."
Do { 
    $ComplianceSearch = Get-ComplianceSearch $SearchName
    $Status = $ComplianceSearch.Status 
    Start-Sleep 3 
} While ($Status -ne 'Completed') 
[int]$Count = 0 

[int]$ItemsFound = $ComplianceSearch.Items
If (!($ItemsFound)) { 
    Write-Host "No items found in compliance search" -ForegroundColor Red 
    break 
}

# Use regex to find all instances where item count is greater than zero
$LocationsWithItemCount = [regex]::Matches($ComplianceSearch.SuccessResults, "Item count: (\d+)")
[array]$Locations = $LocationsWithItemCount | Where-Object { $_.Groups[1].Value -gt 0 }

# Use regex to extract email addresses, item counts, and total sizes
$LocationsWithEmail = [regex]::Matches($ComplianceSearch.SuccessResults, "Location: (\S+@\S+\.\S+), Item count: (\d+), Total size: (\d+)")
# Extract email addresses and item counts where item count is greater than 0
[array]$LocationsWithItems = $null
foreach ($Match in $LocationsWithEmail) {
    $Email = $match.Groups[1].Value
    $ItemCount = [int]$match.Groups[2].Value
    if ($ItemCount -gt 0) {
        $LocationsWithItems += [PSCustomObject]@{
            Email = $email
            ItemCount = $itemCount
        }
    }
}

# Figure out how many loops might be needed to remove all items
[int]$LocationsGT10 = 0; [int]$HighestValue = 0
ForEach ($Item in $LocationsWithItems) {
       $ItemCount = $Item.ItemCount
       If ($ItemCount -gt 10) {
            $LocationsGT10++
        If ($ItemCount -gt $HighestValue) {
            $HighestValue = $ItemCount
            $HighestLocation = $Item.Email
        }
    }
}
$LoopsNeeded = [math]::ceiling($HighestValue/10)
Write-Host ("Compliance search completed. {0} items found in {1} locations. {2} iterations are required to remove these items. The mailbox with most items is {3} with {4}." -f $ItemsFound, $Locations.Count, $LoopsNeeded, $HighestLocation, $HighestValue) -ForeGroundColor Yellow

While ($Count -lt $LoopsNeeded) {
    $Count++ 
    # Remove old purge action
    Write-Host "Removing previous compliance search action (if one exists)" -ForegroundColor Yellow 
    Remove-ComplianceSearchAction $SearchNameAction -Confirm:$False -ErrorAction SilentlyContinue
    Write-Host ("Submitting search purge action for run {0}" -f $Count)
    New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType SoftDelete -Confirm:$False | Out-Null
    Write-Host "Waiting for compliance purge action to finish..."
    Do { 
        $Action = Get-ComplianceSearchAction $SearchNameAction 
        $Status = $Action.Status    
        Start-Sleep -Seconds 3
    } While ($Status -ne 'Completed') 
    # Figure out what happened during the purge action
    $Action = Get-ComplianceSearchAction $SearchNameAction 
    
    # Parse the purge action results to extract the results of all locations where items were found
    If ($Action.Results -match "Item count: (\d+);") {
        $Results = [int]$matches[1]
    } Else {
        Write-Host "No items were removed by the purge"
    }
    $TotalResults = $TotalResults + $Results
    Write-Host ("Loop {0} purged {1} items" -f $Count, $Results) -ForegroundColor Green
    $Action = $null; $Status = $null; $Results = $null
}

Write-Host ("Purge completed. {0} items removed from {1} locations" -f $TotalResults, $Locations.Count) -ForegroundColor Green


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
