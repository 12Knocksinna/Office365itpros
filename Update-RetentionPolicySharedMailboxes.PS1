# Update-RetentionPolicySharedMailboxes.ps1
# New-RetentionPolicySharedMailboxes.PS1 creates a Microsoft 365 retention policy for shared mailboxes.
# This script updates the settings of the retention policy to add new shared mailbpxes and remove old ones.

# V1.0 3-Aug-2025
# GitHub Link:

# The script can be run interactively or as a scheduled Azure Automation runbook.

# Change this value to the tenant name for your organization
$TenantName = "redmondassociates.onmicrosoft.com"
# Change this value to the name of the retention policy for shared mailboxes
$PolicyName = "Shared Mailboxes Retention Policy"

If ([Environment]::UserInteractive) { 
    # We're running interactively...
    Write-Host "Script running interactively... connecting to Exchange Online PowerShell" -ForegroundColor Yellow
    [array]$Modules = Get-Module | Select-Object -ExpandProperty Name
    If ("ExchangeOnlineManagement" -Notin $Modules) {
        Write-Output "Connecting to Exchange Online..." 
        Connect-ExchangeOnline -showBanner:$false 
    }
    Write-Output "Connecting to Security & Compliance Center PowerShell..." 
    Connect-IPPSSession -ShowBanner:$false
} Else { 
    # We're not, so likely in Azure Automation
    Write-Output "Running the update retention policy for Microsoft 365 shared mailboxes in Azure Automation"
    Connect-ExchangeOnline -ManagedIdentity -Organization $TenantName 
    Connect-IPPSSession -ManagedIdentity -Organization $TenantName 
}

# Look for the Microsoft 365 retention policy for shared mailboxes
$Policy = Get-RetentionCompliancePolicy -Identity $PolicyName -DistributionDetail -ErrorAction SilentlyContinue
If ($Policy) {
    Write-Output ("Found retention policy to update: {0}" -f $Policy.Name) 
} Else {
    Write-Output ("No retention policy named {0} found" -f $PolicyName)
    Break
}

Write-Output "Looking for shared mailboxes in the tenant $TenantName    "
[array]$SharedMailboxes = Get-ExoMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
If ($SharedMailboxes) {
    Write-Output ("{0} shared mailboxes found" -f $SharedMailboxes.Count) 
} Else {
    Write-Output "No shared mailboxes found" -ForegroundColor Green
    Break
}   

# Find the current locations for the retention policy
[array]$Locations = $Policy.ExchangeLocation.ImmutableIdentity | Sort-Object
[array]$NewLocations = $null
[array]$RemovedLocations = $null

# Check each shared mailbox to see if it is already in the retention policy and add it if not
ForEach ($Mbx in $SharedMailboxes.ExternalDirectoryObjectId) {
    If ($Locations -notcontains $Mbx) {
        Write-Output ("Shared mailbox will be addedto retention policy: {0}" -f $Mbx)
        $NewLocations += $Mbx
    }
}

# Check each location in the retention policy to see if it is still a shared mailbox and remove it if not
ForEach ($Location in $Locations) {
    If ($SharedMailboxes.ExternalDirectoryObjectId -notcontains $Location) {
        Write-Output ("Shared mailbox will be removed from retention policy: {0}" -f $Location)
        $RemovedLocations += $Location
    }
}

[int]$NewLocationsCount = ($NewLocations.Count + $Locations.Count - $RemovedLocations.Count)
# Check if the additions and removals will create more than 1,000 locations in the retention policy
If ($NewLocationsCount -gt 1000) {
    Write-Output "The retention policy will have more than 1,000 locations after the update."
    Break
} Else {
    Write-Output "The retention policy will have $NewLocationsCount locations after the update." 
}

# Do actual work to remove and add locations
If ($RemovedLocations) {
    Write-Output "Removing locations from retention policy $PolicyName"
    Set-RetentionCompliancePolicy -Identity $Policy.Name -RemoveExchangeLocation $RemovedLocations
} Else {
    Write-Output "No locations to remove from retention policy"
}
# Let the changes propagate
Start-Sleep -Seconds 20
If ($NewLocations) {
    Write-Output "Adding locations to retention policy $PolicyName"
    Set-RetentionCompliancePolicy -Identity $Policy.Name -AddExchangeLocation $NewLocations
} Else {
    Write-Output "No locations to add to retention policy"
}

# And remove any Exchange retention policies from the added shared mailboxes
If ($NewLocations) {
    Write-Output "Removing Exchange retention policies from added shared mailboxes"
    ForEach ($Mbx in $NewLocations) {
       Set-Mailbox -Identity $Mbx -RetentionPolicy $null
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.