# Update-RetentionPolicySharedMailboxes.ps1
# New-RetentionPolicySharedMailboxes.PS1 creates a Microsoft 365 retention policy for shared mailboxes.
# This script updates the settings of the retention policy to add new shared mailbpxes and remove old ones.

# V1.0.2 3-Aug-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Update-RetentionPolicySharedMailboxes.PS1

# The script can be run interactively or as a scheduled Azure Automation runbook. The script uses the Exchange Online PowerShell
# module and the Security & Compliance Center PowerShell module to fetch shared mailboxes and update the retention policy.

# Change this value to the name of the retention policy for shared mailboxes as defined for the tenant.
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
    # Fetch the account credentials to use
    $O365Cred = Get-AutomationPSCredential -Name "AzureExchangeOperations"
    Connect-ExchangeOnline -Credential $O365Cred -DisableWAM
    Connect-IPPSSession -Credential $O365Cred
}

# Look for the Microsoft 365 retention policy for shared mailboxes
$Policy = Get-RetentionCompliancePolicy -Identity $PolicyName -DistributionDetail -ErrorAction SilentlyContinue
If ($Policy) {
    Write-Output ("Found retention policy to update: {0}" -f $Policy.Name) 
} Else {
    Write-Output ("No retention policy named {0} found" -f $PolicyName)
    Break
}

Write-Output "Looking for shared mailboxes in the tenant"
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
        $DisplayName = $SharedMailboxes | Where-Object { $_.ExternalDirectoryObjectId -eq $Mbx } | Select-Object -ExpandProperty DisplayName
        Write-Output ("Shared mailbox will be added to retention policy: {0}" -f $DisplayName)
        $NewLocations += $Mbx
    }
}

# Check each location in the retention policy to see if it is still a shared mailbox and remove it if not
ForEach ($Location in $Locations) {
    If ($SharedMailboxes.ExternalDirectoryObjectId -notcontains $Location) {
        $DisplayName = Get-Recipient -Identity $Location -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DisplayName
        If ($DisplayName) {
            Write-Output ("Shared mailbox will be removed from retention policy: {0}" -f $DisplayName)
        } Else {
            Write-Output ("Shared mailbox with identifier {0} not found in the tenant and will be removed from the retention policy" -f $Location)
        }
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

# Do the actual work to remove and add locations - depending on what we found, we issue the appropriate command
If ($RemovedLocations.count -gt 0 -and $NewLocations.count -gt 0) {
    Write-Output "Updating locations for retention policy $PolicyName"
    Set-RetentionCompliancePolicy -Identity $Policy.Name -RemoveExchangeLocation $RemovedLocations -AddExchangeLocation $NewLocations
} Elseif ($RemovedLocations.count -eq 0 -and $NewLocations.count -gt 0) {
    Write-Output "Adding locations to retention policy $PolicyName"
    Set-RetentionCompliancePolicy -Identity $Policy.Name -AddExchangeLocation $NewLocations
} ElseIf ($RemovedLocations.count -gt 0 -and $NewLocations.count -eq 0) {
    Write-Output "Removing locations from retention policy $PolicyName"
    Set-RetentionCompliancePolicy -Identity $Policy.Name -RemoveExchangeLocation $RemovedLocations
} Else {
    Write-Output "No changes to make to retention policy $PolicyName"
}

# And remove any Exchange retention policies from the added shared mailboxes
If ($NewLocations) {
    Write-Output "Updating mailbox settings to remove Exchange MRM policies from shared mailboxes newly added to the retention policy"
    ForEach ($Mbx in $NewLocations) {
       Set-Mailbox -Identity $Mbx -RetentionPolicy $null
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.