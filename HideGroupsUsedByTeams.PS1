# HideGroupsUsedByTeams.PS1
# Hide the Microsoft 365 Groups used by Teams which might be still visible to Exchange clients and the Exchange address lists
# https://github.com/12Knocksinna/Office365itpros/blob/master/HideGroupsUsedByTeams.PS1

Clear-Host
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}

$HiddenGroups = 0
Write-Host "Finding team-enabled Microsoft 365 Groups and checking for any which are visible to Exchange clients"
[array]$Groups = Get-UnifiedGroup -Filter {ResourceProvisioningOptions -eq "Team"} -ResultSize Unlimited 
# Reduce to the set visible to Exchange clients
[array]$Groups = $Groups | Where-Object {$_.HiddenFromExchangeClientsEnabled -eq $False}

# Process the remaining groups and hide them from Exchange
If ($Groups.Count -ne 0) {
  ForEach ($Group in $Groups) { 
     Write-Host "Hiding" $Group.DisplayName
     $HiddenGroups++
     Set-UnifiedGroup -Identity $Group.ExternalDirectoryObjectId -HiddenFromExchangeClientsEnabled:$True -HiddenFromAddressListsEnabled:$True
  }
} Else { 
  Write-Host "No team-enabled Microsoft 365 Groups are visible to Exchange clients and address lists" 
}

Write-Host ("All done. {0} team-enabled groups hidden from Exchange clients" -f $HiddenGroups)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
