# Populate-TeamsExternalAccess.PS1
# Finds the domains used by guest accounts and uses them to build an external access allow list
# https://github.com/12Knocksinna/Office365itpros/blob/master/Populate-TeamsExternalAccess.PS1

Connect-MgGraph -NoWelcome -Scopes User.Read.All
[array]$Guests = Get-MgUser -All -Filter "usertype eq 'Guest'"
Write-Host ("{0} guest accounts found" -f $Guests.Count)
$GuestList = [System.Collections.Generic.List[Object]]::new()

ForEach ($Guest in $Guests) {
    $Domain = $Guest.Mail.Split("@")[1]
    $ReportLine = [PSCustomObject][Ordered]@{  
      Guest    = $Guest.Mail
      Domain   = $Domain
      Name     = $Guest.DisplayName }
    $GuestList.Add($ReportLine)
}
Write-Host ""
Write-Host "Guest accounts found for the following domains"
Write-Host "----------------------------------------------"
$GuestList | Group-Object Domain | Sort-Object Name | Select-Object Name, Count
$Domains = $GuestList | Sort-Object Domain -Unique | Select-Object -ExpandProperty Domain

Write-Host "Connecting to Microsoft Teams to check current external access configuration"
Connect-MicrosoftTeams

# Get current set of domains configured for Teams extrenal access
$DomainConfiguration  = Get-CsTenantFederationConfiguration  | Select-Object -ExpandProperty AllowedDomains
# Check the set of domains that aren't in the current configuration
[array]$DomainsToAdd = $Domains | Where-Object {$_ -notin $DomainConfiguration.AllowedDomain.Domain}

$Prompt = "Do you want to add the following domains to the list allowed for Teams external access? " + $DomainsToAdd -join ", "
$Choice = Read-Host $Prompt

If (($Choice.ToUpper()) -eq "Y") {
 $i = 0
 ForEach ($Domain in $DomainsToAdd) {
  $i++
  Write-Host ("Adding {0} to the allowed domains list... ({1}/{2})" -f $Domain, $i, $DomainsToAdd.Count)
  Set-CsTenantFederationConfiguration -AllowedDomainsAsAList @{Add=$Domain} -ErrorAction SilentlyContinue
 }
  $DomainConfiguration  = Get-CsTenantFederationConfiguration  | Select-Object -ExpandProperty AllowedDomains
  Write-Host ("External access for Teams now includes {0} domains" -f $DomainConfiguration.AllowedDomain.Domain.count)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
