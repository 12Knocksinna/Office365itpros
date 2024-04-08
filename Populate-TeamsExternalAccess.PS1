# Populate-TeamsExternalAccess.PS1
# Finds the domains used by guest accounts and uses them to build an external access allow list
# https://github.com/12Knocksinna/Office365itpros/blob/master/Populate-TeamsExternalAccess.PS1
# V1.1 8-Apr-2024 Updated to introduce check of domain to make sure that it's valid

function Get-DomainByCheck {
# Check a domain name to make sure that it's active
  param (
      [parameter(Mandatory = $true)]
      $Domain
  )

  $Uri = ("https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByDomainName(domainName='{0}')" -f $Domain) 
  Try {	
    [array]$Global:DomainData = Invoke-MgGraphRequest -Uri $Uri -Method Get -ErrorAction Stop
    If ($DomainData.displayname -in $UnwantedRealms) {
      Return $false
    } Else {
      Return $true
    }
  } Catch {
    Return $false
  }
}

# Define the display name values we don't want to accept as valid domains. MSA Realms is returned by domains like Yahoo.com. Test_Test_Microsoft
# is returned by domains like amer.teams.ms used for inbound email to Teams channels
$Global:UnwantedRealms = "MSA Realms", "Test_Test_Microsoft"
Write-Host "Connecting to Microsoft Teams and the Microsoft Graph..."
Connect-MicrosoftTeams
Connect-MgGraph -NoWelcome -Scopes User.Read.All
Write-Host "Checking for guest accounts in the tenant..."

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

# Get current set of domains configured for Teams extrenal access
$DomainConfiguration  = Get-CsTenantFederationConfiguration  | Select-Object -ExpandProperty AllowedDomains
# Check the set of domains that aren't in the current configuration
[array]$DomainsToAdd = $Domains | Where-Object {$_ -notin $DomainConfiguration.AllowedDomain.Domain}

Write-Host ""
Write-Host "Checking domains used by guest accounts to make sure that they are Microsoft 365 tenants..."
[array]$Microsoft365Tenants = $null
ForEach ($Domain in $DomainsToAdd) {
	$Check = Get-DomainByCheck $Domain
	If ($Check -eq $false) {
    Write-Host ("Domain {0} is not a Microsoft 365 tenant" -f $Domain)
  } Else {
    $Microsoft365Tenants += $Domain
  }
}

$Prompt = ("Do you want to add the following {0} domains to the list allowed for Teams external access? {1}" -f $Microsoft365Tenants.count, ($Microsoft365Tenants -join ", "))
$Choice = Read-Host $Prompt

If (($Choice.ToUpper()) -eq "Y") {
  $i = 0
  ForEach ($Domain in $Microsoft365Tenants) {
    $i++
    Write-Host ("Adding {0} to the allowed domains list... ({1}/{2})" -f $Domain, $i, $Microsoft365Tenants.Count)
    Set-CsTenantFederationConfiguration -AllowedDomainsAsAList @{Add=$Domain} -ErrorAction SilentlyContinue
  }
    $DomainConfiguration  = Get-CsTenantFederationConfiguration  | Select-Object -ExpandProperty AllowedDomains
    Write-Host ("External access for Teams now includes {0} domains" -f $DomainConfiguration.AllowedDomain.Domain.count)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
