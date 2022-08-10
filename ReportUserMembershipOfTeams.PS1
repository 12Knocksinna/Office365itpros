# ReportUserMembershipOfTeams.PS1
# Example script to show how to use the Get-AssociatedTeam cmdlet to fetch team membership of individual user accounts and
# report them. The team membership includes direct membership of a shared channel in a local or external team.
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportUserMembershipOfTeams.PS1

$Modules = @( "MicrosoftTeams", "Microsoft.Graph" )
# Requires -Modules $Modules

$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "MicrosoftTeams")) {Write-Host "Please connect to the Microsoft Teams module and then restart the script"; break}

Connect-MgGraph -Scopes User.Read.All, Directory.Read.All
Select-MgProfile Beta

$Tenant = (Get-MgOrganization)
$TenantId = $Tenant.Id
$TenantName = $Tenant.DisplayName

# Find user accounts with licenses
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All
If (!($Users)) { Write-Host "No user accounts found - exiting" ; break }
$Users = $Users | Sort-Object DisplayName

[int]$i = 0
$UserTeamInfo =  [System.Collections.Generic.List[Object]]::new()

ForEach ($User in $Users) {
  $i++
  Write-Host ("Processing team membership for {0} ({1}/{2})..." -f $User.DisplayName, $i, $Users.Count)
  [array]$TeamInfo = Get-AssociatedTeam -User $User.UserPrincipalName
  ForEach ($Team in $TeamInfo) {
   If ($Team.TenantId -eq $TenantId) { # Resolve the tenant identifier to a name
      $Name = $TenantName }
   Else {
      $LookUpId = $Team.TenantId.toString()
      $Uri = "https://graph.microsoft.com/beta/tenantRelationships/findTenantInformationByTenantId(tenantId='$LookUpId')"
      $ExternalTenantData = Invoke-MgGraphRequest -Uri $Uri -Method Get
      $Name = $ExternalTenantData.DisplayName 
   }
   $TeamData = [PSCustomObject][Ordered]@{  # Write out details of the team
       Id          = $User.Id
       DisplayName = $User.DisplayName 
       UPN         = $User.UserPrincipalName
       Team        = $Team.DisplayName
       TeamId      = $Team.GroupId
       Tenant      = $Name
       TenantId    = $Team.TenantId}
     $UserTeamInfo.Add($TeamData)
  } #End ForEach Team
} # End ForEach User

[array]$ExternalTeams = $UserTeamInfo | Where-Object {$_.TenantId -ne $TenantId} | Sort-Object TeamId -Unique
$ExternalPeople = $UserTeamInfo | Where-Object {$_.TenantId -ne $TenantId} | Sort-Object UPN -Unique
$ExternalPeople = $ExternalPeople.DisplayName -Join ", "
$ExternalTenants = $ExternalTeams.Tenant | Sort-Object -Unique
$AvgTeams = [math]::round(($UserTeamInfo.Count/$Users.Count),2)

Write-Host ""
Write-Host ("Each of the {0} users belongs to an average of {1} teams" -f $Users.Count, $AvgTeams)
Write-Host ("Membership of {0} teams found in {1} external tenant(s)" -f $ExternalTeams.Count, $ExternalTenants.Count)
Write-Host ("These accounts have membership of external teams: {0}" -f $ExternalPeople)

$UserTeamInfo | Out-GridView
# Code to generate a PDF report using the PSWriteHTML module
Import-Module PSWriteHTML.psd1 -Force
$UserTeamInfo | Out-HtmlView -HideFooter -Title "User Membership in Teams Report"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
