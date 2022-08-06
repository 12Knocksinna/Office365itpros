# ReportTeamsChannels.PS1
# Report the set of channels existing in Teams for a tenant
#
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportTeamsChannels.PS1
# Uses the MicrosoftTeams PowerShell module and the Microsoft Graph PowerShell SDK

$Outputfile = "C:\temp\OutputChannels.csv"
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "MicrosoftTeams")) {Write-Host "Please connect to the Microsoft Teams module and then restart the script"; break}

Connect-MgGraph -Scopes Group.Read.All, Directory.Read.All
Select-MgProfile Beta
$Tenant = (Get-MgOrganization)
$TenantId = $Tenant.Id
$TenantName = $Tenant.DisplayName

[array]$Teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All
If (!($Teams)) {Write-Host "Can't find any teams - exiting"; break} 
$Teams = $Teams | Sort-Object DisplayName

$ChannelsList = [System.Collections.Generic.List[Object]]::new()
[int]$i = 0
ForEach ($Team in $Teams) {
   $i++
   Write-Host ("Processing {0} ({1}/{2})" -f $Team.DisplayName, $i, $Teams.Count)
   [array]$Channels = Get-TeamAllChannel -GroupId $Team.Id
   ForEach ($Channel in $Channels) {
   If ($Channel.TenantId -eq $TenantId) {
      $Name = $TenantName }
   Else {
      $LookUpId = $Channel.TenantId.toString()
      $Uri = "https://graph.microsoft.com/beta/tenantRelationships/findTenantInformationByTenantId(tenantId='$LookUpId')"
      $ExternalTenantData = Invoke-MgGraphRequest -Uri $Uri -Method Get
      $Name = $ExternalTenantData.DisplayName }

   $ChannelLine = [PSCustomObject][Ordered]@{  # Write out details of the group
          Team                = $Team.DisplayName
          Channel             = $Channel.DisplayName
          Description         = $Channel.Description
          MembershipType      = $Channel.MembershipType
          HostTeam            = $Channel.HostTeamId
          TenantId            = $Channel.TenantId
          Tenant              = $Name
          Id                  = $Channel.Id }
       $ChannelsList.Add($ChannelLine) } 
}

[array]$TeamsStandardChannels = $ChannelsList | ? {$_.MembershipType -eq "Standard"}
[array]$TeamsPrivateChannels = $ChannelsList | ? {$_.MembershipType -eq "Private"}
[array]$TeamsSharedChannels = $ChannelsList | ? {$_.MembershipType -eq "Shared"}
[array]$TenantsWithChannels = $ChannelsList | Sort-Object Tenant -Unique | Select-Object -ExpandProperty Tenant
$AvgChannels = [math]::round(($ChannelsList.Count/$Teams.Count),2)

Write-Host ""
Write-Host "Total Teams processed:     " $Teams.Count
Write-Host "Total channels:            " $ChannelsList.Count
Write-Host "Average channels per team: " $AvgChannels
Write-Host "Total standard channels:   " $TeamsStandardChannels.Count
Write-Host "Total private channels:    " $TeamsPrivateChannels.Count
Write-Host "Total shared channels:     " $TeamsSharedChannels.Count
Write-Host ""
Write-Host "Channels found in the following tenants:" ($TenantsWithChannels -join ", ")
Write-Host ""

$ChannelsList | Export-CSV -NoTypeInformation $Outputfile
Write-Host "Channels list is in $Outputfile"
