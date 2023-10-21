# Populate-TeamsDirectorySPOList-Graph.PS1
# The Microsoft Graph PowerShell SDK Version of how to populate the contents of a SharePoint List using the Teams directory data
# created by GenerateTeamsDirectory.PS1

Connect-MgGraph -Scopes Sites.ReadWrite.All, Sites.Manage.All -NoWelcome
# Get target site 
$Site =  Get-MgSite -Search 'Office 365 for IT Pros Communications'
# Get List
$List = Get-MgSiteList -SiteId $Site.Id -Filter "displayName eq 'Teams Directory'"
If ($List) {
    # Delete the list
    Write-Host ("Removing previous version of list {0}" -f $List.DisplayName)
    Remove-MgSiteList -SiteId $Site.Id -ListId $List.Id
}
# Define parameters for the new list
$Uri = ("https://graph.microsoft.com/v1.0/sites/{0}/Lists" -f $Site.Id)
$ListDetails = '{
    "displayName": "Teams Directory",
    "description": "Discover teams to join in Office 365 for IT Pros",
    "columns": [
      {
        "name": "Description",
        "text": { }
      },
      {
        "name": "Owner",
        "text": { }
      },      
      {
        "name": "OwnerSMTP",
        "text": { }
      },
      {
        "name": "Members",
        "number": { }
      },
      {
        "name": "ExternalGuests",
        "description": "Number of external guest menbers",
        "number": { }
      },
      {
        "name": "Access",
        "text": { }
      },
    ],
    "list": {
      "template": "Links"
    }
  }'
Invoke-MgGraphRequest -Uri $Uri -Method Post -Body $ListDetails | Out-Null
# Rename the Notes column that's inherited from the Links template
$List = Get-MgSiteList -SiteId $Site.Id -Filter "displayName eq 'Teams Directory'"
$ColumnId = (Get-MgSiteListColumn -SiteId  $Site.Id -ListId $List.Id | Where-Object {$_.displayName -eq 'Comments'}).Id
Update-MgBetaSiteListColumn -ColumnDefinitionId $ColumnId -SiteId $Site.Id -ListId $List.Id `
  -Description 'Name of the team' -DisplayName 'Team Name' -Name 'TeamName' | Out-Null


$Uri = ("https://graph.microsoft.com/v1.0/sites/{0}/Columns/{1}" -f $Site.Id, $ColumnId)
Invoke-MgGraphRequest -Uri $Uri -Method DELETE
