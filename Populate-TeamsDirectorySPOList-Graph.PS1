# Populate-TeamsDirectorySPOList-Graph.PS1
# The Microsoft Graph PowerShell SDK Version of how to populate the contents of a SharePoint List using the Teams directory data
# created by GenerateTeamsDirectory.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/Populate-TeamsDirectorySPOList-Graph.PS1

# Check that we can access CSV file created by GenerateTeamsDirectory.PS1
$CSVFile = 'C:\temp\ListofTeams.csv'
If (!(Get-Item -Path $CSVFile)) {
    Write-Host ("Can't load Teams directory data from {0}" -f $CSVFile)
    Break
}

Write-Host "Connecting to the Graph..."
Connect-MgGraph -Scopes Sites.ReadWrite.All, Sites.Manage.All -NoWelcome
$ListName = "Teams Directory - Graph"
# Get target site 
Write-Host "Fetching details of the target site and list..."
$Site =  Get-MgSite -Search 'Office 365 for IT Pros Communications'
# Get List
$List = Get-MgSiteList -SiteId $Site.Id -Filter "displayName eq 'Teams Directory - Graph'"
If ($List) {
    # Delete the list
    Write-Host ("Removing previous version of list {0}" -f $List.DisplayName)
    Remove-MgSiteList -SiteId $Site.Id -ListId $List.Id
}
# Define parameters for the new list
Write-Host "Defining the new list"
$Uri = ("https://graph.microsoft.com/v1.0/sites/{0}/Lists" -f $Site.Id)
$ListDetails = '{
    "displayName": "Teams Directory - Graph",
    "description": "Discover teams to join in Office 365 for IT Pros",
    "columns": [
      {
        "name": "Deeplink",
        "description": "Link to access the team",
        "text": { }
      },{
        "name": "Description",
        "description": "Purpose of the team",
        "text": { }
      },
      {
        "name": "Owner",
        "description": "Team owner",
        "text": { }
      },      
      {
        "name": "OwnerSMTP",
        "description": "Primary SMTP address for owner",
        "text": { }
      },
      {
        "name": "Members",
        "description": "Number of tenant menbers",
        "number": { }
      },
      {
        "name": "ExternalGuests",
        "description": "Number of external guest menbers",
        "number": { }
      },
      {
        "name": "Access",
        "description": "Public or Private access",
        "text": { }
      },
    ],
  }'
Invoke-MgGraphRequest -Uri $Uri -Method Post -Body $ListDetails | Out-Null
# Rename the Notes column that's inherited from the Links template
$List = Get-MgSiteList -SiteId $Site.Id -Filter "displayName eq 'Teams Directory - Graph'"
$ColumnId = (Get-MgSiteListColumn -SiteId  $Site.Id -ListId $List.Id | `
    Where-Object {$_.Name -eq 'Title'}).Id
Update-MgSiteListColumn -ColumnDefinitionId $ColumnId -SiteId $Site.Id -ListId $List.Id `
  -Description 'Name of the team' -DisplayName 'Team Name' -Name 'TeamName' | Out-Null

# Add records to the list
Write-Host ("Populating the {0} list with data extracted from Teams" -f $ListName)
[array]$TeamsData = Import-CSV -Path $CSVFile
[int]$i = 0
$Uri = ("https://graph.microsoft.com/v1.0/sites/{0}/lists/{1}/items" -f $Site.Id, $List.Id)
ForEach ($Team in $TeamsData) {
  Write-Host ("Adding directory record for team {0} {1}/{2}" -f $Team.Team, $i, $TeamsData.Count)
  $i++
  # Note that the Title field inherited from the blank template retains its name even if
  # it has a display name of 'Team Name'
  $FieldsDataObject  = [PSCustomObject] @{
        Title          = $Team.Team
        Deeplink       = $Team.Deeplink
        Description    = $Team.Description
        Owner          = $Team.Owner
        OwnerSMTP      = $Team.OwnerSMTP
        Members        = $Team.Members
        ExternalGuests = $Team.ExternalGuests
        Access         = $Team.Access
  }
  $NewItem = [PSCustomObject] @{
        fields         = $FieldsDataObject
  } 
  $NewItem = $NewItem | ConvertTo-Json
  $Status = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $NewItem
  If ($Status.Id) {
     Write-Host ("Record added to list with id {0}" -f $Status.Id)
  }
}   

Write-Host "All done"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
