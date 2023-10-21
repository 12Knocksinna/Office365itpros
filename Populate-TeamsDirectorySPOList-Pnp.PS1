# Populate-TeamsDirectorySPOList-PnP.PS1
# Example of how to populate the contents of a SharePoint List using the Teams directory data
# created by GenerateTeamsDirectory.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/Populate-TeamsDirectorySPOList-Pnp.PS1

# Check that we can access CSV file created by GenerateTeamsDirectory.PS1
$CSVFile = 'C:\temp\ListofTeams.csv'
If (!(Get-Item -Path $CSVFile)) {
    Write-Host ("Can't load Teams directory data from {0}" -f $CSVFile)
    Break
}
# Get credentials for Pnp SharePoint
$Credentials = Get-Credential
# Define target site where the list is created
$Site =  "https://o365maestro.sharepoint.com/sites/Office365forITProsCommunications"
$ListName = "Teams Directory"
Import-Module Pnp.PowerShell
$Connection = Connect-PnPOnline -Url $Site -Credentials $Credentials
# Check if list already exists and if so, remove it
$List = Get-PnPList | Where-Object {$_.DisplayName -eq $ListName}
If ($List) {
    Write-Host ("Removing previous version of list {0}" -f $List.Title)
    Remove-PnPList -Identity $ListName -Force
}
# Create new list
Write-Host ("Creating new list for {0}" -f $ListName)
New-PnpList -Title $ListName -Template Links -EnableVersioning -Connection $Connection | Out-Null
# Add fields
Add-PnpField -List $ListName -DisplayName 'Team Name' -Internalname TeamName -Type Text -AddToDefaultView | Out-Null
Add-PnpField -List $ListName -DisplayName 'Description' -Internalname Description -Type Text -AddToDefaultView | Out-Null
Add-PnpField -List $ListName -DisplayName 'Owner' -Internalname Owner -Type Text -AddToDefaultView | Out-Null
Add-PnpField -List $ListName -DisplayName 'Owner SMTP Address' -Internalname OwnerSMTP -Type Text -AddToDefaultView | Out-Null
Add-PnpField -List $ListName -DisplayName 'Member count' -Internalname MemberCount -Type Number -AddToDefaultView | Out-Null
Add-PnpField -List $ListName -DisplayName 'External count' -Internalname ExternalCount -Type Number -AddToDefaultView | Out-Null
Add-PnpField -List $ListName -DisplayName 'Access' -Internalname AccessMode -Type Text -AddToDefaultView | Out-Null
# Remove the Notes field inherited from the Links template
Remove-PnPField -List $ListName -Identity Notes -Force

Write-Host ("Populating the {0} list with data extracted from Teams" -f $ListName)
[array]$TeamsData = Import-CSV -Path $CSVFile
[int]$i = 0
ForEach ($Team in $TeamsData) {
    $i++
    Write-Host ("Adding record for team {0} {1}/{2}" -f $Team.Team, $i, $TeamsData.count)
    Add-PnPListItem -List $ListName -Values @{
        "URL" = $($Team.Deeplink);
        "TeamName" = $($Team.Team);
        "Description" = $($Team.Description);
        "Owner" = $($Team.Owner);
        "OwnerSMTP" = $($Team.OwnerSMTP);
        "MemberCount" = $($Team.Members);
        "ExternalCount" = $($Team.ExternalGuests);
        "AccessMode" = $($Team.Access);
    } | Out-Null
}

Write-Host "All done"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
