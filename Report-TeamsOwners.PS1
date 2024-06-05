# Report-TeamsOwners.PS1
# A quick script to show how to use the Microsoft Graph PowerShell SDK to generate a list of Teams and their owners
# 

Connect-MgGraph -NoWelcome -Scopes Directory.Read.All
Write-Host "Finding teams..."
[array]$Teams = Get-MgGroup -Filter "resourceProvisioningOptions/any(x:x eq 'Team')" -All -PageSize 999 | Sort-Object DisplayName
If ($Teams) {
    Write-Host ("Found {0} teams. Now processing..." -f $Teams.count)
} Else {
    Write-Host "No teams found - exiting"; break
}
$Report = [System.Collections.Generic.List[Object]]::new()
[int]$i = 0
ForEach ($Team in $Teams) {
    $i++
    Write-Host ("Processing {0} ({1}/{2})..." -f $Team.DisplayName, $i, $Teams.count)
    [array]$Owners = Get-MgGroupOwner -GroupId $Team.Id
    If ($Owners) {
        $TeamOwners = $Owners.AdditionalProperties.displayName -Join ", "
    } Else {
        $TeamOwners = "No team owners"
    }
    $DataLine = [PSCustomObject][Ordered]@{
        Team        = $Team.displayName
        Owners      = $TeamOwners
        Visibility  = $Team.Visibility
        ID          = $Team.Id
    }
    $Report.Add($DataLine)
}
$Report | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.