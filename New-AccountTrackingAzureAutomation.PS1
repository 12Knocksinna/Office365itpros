# New-AccountTrackingAzureAutomation.PS1
# Test script to show how to use Azure Automation with the Microsoft Graph PowerShell SDK
# Used with this article: https://office365itpros.com/2025/01/21/azure-automation-runbook-primer/

# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/New-AccountTrackingAzureAutomation.PS1

# An Azure Automation runbook to report recently created user accounts
# Requires User.Read.All permission to access information about user accounts
# Requires Group.Read.All permission to access information about groups
# Requires Sites.ReadWrite.All to update a list (or use Sites.Selected to restrict access to the specific target site)

Connect-MgGraph -Identity -NoWelcome

# Get the site identifier for the target SharePoint site. Your site URI will be different!
$Uri = "https://office365itpros.sharepoint.com/sites/Office365Adoption"
$SiteId = $Uri.Split('//')[1].split("/")[0] + ":/sites/" + $Uri.Split('//')[1].split("/")[2]
$Site = Get-MgSite -SiteId $SiteId
If (!$Site) {
    Write-Output ("Unable to connect to site {0} with id {1}" -f $Uri, $SiteId) 
    Exit
}
$List = Get-MgSiteList -SiteId $Site.Id -Filter "displayName eq 'Tenant Statistics'"
If (!$List) {
    Write-Output ("Unable to find list 'Tenant Statistics' in site {0}" -f $Site.DisplayName)
    Exit
}   

$Date = (Get-Date).ToUniversalTime().AddDays(-30).ToString("yyyy-MM-ddTHH:mm:ssZ")
[array]$RecentUsers = Get-MgUser -Filter "createdDateTime ge $Date" -Property Id, displayName, UserType, CreatedDateTime |Sort-Object UserType
If ($RecentUsers) {
    Write-Output "User Accounts added in the last 30 days"
    Write-Output "====================================="
	$RecentUsers | Format-Table DisplayName, UserType
    Write-Output ""
}

[array]$UserAccounts = Get-MgUser -All -PageSize 500 -Filter "userType eq 'Member'"
[array]$M365Groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'unified')" -All -PageSize 500
$RunDate = Get-Date -format 'dd-MMM-yyyy HH:mm:ss'
# Can only store 255 characters in a list text field
$RecentUserAccounts = $RecentUsers.DisplayName -join ', '
If ($RecentUserAccounts.length -gt 255) {
    $RecentUserAccounts = $RecentUserAccounts.Substring(0, 252) + "..."
}

$NewItemParameters = @{
	fields = @{
	    Title               = 'Azure Automation Check'
        Rundate             = $RunDate
        NumberM365Groups    = $M365Groups.Count
        NumberUserAccounts  = $UserAccounts.Count
        RecentUserAccounts  = $RecentUserAccounts
        }
}

$NewItem = New-MgSiteListItem -SiteId $Site.Id -ListId $List.Id -BodyParameter $NewItemParameters
If ($NewItem) {
    Write-Output ("Added item to list {0}" -f $List.DisplayName)
} Else {
    Write-Output "Failed to add item to list"
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.


# Add Graph permissions to the service principal for the automation account
$GraphApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
$TargetSP = Get-MgServicePrincipal -filter "displayname eq 'M365Automation'"

[array]$Permissions = "Group.Read.All", "Sites.ReadWrite.All"

ForEach ($Permission in $Permissions){
    $Role = $GraphApp.AppRoles | Where-Object {$_.Value -eq $Permission}

    # Create the parameters for the new assignment
    $AppRoleAssignment = @{}
    $AppRoleAssignment.Add("PrincipalId",$TargetSP.Id)
    $AppRoleAssignment.Add("ResourceId",$GraphApp.Id)
    $AppRoleAssignment.Add("AppRoleId",$Role.Id)

    $RoleAssignment = New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $TargetId -BodyParameter $AppRoleAssignment
    If ($RoleAssignment.AppRoleId) {
        Write-Host ("{0} permission granted to {1}" -f $Role.Value, $TargetSP.DisplayName)
    }
}

