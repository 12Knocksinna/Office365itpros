# Update-DynamicM365GroupAzureAutomation
# A script demonnstrating how to update a pseduo-dynamic Microsoft 365 group in Azure Automation.
# This script finds users with manager status (as markeed by CustomAttribuute15 = 'Manager') and adds them to a dynamic group. People who
# lose their manager status are removed from the group.

# V1.0 22-Jul-2025
# GitHub Link: https://github.com/Office365ITPros/Update-DynamicM365GroupAzureAutomation

If ([Environment]::UserInteractive) { 
    # We're running interactively...
    Write-Host "Running interactively..."
    Connect-MgGraph -Scopes Group.ReadWrite.All, GroupMember.ReadWrite.All, User.Read.All -NoWelcome
} Else { 
    # We're not, so likely in Azure Automation
    Write-Output "Running the Update Managers Group script in Azure Automation..."
    Connect-MgGraph -Identity
}

# Find the set of users who have a manager status. This is done by looking for the CustomAttribute15 attribute set to "Manager".
[array]$ManagerAccounts = Get-MgUser -Filter "onPremisesExtensionAttributes/extensionAttribute15 eq 'Manager' and userType eq 'Member'" -ConsistencyLevel eventual -All -CountVariable Managers

# Define the group to update. It can have whatever name you like, but it's not a dynamic group
$TargetGroup = Get-MgGroup -Filter "displayname eq 'Employees with a Manager title'" 

[array]$MembersTargetGroup = Get-MgGroupMember -GroupId $TargetGroup.Id -All | Select-Object -ExpandProperty Id

# Make sure that we only add members who are not already in the group.
[array]$MembersToAdd = $ManagerAccounts.Id | Where-Object { $MembersTargetGroup -notcontains $_ }
# Find the members who are in the group, but no longer have a manager status.
[array]$MembersToRemove = $MembersTargetGroup | Where-Object { $ManagerAccounts.Id -notcontains $_ }  

# Create the list of users to add to the group. We need to use the directory object ID for each user.
$Data = [System.Collections.Generic.List[Object]]::new()
# Populate the list with the directory object IDs of the users to add
$MembersToAdd | ForEach-Object {$Data.Add("https://graph.microsoft.com/V1.0/directoryobjects/{0}" -f $_)}

Write-Output "Adding new managers to our dynamic group"
# Add the users to the group in batches of 20. The API only allows 20 members to be added at a time.
While ($Data.count -ne 0) {
    $Parameters = @{"members@odata.bind" = $Data[0..19] }
    Try {
        Update-MgGroup -GroupId $TargetGroup.Id -BodyParameter $Parameters -ErrorAction Stop
    } Catch {
        Write-Error "Failed to update group: $_"
    }
    If ($Data.count -gt 20) {
        $Data.RemoveRange(0,20)
    } Else {
        $Data.RemoveRange(0,$Data.count)
    }
}

# Remove the users
Write-Output "Removing users who are no longer managers from the group"
ForEach ($Id in $MembersToRemove) {
    Try {
       # Remove-MgGroupMember -GroupId $TargetGroup.Id -DirectoryObjectId $Id -ErrorAction Stop
        Write-Host "Removed user with ID $Id from group $($TargetGroup.DisplayName)"
    } Catch {
        Write-Error "Failed to remove user with ID $Id from group: $_"
    }
}

[array]$MembersTargetGroup = Get-MgGroupMember -GroupId $TargetGroup.Id -All
Write-Output ""
Write-Output "The following are now members of the group $($TargetGroup.DisplayName):"
Write-Output ""
$MembersTargetGroup.additionalProperties.displayName

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.