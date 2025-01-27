# Report-ServicePrincipalsWithHighPermissions.PS1
# An example of how to scan the assignments for highly-privileged Entra ID roles and report the service principals (apps)
# and groups that have these roles assigned.

# Find the Entra roles marked as highly privileged
[array]$PrivilegedRoles = Get-MgBetaRoleManagementDirectoryRoleDefinition -Filter "isPrivileged eq true" -All -PageSize 500
If (!$HighlyPrivilegedRoles) {
    Write-Host "No highly privileged roles found"
    Exit
}

# Create a hash table to store the highly privileged roles - it's faster to search a hash table than an array
$HighlyPrivilegedRoles = @{}
ForEach ($Role in $PrivilegedRoles) {
    $HighlyPrivilegedRoles.Add($Role.Id, $Role.DisplayName)
}   

# Create arrays of service principals, role assignments, and groups
Write-Host "Fetching details of service principals, groups, and role assignments..."
[array]$ServicePrincipals = Get-MgServicePrincipal -All -PageSize 500
[array]$RoleAssignments = Get-MgRoleManagementDirectoryRoleAssignment -All -PageSize 500
[array]$GroupArray = Get-MgGroup -All -Property Id, displayName -PageSize 500 

# Create the output report
$Report = [System.Collections.Generic.List[Object]]::new()

# Check each role assignment to see if it's for a highly privileged role
Write-Host "Starting to check service principals for highly privileged roles..."
ForEach ($Assignment in $RoleAssignments) {
    If ($HighlyPrivilegedRoles.ContainsKey($Assignment.RoleDefinitionId)) {
        $ServicePrincipal = $ServicePrincipals | Where-Object {$_.Id -eq $Assignment.PrincipalId}
        If ($ServicePrincipal) {
            $DataLine = [PSCustomObject][Ordered]@{
                ServicePrincipalId   = $ServicePrincipal.Id
                ServicePrincipalName = $ServicePrincipal.DisplayName
                RoleDefinitionId     = $Assignment.RoleDefinitionId
                RoleDefinitionName   = $HighlyPrivilegedRoles[$Assignment.RoleDefinitionId]
                ServicePrincipalType = $ServicePrincipal.ServicePrincipalType
            }
            $Report.Add($DataLine)
        }
        # Check if the assignment is for a group
        $PrivilegedGroup = $GroupArray | Where-Object {$_.Id -contains $Assignment.PrincipalId}
        If ($PrivilegedGroup) {
            # Get the membership of the group - it could be transitive, so that's what we use
            $PrivilegedGroupMembers = Get-MgGroupTransitiveMember -GroupId $PrivilegedGroup.Id -All

            $DataLine = [PSCustomObject][Ordered]@{
                ServicePrincipalId   = $PrivilegedGroup.Id
                ServicePrincipalName = $PrivilegedGroup.DisplayName
                RoleDefinitionId     = $Assignment.RoleDefinitionId
                RoleDefinitionName   = $HighlyPrivilegedRoles[$Assignment.RoleDefinitionId]
                ServicePrincipalType = "Group"
                Members              = $PrivilegedGroupMembers.additionalProperties.displayName -join ", "
            }
            $Report.Add($DataLine)
        }   
    }
}

$Report | Format-table ServicePrincipalId, ServicePrincipalName, RoleDefinitionName, ServicePrincipalType, Members -AutoSize