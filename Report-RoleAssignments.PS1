# Report-RoleAssignments.PS1
# Example script to show how to report Entra ID role assignments (PIM or direct) using the 
# Microsoft Graph PowerShell SDK

Connect-MgGraph -Scopes 'RoleAssignmentSchedule.Read.Directory','RoleEligibilitySchedule.Read.Directory' -NoWelcome

[array]$AdminUnits = Get-MgDirectoryAdministrativeUnit -Property Id, DisplayName | Sort-Object DisplayName
$AdminUnitsHash = @{}
ForEach ($AU in $AdminUnits) {
    $AdminUnitsHash.Add($AU.Id, $AU.DisplayName)
}
$Report = [System.Collections.Generic.List[Object]]::new()

Write-Host "Checking for PIM active assignments..."
# Get active assignments
[array]$ActiveAssignments = Get-MgBetaRoleManagementDirectoryRoleAssignmentSchedule `
    -ExpandProperty RoleDefinition, Principal, DirectoryScope -All 

Write-Host ("Found {0} PIM active assignments" -f $ActiveAssignments.Count)

ForEach ($Assignment in $ActiveAssignments) {
    $AdminUnitId = $null; $AdminUnitName = $null; $ServicePrincipal = $null; $AppId = $null
    # Check scoping for assignment
    If ($Assignment.DirectoryScopeId -ne "/") {
        $AdminUnitId = $Assignment.DirectoryScopeId.SubString(21,$Assignment.DirectoryScopeId.Length-21)
        $AdminUnitName = $AdminUnitsHash[$AdminUnitId] 
    } Else {
        $AdminUnitName = "Complete directory"
    }
    $RoleName = $Assignment.RoleDefinition.DisplayName

    Switch ($Assignment.Principal.AdditionalProperties."@odata.type") {
        "#microsoft.graph.user" {     

            $ReportLine = [PSCustomObject][Ordered]@{
                RoleName            = $RoleName
                UserPrincipalName   = $Assignment.Principal.AdditionalProperties.userPrincipalName
                Created             = $Assignment.CreatedDateTime
                DirectoryScope      = $adminUnitName
                OnPremisesUser      = $Assignment.Principal.AdditionalProperties.onPremisesSyncEnabled
                AssignmentType      = "Active"
                AssignmentVia       = "User"
                MemberType          = $Assignment.MemberType
            }
            $Report.Add($ReportLine)
        }
    # Process group assignments
        "#microsoft.graph.group" {
            
            [array]$Members = (Get-MgGroupMember -GroupId $Assignment.Principal.Id)
            If ($Members) { 
                ForEach ($Member in $Members) {
                    $ReportLine = [PSCustomObject][Ordered]@{
                        RoleName            = $RoleName
                        UserPrincipalName   = $Member.AdditionalProperties.userPrincipalName
                        Created             = $Assignment.CreatedDateTime
                        DirectoryScope      = $AdminUnitName
                        OnPremisesUser      = $Assignment.Principal.AdditionalProperties.onPremisesSyncEnabled
                        AssignmentType      = "Active"
                        AssignmentVia       = "Group"
                        MemberType          = $Assignment.MemberType
                    }
                    $Report.Add($ReportLine)
                }
            }
        }   
        "#microsoft.graph.servicePrincipal" {

            $AppId = $Assignment.Principal.AdditionalProperties.appId
            $ServicePrincipal = (Get-MgServicePrincipal -Filter "AppId eq '$AppId'").DisplayName
            $ReportLine = [PSCustomObject][Ordered]@{
                RoleName            = $RoleName
                UserPrincipalName   = $Assignment.Principal.AdditionalProperties.displayName
                Created             = $Assignment.CreatedDateTime
                DirectoryScope      = $AdminUnitName
                OnPremisesUser      = "N/A"
                AssignmentType      = "Active"
                AssignmentVia       = "Service Principal"
                MemberType          = $Assignment.MemberType
                SPName              = $ServicePrincipal
            }
            $Report.Add($ReportLine)
        }
    }

}    

Write-Host "Checking for PIM eligible assignments..."

# Get eligible assignments
[array]$EligibleAssignments = Get-MgBetaRoleManagementDirectoryRoleEligibilitySchedule `
    -ExpandProperty RoleDefinition,Principal,DirectoryScope -All 
Write-Host ("Found {0} PIM eligible assignments" -f $EligibleAssignments.Count)

ForEach ($Assignment in $EligibleAssignments) {
    $AdminUnitId = $null; $AdminUnitName = $null; $ServicePrincipal = $null; $AppId = $null
    # Check scoping for assignment
    If ($Assignment.DirectoryScopeId -ne "/") {
        $AdminUnitId = $Assignment.DirectoryScopeId.SubString(21,$Assignment.DirectoryScopeId.Length-21)
        $AdminUnitName = $AdminUnitsHash[$AdminUnitId] 
    } Else {
        $AdminUnitName = "Complete directory"
    }
    $RoleName = $Assignment.RoleDefinition.DisplayName

    Switch ($Assignment.Principal.AdditionalProperties."@odata.type") {
        "#microsoft.graph.user" {  
            $ReportLine = [PSCustomObject][Ordered]@{
                RoleName            = $RoleName
                UserPrincipalName   = $Assignment.Principal.AdditionalProperties.userPrincipalName
                Created             = $Assignment.CreatedDateTime
                DirectoryScope      = $adminUnitName
                OnPremisesUser      = $Assignment.Principal.AdditionalProperties.onPremisesSyncEnabled
                AssignmentType      = "Eligible"
                AssignmentVia       = "User"
                MemberType          = $Assignment.MemberType
            }
            $Report.Add($ReportLine)
        }
        # Process group assignments
        "#microsoft.graph.group" {
            
            [array]$Members = (Get-MgGroupMember -GroupId $Assignment.Principal.Id)
            If ($Members) { 
                ForEach ($Member in $Members) {
                    $ReportLine = [PSCustomObject][Ordered]@{
                        RoleName            = $RoleName
                        UserPrincipalName   = $Member.AdditionalProperties.userPrincipalName
                        Created             = $Assignment.CreatedDateTime
                        DirectoryScope      = $AdminUnitName
                        OnPremisesUser      = $Assignment.Principal.AdditionalProperties.onPremisesSyncEnabled
                        AssignmentType      = "Eligible"
                        AssignmentVia       = "Group"
                        MemberType          = $Assignment.MemberType
                    }
                    $Report.Add($ReportLine)
                }
            }
        }   
        "#microsoft.graph.servicePrincipal" {

            $AppId = $Assignment.Principal.AdditionalProperties.appId
            $ServicePrincipal = (Get-MgServicePrincipal -Filter "AppId eq '$AppId'").DisplayName
            $ReportLine = [PSCustomObject][Ordered]@{
                RoleName            = $RoleName
                UserPrincipalName   = $Assignment.Principal.AdditionalProperties.displayName
                Created             = $Assignment.CreatedDateTime
                DirectoryScope      = $AdminUnitName
                OnPremisesUser      = "N/A"
                AssignmentType      = "Eligible"
                AssignmentVia       = "Service Principal"
                MemberType          = $Assignment.MemberType
                SPName              = $ServicePrincipal
            }
            $Report.Add($ReportLine)
        }
    }
}

$PIMAssignments = $ActiveAssignments.count + $EligibleAssignments.count


    Get-MgDirectoryRole -All | ForEach-Object {
        $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $_.Id | Where-Object {
        $_.AdditionalProperties.userPrincipalName -ne $null -and $_.AdditionalProperties.userPrincipalName -notmatch ".onmicrosoft.com"
        }
        if ($roleMembers.Count -ne 0) {
        $roleName = $_.DisplayName
        $roleMembers | ForEach-Object { if ((Get-MgUser -UserId $_.Id -Property OnPremisesSyncEnabled).OnPremisesSyncEnabled -ne $null) {
        [array]$members += (Get-MgUser -UserId $_.Id).UserPrincipalName
        }}
        if ($members.Count -ne 0) { $members | Foreach-Object { Write-Output "$roleName,$_" }
        Remove-Variable members
        }
        }
        }


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.