# Report-ManagedIdentityPermissions.PS1
# Report the Graph and other permissions assigned to managed identities
# V1.0 30-Jan-2024
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ManagedIdentityPermissions.PS1

Connect-MgGraph -NoWelcome -Scopes Directory.Read.All

$GraphApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
$AppRoleTable = @{}
ForEach ($AppRole in $GraphApp.AppRoles) {
    $AppRoleTable.Add($AppRole.Id,$AppRole.Value)
}
# Add the roles for Exchange Online. The one we're really interested in is Exchange.ManageAsApp
# Graph roles exist in many cases, so we try to add the role to the table and don't care if a failure occurs
# because the role alreadt exists
$ExoApp = Get-MgServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'"
ForEach ($AppRole in $ExoApp.AppRoles) {
    try {
        $RoleName = ("{0} [Exchange Online]" -f $AppRole.Value)
        $AppRoleTable.Add($AppRole.Id,$RoleName) 
    }
    catch {
        Out-Null
    }
}
# Do the same for Teams. In this case, we're interested in application_access
$TeamsApp = Get-MgServicePrincipal -Filter "AppId eq '48ac35b8-9aa8-4d74-927d-1f4a14a0b239'"  
ForEach ($AppRole in $TeamsApp.AppRoles) {
    try {
        $RoleName = ("{0} [Teams]" -f $AppRole.Value)
        $AppRoleTable.Add($AppRole.Id,$RoleName) 
    }
    catch {
        Out-Null
    }
}
# And to be complete, do the same for SharePoint
$SPOApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0ff1-ce00-000000000000'"
ForEach ($AppRole in $SPOApp.AppRoles) {
    try {
        $RoleName = ("{0} [SharePoint Online]" -f $AppRole.Value)
        $AppRoleTable.Add($AppRole.Id,$RoleName) 
    }
    catch {
        Out-Null
    }
}

# Find the set of managed identities in the tenant
[array]$ManagedIdentities = Get-MgServicePrincipal -Filter "servicePrincipalType eq 'ManagedIdentity'" | Sort-Object DisplayName
$Report = [System.Collections.Generic.List[Object]]::new()

# Process each managed identity and find the set of application roles assigned to it
ForEach ($ManagedIdentity in $ManagedIdentities) {
    [array]$RoleNames = $null
    [array]$AssignedRoles = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentity.Id | `
        Select-Object -ExpandProperty AppRoleId
    ForEach ($Role in $AssignedRoles) {
        $GraphRoleName = $AppRoleTable[$Role]
        $RoleNames += $GraphRoleName
    }
    $RoleNames = $RoleNames | Sort-Object 
   
    $Reportline = [PsCustomObject]@{
        ManagedIdentity = $ManagedIdentity.DisplayName
        Id              = $ManagedIdentity.id
        Created         = (Get-Date $ManagedIdentity.additionalProperties.createdDateTime -format 'dd-MMM-yyyy HH:mm')
        Roles           = ($RoleNames -join ", ")
    }
    $Report.Add($ReportLine)
}

$Report | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.

 