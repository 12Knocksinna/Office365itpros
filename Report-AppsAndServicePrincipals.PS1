# Report-AppsAndServicePrincipals.PS1

# Use app-only mode to connect to Microsoft Graph or make sure that the signed in account holds one of the roles
# mentioned in https://learn.microsoft.com/en-us/graph/api/serviceprincipal-list-approleassignments

Connect-MgGraph Application.Read.All, Directory.Read.All, AuditLog.Read.All

Write-Output "Fetching details of app roles (permissions)"
$GraphApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
# Populate hash table with Graph permissions
$GraphRoles = @{}
ForEach ($Role in $GraphApp.AppRoles) { $GraphRoles.Add([string]$Role.Id, [string]$Role.Value) }
# Populate hash table with Exchange Online permissions
$ExoPermissions = @{}
$ExoApp = Get-MgServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'"
ForEach ($Role in $ExoApp.AppRoles) { $ExoPermissions.Add([string]$Role.Id, [string]$Role.Value) }
$O365Permissions = @{}
$O365API = Get-MgServicePrincipal -Filter "DisplayName eq 'Office 365 Management APIs'"
ForEach ($Role in $O365API.AppRoles) { $O365Permissions.Add([string]$Role.Id, [string]$Role.Value) }
$AzureADPermissions = @{}
$AzureAD = Get-MgServicePrincipal -Filter "DisplayName eq 'Windows Azure Active Directory'"
ForEach ($Role in $AzureAD.AppRoles) { $AzureADPermissions.Add([string]$Role.Id, [string]$Role.Value) }
$TeamsPermissions = @{}
$TeamsApp = Get-MgServicePrincipal -Filter "DisplayName eq 'Skype and Teams Tenant Admin API'"
ForEach ($Role in $TeamsApp.AppRoles) { $TeamsPermissions.Add([string]$Role.Id, [string]$Role.Value) }
$RightsManagementPermissions = @{}
$RightsManagementApp = Get-MgServicePrincipal -Filter "DisplayName eq 'Microsoft Rights Management Services'"
ForEach ($Role in $RightsManagementApp.AppRoles) { $RightsManagementPermissions.Add([string]$Role.Id, [string]$Role.Value) }

Write-Output "Fetching service principal sign-in activity records"
[array]$SPSignInLogs = Get-MgBetaReportServicePrincipalSignInActivity  -All -PageSize 999
$SpSignInHash = @{}
ForEach ($LogEntry in $SPSignInLogs) {
    $SpSignInHash.Add([string]$LogEntry.AppId, [string]$LogEntry.LastSignInActivity.LastSignInDateTime.DateTime)
}

# Not needed
# Delegated Graph permissions
[array]$GraphDelegatedPermissions = $GraphApp.Oauth2PermissionScopes | Where-Object { $_.IsEnabled -eq $true }
$GraphAppOauth2Permissions = @{}
ForEach ($DelegatedPermission in $GraphDelegatedPermissions) { $GraphAppOauth2Permissions.Add([string]$DelegatedPermission.Id, [string]$DelegatedPermission.Value) }
# Connect with elevated permissions to read app owners

Connect-MgGraph RoleAssignmentSchedule.ReadWrite.Directory, Application.Read.All, Directory.Read.All

Write-Host "Finding service principals..."
[Array]$ServicePrincipals = Get-MgServicePrincipal -All  `
    -Property Id, appId, displayName, Owners, appDisplayName, AppDescription, AppOwnerOrganizationId, AppRoles, AppRoleAssignments, `
    Oauth2PermissionGrants, keyCredentials, VerifiedPublisher, ServicePrincipalType, createdDateTime, KeyCredentials, passwordCredentials, signInAudience

If (!$ServicePrincipals) {
    Write-Output "No service principals found"
    break
} Else {
    $ServicePrincipals = $ServicePrincipals | Sort-Object AppDisplayName
    Write-Output ("{0} service principals found" -f $ServicePrincipals.Count)  
}

# Remove SharePoint helper apps https://learn.microsoft.com/en-us/answers/questions/1187017/sharepoint-online-client-extensibility-web-applica
$ServicePrincipals = $ServicePrincipals | Where-Object DisplayName -notLike "SharePoint Online Client Extensibility Web Application Principal*"

$AppReport = [System.Collections.Generic.List[Object]]::new()

[int]$i=0
ForEach ($SP in $ServicePrincipals) {
    $i++
    Write-Progress -Activity "Processing application $i of $($ServicePrincipals.Count): $($SP.DisplayName)" -PercentComplete (($i / $ServicePrincipals.Count) * 100)    
    $AppOwners = $null; $AppOwnersString = $null
    $isApp = $true
    $AppId = $SP.AppId
    Try {
        $App = Get-MgApplication -filter "appId eq '$AppId'" -Property Id, displayName, AppId, Notes,CreateDateTime, Owners, VerifiedPublisher,Tags,PublisherDomain,passwordCredentials,KeyCredentials, SignInAudience, Web -ErrorAction Stop 
        [array]$AppOwners = Get-MgApplicationOwner -ApplicationId $App.Id -All  
            
        If ($AppOwners) {
            $AppOwnersString = $AppOwners.additionalProperties.displayName -join "; "
        } Else {
            $AppOwnersString = $null
        }
        $AppIdentifierUris = $App.IdentifierUris -join ";"
        $AppRedirectUris = $App.Web.RedirectUris -join ";"
        $CreatedDateTime = Get-Date $App.CreatedDateTime -Format "dd-MMM-yyyy HH:mm"
     
        $AppName = $App.DisplayName
    } Catch {
        # No app found, so this service principal is probably an enterprise app or managed identity
    
        $CreatedDateTime = Get-Date $SP.additionalProperties['createdDateTime'] -Format "dd-MMM-yyyy HH:mm"
        $AppName = $SP.DisplayName

    }
    If (!$App) { $isApp = $false }

    # Get Application role (permission) assignments
    [array]$AppRoles = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $SP.Id
    If ($AppRoles) {
        [array]$Permissions = @()
        ForEach ($AppRole in $AppRoles) { 
            Switch ($AppRole.ResourceDisplayName) {
                "Microsoft Graph" { 
                [string]$Permission = $GraphRoles[$AppRole.AppRoleId] }
                "Office 365 Exchange Online" {
                [string]$Permission = $ExoPermissions[$AppRole.AppRoleId] }
                "Office 365 Management APIs" {
                [string]$Permission = $O365Permissions[$AppRole.AppRoleId] }
                "Windows Azure Active Directory" {
                [string]$Permission = $AzureADPermissions[$AppRole.AppRoleId] }
                "Skype and Teams Tenant Admin API" {
                [string]$Permission = $TeamsPermissions[$AppRole.AppRoleId] }
                "Microsoft Rights Management Services" {
                [string]$Permission = $RightsManagementPermissions[$AppRole.AppRoleId] }
            }
            $Permissions += $Permission
        }
        [string]$PermissionsOutput = $Permissions -join ", "
    }

    # Get delegated (OAuth2) permission grants
    $OAuth2PermissionsOutput = $null; [array]$OAuth2Permissions = $null; $OAuth2PermissionGrants = $null
    [array]$OAuth2PermissionGrants = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $SP.Id
    If ($OAuth2PermissionGrants) {
    
        ForEach ($PermissionGrant in $OAuth2PermissionGrants) {
                $OAuth2Permissions += $PermissionGrant.Scope
        }
        If ($OAuth2Permissions) { 
            [array]$ScopeTokens = @()
            $ScopeTokens = $OAuth2Permissions -split ' ' | Where-Object { $_ } 
        }
        [string]$OAuth2PermissionsOutput = $ScopeTokens -join ", "
    }

    If ($SP.ServicePrincipalType -ne "ManagedIdentity" -and $SP.AppOwnerOrganizationId) {
        If ($SP.AppOwnerOrganizationId -eq $TenantId) { #Resolve tenant name
            $AppTenantName = $TenantName 
        } Else {
            $LookUpId = $SP.AppOwnerOrganizationId.toString()
            $Uri = "https://graph.microsoft.com/beta/tenantRelationships/findTenantInformationByTenantId(tenantId='$LookUpId')"
            $ExternalTenantData = Invoke-MgGraphRequest -Uri $Uri -Method Get
            $AppTenantName = $ExternalTenantData.DisplayName 
        }
    }

    # Find service principal last sign-in activity
    $SPLastSignInDateTime = $SpSignInHash[$SP.AppId]
    If ($SPLastSignInDateTime) {
        $SPLastActivityDateTime = Get-Date $SPLastSignInDateTime -Format "dd-MMM-yyyy HH:mm"
    }

    $AppReportLine = [PSCustomObject]@{
        AppName                      = $AppName
        AppType                      = If ($isApp) { "App registration" } Else { 'Service principal' }
        AppDescription               = If ($isApp) { $App.Notes } Else { $SP.Description }
        AppId                        = $App.Id
        AppObjectId                  = $App.AppId
        AppOwners                    = $AppOwnersString
        AppCreatedDateTime           = $CreatedDateTime
        AppIdentifierUris            = $AppIdentifierUris
        AppRedirectUris              = $AppRedirectUris
        SignInAudience               = $SP.SignInAudience
        'Application permissions'    = $PermissionsOutput
        'Delegated permissions'      = $OAuth2PermissionsOutput
        'Owning tenant'              = $AppTenantName
        Publisher                    = $App.PublisherDomain
        VerifiedPublisher            = $App.VerifiedPublisher.DisplayName
        ServicePrincipalId           = $SP.Id
        ServicePrincipalType         = $SP.ServicePrincipalType
    }
    $AppReport.Add($AppReportLine)
}

$AppReport | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.