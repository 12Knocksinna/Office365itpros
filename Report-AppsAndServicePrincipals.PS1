# Report-AppsAndServicePrincipals.PS1
# A script (or runbook) to scan service principals and registered applications in an Entra ID (Azure AD) tenant
# and produce a report about their configuration, owners, permissions, and last sign-in activity.

# V1.0 21-Oct-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-AppsAndServicePrincipals.PS1

# Use app-only mode to connect to Microsoft Graph or make sure that the signed in account holds one of the roles
# mentioned in https://learn.microsoft.com/en-us/graph/api/serviceprincipal-list-approleassignments
# Requires the Microsoft.Graph module

[array]$RequiredScopes = "RoleAssignmentSchedule.Read.Directory", "Application.Read.All", "CrossTenantInformation.ReadBasic.All", "Mail.Send", "User.ReadBasic.All"
$Interactive = $false

# Determine if we're interactive or not
If ([Environment]::UserInteractive) { 
    # We're running interactively...
    Clear-Host
    Write-Host "Script running interactively... connecting to the Graph" -ForegroundColor Yellow
    Connect-MgGraph -NoWelcome -Scopes $RequiredScopes
    $Interactive = $true
    # Email address to use when sending email from interactive session
    $MsgFrom = (Get-MgContext).Account    
} Else { 
    # We're not, so likely in Azure Automation
    Write-Output "Executing the runbook to create the last activity report for user member accounts..." 
    Connect-MgGraph -Identity -NoWelcome
    # Email address to use when sending email from Azure Automation
    $MsgFrom = "no-reply@office365itpros.com"
}

# Check that we have the right permissions - in Azure Automation, we assume that the automation account has the right permissions
If ($Interactive) {
    [int]$RequiredScopesCount = $RequiredScopes.Count
    [string[]]$CurrentScopes = (Get-MgContext).Scopes
    [string[]]$RequiredScopes = $RequiredScopes

    $CheckScopes =[object[]][Linq.Enumerable]::Intersect($RequiredScopes,$CurrentScopes)
    If ($CheckScopes.Count -ne $RequiredScopesCount ) { 
        Write-Host ("To run this script, you need to connect to Microsoft Graph with the following scopes: {0}" -f $RequiredScopes) -ForegroundColor Red
        Disconnect-Graph
        Break
    }
}

# Set up to run
# Known Traitorware apps list - see https://huntresslabs.github.io/rogueapps/
[array]$TraitorWareApps = "em client", "perfectdata software", "newsletter software supermailer", "cloudsponge", "rclone"

# High-priority permissions that we should flag if found in apps
[array]$HighPriorityPermissions = "User.Read.All", "User.ReadWrite.All", "Mail.ReadWrite", "Mail.Read", "Files.Read.All", "Files.ReadWrite.All",`
    "Calendars.ReadWrite", "Mail.Send", "User.Export.All", "Directory.Read.All", "Exchange.ManageAsApp", "Directory.ReadWrite.All", "Sites.ReadWrite.All", "Domain.ReadWrite.All", `
    "Sites.Read.All", "Sites.FullControl.All", "Sites.Manage.All", "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All", "Group.ReadWrite.All", "Group.Read.All"

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
[array]$SPSignInLogs = Get-MgBetaReportServicePrincipalSignInActivity -All -PageSize 999
$SpSignInHash = @{}
ForEach ($LogEntry in $SPSignInLogs) {
    $SpSignInHash.Add([string]$LogEntry.AppId, [string]$LogEntry.LastSignInActivity.LastSignInDateTime.DateTime)
}

# Define email address for message with report attachment - make sure to change this to the appropriate address for your tenant
$DestinationEmailAddress = "Customer.Services@office365itpros.com"

Write-Host "Finding user details..."
[array]$Users = Get-MgUser -Filter "usertype eq 'Member' and AccountEnabled eq true" -All -Property Id, displayName, userPrincipalName -PageSize 999 
# Build hash tables for user display names and UPNs to use for matching against app names
$UserHash = @{}
$UPNHash = @{}
# Sometimes have multiple users with same display name, so get unique list
[array]$UserDisplayNames = $Users | Sort-Object displayName -Unique
ForEach ($User in $Users) {
    $UPNHash.Add([string]$User.userPrincipalName.ToLower(), [string]$User.DisplayName.toLower())
}
# Also trim any (xxx) suffixes from display names and add lowercase versions for matching
ForEach ($User in $UserDisplayNames) {
    $originalDisplayName = [string]$User.displayName.ToLower()
    $processedDisplayName = [string]$User.displayName.Split("(")[0].Trim().ToLower()
    $userUPN = [string]$User.userPrincipalName.ToLower()

    # Add the original display name as a key
    If (!$UserHash.ContainsKey($originalDisplayName)) {
       $UserHash.Add($originalDisplayName, $userUPN)
    }

    # Only add the processed display name if it's different from the original
    # and if the processed name is not already present
    If ($originalDisplayName -ne $processedDisplayName -and !$UserHash.ContainsKey($processedDisplayName)) {
       $UserHash.Add($processedDisplayName, $userUPN)
    }
}

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
    If ($Interactive) {
        Write-Progress -Activity "Processing application $i of $($ServicePrincipals.Count): $($SP.DisplayName)" -PercentComplete (($i / $ServicePrincipals.Count) * 100)
    }

    # Find service principal last sign-in activity
    $DaysSinceLastSignIn = $null; $SPLastActivityDateTime = $null
    $SPLastSignInDateTime = $SpSignInHash[$SP.AppId]

    # Verbose logging for troubleshooting sign-in date issue
    if ($Interactive) {
        Write-Verbose "Processing: $($SP.DisplayName) (AppId: $($SP.AppId))"
        Write-Verbose "  Hash lookup for AppId '$($SP.AppId)': '$SPLastSignInDateTime'"
        if ($SPLastSignInDateTime) {
            Write-Verbose "  Raw sign-in date from hash: $SPLastSignInDateTime"
            Write-Verbose "  Converted sign-in date: $(Get-Date $SPLastSignInDateTime -Format 'dd-MMM-yyyy HH:mm:ss')"
            $currentDate = Get-Date
            Write-Verbose "  Current date: $($currentDate.ToString('dd-MMM-yyyy HH:mm:ss'))"
            $timeSpan = New-TimeSpan -Start (Get-Date $SPLastSignInDateTime) -End $currentDate
            Write-Verbose "  Timespan: $timeSpan (Days: $($timeSpan.Days))"
        } else {
            Write-Verbose "  No sign-in data found in hash"
        }
    }

    $AppOwners = $null; $AppOwnersString = $null; $AppRedirectUris = $null; $AppIdentifierUris = $null; $App = $null
    $isApp = $true
    $AppId = $SP.AppId
    Try {
        # Check if an app registration exists for this service principal
        $App = Get-MgApplication -filter "appId eq '$AppId'" `
         -Property Id, displayName, AppId, Notes,CreateDateTime, Owners, VerifiedPublisher, Tags, AppRoles, PublisherDomain, passwordCredentials, KeyCredentials, SignInAudience, Web -ErrorAction Stop 
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
    
    # Checks for potentially suspicious apps - idea from https://www.bleepingcomputer.com/news/security/find-hidden-malicious-oauth-apps-in-microsoft-365-using-cazadora/
    # Check for TraitorWare apps
    If ($TraitorWareApps -contains $AppName.ToLower()) {
        $TraitorWareAppWarning = "⚠️ (Potential TraitorWare App)"
    }  else {
        $TraitorWareAppWarning = $null
    }
    # Flag app names composed entirely of non-alphanumeric characters
    If ($AppName -and ($AppName -notmatch '[A-Za-z0-9]')) {
        $NonAlnumNameWarning = "⚠️ App name contains no alphanumeric characters"
    } Else {
        $NonAlnumNameWarning = $null
    }
    # Check if app has "test" in the name
    If ($AppName -and ($AppName.ToLower() -like "*test*")) {
        $TestAppWarning = "⚠️ App name contains 'test'"
    } Else {
        $TestAppWarning = $null
    }
    # Look for ProofPoint MACT campaign 1445 https://www.proofpoint.com/us/blog/cloud-security/revisiting-mact-malicious-applications-credible-cloud-tenants
    # “http://localhost:7823/access/” in the Redirect URIs
    if ($AppRedirectUris -like "*http://localhost:7823/access/*") {
        $MACT1445Warning = "⚠️ App has ProofPoint MACT 1445 redirect URI"
    } Else {
        $MACT1445Warning = $null
    }

    # Check for apps named after account display names or user principal names
    If (($UserHash[$Appname]) -or ($UPNHash[$Appname])) {
        $UserMatchWarning = "⚠️ App name matches user"
    } ElseIf ($UserHash[$Appname.ToLower()]) {
        $UserMatchWarning = "⚠️ App name matches user"
    } Else {
        $UserMatchWarning = $null
    }

    # Get Application role (permission) assignments
    $PermissionsOutput = $null; $Permissions = $null
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

    [array]$HighPermissionsFound = @()
    # Check the permissions against high-priority list
    ForEach ($Permission in $Permissions) {
        If ($Permission -in $HighPriorityPermissions) {
            $HighPermissionsFound += $Permission
        }
    }
    If ($HighPermissionsFound) {
        $HighPermissionsFoundOutput = "⚠️ High-priority permissions: " + ($HighPermissionsFound -join ", ")
    } Else {
        $HighPermissionsFoundOutput = $null
    }

    # Get delegated (OAuth2) permission grants
    [array]$OAuth2PermissionsOutput = $null;  [array]$OAuth2PermissionGrants = $null
    [array]$OAuth2PermissionGrants = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $SP.Id
    If ($OAuth2PermissionGrants) {
        [array]$OAuth2PermissionsOutput = $null; [string]$OAuth2PermissionsFormatted = $Null; [array]$OAuth2Permissions = $null;
        ForEach ($PermissionGrant in $OAuth2PermissionGrants) {
            $OAuth2Permissions = $PermissionGrant.Scope.trim()    
            [array]$ScopeTokens = @()
            $ScopeTokens = $OAuth2Permissions -split ' ' | Where-Object { $_ }             
            [string]$OAuth2PermissionsFormatted = $ScopeTokens -join ", "
            If ($PermissionGrant.ConsentType -eq 'AllPrincipals') {
                $OAuth2PermissionsFormatted = $OAuth2PermissionsFormatted + " (Admin)"
            } Else {
                $OAuth2PermissionsFormatted = $OAuth2PermissionsFormatted + " (User)"
            }
            $OAuth2PermissionsOutput += $OAuth2PermissionsFormatted
        }
    }

    If ($SP.ServicePrincipalType -ne "ManagedIdentity" -and $SP.AppOwnerOrganizationId) {
        If ($SP.AppOwnerOrganizationId -eq $TenantId) { #Resolve tenant name
            $AppTenantName = $TenantName
        } Else {
            $LookUpId = $SP.AppOwnerOrganizationId.toString()
            Try {
                $ExternalTenantData = Find-MgTenantRelationshipTenantInformationByTenantId -TenantId $LookUpId -ErrorAction Stop
                $AppTenantName = $ExternalTenantData.DisplayName
            } Catch {
                $AppTenantName = "Unknown"
                Write-Warning "Could not retrieve tenant details for $LookUpId : $($_.Exception.Message)"
            }
        }
    }

    Switch ($SP.SignInAudience) {
        "AzureADMyOrg" { 
            $SignInAudience = "Only this tenant" 
        }
        "AzureADMultipleOrgs" { 
            $SignInAudience = "Accounts from any Entra ID tenant" 
        }
        "AzureADandPersonalMicrosoftAccount" { 
            $SignInAudience = "Accounts from any Entra ID tenant and personal Microsoft accounts" 
        }
        "PersonalMicrosoftAccount" { 
            $SignInAudience = "Personal Microsoft accounts only" 
        }
        Default { $SignInAudience = $SP.SignInAudience }
    }

    # Find service principal last sign-in activity
    $DaysSinceLastSignIn = $null; $SPLastActivityDateTime = $null
    $SPLastSignInDateTime = $SpSignInHash[$SP.AppId]
    If ($SPLastSignInDateTime) {
        $SPLastActivityDateTime = Get-Date $SPLastSignInDateTime -Format "dd-MMM-yyyy HH:mm"
        $timeSpan = New-TimeSpan -Start (Get-Date $SPLastSignInDateTime) -End (Get-Date)
        $DaysSinceLastSignIn = $timeSpan.Days

        # Additional logging for debugging
        if ($Interactive) {
            Write-Verbose "  Final calculation - Sign-in date: $(Get-Date $SPLastSignInDateTime -Format 'dd-MMM-yyyy HH:mm:ss'), Days since: $DaysSinceLastSignIn"
        }
    } Else {
        $SPLastActivityDateTime = "Never"
        if ($Interactive) {
            Write-Verbose "  No sign-in data found for AppId '$($SP.AppId)', setting to 'Never'"
        }
    }   

    # Check password credentials (app secrets)
    $PasswordOutput = $null; $PasswordReportOutput = $null; [int]$ValidAppPwd = 0
    If ($App.PasswordCredentials) {
        [array]$PasswordOutput = @()
        ForEach ($AppPwd in $App.PasswordCredentials) {
            If ($AppPwd.EndDateTime -gt (Get-Date).AddDays(30)) {
                $ValidAppPwd++
                $PasswordOutput += ("👍 App Password {0} valid. End date {1}" -f $AppPwd.Hint, (Get-Date $AppPwd.EndDateTime -Format "dd-MMM-yyyy HH:mm"))   
            } Else {
                $PasswordOutput += ("⛔ App Password {0} EXPIRED or expiring SOON! End date {1}" -f $AppPwd.Hint, (Get-Date $AppPwd.EndDateTime -Format "dd-MMM-yyyy HH:mm"))       
            }
            $PasswordReportOutput = $PasswordOutput -join "; "
        }
    }

      # Check X.509 cert credentials
    $CertOutput = $null; $CertReportOutput = $null; [int]$ValidAppCert = 0
      If ($App.KeyCredentials) {
        [array]$CertOutput = @()
        ForEach ($AppCert in $App.KeyCredentials) {
            # Could add code to report on certs if needed
            If ($AppCert.EndDateTime -lt (Get-Date).AddDays(30)) {
                $ValidAppCert++
                $CertOutput += ("⛔ App Cert {0} EXPIRED or expiring SOON! End date {1}" -f $AppCert.DisplayName, (Get-Date $AppCert.EndDateTime -Format "dd-MMM-yyyy HH:mm"))
            }  Else {  
                $CertOutput += ("👍 App Cert {0} valid. End date {1}" -f $AppCert.DisplayName, (Get-Date $AppCert.EndDateTime -Format "dd-MMM-yyyy HH:mm"))
            }   
            $CertReportOutput = $CertOutput -join "; "
        }
    }       

    If ($SP.AppRoleAssignmentRequired) {
        $AccessAllowedToApp = "Assigned users only"
    } Else {
        $AccessAllowedToApp = "All users"
    }

     If ($SP.Tags -contains "HideApp") { 
        $AppUserVisibility="Hidden" 
    } Else { 
        $AppUserVisibility="Visible"
    }

    # Check when service principal was created to highlight any created in the last 10 days
    $SPCreatedDateTimeOutput = $null
    $SPCreatedDate = Get-Date $SP.additionalProperties['createdDateTime']
    If ($SPCreatedDate -gt (Get-Date).AddDays(-10)) {
        $SPCreatedDateTimeOutput = " ⚠️ (Newly created service principal) " + (Get-Date $SPCreatedDate -Format "dd-MMM-yyyy HH:mm") 
    } Else {
        $SPCreatedDateTimeOutput = Get-Date $SPCreatedDate -Format "dd-MMM-yyyy HH:mm"
    }

    $AppReportLine = [PSCustomObject]@{
        AppName                         = $AppName
        AppType                         = If ($isApp) { "App registration" } Else { 'Service principal' }
        AppDescription                  = If ($isApp) { $App.Notes } Else { $SP.Description }
        AppOwners                       = $AppOwnersString
        AppCreatedDateTime              = $CreatedDateTime
        'App Access'                    = $AccessAllowedToApp
        'App Visibility'                = $AppUserVisibility
        'Service Principal last used'   = $SPLastActivityDateTime
        'Days since last sign-in'       = If ($DaysSinceLastSignIn -ne $null) { $DaysSinceLastSignIn } Else { "Never" }
        'App Passwords'                 = If ($PasswordReportOutput) { $PasswordReportOutput } Else { "No app passwords" }
        'Valid App Passwords'           = $ValidAppPwd
        'Invalid App Passwords'         = $App.PasswordCredentials.Count - $ValidAppPwd
        'App Certificates'              = If ($CertReportOutput) { $CertReportOutput } Else { "No app certificates" }
        'Valid App Certificates'        = $ValidAppCert
        'Invalid App Certificates'      = $App.KeyCredentials.Count - $ValidAppCert
        AppIdentifierUris               = $AppIdentifierUris
        AppRedirectUris                 = $AppRedirectUris
        'Sign in audience'              = $SignInAudience
        'Application permissions'       = $PermissionsOutput
        'High-priority permissions'     = $HighPermissionsFoundOutput
        'Delegated permissions'         = $OAuth2PermissionsOutput -join "`n"
        'Owning tenant'                 = $AppTenantName
        'App Publisher Domain'          = $App.PublisherDomain
        'Verified Publisher'            = $SP.VerifiedPublisher.DisplayName
        AppId                           = $App.Id
        AppObjectId                     = $App.AppId
        'Service principal id'          = $SP.Id
        'Service Principal Type'        = $SP.ServicePrincipalType
        'TraitorWare App Warning'       = $TraitorWareAppWarning
        'Non-Alphanumeric Name Warning' = $NonAlnumNameWarning
        'Test App Name Warning'         = $TestAppWarning
        'MACT 1445 Warning'             = $MACT1445Warning
        'User Match warning'            = $UserMatchWarning
        'Service Principal Created Date' =  $SPCreatedDateTimeOutput 
    }
    $AppReport.Add($AppReportLine)
}

$AppReport | Out-GridView -Title 'Service Principal Report'

# Get some statistics
$TotalSPsNotUsed = $AppReport | Where-Object { $_.'Service Principal last used' -eq "Never" } | Measure-Object | Select-Object -ExpandProperty Count
$PercentSPsNotUsed = ($TotalSPsNotUsed / $AppReport.Count).ToString('P')
$TotalApps = $AppReport | Where-Object { $_.AppType -eq "App registration" } | Measure-Object | Select-Object -ExpandProperty Count
$TotalManagedIdentities = $AppReport | Where-Object { $_.'Service Principal Type' -eq "ManagedIdentity" } | Measure-Object | Select-Object -ExpandProperty Count
$TotalLegacySPs = $AppReport | Where-Object { $_.'Service Principal Type' -eq "Legacy" } | Measure-Object | Select-Object -ExpandProperty Count
$TotalAppsNoValidPasswords = $AppReport | Where-Object { $_.AppType -eq 'App registration' -and $_.'Valid App Passwords' -eq 0 -and $_.'App Passwords' -ne "No app passwords" } | Measure-Object | Select-Object -ExpandProperty Count
$TotalAppsNoValidCertificates = $AppReport | Where-Object { $_.AppType -eq 'App registration' -and $_.'Valid App Certificates' -eq 0 -and $_.'App Certificates' -ne "No app certificates" } | Measure-Object | Select-Object -ExpandProperty Count
$TotalAppsNoValidPasswordsOrCertificates = $AppReport | Where-Object { $_.AppType -eq 'App registration' -and $_.'Valid App Passwords' -eq 0 -and $_.'App Certificates' -ne "No app certificates" -and $_.'Valid App Certificates' -eq 0 } | Measure-Object | Select-Object -ExpandProperty Count
$TotalAppsWithHighPriorityPermissions = $AppReport | Where-Object { $_.AppType -eq 'App registration' -and $_.'High-priority permissions' -ne $null } | Measure-Object | Select-Object -ExpandProperty Count
$TotalSPsWithHighPriorityPermissions = $AppReport | Where-Object { $_.AppType -eq 'Service Principal' -and $_.'High-priority permissions' -ne $null } | Measure-Object | Select-Object -ExpandProperty Count
$TotalAppsWithWarnings = $AppReport | Where-Object { $_.AppType -eq "App registration" -and $_.'Traitor Ware App Warning' -ne $null -or $_.'Non-Alphanumeric Name Warning' -ne $null -or $_.'Test App Name Warning' -ne $null -or $_.'MACT 1445 Warning' -ne $null } | Measure-Object | Select-Object -ExpandProperty Count
$TotalMicrosoftSPs = $AppReport | Where-Object { $_.AppType -eq 'Service Principal' -and $_.'Owning tenant' -like "*microsoft*" } | Measure-Object | Select-Object -ExpandProperty Count
$PercentMicrosoftSPs = ($TotalMicrosoftSPs / $AppReport.Count).ToString('P')    
$PercentAppswithHighPriorityPermissions = ($TotalAppsWithHighPriorityPermissions / $TotalApps).ToString('P')
$ServicePrincipalsSignedInLast2Years = $AppReport | Where-Object { $_.'Days since last sign-in' -ne "Never" -and $_.'Days since last sign-in' -le 730 } | Measure-Object | Select-Object -ExpandProperty Count
$PercentServicePrincipalsSignedInLast2Years = ($ServicePrincipalsSignedInLast2Years / $AppReport.Count).ToString('P')
[array]$NewServicePrincipals = $AppReport | Where-Object { $_.'Service Principal Created Date' -like "*Newly created service principal*" } 


Write-Output ""
Write-Output "Service Principals Report Summary"
Write-Output "---------------------------------"
Write-Output ("Total service principals:                                            {0}" -f $AppReport.Count)
Write-Output ("Total Microsoft-owned service principals:                            {0} ({1})" -f $TotalMicrosoftSPs, $PercentMicrosoftSPs)
Write-Output ("Total non-Microsoft service principals:                              {0}" -f ($AppReport.Count - $TotalMicrosoftSPs))
Write-Output ("Total app registrations:                                             {0}" -f $TotalApps) 
Write-Output ("Total managed identities:                                            {0}" -f $TotalManagedIdentities)
Write-Output ("Total legacy service principals:                                     {0}" -f $TotalLegacySPs)
Write-Output ("Total service principals signed in within last 2 years:              {0} ({1})" -f $ServicePrincipalsSignedInLast2Years, $PercentServicePrincipalsSignedInLast2Years)
Write-Output ("Total service principals never used:                                 {0} ({1})" -f $TotalSPsNotUsed, $PercentSPsNotUsed)
Write-Output ("Total app registrations with no valid app passwords:                 {0}" -f $TotalAppsNoValidPasswords)
Write-Output ("Total app registrations with no valid app certificates:              {0}" -f $TotalAppsNoValidCertificates)
Write-Output ("Total app registrations with no valid app passwords or certificates: {0} " -f $TotalAppsNoValidPasswordsOrCertificates)
Write-Output ("Total app registrations with high-priority permissions:              {0} ({1})" -f $TotalAppsWithHighPriorityPermissions, $PercentAppswithHighPriorityPermissions)
Write-Output ("Total service principals with high-priority permissions:             {0}" -f $TotalSPsWithHighPriorityPermissions)
Write-Output ("Total app registrations with warnings:                               {0}" -f $TotalAppsWithWarnings)
Write-Output ("Total newly created service principals (last 10 days):               {0}" -f $NewServicePrincipals.Count)
Write-Output ("Newly created service principals:                                    {0}" -f ($NewServicePrincipals.AppName -join ", "))
Write-Output ""

Write-Output "Service principals by owning tenant"
Write-Output "-----------------------------------"
$AppReport | Group-Object -Property 'Owning tenant' | Sort-Object Count -Descending | ForEach-Object {
    Write-Output ("{0,-40} {1,5}" -f $_.Name, $_.Count)
}
Write-Output ""

If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $true
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\TenantAppReport.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    } 
    $AppReport | Export-Excel -Path $ExcelOutputFile -WorksheetName "Service Principals" -Title ("Service Principals Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "ServicePrincipals" 
    $AttachmentFile = $ExcelOutputFile
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\TenantAppReport.CSV"
    $AppReport | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
    $AttachmentFile = $CSVOutputFile
}

If ($ExcelGenerated) {
    Write-Output ("Excel worksheet output written to {0}" -f $ExcelOutputFile)
} Else {
    Write-Output ("CSV output file written to {0}" -f $CSVOutputFile)
} 

# Send the spreadsheet as an email attachment
$EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($AttachmentFile))
$MsgAttachments = @(
    @{
        '@odata.type' = '#microsoft.graph.fileAttachment'
        Name = (Split-Path $AttachmentFile -Leaf)
        ContentBytes = $EncodedAttachmentFile
        ContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
)

# Build the array of a single TO recipient detailed in a hash table - change this to the appropriate recipient for your tenant
$ToRecipient = @{}
$ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
[array]$MsgTo = $ToRecipient
# Define the message subject
$MsgSubject = "Important: Service Principals Analysis Report"
# Create the HTML content
$HtmlMsg = "</body></html><p>The output file for the <b>Service Principals Analysis Report</b> is attached to this message. Please review the information at your convenience</p>"

# Add the summary content
$HtmlMsg += "<h2>Service Principals Analysis Summary</h2>"
$HtmlMsg += "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;'>"
$HtmlMsg += "<tr><th align='left'>Metric</th><th align='left'>Value</th></tr>"
$HtmlMsg += ("<tr><td>Total service principals:</td><td>{0}</td></tr>" -f $AppReport.Count)
$HtmlMsg += ("<tr><td>Total Microsoft-owned service principals:</td><td>{0} ({1})</td></tr>" -f $TotalMicrosoftSPs, $PercentMicrosoftSPs)
$HtmlMsg += ("<tr><td>Total non-Microsoft service principals:</td><td>{0}</td></tr>" -f ($AppReport.Count - $TotalMicrosoftSPs))
$HtmlMsg += ("<tr><td>Total app registrations:</td><td>{0}</td></tr>" -f $TotalApps)            
$HtmlMsg += ("<tr><td>Total managed identities:</td><td>{0}</td></tr>" -f $TotalManagedIdentities)
$HtmlMsg += ("<tr><td>Total legacy service principals:</td><td>{0}</td></tr>" -f $TotalLegacySPs)
$HtmlMsg += ("<tr><td>Total service principals signed in within last 2 years:</td><td>{0} ({1})</td></tr>" -f $ServicePrincipalsSignedInLast2Years, $PercentServicePrincipalsSignedInLast2Years)
$HtmlMsg += ("<tr><td>Total service principals never used:</td><td>{0} ({1})</td></tr>" -f $TotalSPsNotUsed, $PercentSPsNotUsed)
$HtmlMsg += ("<tr><td>Total app registrations with no valid app passwords:</td><td>{0}</td></tr>" -f $TotalAppsNoValidPasswords)
$HtmlMsg += ("<tr><td>Total app registrations with no valid app certificates:</td><td>{0}</td></tr>" -f $TotalAppsNoValidCertificates)
$HtmlMsg += ("<tr><td>Total app registrations with no valid app passwords or certificates:</td><td>{0}</td></tr>" -f $TotalAppsNoValidPasswordsOrCertificates)
$HtmlMsg += ("<tr><td>Total app registrations with high-priority permissions:</td><td>{0} ({1})</td></tr>" -f $TotalAppsWithHighPriorityPermissions, $PercentAppswithHighPriorityPermissions)
$HtmlMsg += ("<tr><td>Total service principals with high-priority permissions:</td><td>{0}</td></tr>" -f $TotalSPsWithHighPriorityPermissions)
$HtmlMsg += ("<tr><td>Total app registrations with warnings:</td><td>{0}</td></tr>" -f $TotalAppsWithWarnings)
$HtmlMsg += ("<tr><td>Total newly created service principals (last 10 days):</td><td>{0}</td></tr>" -f $NewServicePrincipals.Count)
$HtmlMsg += ("<tr><td>Newly created service principals:</td><td>{0}</td></tr>" -f ($NewServicePrincipals.AppName -join ", "))
$HtmlMsg += "</table></p>"  

# Construct the message body 	
$MsgBody = @{}
$MsgBody.Add('Content', "$($HtmlMsg)")
$MsgBody.Add('ContentType','html')
# Build the parameters to submit the message
$Message = @{}
$Message.Add('subject', $MsgSubject)
$Message.Add('toRecipients', $MsgTo)
$Message.Add('body', $MsgBody)
$Message.Add("attachments", $MsgAttachments)

$EmailParameters = @{}
$EmailParameters.Add('message', $Message)
$EmailParameters.Add('saveToSentItems', $true)
$EmailParameters.Add('isDeliveryReceiptRequested', $true)

# Send the message
Try {
    Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters -ErrorAction Stop
    Write-Output ("Service Principals analysis report emailed to {0}" -f $ToRecipient.emailAddress.address)
    Write-Output "All done!"
} Catch {
    Write-Output "Unable to send email"
    Write-Output $_.Exception.Message
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
