# Report-ConditionalAccessPolicies.PS1
# A script to show how to report conditional access policy settings
# V1.0 16-Dec-2023

Connect-MgGraph -NoWelcome -Scopes User.Read.All, Policy.Read.All, Group.Read.All, Agreement.Read.All, Application.Read.All
Write-Host 'Finding conditional access policies'
[array]$CAPolicies = Get-MgIdentityConditionalAccessPolicy | Sort-Object displayName
If (!($CAPolicies)) {
    Write-Host "Unable to retrieve conditional access policies - exiting"; break
}

# Hash lookup table for service principals
$HashAppLookup = @{}
[array]$SPs = Get-MgServicePrincipal -All
ForEach ($SP in $SPs) {
    $HashAppLookup.Add($SP.AppId, $SP.displayName)
}

# Hash lookup table for admin roles
$HashAdminRoles = @{}
[array]$AdminRoles = Get-MgDirectoryRoleTemplate -All | Select-Object Id, DisplayName | Sort-Object Id
ForEach ($Role in $AdminRoles) {
    $HashAdminRoles.Add($Role.Id, $Role.DisplayName)
}

# Loop through the CA policies and extract the values for each before generating a report
$Report = [System.Collections.Generic.List[Object]]::new()
[int]$i = 0
ForEach ($Policy in $CAPolicies) {
    $i++
   
    Write-Host ("Processing conditional access polucy {0} ({1}/{2})" -f $Policy.displayName, $i, $CAPolicies.Count)
    [array]$Conditions = $Policy | Select-Object -ExpandProperty conditions
    [array]$SessionControls = $Policy | Select-Object -ExpandProperty sessionControls
    [array]$GrantControls = $Policy | Select-Object -ExpandProperty grantControls
    [array]$Applications = $Conditions | Select-Object -ExpandProperty Applications

    # Application filter
    [string]$ApplicationFilterMode = $Null; [string]$ApplicationFilterRule = $Null
    [array]$ApplicationFilter = $Applications | Select-Object -ExpandProperty ApplicationFilter
    $ApplicationFilterMode = $ApplicationFilter.Mode 
    $ApplicationFilterRule = $ApplicationFilter.Rule

    # Process app exclusions
    [string]$ExcludedAppOutput = $Null; [string]$IncludedAppOutput = $Null
    If ($Applications.excludeApplications) {
        [array]$ExcludeAppNames = $Null
        ForEach ($ExcludeApp in $Applications.excludeApplications) {
            $ExcludeAppName = $HashAppLookup[$ExcludeApp]
            $ExcludeAppNames += $ExcludeAppName
        }
        [string]$ExcludedAppOutput = ($ExcludeAppNames -join ", ")
    }

    # Process app inclusions
    If ($Applications.includeApplications) {
        [array]$IncludeAppNames = $Null;
        ForEach ($IncludeApp in $Applications.includeApplications) {
            $IncludeAppName = $HashAppLookup[$IncludeApp]
            $IncludeAppNames += $IncludeAppName
        }
        [string]$IncludedAppOutput = ($IncludeAppNames -join ", ")
    }

    [string]$IncludeAuthenticationContextClassReferences = $Applications.IncludeAuthenticationContextClassReferences
    [string]$IncludeUserActions = $Applications.IncludeUserActions

    [string]$ServicePrincipalRiskLevels = $Conditions.ServicePrincipalRiskLevels
    [string]$SignInRiskLevels = $Conditions.SignInRiskLevels
    [string]$UserRiskLevels = $Conditions.UserRiskLevels

    # Client app types
    [string]$ClientAppTypes = ($Conditions.ClientAppTypes -join ", ")
    
    # Client applications
    [array]$ClientApplications = $Conditions | Select-Object -ExpandProperty ClientApplications
    [array]$ExcludedClientApps = $Null; [string]$ExcludedClientAppNames = $Null
    [array]$IncludedClientApps = $Null; [string]$IncludedClientAppNames = $Null
    
    If ($ClientApplications.ExcludeServicePrincipals) {
        ForEach ($ClientApp in $ClientApplications.ExcludeServicePrincipals) {
            $ClientAppName = $HashLookUp[$ClientApp]
            $ExcludedClientApps += $ClientAppName
        }
        $ExcludedClientAppNames = $ExcludedClientApps -join ", "
    }
    If ($ClientApplications.IncludeServicePrincipals) {
        ForEach ($ClientApp in $ClientApplications.IncludeServicePrincipals) {
            $ClientAppName = $HashLookUp[$ClientApp]
            $IncludedClientApps += $ClientAppName
        }
        $IncludedClientAppNames = $IncludedClientApps -join ", "
    }

    [array]$Devices = $Conditions | Select-Object -ExpandProperty Devices
    [string]$DeviceFilterMode = $Devices.DeviceFilter.Mode
    [string]$DeviceFilterRule = $Devices.DeviceFilterRule
    [array]$Locations = $Conditions | Select-Object -ExpandProperty Locations

    # Process locations - included and excluded
    [string]$IncludeLocationsOutput = $Null; [array]$IncludeLocationNames = $Null 
    If ($Locations.IncludeLocations -eq 'All') {
            [string]$IncludeLocationsOutput = 'All'
    } Else {
            ForEach ($Location in $Locations.IncludeLocations) {
                $IncludeLocationName = (Get-MgIdentityConditionalAccessNamedLocation -NamedLocationId $Location).DisplayName
                $IncludeLocationNames += $IncludeLocationName
            }
            [string]$IncludeLocationsOutput = ($IncludeLocationNames -join ", ")
    }

    [string]$ExcludeLocationsOutput = $Null; [array]$ExcludeLocationNames = $Null 
    If ($Locations.ExcludeLocations -eq 'All') {
            [string]$ExcludeLocationsOutput = 'All'
    } Else {
            ForEach ($Location in $Locations.ExcludeLocations) {
                $ExcludeLocationName = (Get-MgIdentityConditionalAccessNamedLocation -NamedLocationId $Location).DisplayName
                $ExcludeLocationNames += $ExcludeLocationName
            }
            [string]$ExcludeLocationsOutput = ($ExcludeLocationNames -join ", ")
    }

    # Process platforms
    [array]$Platforms = $Conditions | Select-Object -ExpandProperty Platforms
    [array]$ExcludedPlatforms = $Platforms.ExcludePlatforms
    [array]$IncludedPlatforms = $Platforms.IncludePlatforms
    [string]$IncludedPlatformsOutput = $IncludedPlatforms -join ", "
    [string]$ExcludedPlatformsOutput = $ExcludedPlatforms -join ", "

    # Process users, groups, and roles
    [array]$Users = $Conditions | Select-Object -ExpandProperty Users
    [array]$ExcludedRoles = $Users.ExcludeRoles
    [array]$IncludedRoles = $Users.IncludeRoles
    [array]$ExcludedUsers = $Users.ExcludeUsers
    [array]$IncludedUsers = $Users.IncludeUsers
    [array]$IncludedGroups = $Users.IncludeGroups
    [array]$ExcludedGroups = $Users.ExcludeGroups
    $ExcludeGuests = $Conditions.Users | Select-Object -ExpandProperty ExcludeGuestsOrExternalUsers
    $IncludeGuests = $Conditions.Users | Select-Object -ExpandProperty IncludeGuestsOrExternalUsers
    
    # Included groups
    [string]$IncludedGroupsOutput = $null
    If ($IncludedGroups) {
        [array]$IncludedGroupNames = $null
        ForEach ($GroupId in $IncludedGroups) {
            $GroupName = (Get-MgGroup -GroupId $GroupId).DisplayName
            $IncludedGroupNames += $GroupName
        }
        $IncludedGroupsOutput = $IncludedGroupNames -join ", "
    }

    # Excluded groups
    [string]$ExcludedGroupsOutput = $null
    If ($ExcludedGroups) {
        [array]$ExcludedGroupNames = $null
        ForEach ($GroupId in $ExcludedGroups) {
            $GroupName = (Get-MgGroup -GroupId $GroupId).DisplayName
            $ExcludedGroupNames += $GroupName
        }
        $ExcludedGroupsOutput = $ExcludedGroupNames -join ", "
    }

    # Excluded admin roles
    [string]$ExcludedRolesOutput = $null
    If ($ExcludedRoles) {
        [array]$ExcludedRoleNames = $null
        ForEach ($Role in $ExcludedRoles) {
            $RoleName = $HashAdminRoles[$Role]
            $ExcludedRoleNames += $RoleName
        }
        $ExcludedRolesOutput = $ExcludedRoleNames -join ", "
    }

    # Included admin roles
    [string]$IncludedRolesOutput = $null
    If ($IncludedRoles) {
        [array]$IncludedRoleNames = $null
        ForEach ($Role in $IncludedRoles) {
            $RoleName = $HashAdminRoles[$Role]
            $IncludedRoleNames += $RoleName
        }
        $IncludedRolesOutput = $IncludedRoleNames -join ", "
    }

    # Included users - the Regex check for a GUID is to make sure that we don't try to use values like GuestsOrExternalUsers with Get-MgUser
    [string]$IncludedUsersOutput = $null
    If ($IncludedUsers) {
        [array]$IncludedUserNames = $Null
        If ($IncludedUsers[0] -match("^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$")) {
            ForEach ($U in $IncludedUsers) {
                $UserName = (Get-MgUser -UserId $U).UserPrincipalName
                $IncludedUserNames += $UserName
            } 
            $IncludedUsersOutput = $IncludedUserNames -join ", "
        } Else {
            $IncludedUsersOutput = $IncludedUsers[0]
        }
    }   
    
    # Excluded users
    [string]$ExcludedUsersOutput = $null
    If ($ExcludedUsers) {
        [array]$ExcludedUserNames = $Null
        ForEach ($U in $ExcludedUsers) {
            $UserName = (Get-MgUser -UserId $U).UserPrincipalName
            $ExcludedUserNames += $UserName

        }
        $ExcludedUsersOutput = $ExcludedUserNames -join ", "
    }

    If ($Policy.State -eq 'enabledForReportingButNotEnforced') {
        $PolicyState = "Report only"
    } Else {
        $PolicyState = $Policy.State
    }

    # Session Controls
    $AppRestrictions = $SessionControls | Select-Object -ExpandProperty ApplicationEnforcedRestrictions
    $AppRestrictionsEnabled = $AppRestrictions.IsEnabled
    $CloudAppSecurity = $SessionControls | Select-Object -ExpandProperty CloudAppSecurity
    $PersistentBrowser = $SessionControls | Select-Object -ExpandProperty PersistentBrowser
    $SignInFrequency = $SessionControls | Select-Object -ExpandProperty SignInFrequency

    # Grant controls
    $AuthenticationStrength = $GrantControls | Select-Object -ExpandProperty AuthenticationStrength
    $BuiltInControls = $GrantControls | Select-Object -ExpandProperty builtInControls
    $GrantOperator = $GrantControls | Select-Object -ExpandProperty Operator
    $TermsOfUse = $GrantControls | Select-Object -ExpandProperty TermsOfUse

    # Output the data
    $DataLine = [PSCustomObject][Ordered]@{
        'Policy'                        = $Policy.ID
        'Policy name'                   = $Policy.displayName
        'State'                         = $PolicyState
        'Created'                       = $Policy.CreatedDateTime
        'Last modified'                 = $Policy.ModifiedDateTime
        'Included apps'                 = $IncludedAppOutput
        'Excluded apps'                 = $ExcludedAppOutput
        'App filter mode'               = $ApplicationFilterMode
        'App filter rule'               = $ApplicationFilterRule
        'Client app types'              = $ClientAppTypes
        'Included client apps'          = $IncludedClientAppNames
        'Excluded client apps'          = $ExcludedClientAppNames
        'Included Locations'            = $IncludeLocationsOutput
        'Excluded locations'            = $ExcludeLocationsOutput
        'Included users'                = $IncludedUsersOutput
        'Excluded users'                = $ExcludedUsersOutput
        'Included groups'               = $IncludedGroupsOutput
        'Excluded groups'               = $ExcludedGroupsOutput
        'Included roles'                = $IncludedRolesOutput
        'Excluded roles'                = $ExcludedRolesOutput
        'Include guests'               = $IncludeGuests
        'Exclude guests'                = $ExcludeGuests
        'User risk levels'              = $UserRiskLevels
        'Sign in risk levels'           = $SignInRiskLevels
        'SP risk levels'                = $ServicePrincipalRiskLevels
        'Auth context'                  = $IncludeAuthenticationContextClassReferences
        'Include User actions'          = $IncludeUserActions
        'Included platforms'            = $IncludedPlatformsOutput
        'Excluded platforms'            = $ExcludedPlatformsOutput
        'Device filter mode'            = $DeviceFilterMode
        'Device filter rule'            = $DeviceFilterRule    
        'App Restrictions'              = $AppRestrictionsEnabled
        'Cloud App Security'            = $CloudAppSecurity.IsEnabled
        'CAS type'                      = $CloudAppSecurity.Type
        'Persistent browser'            = $PersistentBrowser.IsEnabled
        'Browser mode'                  = $PersistentBrowser.Mode
        'Sign in frequency'             = $SignInFrequency.IsEnabled
        'Sign in frequency auth'        = $SignInFrequency.AuthenticationType
        'Sign in frequency interval'    = $SignInFrequency.FrequencyInterval
        'Sign in frequency type'        = $SignInFrequency.Type
        'Sign in frequency value'       = $SignInFrequency.Value
        'Authentication strength'       = $AuthenticationStrength.displayName
        'Auth strength description'     = $AuthenticationStrength.description
        'Built in controls'             = $BuiltInControls
        'Grant operator'                = $GrantOperator
        'Terms of use'                  = $TermsOfUse
        
    }
    $Report.Add($DataLine)
}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.