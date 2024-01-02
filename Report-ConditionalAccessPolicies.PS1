# Report-ConditionalAccessPolicies.PS1
# A script to show how to report conditional access policy settings
# V1.0 16-Dec-2023
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ConditionalAccessPolicies.PS1

[string]$RunDate = Get-Date -format "dd-MMM-yyyy HH:mm:ss"
$Version = "1.0"
$CSVOutputFile = "c:\temp\CAPoliciesReport.CSV"
$ReportFile = "c:\temp\CAPoliciesReport.html"

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
    $ExcludeGuests = $Conditions.Users.ExcludeGuestsOrExternalUsers | Select-Object -ExpandProperty GuestOrExternalUserTypes
    $IncludeGuests = $Conditions.Users.IncludeGuestsOrExternalUsers | Select-Object -ExpandProperty GuestOrExternalUserTypes

    $ExcludeGuestsOutput = $null
    $IncludeGuestsOutput = $null
    # Excluded Guests
    If ($ExcludeGuests) {
        [array]$GuestTypes = $null
        ForEach ($Guest in $ExcludeGuests) {
            Switch ($Guest) {
                "internalGuest" { 
                    $GuestType = "Local Guest users" }
                "b2bCollaborationGuest" {
                    $GuestType = "B2B Collaboration guest users"
                }
                "b2bCollaborationMember" {
                    $GuestType = "B2B Collaboration member users" 
                }
                "b2bDirectConnectUser" {
                    $GuestType = "B2B Direct connect users"
                }
                "otherExternalUser" {
                    $GuestType = "Other external users"
                }
                "serviceProvider" {
                    $GuestType = "Service Provider users"
                }
            }
            $GuestTypes += $GuestType
        }
        $ExcludeGuestsOutput = $GuestTypes -join ", "
    }

    # included guests
    If ($IncludeGuests) {
        [array]$GuestTypes = $null
        ForEach ($Guest in $IncludeGuests) {
            Switch ($Guest) {
                "internalGuest" { 
                    $GuestType = "Local Guest users" }
                "b2bCollaborationGuest" {
                    $GuestType = "B2B Collaboration guest users"
                }
                "b2bCollaborationMember" {
                    $GuestType = "B2B Collaboration member users" 
                }
                "b2bDirectConnectUser" {
                    $GuestType = "B2B Direct connect users"
                }
                "otherExternalUser" {
                    $GuestType = "Other external users"
                }
                "serviceProvider" {
                    $GuestType = "Service Provider users"
                }
            }
            $GuestTypes += $GuestType
        }
        $IncludeGuestsOutput = $GuestTypes -join ", "
    }
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
        'Include guests'                = $IncludeGuestsOutput
        'Exclude guests'                = $ExcludeGuestsOutput
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

# Now to generate a HTML report
Write-Host "Generating report..."
$OrgName  = (Get-MgOrganization).DisplayName
#  First, define the header.
$HTMLHead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.enabled{background: #B7EB83;}
	   td.disabled{background: #E3242B;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Conditional Access Policies Report</h1></p>
           <p><h2><b>For the " + $Orgname + " tenant</b></h2></p>
           <p><h3>Generated: " + $RunDate + "</h3></p></div>"


# This section highlights whether a conditional access policy is enabled or disabled in the summary.
# Idea from https://stackoverflow.com/questions/37662940/convertto-html-highlight-the-cells-with-special-values
# First, convert the CA Policies report to HTML and then import it into an XML structure
[array]$CAPoliciesTable = $Report | Select-Object 'Policy Name', 'Policy', State, 'Created', 'Last Modified'
$HTMLTable = $CAPoliciesTable | ConvertTo-Html -Fragment
[xml]$XML = $HTMLTable
# Create an attribute class to use, name it, and append to the XML table attributes
$TableClass = $XML.CreateAttribute("class")
$TableClass.Value = "State"
$XML.table.Attributes.Append($TableClass) | Out-Null
# Conditional formatting for the table rows. The number of available units is in table row 6, so we update td[5]
ForEach ($TableRow in $XML.table.SelectNodes("tr")) {
    # each TR becomes a member of class "tablerow"
    $TableRow.SetAttribute("class","tablerow")
    ## If row has TD check the state of the policy
    If (($TableRow.td) -and ([string]$TableRow.td[2] -eq 'enabled'))  {
        ## tag the TD with eirher the color for "warn" or "pass" defined in the heading
        $TableRow.SelectNodes("td")[2].SetAttribute("class","enabled")
    } ElseIf (($TableRow.td) -and ([string]$TableRow.td[2] -eq 'disabled')) {
        $TableRow.SelectNodes("td")[2].SetAttribute("class","disabled")
    }
}
# Wrap the output table with a div tag
$HTMLBody = [string]::Format('<div class="tablediv">{0}</div>',$XML.OuterXml)

# Add separate section for each policy
[string]$HTMLPolicySeparator = "<p><h2>Details of Conditional Access Policies</h2></p>"
[string]$HTMLPolicyOutput = $null

# Foreach policy, extract its details and output
ForEach ($Policy in $Report) {
    $HTMLPolicyHeader = "<h3>Policy Settings for <b>" + $Policy.'Policy Name' + "<b></h3><p>"
    $HTMLPolicyContent = $Policy | `
      Select-Object 'Authentication strength', 'Auth strength description', 'Included apps', 'Excluded apps', 'Included Locations', 'Excluded locations', `
       'Included users', 'Excluded users', 'Included groups', 'Excluded groups', 'Included roles', 'Excluded roles', 'Include guests','Exclude guests'  | ConvertTo-HTML -Fragment
    $HTMLPolicyOutput = $HTMLPolicyOutput + $HTMLPolicyHeader + $HTMLPolicyContent + "<p><p>"
}

# End stuff to output
$HTMLTail = "<p>Report created for the " + $OrgName + " tenant on " + $RunDate + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of conditionl access policies found:     " + $CAPolicies.Count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Entra ID Conditional Access Policies Report<b> " + $Version + "</b>"	

$HTMLReport = $HTMLHead + $HTMLBody + $HTMLPolicySeparator + $HTMLPolicyOutput + $HTMLtail
$HTMLReport | Out-File $ReportFile  -Encoding UTF8

$Report | Export-Csv -NoTypeInformation $CSVOutputFile -Encoding utf8
Write-Host ("HTML format report is available in {0}  and CSV file in {1}" -f $ReportFile, $CSVOutputFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
