# ReportTeamsPolicyAssignments.PS1
# Generate a report about major Teams policies assigned to user accounts
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportTeamsPolicyAssignments.PS1
# V1.1 27-Jan-2024 Add support for group assigned policies
# Only a selected set of Teams policies are included in this report. It's easy to add more by duplicating the code for another
# policy (changing the name etc.)

function Get-GroupNameByRef {   
    Param(
    #GroupId
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$GroupId)
    
    If (!$GroupNames[$GroupId]) {
        $GroupNames[$GroupId] = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
    }
    return $GroupNames[$GroupId]
}

[array]$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "MicrosoftTeams")) {Write-Host "Please connect to the Microsoft Teams module and then restart the script"; break}

Connect-MgGraph -NoWelcome -Scopes Organization.Read.All, Group.Read.All

$OrgName = (Get-MgOrganization).DisplayName
[string]$CreationDate = Get-Date -format 'dd-MMM-yyyy HH:mm'
$Global:GroupNames = @{}
$CSVOutput = "c:\temp\TeamsPolicyAssignments.CSV"
$ReportFile = "c:\temp\TeamsPolicyAssignments.html"

Write-Host "Looking for Teams-enabled users..."
[array]$Users = Get-CsOnlineUser -ResultSize 5000
# Filter the set to get Teams users - this will filter out all but cloud-only Teams users. If you don't want to use the filter, comment it out.
$Users = $Users | Where-Object {$_.InterpretedUserType -eq "PureOnlineTeamsOnlyUser" -or $_.InterpretedUserType -eq "PureOnlineTeamsOnlyUserFailedPublishingToAAD"} | Sort-Object DisplayName
If (!($Users)) 
    {Write-Host "No users found - exiting"; 
    break 
}

# Set up report file
$Report = [System.Collections.Generic.List[Object]]::new()
[int]$i = 0
# Process each user to fetch their policy assignments
$TenantDefaultString = "Tenant Default"
ForEach ($User in $Users) {
    $i++
    Write-Host ("Processing user {0} ({1}/{2})... " -f $User.displayName, $i, $Users.count)
    $TeamsMeetingPolicy = $TenantDefaultString
    $TeamsCallingPolicy = $TenantDefaultString
    $TeamsMessagingPolicy = $TenantDefaultString
    $TeamsAppSetupPolicy = $TenantDefaultString
    $TeamsAppPermissionPolicy = $TenantDefaultString
    $TeamsEncryptionPolicy = $TenantDefaultString
    $TeamsUpdatePolicy = $TenantDefaultString
    $TeamsPolicy = $TenantDefaultString
    $TeamsEventsPolicy = $TenantDefaultString
    $TeamsAudioConferencingPolicy = $TenantDefaultString

    # Meeting policy
    $CurrentAssignment = $null
    If ($User.TeamsMeetingPolicy) {
        $TeamsMeetingPolicy = $User.TeamsMeetingPolicy.Name
    } Else {
        [array]$PolicyAssignments = Get-CsUserPolicyAssignment -Identity $User.UserPrincipalName `
          -PolicyType TeamsMeetingPolicy | Select-Object -ExpandProperty PolicySource
        If ($PolicyAssignments) {
            $CurrentAssignment = $PolicyAssignments[0]
        }
        If ($CurrentAssignment) {
            Switch ($CurrentAssignment.AssignmentType) {
                "Direct" {
                    $TeamsMeetingPolicy = ("{0} (Direct)" -f $CurrentAssignment.PolicyName)
                }
                "Group" {
                    $GroupName = (Get-GroupNameByRef -GroupId $CurrentAssignment.Reference).DisplayName
                    $TeamsMeetingPolicy = ("{0} (Group: {1})" -f $CurrentAssignment.PolicyName, $GroupName)
                }
            }
        }
    }

    # Calling policy
    $CurrentAssignment = $null
    If ($User.TeamsCallingPolicy) {
        $TeamsCallingPolicy = $User.TeamsCallingPolicy.Name
    } Else {
        [array]$PolicyAssignments = Get-CsUserPolicyAssignment -Identity $User.UserPrincipalName `
          -PolicyType TeamsCallingPolicy | Select-Object -ExpandProperty PolicySource
        If ($PolicyAssignments) {
            $CurrentAssignment = $PolicyAssignments[0]
        }
        If ($CurrentAssignment) {
            Switch ($CurrentAssignment.AssignmentType) {
                "Direct" {
                    $TeamsCallingPolicy = ("{0} (Direct)" -f $CurrentAssignment.PolicyName)
                }
                "Group" {
                    $GroupName = (Get-GroupNameByRef -GroupId $CurrentAssignment.Reference).DisplayName
                    $TeamsMeetingPolicy = ("{0} (Group: {1})" -f $CurrentAssignment.PolicyName, $GroupName)
                }
            }
        }
    }

    # Messaging policy
    $CurrentAssignment = $null
    If ($User.TeamsMessagingPolicy) {
        $TeamsMessagingPolicy = $User.TeamsMessagingPolicy.Name
    } Else {
        [array]$PolicyAssignments = Get-CsUserPolicyAssignment -Identity $User.UserPrincipalName `
          -PolicyType TeamsMessagingPolicy | Select-Object -ExpandProperty PolicySource
        If ($PolicyAssignments) {
            $CurrentAssignment = $PolicyAssignments[0]
        }
        If ($CurrentAssignment) {
            Switch ($CurrentAssignment.AssignmentType) {
                "Direct" {
                    $TeamsMessagingPolicy = ("{0} (Direct)" -f $CurrentAssignment.PolicyName)
                }
                "Group" {
                    $GroupName = (Get-GroupNameByRef -GroupId $CurrentAssignment.Reference).DisplayName
                    $TeamsMessagingPolicy = ("{0} (Group: {1})" -f $CurrentAssignment.PolicyName, $GroupName)
                }
            }
        }
    }    
    # App setup policy
    $CurrentAssignment = $null
    If ($User.TeamsAppSetupPolicy) {
        $TeamsAppSetupPolicy = $User.TeamsAppSetupPolicy
    } Else {
        [array]$PolicyAssignments = Get-CsUserPolicyAssignment -Identity $User.UserPrincipalName `
          -PolicyType TeamsAppSetupPolicy | Select-Object -ExpandProperty PolicySource
        If ($PolicyAssignments) {
            $CurrentAssignment = $PolicyAssignments[0]
        }
        If ($CurrentAssignment) {
            Switch ($CurrentAssignment.AssignmentType) {
                "Direct" {
                    $TeamsAppSetupPolicy = ("{0} (Direct)" -f $CurrentAssignment.PolicyName)
                }
                "Group" {
                    $GroupName = (Get-GroupNameByRef -GroupId $CurrentAssignment.Reference).DisplayName
                    $TeamsAppSetupPolicy = ("{0} (Group: {1})" -f $CurrentAssignment.PolicyName, $GroupName)
                }
            }
        }
    }
    # App Permission policy
    $CurrentAssignment = $null
    If ($User.TeamsAppPermissionPolicy) {
        $TeamsAppPermissionPolicy = $User.TeamsAppPermissionPolicy
    } Else {
        [array]$PolicyAssignments = Get-CsUserPolicyAssignment -Identity $User.UserPrincipalName `
          -PolicyType TeamsAppPermissionPolicy | Select-Object -ExpandProperty PolicySource
        If ($PolicyAssignments) {
            $CurrentAssignment = $PolicyAssignments[0]
        }
        If ($CurrentAssignment) {
            Switch ($CurrentAssignment.AssignmentType) {
                "Direct" {
                    $TeamsAppPermissionPolicy = ("{0} (Direct)" -f $CurrentAssignment.PolicyName)
                }
                "Group" {
                    $GroupName = (Get-GroupNameByRef -GroupId $CurrentAssignment.Reference).DisplayName
                    $TeamsAppPermissionPolicy = ("{0} (Group: {1})" -f $CurrentAssignment.PolicyName, $GroupName)
                }
            }
        }
    }
    # Teams policy
    $CurrentAssignment = $null
    If ($User.TeamsChannelsPolicy) {
        $TeamsPolicy = $User.TeamsChannelsPolicy
    } Else {
        [array]$PolicyAssignments = Get-CsUserPolicyAssignment -Identity $User.UserPrincipalName `
          -PolicyType TeamsChannelsPolicy | Select-Object -ExpandProperty PolicySource
        If ($PolicyAssignments) {
            $CurrentAssignment = $PolicyAssignments[0]
        }
        If ($CurrentAssignment) {
            Switch ($CurrentAssignment.AssignmentType) {
                "Direct" {
                    $TeamsPolicy = ("{0} (Direct)" -f $CurrentAssignment.PolicyName)
                }
                "Group" {
                    $GroupName = (Get-GroupNameByRef -GroupId $CurrentAssignment.Reference).DisplayName
                    $TeamsPolicy = ("{0} (Group: {1})" -f $CurrentAssignment.PolicyName, $GroupName)
                }
            }
        }
    }   
    # Teams encryption policy
    $CurrentAssignment = $null
    If ($User.TeamsEnhancedEncryptionPolicy) {
        $TeamsEncryptionPolicy = $User.TeamsEnhancedEncryptionPolicy
    } Else {
        [array]$PolicyAssignments = Get-CsUserPolicyAssignment -Identity $User.UserPrincipalName `
          -PolicyType TeamsEnhancedEncryptionPolicy | Select-Object -ExpandProperty PolicySource
        If ($PolicyAssignments) {
            $CurrentAssignment = $PolicyAssignments[0]
        }
        If ($CurrentAssignment) {
            Switch ($CurrentAssignment.AssignmentType) {
                "Direct" {
                    $TeamsEncryptionPolicy = ("{0} (Direct)" -f $CurrentAssignment.PolicyName)
                }
                "Group" {
                    $GroupName = (Get-GroupNameByRef -GroupId $CurrentAssignment.Reference).DisplayName
                    $TeamsEncryptionPolicy = ("{0} (Group: {1})" -f $CurrentAssignment.PolicyName, $GroupName)
                }
            }
        }
    }   
    # Teams Update management policy
    $CurrentAssignment = $null
    If ($User.TeamsUpdateManagementPolicy) {
        $TeamsUpdatePolicy = $User.TeamsUpdateManagementPolicy
    } Else {
        [array]$PolicyAssignments = Get-CsUserPolicyAssignment -Identity $User.UserPrincipalName `
          -PolicyType TeamsUpdateManagementPolicy | Select-Object -ExpandProperty PolicySource
        If ($PolicyAssignments) {
            $CurrentAssignment = $PolicyAssignments[0]
        }
        If ($CurrentAssignment) {
            Switch ($CurrentAssignment.AssignmentType) {
                "Direct" {
                    $TeamsUpdatePolicy = ("{0} (Direct)" -f $CurrentAssignment.PolicyName)
                }
                "Group" {
                    $GroupName = (Get-GroupNameByRef -GroupId $CurrentAssignment.Reference).DisplayName
                    $TeamsUpdatePolicy = ("{0} (Group: {1})" -f $CurrentAssignment.PolicyName, $GroupName)
                }
            }
        }
    }   
    # Teams Events policy
    $CurrentAssignment = $null
    If ($User.TeamsEventsPolicy) {
        $TeamsEventsPolicy = $User.TeamsEventsPolicy
    } Else {
        [array]$PolicyAssignments = Get-CsUserPolicyAssignment -Identity $User.UserPrincipalName `
          -PolicyType TeamsEventsPolicy | Select-Object -ExpandProperty PolicySource
        If ($PolicyAssignments) {
            $CurrentAssignment = $PolicyAssignments[0]
        }
        If ($CurrentAssignment) {
            Switch ($CurrentAssignment.AssignmentType) {
                "Direct" {
                    $TeamsEventsPolicy = ("{0} (Direct)" -f $CurrentAssignment.PolicyName)
                }
                "Group" {
                    $GroupName = (Get-GroupNameByRef -GroupId $CurrentAssignment.Reference).DisplayName
                    $TeamsEventsPolicy = ("{0} (Group: {1})" -f $CurrentAssignment.PolicyName, $GroupName)
                }
            }
        }
    }   
    # Teams Audio Conferencing Policy
    $CurrentAssignment = $null
    If ($User.TeamsAudioConferencingPolicy) {
        $TeamsAudioConferencingPolicy = $User.TeamsAudioConferencingPolicy
    } Else {
        [array]$PolicyAssignments = Get-CsUserPolicyAssignment -Identity $User.UserPrincipalName `
          -PolicyType TeamsAudioConferencingPolicy | Select-Object -ExpandProperty PolicySource
        If ($PolicyAssignments) {
            $CurrentAssignment = $PolicyAssignments[0]
        }
        If ($CurrentAssignment) {
            Switch ($CurrentAssignment.AssignmentType) {
                "Direct" {
                    $TeamsAudioConferencingPolicy = ("{0} (Direct)" -f $CurrentAssignment.PolicyName)
                }
                "Group" {
                    $GroupName = (Get-GroupNameByRef -GroupId $CurrentAssignment.Reference).DisplayName
                    $TeamsAudioConferencingPolicy = ("{0} (Group: {1})" -f $CurrentAssignment.PolicyName, $GroupName)
                }
            }
        }
    }   
    
    # Output a report line
    $ReportLine = [PSCustomObject][Ordered]@{  
        User                         = $User.DisplayName
        UPN                          = $User.UserPrincipalName
        "Messaging Policy"           = $TeamsMessagingPolicy
        "Meeting Policy"             = $TeamsMeetingPolicy
        "Calling Policy"             = $TeamsCallingPolicy
        "Audio conferencing policy"  = $TeamsAudioConferencingPolicy
        "App Setup Policy"           = $TeamsAppSetupPolicy
        "App Permissions Policy"     = $TeamsAppPermissionPolicy
        "Enhanced Encryption Policy" = $TeamsEncryptionPolicy
        "Update Policy"              = $TeamsUpdatePolicy
        "Teams Policy"               = $TeamsPolicy
        "Teams Events"               = $TeamsEventsPolicy
	    "InterpretedUserType"        = $User.InterpretedUserType
    }
 #  Add the line to the report
    $Report.Add($ReportLine) 
}

# Create the HTML report
$Version = "1.1"
$htmlhead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.pass{background: #B7EB83;}
	   td.warn{background: #FFF275;}
	   td.fail{background: #FF2626; color: #ffffff;}
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Teams Policy Assignment Report</h1></p>
           <p><h2><b>For the " + $OrgName + " organization</b></h2></p>
           <p><h3>Generated: " + (Get-Date -format g) + "</h3></p></div>"

$htmlbody1 = $Report | ConvertTo-Html -Fragment
$htmltail = "<p>Report created for: " + $OrgName + "</p>" +
             "<p>Created: " + $CreationDate + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of Teams users found:    " + $Users.Count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Teams Policy Assignment Report<b> " + $Version + "</b>"	
# Generate the HTML file
$htmlreport = $htmlhead + $htmlbody1 + $htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8
Write-Host ("All done. Teams policies for {0} users analyzed. CSV file is available at {1} and a HTML report at {2}" -f $Users.Count, $CSVOutput, $ReportFile)
$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation $CSVOutput

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
