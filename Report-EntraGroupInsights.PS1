# Report-EntraGroupInsights.PS1
# A script to show how to extract the Entra Groups insights from the Graph to report them in a readable format.

# Github link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-EntraGroupInsights.PS1
# V1.0 27-Mar-2026

Connect-MgGraph -Scopes "Group.Read.All", "GroupMember.Read.All", "User.Read.All", "Organization.Read.All" -NoWelcome
Write-Host "Starting up by finding groups and insights..." -ForegroundColor Green

# Can we get insights?
$Uri = "https://graph.microsoft.com/beta/reports/identityAnalytics/groups"
[array]$GroupInsights = Invoke-MgGraphRequest -Method Get -Uri $Uri -OutputType PSObject | Select-Object -ExpandProperty value

If ($GroupInsights) {
    Write-Host "Group insights retrieved successfully." -ForegroundColor Green
} Else {
    Write-Host "Failed to retrieve group insights." -ForegroundColor Red
    Break
}
$InsightsDate = Get-Date $GroupInsights[0].CalculatedDateTime -format "dd-MMM-yyyy"

[array]$Groups = Get-MgGroup -All -Property "id,displayName,description,createdDateTime,securityEnabled,mailEnabled,groupTypes,owners" -PageSize 500 -ExpandProperty Owners
If ($Groups) {
    Write-Host ("{0} groups retrieved successfully." -f $Groups.count)   -ForegroundColor Green
} Else {
    Write-Host "Failed to retrieve groups." -ForegroundColor Red
    Break
}   

$Report =  [System.Collections.Generic.List[Object]]::new()

ForEach ($Group in $Groups) {
    $Insight = $null
    $Insight = $GroupInsights | Where-Object { $_.id -eq $Group.id }
    If ($Insight) {
        $GroupSoftDeletionDateTime = $null; $GroupRestorationDateTime = $null; $GroupExpirationDateTime = $null
        If ($Insight.groupExpirationDateTime) {
            $GroupExpirationDateTime = Get-Date($Insight.groupExpirationDateTime) -format 'dd-MMM-yyyy HH:mm'
        }
        If ($insight.softDeletionDateTime) {
            $GroupSoftDeletionDateTime = Get-Date($Insight.softDeletionDateTime) -format 'dd-MMM-yyyy HH:mm'
        }
        If ($INsight.GroupRestorationDateTime) {
            $GroupRestorationDateTime = Get-Date($Insight.GroupRestorationDateTime) -format 'dd-MMM-yyyy HH:mm'
        }   
        If ($Group.GroupTypes -contains "DynamicMembership") {
             $DynamicGroup = $true
        } Else {
             $DynamicGroup = $false
        }
        If ($Group.GroupTypes -contains "Unified") {
             $M365Group = $true
        } Else {
             $M365Group = $false
        }
        $ReportLine = [PSCustomObject]@{ 
            Group                          = $Group.displayName
            Id                             = $Group.id
            Created                        = Get-Date($Group.createdDateTime) -format 'dd-MMM-yyyy HH:mm'
            expirationDateTime             = $GroupExpirationDateTime
            softDeletionDateTime           = $GroupSoftDeletionDateTime
            lastRestorationDateTime        = $GroupRestorationDateTime 
            SecurityEnabled                = $Group.securityEnabled
            MailEnabled                    = $Group.mailEnabled
            GroupTypes                     = ($Group.groupTypes -join ",")
            TotalMembers                   = $Insight.memberTransitiveUserCount
            TotalOwners                    = $Group.owners.Count
            Owners                         = $Group.owners.additionalProperties.displayName -join ", "
            TotalGuests                    = $Insight.guestTransitiveUserCount
            GuestOwnerCount                = $Insight.guestOwnerCount
            MemberOwnerCount               = $Insight.memberOwnerCount
            TotalMembership                = $Insight.transitiveUserCount
            servicePrincipalOwnerCount     = $Insight.servicePrincipalOwnerCount
            transitiveServicePrincipalCount = $Insight.transitiveServicePrincipalCount
            DynamicMembership              = $DynamicGroup
            'Microsoft 365 Group'          = $M365Group
            MembershipRuleProcessingState  = $Insight.membershipRuleProcessingState
            MembershipRuleExpressionCount  = $Insight.membershipRuleExpressionCount
            membershipRuleContainsCount    = $Insight.membershipRuleContainsCount
            membershipRuleMatchCount       = $Insight.membershipRuleMatchCount
            sensitivityLabelCount          = $Insight.sensitivityLabelCount
            Grouptype                      = $Insight.groupType 
            Visibility                     = $Group.visibility
            CloudDL                        = $Insight.isCloudDistributionListGroup
            OnPremisesDL                   = $Insight.isOnPremisesDistributionListGroup  
            CloudSecurityGroup             = $Insight.isCloudSecurityGroup   
            OnPremisesSecurityGroup        = $Insight.isOnPremisesSecurityGroup
            CloudMailEnabledSecurityGroup = $Insight.isCloudMailEnabledSecurityGroup
            OnPremisesMailEnabledSecurityGroup = $Insight.isOnPremisesMailEnabledSecurityGroup 
        } 
        $Report.Add($ReportLine) 
    } Else {
        $ReportLine = [PSCustomObject]@{ 
            Group                      = $Group.displayName
            Id                         = $Group.id
            Created                    = Get-Date($Group.createdDateTime) -format 'dd-MMM-yyyy HH:mm'
            SecurityEnabled            = $Group.securityEnabled
            MailEnabled                = $Group.mailEnabled
            GroupTypes                 = ($Group.groupTypes -join ",")
            TotalMembers               = "unknown"
            TotalOwners                = $Group.owners.Count
            Owners                     = $Group.owners.additionalProperties.displayName -join ", "
        } 
        $Report.Add($ReportLine) 
    }

}

$Report = $Report | Sort-Object -Property Group

# Now we can generate some analytics from the insights

[datetime]$30DaysAgo = (Get-Date).AddDays(-30)
[datetime]$30DaysFromNow = (Get-Date).AddDays(30)

# Graph API 
# https://graph.microsoft.com/beta/reports/identityAnalytics/groups?$filter=guestOwnerCount eq 0 and memberOwnerCount eq 0 and servicePrincipalOwnerCount eq 0
[array]$GroupsWithNoOwners = $Report | Where-Object { $_.GuestOwnerCount -eq 0 -and $_.MemberOwnerCount -eq 0 -and $_.servicePrincipalOwnerCount -eq 0 }
Write-Host ("Groups with no owners: {0}" -f $GroupsWithNoOwners.count) -ForegroundColor Yellow
[array]$GroupsWithServicePrincipalOwners = $Report | Where-Object { $_.servicePrincipalOwnerCount -gt 0 }
Write-Host ("Groups with service principals as owners: {0}" -f $GroupsWithServicePrincipalOwners.count) -ForegroundColor Yellow
[array]$GroupsWithServicePrincipalsAsOwnersOrMembers = $Report | Where-Object { $_.servicePrincipalOwnerCount -gt 0 -or $_.transitiveServicePrincipalCount -gt 0 }
Write-Host ("Groups with service principals as owners or members: {0}" -f $GroupsWithServicePrincipalsAsOwnersOrMembers.count) -ForegroundColor Yellow
[array]$GroupsWithGuestsAsOwners = $Report | Where-Object { $_.GuestOwnerCount -gt 0 }
Write-Host ("Groups with guest owners: {0}" -f $GroupsWithGuestsAsOwners.count) -ForegroundColor Yellow
[int]$ServicePrincipalsAsMembers =  $GroupsWithServicePrincipalsAsOwnersOrMembers.count - $GroupsWithServicePrincipalOwners.count
Write-Host ("Groups with service principals as members but not owners: {0}" -f $ServicePrincipalsAsMembers) -ForegroundColor Yellow

[array]$GroupsWithDynamicMembership = $Report | Where-Object { $_.DynamicMembership -eq $true }
Write-Host ("Groups with dynamic membership: {0}" -f $GroupsWithDynamicMembership.count) -ForegroundColor Yellow
[array]$GroupsWithPendingExpiration = $Report | Where-Object {$_.expirationDateTime -ne $null -and $_.expirationDateTime -ne "01-Jan-0001 00:00" -and $_.expirationDateTime -as [datetime] -lt $30DaysFromNow}
Write-Host ("Groups with pending expiration in the next 30 days: {0}" -f $GroupsWithPendingExpiration.count) -ForegroundColor Yellow
[array]$GroupsInSoftDeletion = $Report | Where-Object { $_.softDeletionDateTime -ne $null -and $_.softDeletionDateTime -as [datetime] -gt $30DaysAgo }
Write-Host ("Soft-deleted groups: {0}" -f $GroupsInSoftDeletion.count) -ForegroundColor Yellow
[array]$GroupsRestored = $Report | Where-Object { $_.lastRestorationDateTime -as [datetime] -gt $30DaysAgo }
Write-Host ("Groups restored in the last 30 days: {0}" -f $GroupsRestored.count) -ForegroundColor Yellow
[array]$NewlyCreatedGroups = $Report | Where-Object { $_.Created -as [datetime] -gt (Get-Date).AddDays(-30) }
Write-Host ("Newly created groups in the last 30 days: {0}" -f $NewlyCreatedGroups.count) -ForegroundColor Yellow
[array]$GroupsWithNoSensitivityLabel = $Report | Where-Object { $_.sensitivityLabelCount -eq 0 }
Write-Host ("Groups with no sensitivity label: {0}" -f $GroupsWithNoSensitivityLabel.count) -ForegroundColor Yellow
$M365GroupCount = ($Report | Where-Object { $_.'Microsoft 365 Group' -eq $true }).Count
Write-Host ("Number of Microsoft 365 Groups: {0}" -f $M365GroupCount) -ForegroundColor Yellow
$DynamicMicrosoft365GroupCount = ($Report | Where-Object { $_.'Microsoft 365 Group' -eq $true -and $_.DynamicMembership -eq $true }).Count
Write-Host ("Number of Dynamic Microsoft 365 Groups: {0}" -f $DynamicMicrosoft365GroupCount) -ForegroundColor Yellow

# Largest group
$LargestGroup = $Report | Sort-Object -Property TotalMembership -Descending | Select-Object -First 1
Write-Host ("The largest group is {0} with {1} members." -f $LargestGroup.Group, $LargestGroup.TotalMembership) -ForegroundColor Yellow
# Group with largest amount of guest accounts
$LargestGroupWithGuests = $Report | Where-Object { $_.TotalGuests -gt 0 } | Sort-Object -Property TotalMembership -Descending | Select-Object -First 1
Write-Host ("Group with largest number of guest accounts is {0} with {1} members." -f $LargestGroupWithGuests.Group, $LargestGroupWithGuests.TotalMembership) -ForegroundColor Yellow
# Group with largest amount of member accounts
$LargestGroupWithMembers = $Report | Where-Object { $_.TotalMembers -gt 0 } | Sort-Object -Property TotalMembers -Descending | Select-Object -First 1
Write-Host ("Group with largest number of member accounts is {0} with {1} members." -f $LargestGroupWithMembers.Group, $LargestGroupWithMembers.TotalMembership) -ForegroundColor Yellow

# Groups with complicated rules
$GroupsComplicatedRules = $Report | Where-Object { $_.MembershipRuleExpressionCount -gt 10 } | Sort-Object -Property MembershipRuleExpressionCount -Descending
Write-Host ("Dynamic groups with complicated membership rules: {0}" -f $GroupsComplicatedRules.count) -ForegroundColor Yellow
# Dynamic groups using ‘contains’ or ‘match’ operators in their membership rules, which can impact performance.
$GroupInefficientRules = $Report | Where-Object { $_.MembershipRuleMatchCount -gt 0 -or $_.MembershipRuleContainsCount -gt 0 } | Sort-Object -Property MembershipRuleExpressionCount -Descending
Write-Host ("Dynamic groups with inefficient membership rules: {0}" -f $GroupInefficientRules.count) -ForegroundColor Yellow

$CloudDLCount = ($Report | Where-Object { $_.CloudDL -eq $true }).Count     
Write-Host ("Cloud distribution lists: {0}" -f $CloudDLCount) -ForegroundColor Yellow
$OnPremisesDLCount = ($Report | Where-Object { $_.OnPremisesDL -eq $true }).Count
Write-Host ("On-premises distribution lists : {0}" -f $OnPremisesDLCount) -ForegroundColor Yellow      
$CloudSecurityGroupCount = ($Report | Where-Object { $_.CloudSecurityGroup -eq $true }).Count
Write-Host ("Cloud security groups: {0}" -f $CloudSecurityGroupCount) -ForegroundColor Yellow
$OnPremisesSecurityGroupCount = ($Report | Where-Object { $_.OnPremisesSecurityGroup -eq $true }).Count
Write-Host ("On-premises security groups: {0}" -f $OnPremisesSecurityGroupCount) -ForegroundColor Yellow
$CloudMailEnabledSecurityGroupCount = ($Report | Where-Object { $_.CloudMailEnabledSecurityGroup -eq $true }).Count
Write-Host ("Cloud mail-enabled security groups: {0}" -f $CloudMailEnabledSecurityGroupCount) -ForegroundColor Yellow
$OnPremisesMailEnabledSecurityGroupCount = ($Report | Where-Object { $_.OnPremisesMailEnabledSecurityGroup -eq $true }).Count   
Write-Host ("On-premises mail-enabled security groups: {0}" -f $OnPremisesMailEnabledSecurityGroupCount) -ForegroundColor Yellow

# OK. Output the report files
Write-Host "Generating HTML report..." -ForegroundColor Green
$CreationDate = Get-Date -format 'dd-MMM-yyyy HH:mm:ss'
$Version = "1.0"
$Organization = (Get-MgOrganization).DisplayName

# Create the HTML report
$HTMLHead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 36px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 24px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Entra ID Groups Insights</h1></p>
           <p><h2><b>For the " + $Organization + " organization</b></h2></p>
           <p><h3>Generated: " + (Get-Date -format "dd-MMM-yyyy") + " (Insights from " + $InsightsDate + ")</h3></p></div>"

# Add a section for groups with no owners
If ($GroupsWithNoOwners.count -gt 0) {
    $HTMLGroupsWithNoOwnersSection = "<p><h2>Groups with no owners" + " (" + $GroupsWithNoOwners.count + ")</h2></p>" + ($GroupsWithNoOwners | ConvertTo-Html -Property Group, Id, Created, SecurityEnabled, MailEnabled, GroupTypes, TotalMembers, TotalOwners, Owners -Fragment )
} Else {
    $HTMLGroupsWithNoOwnersSection = "<p><h2>Groups with no owners</h2></p><p>No groups without owners found.</p>"
}
$HTMLBody = $HTMLHead + $HTMLGroupsWithNoOwnersSection

# Groups with service principals as owners
If ($GroupsWithServicePrincipalOwners.count -gt 0) {
    $HTMLGroupsWithServicePrincipalOwnersSection = "<p><h2>Groups with service principals as owners" + " (" + $GroupsWithServicePrincipalOwners.count + ")</h2></p>" + ($GroupsWithServicePrincipalOwners | ConvertTo-Html -Property Group, Id, Created, SecurityEnabled, MailEnabled, GroupTypes, TotalMembers, TotalOwners, Owners -Fragment )
} Else {
    $HTMLGroupsWithServicePrincipalOwnersSection = "<p><h2>Groups with service principals as owners</h2></p><p>No groups with service principals as owners found.</p>"
}   
$HTMLBody = $HTMLBody + $HTMLGroupsWithServicePrincipalOwnersSection

# Groups with service principals as owners or members
If ($GroupsWithServicePrincipalsAsOwnersOrMembers.count -gt 0) {
    $HTMLGroupsWithServicePrincipalsAsOwnersOrMembersSection = "<p><h2>Groups with service principals as owners or members" + " (" + $GroupsWithServicePrincipalsAsOwnersOrMembers.count + ")</h2></p>" + ($GroupsWithServicePrincipalsAsOwnersOrMembers | ConvertTo-Html -Property Group, Id, Created, SecurityEnabled, MailEnabled, GroupTypes, TotalMembers, TotalOwners, Owners -Fragment )
} Else {
    $HTMLGroupsWithServicePrincipalsAsOwnersOrMembersSection = "<p><h2>Groups with service principals as owners or members</h2></p><p>No groups with service principals as owners or members found.</p>"
}
$HTMLBody = $HTMLBody + $HTMLGroupsWithServicePrincipalsAsOwnersOrMembersSection

# Groups with service principals as members but not owners
If ($ServicePrincipalsAsMembers -gt 0) {
    $HTMLGroupsWithServicePrincipalsAsMembersSection = "<p><h2>Groups with service principals as members but not owners" + " (" + $ServicePrincipalsAsMembers + ")</h2></p>" + ($Report | Where-Object { $_.transitiveServicePrincipalCount -gt 0 -and $_.servicePrincipalOwnerCount -eq 0 } | ConvertTo-Html -Property Group, Id, Created, SecurityEnabled, MailEnabled, GroupTypes, TotalMembers, TotalOwners, Owners -Fragment )
} Else {        
    $HTMLGroupsWithServicePrincipalsAsMembersSection = "<p><h2>Groups with service principals as members but not owners</h2></p><p>No groups with service principals as members but not owners found.</p>"
}
$HTMLBody = $HTMLBody + $HTMLGroupsWithServicePrincipalsAsMembersSection

# Groups with guest owners
If ($GroupsWithGuestsAsOwners.count -gt 0) {
    $HTMLGroupsWithGuestsAsOwnersSection = "<p><h2>Groups with guest owners" + " (" + $GroupsWithGuestsAsOwners.count + ")</h2></p>" + ($GroupsWithGuestsAsOwners | ConvertTo-Html -Property Group, Id, Created, SecurityEnabled, MailEnabled, GroupTypes, TotalMembers, TotalOwners, Owners -Fragment )
} Else {
    $HTMLGroupsWithGuestsAsOwnersSection = "<p><h2>Groups with guest owners</h2></p><p>No groups with guest owners found.</p>"
}
$HTMLBody = $HTMLBody + $HTMLGroupsWithGuestsAsOwnersSection

# Groups with dynamic membership       
If ($GroupsWithDynamicMembership.count -gt 0) {
    $HTMLGroupsWithDynamicMembershipSection = "<p><h2>Groups with dynamic membership" + " (" + $GroupsWithDynamicMembership.count + ")</h2></p>" + ($GroupsWithDynamicMembership | ConvertTo-Html -Property Group, Id, Created, SecurityEnabled, MailEnabled, GroupTypes, TotalMembers, TotalOwners, Owners -Fragment )
} Else {
    $HTMLGroupsWithDynamicMembershipSection = "<p><h2>Groups with dynamic membership</h2></p><p>No groups with dynamic membership found.</p>"
}
$HTMLBody = $HTMLBody + $HTMLGroupsWithDynamicMembershipSection

# Dynamic groups with complicated rules
If ($GroupsComplicatedRules.count -gt 0) {
    $HTMLGroupsWithComplicatedRulesSection = "<p><h2>Dynamic groups with complicated membership rules" + " (" + $GroupsComplicatedRules.count + ")</h2></p>" + ($GroupsComplicatedRules | ConvertTo-Html -Property Group, Id, Created, MembershipRuleExpressionCount, SecurityEnabled, MailEnabled, GroupTypes, TotalMembers, TotalOwners, Owners -Fragment )
} Else {
    $HTMLGroupsWithComplicatedRulesSection = "<p><h2>Dynamic groups with complicated membership rules</h2></p><p>No dynamic groups with complicated membership rules found.</p>"
}   
$HTMLBody = $HTMLBody + $HTMLGroupsWithComplicatedRulesSection

# Dynamic groups with inefficient rules
If ($GroupInefficientRules.count -gt 0) {
    $HTMLGroupsWithInefficientRulesSection = "<p><h2>Dynamic groups with inefficient membership rules" + " (" + $GroupInefficientRules.count + ")</h2></p>" + ($GroupInefficientRules | ConvertTo-Html -Property Group, Id, Created, MembershipRuleMatchCount, MembershipRuleContainsCount, SecurityEnabled, MailEnabled, GroupTypes, TotalMembers, TotalOwners, Owners -Fragment )
} Else {
    $HTMLGroupsWithInefficientRulesSection = "<p><h2>Dynamic groups with inefficient membership rules</h2></p><p>No dynamic groups with inefficient membership rules found.</p>"
}
$HTMLBody = $HTMLBody + $HTMLGroupsWithInefficientRulesSection

# Groups with pending expiration in the next 30 days
If ($GroupsWithPendingExpiration.count -gt 0) {
    $HTMLGroupsWithPendingExpirationSection = "<p><h2>Groups with pending expiration in the next 30 days" + " (" + $GroupsWithPendingExpiration.count + ")</h2></p>" + ($GroupsWithPendingExpiration | ConvertTo-Html -Property Group, Id, Created, expirationDateTime, SecurityEnabled, MailEnabled, GroupTypes, TotalMembers, TotalOwners, Owners -Fragment )
} Else {
    $HTMLGroupsWithPendingExpirationSection = "<p><h2>Groups with pending expiration in the next 30 days</h2></p><p>No groups with pending expiration in the next 30 days found.</p>"
}
$HTMLBody = $HTMLBody + $HTMLGroupsWithPendingExpirationSection

# Groups in soft deletion status
If ($GroupsInSoftDeletion.count -gt 0) {
    $HTMLGroupsInSoftDeletionSection = "<p><h2>Groups in soft deletion" + " (" + $GroupsInSoftDeletion.count + ")</h2></p>" + ($GroupsInSoftDeletion | ConvertTo-Html -Property Group, Id, Created, softDeletionDateTime, SecurityEnabled, MailEnabled, GroupTypes, TotalMembers, TotalOwners, Owners -Fragment )
} Else {
    $HTMLGroupsInSoftDeletionSection = "<p><h2>Groups in soft deletion</h2></p><p>No groups in soft deletion found.</p>"
}
$HTMLBody = $HTMLBody + $HTMLGroupsInSoftDeletionSection

# Groups restored in the last 30 days
If ($GroupsRestored.count -gt 0) {
    $HTMLGroupsRestoredSection = "<p><h2>Groups restored in the last 30 days" + " (" + $GroupsRestored.count + ")</h2></p>" + ($GroupsRestored | ConvertTo-Html -Property Group, Id, Created, lastRestorationDateTime, SecurityEnabled, MailEnabled, GroupTypes, TotalMembers, TotalOwners, Owners -Fragment )
} Else {
    $HTMLGroupsRestoredSection = "<p><h2>Groups restored in the last 30 days</h2></p><p>No groups restored in the last 30 days found.</p>"
}
$HTMLBody = $HTMLBody + $HTMLGroupsRestoredSection

# Newly created groups in the last 30 days
If ($NewlyCreatedGroups.count -gt 0) {
    $HTMLNewlyCreatedGroupsSection = "<p><h2>Newly created groups in the last 30 days" + " (" + $NewlyCreatedGroups.count + ")</h2></p>" + ($NewlyCreatedGroups | ConvertTo-Html -Property Group, Id, Created, SecurityEnabled, MailEnabled, GroupTypes, TotalMembers, TotalOwners, Owners -Fragment )
} Else {
    $HTMLNewlyCreatedGroupsSection = "<p><h2>Newly created groups in the last 30 days</h2></p><p>No newly created groups in the last 30 days found.</p>"
}   
$HTMLBody = $HTMLBody + $HTMLNewlyCreatedGroupsSection

If ($GroupsWithNoSensitivityLabel.count -gt 0) {
    $HTMLGroupsWithNoSensitivityLabelSection = "<p><h2>Groups with no sensitivity label" + " (" + $GroupsWithNoSensitivityLabel.count + ")</h2></p>" + ($GroupsWithNoSensitivityLabel | ConvertTo-Html -Property Group, Id, Created, sensitivityLabelCount, SecurityEnabled, MailEnabled, GroupTypes, TotalMembers, TotalOwners, Owners -Fragment )
} Else {
    $HTMLGroupsWithNoSensitivityLabelSection = "<p><h2>Groups with no sensitivity label</h2></p><p>No groups with no sensitivity label found.</p>"
}   
$HTMLBody = $HTMLBody + $HTMLGroupsWithNoSensitivityLabelSection


$HTMLTail = "<p><h2>Summary of Entra ID Groups Insights created for: " + $Organization + "</h2></p>" +
    "<p>Created: " + $CreationDate + " Entra insights data from " + $InsightsDate + "</p>" +
    "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
    "<p>Total groups found:                                                 " + $Report.Count + "</p>" +
    "<p>Number of groups without a manager:                                 " + $GroupsWithNoOwners.count+ "</p>" +
    "<p>Number of groups with no members:                                   " + ($Report | Where-Object { $_.TotalMembers -eq 0 }).Count + "</p>" +
    "<p>NUmber of groups with service principals as owners:                 " + $GroupsWithServicePrincipalOwners.count + "</p>" +
    "<p>Number of groups with service principals as owners or members:      " + $GroupsWithServicePrincipalsAsOwnersOrMembers.count + "</p>" +
    "<p>Number of groups with service principals as members but not owners: " + $ServicePrincipalsAsMembers + "</p>" +
    "<p>Number of groups with guest owners:                                 " + $GroupsWithGuestsAsOwners.count + "</p>" +
    "<p>Number of Microsoft 365 Groups:                                     " + $M365GroupCount + "</p>" +
    "<p>Number of groups with dynamic membership:                           " + $GroupsWithDynamicMembership.count + "</p>" +
    "<p>Number of dynamic Microsoft 365 Groups:                             " + $DynamicMicrosoft365GroupCount + "</p>" +
    "<p>Number of groups pending expiration in next 30 days:                " + $GroupsWithPendingExpiration.Count + "</p>" +
    "<p>Number of soft-deleted groupss:                                     " + $GroupsInSoftDeletion.Count + "</p>" +
    "<p>Number of groups restored in the last 30 days:                      " + $GroupsRestored.Count + "</p>" +
    "<p>Number of groups created in the last 30 days:                       " + $NewlyCreatedGroups.Count + "</p>" +
    "<p>Number of groups with no sensitivity label:                         " + $GroupsWithNoSensitivityLabel.Count + "</p>" +
    "<p>Number of dynamic groups with complicated membership rules:         " + $GroupsComplicatedRules.Count + "</p>" +
    "<p>Number of dynamic groups with inefficient membership rules:         " + $GroupInefficientRules.Count + "</p>" + 
    "<p>Number of cloud distribution lists:                                 " + $CloudDLCount + "</p>" +
    "<p>Number of on-premises distribution lists:                           " + $OnPremisesDLCount + "</p>" +
    "<p>Number of cloud security groups:                                    " + $CloudSecurityGroupCount + "</p>" +    
    "<p>Number of on-premises security groups:                              " + $OnPremisesSecurityGroupCount + "</p>" +    
    "<p>Number of cloud mail-enabled security groups:                       " + $CloudMailEnabledSecurityGroupCount + "</p>" +
    "<p>Number of on-premises mail-enabled security groups:                 " + $OnPremisesMailEnabledSecurityGroupCount + "</p>" +
    "<p>The largest group is <b>" + $LargestGroup.Group + "</b> with <b>" + $LargestGroup.TotalMembership + "</b> members.</p>" +
    "<p>The group with the largest number of guest accounts is <b>" + $LargestGroupWithGuests.Group + "</b> with <b>" + $LargestGroupWithGuests.TotalMembership + "</b> members.</p>" +
    "<p>The group with the largest number of member accounts is <b>" + $LargestGroupWithMembers.Group + "</b> with <b>" + $LargestGroupWithMembers.TotalMembership + "</b> members.</p>" +
    "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
    "<p>Entra ID Group Insights<b> " + $Version + "</b>"	

$HTMLReport = $HTMLBody + $HTMLTail
$HTMLReportFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Entra ID Groups Insights.html"
$HTMLReport | Out-File $HTMLReportFile -Encoding UTF8

# Generate the report in either Excel worksheet or CSV format, d    epending on if the ImportExcel module is available
If (Get-Module ImportExcel -ListAvailable) {
   $ExcelGenerated = $True
   Import-Module ImportExcel -ErrorAction SilentlyContinue
   $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Entra Group Insights.xlsx"
   $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Entra Group Insights" -Title ("Entra Group Insights {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "EntraGroupInsights" -AutoSize -AutoFilter
} Else {
   $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Entra Group Insights.CSV"
   $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}   

If ($ExcelGenerated -eq $true) {
   Write-Host ("Entra Group Insights report is available in Excel workbook {0}" -f $ExcelOutputFile) -ForegroundColor Green
} Else {
   Write-Host ("Entra Group Insights report is available in CSV file {0}" -f $CSVOutputFile) -ForegroundColor Green
}
Write-Host ("HTML report is available in file {0}" -f $HTMLReportFile) -ForegroundColor Green

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.