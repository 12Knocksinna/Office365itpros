# Report-EntraGroupInsights.PS1
# A script to show how to extract the Entra Groups insights from the Graph to report them in a readable format.

Connect-MgGraph -Scopes "Group.Read.All", "GroupMember.Read.All", "User.Read.All"
Write-Host "Starting up by finding groups and insights..." -ForegroundColor Green

[array]$Groups = Get-MgGroup -All -Property "id,displayName,description,createdDateTime,securityEnabled,mailEnabled,groupTypes,owners" -PageSize 500 -ExpandProperty Owners
If ($Groups) {
    Write-Host ("{0} groups retrieved successfully." -f $Groups.count)   -ForegroundColor Green
} Else {
    Write-Host "Failed to retrieve groups." -ForegroundColor Red
    Break
}   

$Uri = "https://graph.microsoft.com/beta/reports/identityAnalytics/groups"
[array]$GroupInsights = Invoke-MgGraphRequest -Method Get -Uri $Uri -OutputType PSObject | Select-Object -ExpandProperty value

If ($GroupInsights) {
    Write-Host "Group insights retrieved successfully." -ForegroundColor Green
} Else {
    Write-Host "Failed to retrieve group insights." -ForegroundColor Red
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
            TotalMembers                   = $Insight.directGroupMemberCount
            TotalOwners                    = $Group.owners.Count
            Owners                         = $Group.owners.additionalProperties.displayName -join ", "
            TotalGuests                    = $Insight.totalGuests
            GuestOwnerCount                = $Insight.guestOwnerCount
            servicePrincipalOwnerCount     = $Insight.servicePrincipalOwnerCount
            transitiveServicePrincipalCount = $Insight.transitiveServicePrincipalCount
            DynamicMembership              = $DynamicGroup
            MembershipRuleProcessingState  = $Insight.membershipRuleProcessingState
            MembershipRuleExpressionCount  = $Insight.membershipRuleExpressionCount
            membershipRuleContainsCount    = $Insight.membershipRuleContainsCount
            membershipRuleMatchCount       = $Insight.membershipRuleMatchCount
            sensitivityLabelCount          = $Insight.sensitivityLabelCount
            Grouptype                      = $Insight.groupType 
            Visibility                     = $Group.visibility
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
            TotalMembers               = $Insight.directGroupMemberCount
            TotalOwners                = $Group.owners.Count
            Owners                     = $Group.owners.additionalProperties.displayName -join ", "
        } 
        $Report.Add($ReportLine) 
    }

}

$Report = $Report | Sort-Object -Property Group -Descending

# Now we can generate some analytics from the insights

[datetime]$30DaysAgo = (Get-Date).AddDays(-30)
[datetime]$30DaysFromNow = (Get-Date).AddDays(30)

[array]$GroupsWithNoOwners = $Report | Where-Object { $_.TotalOwners -eq 0 }
Write-Host ("Groups with no owners: {0}" -f $GroupsWithNoOwners.count) -ForegroundColor Yellow
[array]$GroupsWithServicePrincipalOwners = $Report | Where-Object { $_.servicePrincipalOwnerCount -gt 0 }
Write-Host ("Groups with service principals as owners: {0}" -f $GroupsWithServicePrincipalOwners.count) -ForegroundColor Yellow
[array]$GroupsWithServicePrincipalsAsOwnersOrMembers = $Report | Where-Object { $_.servicePrincipalOwnerCount -gt 0 -or $_.transitiveServicePrincipalCount -gt 0 }
Write-Host ("Groups with service principals as owners or members: {0}" -f $GroupsWithServicePrincipalsAsOwnersOrMembers.count) -ForegroundColor Yellow
[array]$GroupsWithGuestsAsOwners = $Report | Where-Object { $_.GuestOwnerCount -gt 0 }

[int]$ServicePrincipalsAsMembers =  $GroupsWithServicePrincipalsAsOwnersOrMembers.count - $GroupsWithServicePrincipalOwners.count
Write-Host ("Groups with service principals as members but not owners: {0}" -f $ServicePrincipalsAsMembers) -ForegroundColor Yellow

Write-Host ("Groups with guest owners: {0}" -f $GroupsWithGuestsAsOwners.count) -ForegroundColor Yellow
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

# OK. Output the report 
