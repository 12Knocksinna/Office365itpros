# Report-EntraIDGovernanceCosts
# Read Entra ID audit logs to find records where guest accounts use ID Governance chargeable features and report same.

# V1.0 22-Jul-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-EntraIDGoveranceCosts.PS1

Connect-MgGraph -Scopes AuditLog.Read.All, User.Read.All -NoWelcome

Write-Host "Fetching audit records for the current month to check usage of Entra ID Governance features"
[string]$FirstDayOfMonth = (Get-Date -Day 1).ToString('yyyy-MM-ddT00:00:00Z')
[array]$AuditRecords = Get-MgAuditLogDirectoryAudit -All -Filter "activityDateTime ge $FirstDayOfMonth and result eq 'success'"

If ($AuditRecords.Count -eq 0) {
    Write-Host "No audit records found for the current month"
    Break
} Else {
    Write-Host ("Found {0} audit records for the current month" -f $AuditRecords.Count)
}   

# Extract the audit records that contain information about the use of Entra ID Governance features.
[array]$GovernanceRecords = $AuditRecords | Where-Object { $_.additionalDetails.key -eq "GovernanceLicenseFeatureUsed"}

If ($GovernanceRecords.Count -eq 0) {
    Write-Host "No governance records found"
    Break
} Else {
    Write-Host ("Found {0} governance records" -f $GovernanceRecords.Count)
}

# Create a report of the users who used Entra ID Governance features
$GovernanceReport = [System.Collections.Generic.List[Object]]::new()    
ForEach ($Record in $GovernanceRecords) {
    $UserDisplayName = $null; $UserEmail = $null; $UserUPN = $null; $UserId = $null

    If  ('"Guest"' -in $Record.TargetResources.ModifiedProperties.NewValue) {
        $UserDisplayName = ($Record.TargetResources.ModifiedProperties | Where-Object {$_.DisplayName -eq "DisplayName"}).NewValue
        $UserEmail = ($Record.TargetResources.ModifiedProperties | Where-Object {$_.DisplayName -eq "Email"}).NewValue 
        $UserUPN = ($Record.TargetResources.ModifiedProperties | Where-Object {$_.DisplayName -eq "PrincipalName"}).NewValue
        $UserId = ($Record.TargetResources.ModifiedProperties | Where-Object {$_.DisplayName -eq "TargetId"}).NewValue

        $UserDisplayName = $UserDisplayName -replace '"', ''
        $UserEmail = $UserEmail -replace '"', ''
        $UserUPN = $UserUPN -replace '"', ''
        $UserId = $UserId -replace '"', ''

        }
        $ReportLine = [PSCustomObject]@{
            Timestamp = $Record.ActivityDateTime
            UserDisplayName = $UserDisplayName
            UserEmail = $UserEmail
            UserUPN = $UserUPN
            UserId = $UserId
            Feature = $Record.Id
            FeatureUsed = $Record.ActivityDisplayName
        }
        $GovernanceReport.Add($ReportLine)
}   

# Reduce the set of records to unique users
[array]$GovernanceUsers = $GovernanceReport | Sort-Object UserId -Unique | Select-Object UserId, UserDisplayName, UserEmail, UserUPN
[float]$TotalCosts = $GovernanceUsers.Count * 0.75 # $0.75 per user per month for Entra ID Governance features
# Format the total cost to 2 decimal places
$TotalCosts = "{0:F2}" -f $TotalCosts

# Tell people what we found
Write-Host ""
Write-Host "The following guest accounts used Entra ID Governance features:"
$GovernanceUsers | Format-Table -AutoSize
Write-Host ""
Write-Host ("The monthly cost of Entra ID Governance features is {0} per user, so the total cost for the current month is `${1}" -f "`$0.75", $TotalCosts)


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.