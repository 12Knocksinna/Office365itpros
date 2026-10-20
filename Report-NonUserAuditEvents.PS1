# Report-NonUserAuditEvents.PS1
# This script generates a report of non-user audit events logged for a Microsoft 365 tenant over the last 7 days
# Two sets of code are in the script. The first uses the Search-UnifiedAuditLog cmdlet to fetch audit events, and the second uses 
# the AuditLogQuery Graph API to do the same. Both methods are included for comparison and learning purposes.

# V1.0 19-Oct-2026

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-NonUserAuditEvents.PS1

Connect-ExchangeOnline -ShowBanner:$false

[array]$BackEndSIDS = 'S-1-8-618832045-1217504451-1955385509-1851076329', 'S-1-8-618832045-1217504451-1955385509-1851076329'
[array]$NotFound = 'NOT-FOUND', 'Not Available'

$WellKnownApps = @{}
$WellKnownApps.Add('0d8921ab-7749-4a55-846e-67f0fef6b953', 'Entra ID Strong Authentication Service')
$WellKnownApps.Add('1169fdeb-9f54-484f-898a-5b84bc47c751', 'Entra ID Terms of Use Service')
$WellKnownApps.Add('181a9527-b9c9-41b6-b374-8928ee2ad9af', 'Groups Configuration Service')
$WellKnownApps.Add('9008b935-c511-4231-b84c-8ac6d4920f4d', 'Office 365 Admin Portal')
$WellKnownApps.Add('b251bbb7-4e3d-416c-8f3d-c065d378f909', 'Device Configuration Service')
$WellKnownApps.Add('dacf6086-a190-467a-aadd-d519472b8d1d', 'Exchange Online')
$WellKnownApps.Add('e6ff64fa-aad6-4944-8e6c-c746c7b613a6', 'Microsoft 365 Substrate Management')
$WellKnownApps.Add('1342cefb-7a89-4ee2-af90-c8443053e1e8', 'Microsoft Approval Management')

$MbxHash = [ordered]@{}
Write-Output "Finding all user mailboxes in the tenant to help identify non-user audit events..."
# Get the set of user mailboxes in the tenant to help identify non-user audit events and populate a hash table for later use
[array]$Mbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | `
    Select-Object DisplayName, userPrincipalName | Sort-Object userPrincipalName
$Mbx | ForEach-Object { $MbxHash.Add($_.userPrincipalName.ToString().ToLower(), $_.DisplayName) }

Write-Output "Starting search for non-user audit events for the last 7 days..."
# Example of fetching all audit events for the last week (well, limited to 5000 records)
$Records = @()
$SessionId = [guid]::NewGuid()
[int]$i = 0
# Search for the last seven days
$StartDate = (Get-Date).AddDays(-7)
$EndDate = Get-Date

Do {
    $i++
    Write-Output ("Fetching page {0} of audit records" -f $i)
    [array]$PageRecords = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -SessionId $SessionId `
         -SessionCommand ReturnLargeSet -ResultSize 5000 -Formatted
    If ($PageRecords) {
        # Add the set to the overall array of audit records fetched by the search
        $Records += $PageRecords
        Write-Output ("{0} audit records found so far…" -f $Records.count)
        # Any more records?
        $More =  ($PageRecords[0].AuditSearchRequestMetadata | ConvertFrom-Json).moreRecordsAvailable
    } Else {
        $More = $false
    }
} While ($More -eq $true -and $Records.count -le 50000)

# Make sure records are deduplicated and sorted in date order
$Records = $Records | Sort-Object Identity -Unique | Sort-Object {$_.CreationDate -as [datetime]}

Write-Output ("All done {0} records found in {1} minutes." -f $Records.Count, [math]::Round((New-TimeSpan -Start $EndDate).TotalMinutes,3))

#[array]$Records = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date).AddDays(1) -ResultSize 5000 -Formatted -SessionCommand ReturnLargeSet
#$Records = $Records | Sort-Object identity -Unique

#If (!$Records) {
#    Write-Host "No audit records found for the last 7 days"
#    Return
#} Else {
#    $EndTimeSearch = Get-Date
#    $SearchDuration = $EndTimeSearch - $StartTimeSearch     
#    Write-Host "Found $($Records.Count) audit records for the last 7 days"
#    Write-Host "Search duration: $SearchDuration"
#}

# Remove any records that have a user principal name that is in the mailbox hash table (i.e. a user mailbox)
$Records = $Records | Where-Object {$_.UserIds.tolower() -notin $MbxHash.Keys}

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {

    # Drop guest user activities
    If ($Rec.UserIds -Like  "*#ext#*") {
        Continue
    }
    $SystemAppName = $Null; $AppOperation = $Null; $ClientApp = $Null

    $AuditData = $Rec.AuditData | ConvertFrom-Json

    If ($Rec.Operations -eq 'ApplicableAdaptivePolicyChange') {
        $SystemAppName = 'Microsoft Purview Adaptive Policy Processing Service'
        $AppOperation = 'Adaptive Policy Change'
    }

    If ($Rec.UserIds -in $BackEndSIDS ) {
        $SystemAppName = 'Microsoft Service (Internal / Backend)'
        $AppOperation = 'Back-end Service Operation'
    }

    If ($Rec.UserIds -eq 'Certificate') {
        $SystemAppName = 'Microsoft Entra ID'
        $AppOperation = 'JIT Provisioning Operation'
    }

    If ($Rec.UserIds -eq 'UpdateGroupPropertiesFromTenantSensitivityLabelsComplianceAssistant') {
        $SystemAppName = 'Microsoft Group Assistant'
        $AppOperation = 'Update Group Properties from Tenant Sensitivity Labels'
    }

    If ($Rec.UserIds -eq 'Microsoft Teams Sync') {
        $SystemAppName = 'Microsoft Teams'
        $AppOperation = 'Microsoft Teams Sync Operation'
    }

    If ($Rec.UserIds -eq 'SecurityComplianceAlerts') {
        $SystemAppName = 'Microsoft Purview'
        $AppOperation = 'Security and Compliance Alert'
    }

    If ($Rec.UserIds -eq 'SubmissionTriageJob') {
        $SystemAppName = 'Microsoft 365 Admin'
        $AppOperation = 'Grading Pipeline'
    }

    If ($Rec.UserIds -eq 'UploadDefaultPhotoComplianceAssistant') {
        $SystemAppName = 'Microsoft Entra ID'
        $AppOperation = 'Update Photo for Microsoft 365 Groups'
    }

    If ($Rec.UserIds -like "*NT SERVICE\MSExchange*") {
        $SystemAppName = 'Microsoft Exchange Online'
        $AppOperation = 'Exchange Online Admin Operation'
    }

    If ($Rec.UserIds -eq 'Admin' -and $Rec.RecordType -eq 'InformationBarrierPolicyApplication') {
        $SystemAppName = 'Microsoft Purview'
        $AppOperation = 'Information Barrier Policy Application'
    }

    # SharePoint Online administrative operations
    If ($Rec.UserIds -eq 'app@sharepoint' -or $Rec.UserIds -eq "SHAREPOINT\System") {
        $SystemAppName = ' SharePoint Online'
        $AppOperation = 'SharePoint Administrative Operation'
    }

    If ($Rec.UserIds -eq 'ThreatIntel') {
        $SystemAppName = 'Threat Intelligence'
        $AppOperation = 'Threat Intelligence Update'
    }
  
    # Some events have an odd Not found or Not available user Id
    If ($Rec.UserIds -in $NotFound) {
        $SystemAppName = 'Unknown App'
        $AppOperation = 'Unknown System Operation'  

    }

    # A service principal
    If ($Rec.UserIds -like '*ServicePrincipal_*') {
        $SPId = $Rec.UserIds.Substring(17,36)
        If ($WellKnownApps.ContainsKey($SPId)) {
            $SystemAppName = $WellKnownApps[$SPId]
            $AppOperation = 'Well-known Microsoft App Operation'
        }
        Else {
             $SystemAppName = 'Unknown Service Principal'
             $AppOperation = 'Unknown Service Principal Operation'
        }
    }

    IF ($Rec.Operations -eq 'IdCrlBlockedDueToSoftEnforcement') {
        $SystemAppName = 'SharePoint Online'
        $AppOperation = 'Block against IDCRL'   
    }

    $ReportLine = [PSCustomObject] @{
        Timestamp               = $Rec.CreationDate
        UserIds                 = $Rec.UserIds
        SystemAppName           = $SystemAppName
        'App Operation'         = $AppOperation
        Operations              = $Rec.Operations
        Workload                = $Rec.RecordType
        GuestAccountFlag        = $GuestAccountFlag
        ClientApp               = $AuditData.ClientProcessName
    }
    $Report.Add($ReportLine)
} 

$Report= $Report | Sort-Object {$_.timestamp -as [datetime]}
Write-Output "Found $($Report.Count) non-user audit events for the last 7 days"
$Report | Out-GridView -Title "Non-user audit events for the last week"

# Steps to do the same with the AuditLogQuery Graph API.

# Step 1 - Connect to Microsoft Graph with the required permissions to read audit logs. This is necessary to access the audit data for non-user events.
Connect-MgGraph -Scopes AuditLogsQuery.Read.All -NoWelcome

# Create the same kind of hash table holding details of user mailboxes to help identify non-user audit events. 
# This is done by fetching all users with an Exchange Online license assigned and enabled. This isn't necessary to run audit jobs
# but we need the data to filter out user mailbox events from the audit log results. 
$ExoServicePlan1 = "9aaf7827-d63c-4b61-89c3-182f06f82e5c"
$ExoServicePlan2 = "efb87545-963c-4e0d-99df-69c6916d9eb0" 

$MbxHash = [ordered]@{}
Write-Output "Finding all user mailboxes in the tenant to help identify non-user audit events..."
[array]$Mbx = Get-MgUser -Filter "assignedPlans/any(c:c/servicePlanId eq $ExoServicePlan1 and capabilityStatus eq 'Enabled') or assignedPlans/any(c:c/servicePlanId eq $ExoServicePlan2 and capabilityStatus eq 'Enabled') and accountEnabled eq true" `
   -ConsistencyLevel eventual -CountVariable Test -All -PageSize 500  -Property DisplayName, UserPrincipalName -Sort UserPrincipalName 
If (!$Mbx) {
    Write-Output "No user mailboxes found in the tenant. Exiting script."
    Return
} Else {
    Write-Output ("Found {0} user mailboxes in the tenant" -f $Mbx.Count)
}   
$Mbx | ForEach-Object { $MbxHash.Add($_.UserPrincipalName.ToString().ToLower(), $_.DisplayName) }

# Step 2: Create a hash table for the request body for the audit log query job.
$AuditJobName = ("Audit job to fetch all events for the last week created at {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm'))
$EndDate = (Get-Date).AddHours(1)
$StartDate = (Get-Date $EndDate).AddDays(-7)
$AuditQueryStart = (Get-Date $StartDate -format s)
$AuditQueryEnd = (Get-Date $EndDate -format s)

$AuditQueryParameters = @{}
$AuditQueryParameters.Add("@odata.type","#microsoft.graph.security.auditLogQuery")
$AuditQueryParameters.Add("displayName", $AuditJobName)
$AuditQueryParameters.Add("filterStartDateTime", $AuditQueryStart)
$AuditQueryParameters.Add("filterEndDateTime", $AuditQueryEnd)

# Step 3: Submit the audit log query job.
$AuditJob = New-MgSecurityAuditLogQuery -BodyParameter $AuditQueryParameters
If (!$AuditJob) {
    Write-Output "Failed to create audit log query job. Exiting script."
    Return
} Else {
    Write-Output ("Audit log query job created successfully with ID: {0}" -f $AuditJob.Id)
    $AuditJobDate = (Get-Date)
}

# Step 4. Monitor the audit job until it comepletes. This may take some time depending on the volume of audit data.
[int]$i = 1
[int]$SleepSeconds = 20
$SearchFinished = $false; [int]$SecondsElapsed = 20
Write-Host "Checking audit query status..."
# Initial wait to let the audit job spin up
Start-Sleep -Seconds 30
$AuditQueryStatus = Get-MgSecurityAuditLogQuery -AuditLogQueryId $AuditJob.Id

While ($SearchFinished -eq $false) {
   $i++
   Write-Host ("Waiting for audit search to complete. Check {0} after {1} seconds. Current state {2}" -f $i, $SecondsElapsed, $AuditQueryStatus.status)
   If ($AuditQueryStatus.status -eq 'succeeded') {
      $SearchFinished = $true
   } Else {
      Start-Sleep -Seconds $SleepSeconds
      $SecondsElapsed = $SecondsElapsed + $SleepSeconds
      $AuditQueryStatus =  Get-MgSecurityAuditLogQuery -AuditLogQueryId $AuditJob.Id    
   }
}

# Step 5. Fetch the results of the audit log query job. Use the -All parameter to fetch all results if the job has multiple pages of results.
[array]$AuditRecords = Get-MgSecurityAuditLogQueryRecord -All -AuditLogQueryId $AuditJob.Id

Write-Output ("All done {0} records found in {1} minutes." -f $AuditRecords.Count, [math]::Round((New-TimeSpan -Start $AuditJobDate).TotalMinutes,3))

# Step 6. Process the audit records to identify non-user events and generate a report. This involves filtering out user mailboxes and identifying system applications and operations.
$AuditRecords = $AuditRecords | Where-Object {$_.UserPrincipalName.ToLower() -notin $MbxHash.Keys}

$AuditJobReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $AuditRecords) {

    # Drop guest user activities
    If ($Rec.UserPrincipalName -Like  "*#ext#*") {
        Continue
    }
    $SystemAppName = $Null; $AppOperation = $Null; $ClientApp = $Null

    $AuditData = $Rec.auditData.additionalProperties

    If ($Rec.Operation -eq 'ApplicableAdaptivePolicyChange') {
        $SystemAppName = 'Microsoft Purview Adaptive Policy Processing Service'
        $AppOperation = 'Adaptive Policy Change'
    }

    If ($Rec.UserPrincipalName -in $BackEndSIDS ) {
        $SystemAppName = 'Microsoft Service (Internal / Backend)'
        $AppOperation = 'Back-end Service Operation'
    }

    If ($Rec.UserPrincipalName -eq 'Certificate') {
        $SystemAppName = 'Microsoft Entra ID'
        $AppOperation = 'JIT Provisioning Operation'
    }

    If ($Rec.UserPrincipalName -eq 'UpdateGroupPropertiesFromTenantSensitivityLabelsComplianceAssistant') {
        $SystemAppName = 'Microsoft Group Assistant'
        $AppOperation = 'Update Group Properties from Tenant Sensitivity Labels'
    }

    If ($Rec.UserPrincipalName -eq 'Microsoft Teams Sync') {
        $SystemAppName = 'Microsoft Teams'
        $AppOperation = 'Microsoft Teams Sync Operation'
    }

    If ($Rec.UserPrincipalName -eq 'SecurityComplianceAlerts') {
        $SystemAppName = 'Microsoft Purview'
        $AppOperation = 'Security and Compliance Alert'
    }

    If ($Rec.UserPrincipalName -eq 'SubmissionTriageJob') {
        $SystemAppName = 'Microsoft 365 Admin'
        $AppOperation = 'Grading Pipeline'
    }

    If ($Rec.UserPrincipalName -eq 'UploadDefaultPhotoComplianceAssistant') {
        $SystemAppName = 'Microsoft Entra ID'
        $AppOperation = 'Update Photo for Microsoft 365 Groups'
    }

    If ($Rec.UserPrincipalName -like "*NT SERVICE\MSExchange*") {
        $SystemAppName = 'Microsoft Exchange Online'
        $AppOperation = 'Exchange Online Admin Operation'
    }

    If ($Rec.UserPrincipalName -eq 'Admin' -and $Rec.RecordType -eq 'InformationBarrierPolicyApplication') {
        $SystemAppName = 'Microsoft Purview'
        $AppOperation = 'Information Barrier Policy Application'
    }

    # SharePoint Online administrative operations
    If ($Rec.UserPrincipalName -eq 'app@sharepoint' -or $Rec.UserPrincipalName -eq "SHAREPOINT\System") {
        $SystemAppName = ' SharePoint Online'
        $AppOperation = 'SharePoint Administrative Operation'
    }

    If ($Rec.UserPrincipalName -eq 'ThreatIntel') {
        $SystemAppName = 'Threat Intelligence'
        $AppOperation = 'Threat Intelligence Update'
    }

    # Some events have an odd Not found or Not available user Id
    If ($Rec.UserPrincipalName -in $NotFound) {
        $SystemAppName = 'Unknown App'
        $AppOperation = 'Unknown System Operation'  
    }

    # A service principal
    If ($Rec.UserPrincipalName -like '*ServicePrincipal_*') {
        $SPId = $Rec.UserPrincipalName.Substring(17,36)
        If ($WellKnownApps.ContainsKey($SPId)) {
            $SystemAppName = $WellKnownApps[$SPId]
            $AppOperation = 'Well-known Microsoft App Operation'
        }
        Else {
             $SystemAppName = 'Unknown Service Principal'
             $AppOperation = 'Unknown Service Principal Operation'
        }
    }

    IF ($Rec.Operations -eq 'IdCrlBlockedDueToSoftEnforcement') {
        $SystemAppName = 'SharePoint Online'
        $AppOperation = 'Block against IDCRL'   
        $ClientApp = $AuditData.UserAgent
    }

    $ReportLine = [PSCustomObject] @{
        Timestamp               = $Rec.CreatedDateTime
        UserPrincipalName       = $Rec.UserPrincipalName
        SystemAppName           = $SystemAppName
        'App Operation'         = $AppOperation
        Operations              = $Rec.Operation
        Workload                = $AuditData.Workload
        ClientApp               = $AuditData.ApplicationDisplayName
    }
    $AuditJobReport.Add($ReportLine)
} 

$AuditJobReport = $AuditJobReport | Sort-Object {$_.timestamp -as [datetime]}
Write-Output "Found $($AuditJobReport.Count) non-user audit events for the last 7 days"
$AuditJobReport | Out-GridView -Title "Non-user audit events for the last week (Graph)"

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.