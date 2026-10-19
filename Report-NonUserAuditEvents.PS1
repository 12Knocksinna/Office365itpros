# Report-NonUserAuditEvents.PS1
# This script generates a report of non-user audit events logged for a Microsoft 365 tenant over the last 7 days

# V1.0 5-May-2026

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-NonUserAuditEvents.PS1

Connect-ExchangeOnline -ShowBanner:$false

[array]$BackEndSIDS = 'S-1-8-618832045-1217504451-1955385509-1851076329', 'S-1-8-618832045-1217504451-1955385509-1851076329'
[array]$NotFound = 'NOT-FOUND', 'Not Available'
[array]$Domains = Get-AcceptedDomain | Select-Object -ExpandProperty DomainName

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
# Get the set of user mailboxes in the tenant to help identify non-user audit events and populate a hash table for later use
[array]$Mbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | `
    Select-Object DisplayName, PrimarySmtpAddress | Sort-Object PrimarySmtpAddress 
$Mbx | ForEach-Object { $MbxHash.Add($_.PrimarySmtpAddress.ToString().ToLower(), $_.DisplayName) }

Write-Output "Starting search for non-user audit events for the last 7 days..."
$StartTimeSearch = Get-Date
# Example of fetching all audit events for the last week (well, limited to 5000 records)
[array]$Records = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date).AddDays(1) -ResultSize 5000 -Formatted -SessionCommand ReturnLargeSet
$Records = $Records | Sort-Object identity -Unique

If (!$Records) {
    Write-Host "No audit records found for the last 7 days"
    Return
} Else {
    $EndTimeSearch = Get-Date
    $SearchDuration = $EndTimeSearch - $StartTimeSearch     
    Write-Host "Found $($Records.Count) audit records for the last 7 days"
    Write-Host "Search duration: $SearchDuration"
}

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records | Where-Object {$_.UserIds -notin $MbxHash.Keys}) {

    # Drop guest user activities
    If ($Rec.UserIds -Like  "*#ext#*") {
        Continue
    }
    $GuestAccountFlag = $False
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
     # Guest Account Flag
    If ($Rec.UserIds -Like  "*#ext#*") {
        $GuestAccountFlag = $True
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
        $ClientApp = $AuditData.UserAgent
    }


    $ReportLine = [PSCustomObject] @{
        Timestamp               = $Rec.CreationDate
        UserIds                 = $Rec.UserIds
        SystemAppName           = $SystemAppName
        'App Operation'         = $AppOperation
        Operations              = $Rec.Operations
        Workload                = $Rec.RecordType
        GuestAccountFlag        = $GuestAccountFlag
        ClientApp               = $ClientApp
    }
    $Report.Add($ReportLine)
} 

$Report= $Report | Sort-Object {$_.timestamp -as [datetime]}
Write-Output "Found $($Report.Count) non-user audit events for the last 7 days"
$Report | Out-GridView -Title "Non-user audit events for the last week"

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.