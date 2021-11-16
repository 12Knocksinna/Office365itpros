# ReportExpiringLinkExtended.PS1
# Check modules - Exchange Online used for Search-UnifiedAuditLog. Azure AD is to check the guest account and report a display name
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
If (!($ModulesLoaded -match "AzureAD*")) {Write-Host "Please connect to the Azure AD module and then restart the script"; break}
# Search for the last 90 days
$StartDate = (Get-Date).AddDays(-90); $EndDate = (Get-Date) 
[array]$Records = Search-UnifiedAuditLog -Operations UserExpirationChanged, SharingInvitationCreated, SharingSet, SecureLinkCreated, SecureLinkUsed -StartDate $StartDate -EndDate $EndDate -Formatted -ResultSize 5000
# If we find some records, process them
If (!$Records) {  Write-Host "No audit records for extending sharing links found."; break }
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
# Process the records
ForEach ($Rec in $Records) {
    $DisplayName = $Null
    $AuditData = $Rec.AuditData | ConvertFrom-Json
    Switch ($Rec.Operations) {
        "SecureLinkUsed" {
            $Target = $AuditData.SourceFileName
            $DisplayName = "N/A"
        }
        "SecureLinkCreated" {
            $Target = $AuditData.SourceFileName
            $DisplayName = "N/A"
        }
        "SharingSet" {
            $Target = $AuditData.SourceFileName
            $DisplayName = "N/A"
        }
        "SharingInvitationCreated" {
            $Target = $AuditData.SourceFileName
            If ($AuditData.TargetUserOrGroupType -eq "Guest") {
                $DisplayName = (Get-AzureADUser -ObjectId $AuditData.TargetUserOrGroupName).DisplayName }
        }
        "UserExpirationChanged" {
            If ($AuditData.TargetUserOrGroupType -eq "Guest") {
                $DisplayName = (Get-AzureADUser -ObjectId $AuditData.TargetUserOrGroupName).DisplayName }
            $Target = $AuditData.TargetUserOrGroupName
            }
    }
        
        $ReportLine = [PSCustomObject] @{
        TimeStamp   = $Rec.CreationDate
        UPN         = $Rec.UserIds
        Name        = $DisplayName
        Action      = $AuditData.Operation
        Source      = $AuditData.EventSource
        Target      = $Target
        Type        = $AuditData.TargetUserOrGroupType
        Site        = $AuditData.SiteUrl 
        Correlation = $auditData.CorrelationId}        
   $Report.Add($ReportLine) 
} #End ForEach Records

$Report = $Report | Sort-Object {$_.TimeStamp -as [datetime]}, Target
$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation c:\temp\SPOSharingEvents.CSV
