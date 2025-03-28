# Report-RetentionLabelAssignments.PS1
# Report on retention label assignments in a tenant for the last 30 days
# V1.0 27-Mar-2025

# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-RetentionLabelAssignments.PS1

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ($Modules -notcontains "ExchangeOnlineManagement") {
    Connect-ExchangeOnline -ShowBanner:$false
}

Write-Host "Connecting to the Compliance endpoint to fetch retention policy information..."
Connect-IPPSSession -ShowBanner:$false
# Get retention policy information
[array]$RetentionPolicies = Get-RetentionCompliancePolicy | Sort-Object Name
$RetentionPolicyHash = @{}
ForEach ($Policy in $RetentionPolicies) {
    $RetentionPolicyHash.Add([string]$Policy.Guid, $Policy.Name)
}
# Get retention policy rule information
[array]$RetentionPolicyRules = Get-RetentionComplianceRule
$RetentionRuleHash = @{}
ForEach ($Rule in $RetentionPolicyRules) {
    $RetentionRuleHash.Add([string]$Rule.Guid, [string]$Rule.Policy)
}

[array]$SharePointAutoLabelId = 'eba15bfd-c28e-4433-a20e-0278888c5825', 'a405a596-28e7-43c6-8ac7-76e0fc13ee0f'
$StartDate = (Get-Date).AddDays(-30)
$EndDate = Get-Date
# Search for TagAdded events
Write-Output "Searching for TagApplied events..."
[array]$Records = Search-UnifiedAuditLog -Operations TagApplied -StartDate $StartDate -EndDate $EndDate -Formatted -SessionCommand ReturnLargeSet -ResultSize 5000

If ($Records.Count -eq 0) {
    Write-Output "No TagApplied events found in the last 30 days"
    Break
}   Else {
    # Get rid of duplicates and make sure that audit records are sorted by date
    $Records = $Records | Sort-Object Identity -Unique
    $Records = $Records | Sort-Object {$_.CreationDate -as [datetime]} -Descending
    Write-Output ("Processing {0} retention label assignment events..." -f $Records.Count)
}

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
    $UserName = $null
    $AuditData = $Rec.AuditData | ConvertFrom-Json

    If ($AuditData.UserType -eq "CustomPolicy") {
        $RuleId = $Rec.UserIds
        If ($RuleId -in $SharePointAutoLabelId) {
            $UserName = "SharePoint Auto Label Process"
        } Else {
            $PolicyId = $RetentionRuleHash[$RuleId]
            If ($null -eq $PolicyId) {
                $UserName = "Unknown"
            } Else {
                $UserName = $RetentionPolicyHash[$PolicyId]
            }
        }
    } Else {
        $UserName = $Rec.UserIds
    }
  
    $ReportLine = [PSCustomObject][Ordered]@{
        TimeStamp   = Get-Date $Rec.CreationDate -format 'dd-MMM-yyyy HH:mm:ss'
        User        = $UserName
        Assignment  = $AuditData.UserType
        File        = $AuditData.SourceFileName
        Label       = $AuditData.DestinationLabel
        SiteUrl     = $AuditData.SiteUrl
        Folder      = $AuditData.SourceRelativeUrl
    }
    $Report.Add($ReportLine)
}

Write-Host "All done!"
$Report | Out-Gridview -Title 'Retention Label Assignment events'

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.