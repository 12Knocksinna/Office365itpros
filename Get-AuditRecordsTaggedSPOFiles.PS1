# Get-AuditRecordsTaggedSPOFiles.PS1
# Show how to report audit records generated when SharePoint Online files are tagged with a retention label
# V1.0 24 December 2024

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Get-AuditRecordsTaggedSPOFiles.PS1

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Please connect to the Exchange Online Management module and then restart the script" 
    break
}

# Start and end date for the audit scan. By default, we look for 3 days, but you can choose any value you like up to 365 (assuming Office 365 E5)
$StartDate = (Get-Date).AddDays(-3); $EndDate = (Get-Date).AddDays(1)
# AppId for the Microsoft Graph PowerShell SDK
$AppId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'

# Find the audit records
[array]$Records = Search-UnifiedAuditLog -Operations TagApplied -StartDate $StartDate -EndDate $EndDate -Formatted -ResultSize 5000 -SessionCommand ReturnLargeSet

If (!$Records) {
   Write-Host "No audit records found - exiting!"; break
}

# Sort to remove duplicate audit records
$Records = $Records | Sort-Object Identity -Unique

$TaggedFilesReport = [System.Collections.Generic.List[Object]]::new() 	

ForEach ($Rec in $Records) {
    $AuditData = $Rec.AuditData | ConvertFrom-Json
    If ($AuditData.AppAccessContext.ClientAppId -eq $AppId) {
        # Audit record is for a SharePoint Online file tagged with a retention label by the Microsoft Graph PowerShell 
        $AuditReportLine = [PSCustomObject] @{
            Workload            = $AuditData.Workload
            File                = $AuditData.DestinationFileName
            "Retention Label"   = $AuditData.DestinationLabel
            "Tagging Date"      = Get-Date($AuditData.CreationTime) -format 'dd-MMM-yyyy HH:mm:ss'
            Site                = $AuditData.SiteURL
            FullURL             = $AuditData.ObjectId
        }
        $TaggedFilesReport.Add($AuditReportLine)
    }
}

Write-Host ("{0} audit records found for files tagged with a retention label by the Microsoft Graph PowerShell SDK" -f $TaggedFilesReport.Count)
$TaggedFilesReport = $TaggedFilesReport | Sort-Object {$_."Tagging Date" -as [datetime]} -Descending
$TaggedFilesReport | Out-GridView -Title 'SharePoint Files tagged by Graph SDK script'

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.