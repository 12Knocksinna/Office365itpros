# --- Find Priority Cleanup audit records for the documents the policy tags and then moves to the second stafge recycle bin ---
# V1.0 27-Nov-2025

# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-PriorityCleanupSPOAuditRecords.PS1

[array]$Modules = Get-Module | Select-Object Name
If ("ExchangeOnlineManagement" -notin $Modules.Name) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
}
    
Write-Output "Scanning for Priority Cleanup audit records..."
[array]$Records = Search-UnifiedAuditLog -StartDate '1-Nov-2025' -EndDate '28-Nov-2025' -Formatted -SessionCommand ReturnLargeSet -ResultSize 5000 -Operations 'PriorityCleanupFileRecycled', 'PriorityCleanupTagApplied'
$Records = $Records | Sort-Object Identity -Unique
If ($Records.Count -eq 0) {
    Write-Host "No audit records found for Priority Cleanup operations"
    Break
} Else {
    $Records = $records | Sort-Object Identity -Unique | Sort-Object { $_.CreationDate -as [datetime]} -Descending
    Write-Host ("Processing {0} audit records..." -f $Records.Count)
}   

$Records | Group-Object Operations -NoElement | Sort-Object Count | Format-Table Name, Count -AutoSize

# Report removed files
$PriorityCleanupReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Item in $Records) {
    $AuditData = $Item.AuditData | ConvertFrom-Json
    $ReportLine = [PSCustomObject][Ordered]@{
        TimeStamp       = Get-Date ($AuditData.CreationTime) -format 'dd-MMM-yyyy HH:mm'
        User            = $AuditData.UserId
        Action          = $AuditData.Operation
        ItemType        = $AuditData.ItemType
        SiteURL         = $AuditData.SiteURL
        Item            = $AuditData.SourceFileName
        Extension       = $AuditData.SourceFileExtension
        ObjectId        = $AuditData.ObjectId
        FileSizeBytes   = $AuditData.FileSizeBytes
        LabelName       = $AuditData.PolicyMatchInfo.LabelName
        LabelID         = $AuditData.PolicyMatchInfo.LabelId
    }
    $PriorityCleanupReport.Add($ReportLine)
}

$TaggedOperations = $PriorityCleanupReport | Where-Object { $_.Action -eq 'PriorityCleanupTagApplied' }
$MovedOperations = $PriorityCleanupReport | Where-Object { $_.Action -eq 'PriorityCleanupFileRecycled' }
Write-Host ""
Write-Host ("Found {0} Priority Cleanup Tag Applied operations" -f $TaggedOperations.Count)
$TaggedOperations | Group-Object SiteURL -NoElement | Sort-Object Count -Descending | Format-Table Name, Count
Write-Host ""
Write-Host ("Found {0} Priority Cleanup File Recycled operations" -f $MovedOperations.Count)
$MovedOperations | Group-Object SiteURL -NoElement | Sort-Object Count -Descending | Format-Table Name, Count


# Export the report as an Excel worksheet or CSV file
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $true
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\PriorityCleanupReport.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    } 
    $PriorityCleanupReport | Export-Excel -Path $ExcelOutputFile -WorksheetName "Priority Cleanup" -Title ("Priority Cleanup Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "PriorityCleanup" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\PriorityCleanupReport.CSV"
    $PriorityCleanupReport  | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}

If ($ExcelGenerated) {
    Write-Output ("Excel worksheet output written to {0}" -f $ExcelOutputFile)
} Else {
    Write-Output ("CSV output file written to {0}" -f $CSVOutputFile)
} 

Write-Output "All done - enjoy the Priority Cleanup data!"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.