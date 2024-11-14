# Find-LastAccessedDateDocuments.PS1
# An example of finding the last accessed date for documents in a SharePoint Online site using audit records
# V1.0 14-Nov-2024

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-LastAccessedDateDocuments.PS1

# Connect to Exchange Online if necessary
[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules ) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
}

[string]$TargetSite = "https://redmondassociates.sharepoint.com/sites/confidentialstuff/"
[string]$TargetSearchSite = $TargetSite + "*"
$StartDate = (Get-Date).AddDays(-180)
$EndDate = (Get-Date).AddDays(1)

Write-Host ("Searching for SharePoint file operations in the {0} site between {1} and {2}" -f $TargetSite, $StartDate, $EndDate)
[array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType "SharePointFileOperation" `
     -ResultSize 5000 -SessionCommand ReturnLargeSet -Formatted -ObjectIds $TargetSearchSite
If (!($Records)) {
    Write-Host "No audit records found"
    Break
}
# Remove any duplicates
$Records = $Records | Sort-Object Identity -Unique | Sort-Object {$_.CreationDate -as [DateTime]} 
Write-Host ("{0} SharePoint Online file operations audit records found" -f $Records.Count)

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
    $AuditData = $Rec.AuditData | ConvertFrom-Json
    $ReportLine = [PSCustomObject]@{
        Timestamp   = Get-Date $Rec.CreationDate -format 'dd-MMM-yyyy HH:mm:ss'
        User        = $AuditData.UserId
        Operation   = $AuditData.Operation
        File        = $AuditData.SourceFileName
        Folder      = $AuditData.SourceRelativeUrl
        URL         = $AuditData.ObjectId
    }
    $Report.Add($ReportLine)
}

# Filter to find the operations performed by real users
[array]$UserFileOperations = $Report | Where-Object {
    $_.Folder -ne 'SiteAssets' -and
    $_.Folder -notlike '*PreservationHoldLibrary/SharedVersions*' -and
    $_.Folder -notlike '*Shared Documents/Forms*'
}

$UserFileOperations = $UserFileOperations | Where-Object {$_.File -ne 'AllItems.aspx'}

Write-Host "Files found with audit records for file operations"
$UserFileOperations | Group-Object File -NoElement | Sort-Object Count -Descending | Format-Table -AutoSize
Write-Host ""
Write-Host "Users who accessed files"
$UserFileOperations | Group-Object User -NoElement | Sort-Object Count -Descending | Format-Table -AutoSize
Write-Host ""

Write-Host "Latest file access for each file"
[array]$FilesAccessed = $UserFileOperations | Where-Object {$_.Operation -eq 'FileAccessed'} | Sort-Object File, {$_.Timestamp -as [DateTime]} -Descending
# Eliminate odd entries for SharePoint background operations
$FilesAccessed = $FilesAccessed | Where-Object {$_.User -ne 'eba15bfd-c28e-4433-a20e-0278888c5825'}
$FilesAccessed = $FilesAccessed | Where-Object {$_.User -ne 'bdc6105c-4e11-4050-82e6-6549f9b99b89'}
[array]$LatestFileAccess = $null

$FilesAccessed | Group-Object File | ForEach-Object {
    $LatestFileAccess += $_.Group | Sort-Object Timestamp -Descending | Select-Object -First 1
}

$LatestFileAccess | Format-Table File, User, TimeStamp -AutoSize


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.