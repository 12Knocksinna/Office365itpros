# ReportDocumentSensitivityLabelMismatches.PS1
# A script to show how to find audit records for document mismatch events (the sensitivity label applied to an uploaded 
# document is higher than the label assigned to the site) and lookup the user who uploaded the file and the display names
# of the labels
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportDocumentSensitivityLabelMismatches.PS1

$Modules = @( "ExchangeOnlineManagement" )
# Requires -Modules $Modules

$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {
   Write-Host "Please connect to the Exchange Online management module and then restart the script"; break
}

# Connect to the compliance endpoint
Write-Host "Finding details of sensitivity labels defined in the tenant" 
Connect-IPPSSession
[array]$Labels = Get-Label
If (!($Labels)) {
   Write-Host "Error: can't retrieve sensitivity labels - exiting"; break
}

# Create and populate hash table to lookup sensitivity labels
$LabelLookUp = @{}
ForEach ($L in $Labels) { $LabelLookUp.Add([string]$L.ImmutableId, [string]$L.DisplayName) }
Write-Host ("{0} sensitivity labels found" -f $Labels.count)

Write-Host "Finding audit records for document upload and modified events..."
$StartDate = (Get-Date).AddDays(-90)
$EndDate = (Get-Date).AddDays(1)
[array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Formatted -ResultSize 5000 -Operations FileUploaded, DocumentSensitivityMismatchDetected, FileModified -SessionCommand ReturnLargeSet
If (!($Records)) {Write-Host "Error: Can't find any audit records to process - exiting" ; break}
Write-Host ("Processing {0} audit records..." -f $Records.count)

# Split out the two kinds of audit records
$MismatchRecords = $Records | Where-Object {$_.Operations -eq "DocumentSensitivityMismatchDetected"}
$FileUploads = $Records | Where-Object {$_.Operations -ne "DocumentSensitivityMismatchDetected" -and $_.UserIds -ne "app@sharepoint"}

# Build a lookup table of list item unique identifiers and user names
$LookupRecords = [System.Collections.Generic.List[Object]]::new()
ForEach ($F in $FileUploads) {
   $AuditData = $F.AuditData | ConvertFrom-Json
   $LookupLine = [PSCustomObject][Ordered]@{  
     ListId = $AuditData.ListItemUniqueId
     User   = $AuditData.UserId 
     Date   = $F.CreationDate }
   $LookupRecords.Add($LookupLine)  
}
$LookUpRecordsSort = $LookUpRecords | Sort-Object ListId -Unique
$LookUpTable = @{}
ForEach ($L in $LookUpRecordsSort) { $LookUpTable.Add($L.ListId, $L.User) }

# Generate the report
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($M in $MismatchRecords) {
   $AuditData = $M.AuditData | ConvertFrom-Json
   [string]$UploadUser = $LookUpTable[$AuditData.ListItemUniqueId]
   [string]$DocumentLabel = $LabelLookUp[$AuditData.SensitivityLabelId]     
   [string]$SiteLabel = $LabelLookUp[$AuditData.SiteSensitivityLabelId]   
   If ($AuditData.SourceRelativeUrl -eq "PreservationHoldLibrary") { $UploadUser = "SharePoint Online" }

   $ReportLine = [PSCustomObject][Ordered]@{  
     Date             = $M.CreationDate
     Object           = $AuditData.ObjectId
     FileName         = $AuditData.SourceFileName
     User             = $UploadUser 
     SiteLabel        = $SiteLabel
     SitePriority     = $AuditData.SiteSensitivityLabelOrder
     DocumentLabel    = $DocumentLabel
     DocumentPriority = $AuditData.SensitivityLabelOrder }
   $Report.Add($ReportLine)
}

$Report | Sort-Object {$_.Date -as [datetime]} | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
