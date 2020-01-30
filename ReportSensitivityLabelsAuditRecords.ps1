# ReportSensitivityLabelAuditRecords.PS1
# Needs connections to Exchange Online and the compliance endpoints
CLS; 
# First Check for connection to compliance endpoint
$TenantLabels = @{}
Try {
    $Labels = Get-Label }
Catch {
    Write-Host "Your PowerShell session must be connected to the Compliance endpoint to fetch label data" ; break}
# Now Populate hash table with label data  
$Labels.ForEach( {
       $TenantLabels.Add([String]$_.ImmutableId, $_.DisplayName) } )

# Search the Office 365 Audit Log for Sensitivity Label events
Write-Host "Searching Office 365 Audit Records to find auto-expired group deletions"
$StartDate = (Get-Date).AddDays(-90); $EndDate = (Get-Date) 
$GroupLabels = 0; $LabelsChanged = 0; $MisMatches = 0; $NewDocLabels = 0; $LabelsRemoved = 0
$Records = (Search-UnifiedAuditLog -Operations FileSensitivityLabelChanged, FileSensitivityLabelApplied, FileSensitivityLabelRemoved, SensitivityLabelApplied, DocumentSensitivityMismatchDetected -StartDate $StartDate -EndDate $EndDate -ResultSize 1000)
# If we find some records, process them
If ($Records.Count -eq 0) {
    Write-Host "No audit records for group deletions found." }
Else {
    Write-Host "Processing" $Records.Count "team deletion audit records..."
    $Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
    # Scan each audit record to extract information
    ForEach ($Rec in $Records) {
      $Document = $Null; $Site = $Null; $OldLabelId = "None"; $SiteLabelId = $Null; $SiteLabel = $Null; $OldLabel = "None"
      $AuditData = ConvertFrom-Json $Rec.Auditdata
          Switch ($AuditData.Operation)
          {
            "SensitivityLabelApplied" { # Apply sensitivity label to a site, group, or team
                 $GroupLabels++ 
                 $Reason      = "Label applied to site" 
                 $User        = $AuditData.UserId
                 $Target      = $AuditData.ObjectId
                 $LabelId     = $AuditData.ModifiedProperties.NewValue
                 If (([String]::IsNUllOrWhiteSpace($AuditData.ModifiedProperties.OldValue) -eq $False )) {$OldLabelId = $AuditData.ModifiedProperties.OldValue}
                 $SiteLabelId = $Null
                 $Site        = $AuditData.ObjectId}
            "FileSensitivityLabelApplied" { # Label applied to an Office Online document
                 $NewDocLabels++
                 $Reason = "Label applied to document"
                 $User        = $AuditData.UserId
                 $Target      = $AuditData.ObjectId
                 $Document    = $AuditData.DestinationFileName
                 $Site        = $AuditData.SiteURL
                 $LabelId     = $AuditData.DestinationLabel
                 $SiteLabelId = $Null }
            "FileSensitivityLabelChanged" { # Office Online changes a sensitivity label
                 $LabelsChanged++
                 $Reason      = "Label changed in Office app"
                 $User        = $AuditData.UserId
                 $Target      = $AuditData.ObjectId
                 $Document    = $AuditData.SourceFileName
                 $Site        = $AuditData.SiteURL
                 $LabelId     = $AuditData.SensitivityLabelEventData.SensitivityLabelId 
                 $OldLabelId  = $AuditData.SensitivityLabelEventData.OldSensitivityLabelId 
                 $SiteLabelId = $Null  }
            "FileSensitivityLabelRemoved" { # Label removed from an Office document
                 $LabelsRemoved++
                 $Reason      = "Label removed in Office app"
                 $User        = $AuditData.UserId
                 $Target      = $AuditData.ObjectId
                 $Document    = $AuditData.SourceFile
                 $Site        = $AuditData.SiteURL
                 $LabelId     = $AuditData.SensitivityLabelEventData.SensitivityLabelId 
                 $OldLabelId  = $AuditData.SensitivityLabelEventData.OldSensitivityLabelId 
                 $SiteLabelId = $Null }
            "DocumentSensitivityMismatchDetected" { # Mismatch between document label and site label 
                 $MisMatches++
                 $Reason      = "Mismatch between label assigned to document and host site"
                 $User        = $AuditData.UserId
                 $Target      = $AuditData.ObjectId
                 $Document    = $AuditData.SourceFileName
                 $Site        = $AuditData.SiteURL
                 $LabelId     = $AuditData.SensitivityLabelId
                 $SiteLabelId = $AuditData.SiteSensitivityLabelId }
          }   
# Resolve Label identifiers to display name
         If (([String]::IsNUllOrWhiteSpace($LabelId) -eq $False )) { $Label= $TenantLabels.Item($LabelId) }
         If ($SiteLabelId -ne $Null) { $SiteLabel = $TenantLabels.Item($SiteLabelId) }
         If ($OldLabelId -ne "None") { $OldLabel = $TenantLabels.Item($OldLabelId) }
          $ReportLine = [PSCustomObject] @{
           TimeStamp  = Get-Date($AuditData.CreationTime) -format g
           User        = $User
           Target      = $Target
           Reason      = $Reason
           Label       = $Label
           LabelId     = $LabelId
           OldLabel    = $OldLabel
           OldLabelId  = $OldLabelId
           Document    = $Document
           Site        = $Site
           SiteLabel   = $SiteLabel
           SiteLabelId = $SiteLabelId
           Action      = $AuditData.Operation }        
      $Report.Add($ReportLine) }
}
Cls
Write-Host "All done!" $Records.Count "Sensitivity Label records for the last 90 days"
Write-Host "Site Labels Applied:             " $GroupLabels
Write-Host "Labels applied to new documents: " $NewDocLabels
Write-Host "Labels updated on documents:     " $LabelsChanged
Write-Host "Labels removed from documents:   " $LabelsRemoved
Write-Host "Mistmatches detected:            " $MisMatches     
Write-Host "----------------------"

$Report |  Out-GridView

