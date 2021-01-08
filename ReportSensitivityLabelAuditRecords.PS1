# ReportSensitivityLabelAuditRecords.PS1
# Extracts audit events from the Office 365 audit log to generate a report
# https://github.com/12Knocksinna/Office365itpros/edit/master/ReportSensitivityLabelAuditRecords.PS1
# Needs connections to Exchange Online and the compliance endpoints
CLS
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
# Now check for connection to compliance endpoint
$TenantLabels = @{}
Try {
    $Labels = Get-Label }
Catch {
    Write-Host "Your PowerShell session must be connected to the Compliance endpoint to fetch label data" ; break}
# Populate hash table with label data  
$Labels.ForEach( {
       $TenantLabels.Add([String]$_.ImmutableId, $_.DisplayName) } )

# Search the Office 365 Audit Log for Sensitivity Label events
Write-Host "Searching Office 365 Audit Log to find audit records for sensitivity labels"
$StartDate = (Get-Date).AddDays(-90); $EndDate = (Get-Date) 
$OutputCSVFile = "C:\temp\SensitivityLabelsAuditRecords.csv"
$Operations = "SensitivityLabeledFileOpened", "SensitivityLabeledFileRenamed", "SensitivityLabelRemoved", "SensitivityLabelApplied", "FileSensitivityLabelApplied", "FileSensitivityLabelRemoved", "FileSensitivityLabelChanged", "Assign label to group."

$GroupLabels = 0; $LabelsChanged = 0; $MisMatches = 0; $NewDocLabels = 0; $LabelsRemoved = 0; $GroupLabels = 0; $LabelsRenamed = 0; $OfficeFileOpens = 0
$Records = (Search-UnifiedAuditLog -Operations $Operations -StartDate $StartDate -EndDate $EndDate -ResultSize 2000)
# If we find some records, process them
If (!$Records) {
    Write-Host "No audit records for sensitivity labels found." }
Else {
    Write-Host "Processing" $Records.Count "sensitivity labels audit records..."
    $Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
    # Scan each audit record to extract information
    ForEach ($Rec in $Records) {
      $Document = $Null; $Site = $Null; $OldLabelId = "None"; $SiteLabelId = $Null; $SiteLabel = $Null; $OldLabel = "None"; $Device = $Null; $Application = $Null
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      $User        = $AuditData.UserId
      $Target      = $AuditData.ObjectId
      Switch ($AuditData.Operation)            {
            "SensitivityLabelApplied" { # Apply sensitivity label to a site, group, or team
                 If ($Rec.RecordType -eq "SharePoint") { # It's an application of a label to a site rather than a file
                    $GroupLabels++
                    $Reason      = "Label applied to site" 
                    $LabelId     = $AuditData.ModifiedProperties.NewValue
                    If (([String]::IsNUllOrWhiteSpace($AuditData.ModifiedProperties.OldValue) -eq $False )) {$OldLabelId = $AuditData.ModifiedProperties.OldValue}
                    $SiteLabelId = $Null
                    $Site        = $AuditData.ObjectId }
                 ElseIf ($AuditData.EmailInfo.Subject)  { #It's email
                    $NewDocLabels++
                    $Target      = "Email: " + $AuditData.EmailInfo.Subject 
                    $Application = "Outlook"
                    $Document    = $Target
                    $Site        = "Exchange Online mailbox"
                    $Reason      = "Label applied to email" }
                 Else {
                    $NewDocLabels++
                    $Reason      = "Label applied to file (desktop app)" 
                    $LabelId     = $AuditData.SensitivityLabelEventData.SensitivityLabelId 
                    $Document    = $Target
                    $Site        = $Target.SubString(0,$Target.IndexOf("/Shared")+1) }
                 $Device         = $AuditData.DeviceName
                 $Application    = $AuditData.Application 
             }
            "FileSensitivityLabelApplied" { # Label applied to an Office Online document
                 $NewDocLabels++
                 $Reason = "Label applied to document"
                 $Document    = $AuditData.DestinationFileName
                 $Site        = $AuditData.SiteURL
                 $LabelId     = $AuditData.DestinationLabel
                 $SiteLabelId = $Null }
            "FileSensitivityLabelChanged" { # Office Online changes a sensitivity label
                 $LabelsChanged++
                 $Reason      = "Label changed in Office app"
                 $Document    = $AuditData.SourceFileName
                 $Site        = $AuditData.SiteURL
                 $LabelId     = $AuditData.SensitivityLabelEventData.SensitivityLabelId 
                 $OldLabelId  = $AuditData.SensitivityLabelEventData.OldSensitivityLabelId 
                 $SiteLabelId = $Null  }
            "FileSensitivityLabelRemoved" { # Label removed from an Office document
                 $LabelsRemoved++
                 $Reason      = "Label removed in Office app"
                 $Document    = $AuditData.SourceFileName
                 $Site        = $AuditData.SiteURL
                 $LabelId     = $AuditData.SensitivityLabelEventData.SensitivityLabelId 
                 $OldLabelId  = $AuditData.SensitivityLabelEventData.OldSensitivityLabelId 
                 $SiteLabelId = $Null }
            "DocumentSensitivityMismatchDetected" { # Mismatch between document label and site label 
                 $MisMatches++
                 $Reason      = "Mismatch between label assigned to document and host site"
                 $Document    = $AuditData.SourceFileName
                 $Site        = $AuditData.SiteURL
                 $LabelId     = $AuditData.SensitivityLabelId
                 $SiteLabelId = $AuditData.SiteSensitivityLabelId }
            "SensitivityLabeledFileOpened"  { # Office desktop app opens a labeled file
                 $OfficeFileOpens++
                 $Application = $AuditData.Application
                 $Device      = $AuditData.DeviceName
                 $LabelId     = $AuditData.LabelId
                 $Document    = $AuditData.ObjectId
                 $Site        = "Local workstation (" + $AuditData.DeviceName + ")"
                 $Reason      = "Labeled document opened by " + $AuditData.Application             }
            "SensitivityLabeledFileRenamed" { #Label renamed by an Office desktop app
                 $LabelsRenamed++
                 $Application  = $AuditData.Application
                 $Device       = $AuditData.DeviceName
                 $LabelId      = $AuditData.LabelId
                 $Reason       = "Labeled file renamed" }
            "SensitivityLabelRemoved" { #Label removed by an Office desktop app
                 $LabelsRemoved++
                 $Application  = $AuditData.Application
                 $Device       = $AuditData.DeviceName
                 $LabelId      = $AuditData.SensitivityLabelEventData.OldSensitivityLabelId 
                 $Reason       = "Label removed from file with " + $AuditData.Application }
            "Assign label to group." { # Azure Active Directory notes label assignment
                 $GroupLabels++
                 $Reason      = "Label assigned to Azure Active Directory Group"
                 $Target      = $AuditData.Target[3].Id
                 $Application = $AuditData.Actor.Id[0]
                 $LabelId     = $AuditData.ModifiedProperties[2].NewValue }
            } # End Switch
            
# Resolve Label identifiers to display name
         If (([String]::IsNUllOrWhiteSpace($LabelId) -eq $False )) { $Label = $TenantLabels.Item($LabelId) }
         If ($SiteLabelId -ne $Null) { $SiteLabel = $TenantLabels.Item($SiteLabelId) }
         If ($OldLabelId -ne "None") { $OldLabel = $TenantLabels.Item($OldLabelId) }
         $ReportLine = [PSCustomObject] @{
           TimeStamp   = Get-Date($AuditData.CreationTime) -format g
           User        = $AuditData.UserId
           Target      = $Target
           Reason      = $Reason
           Label       = $Label
           LabelId     = $LabelId
           OldLabel    = $OldLabel
           OldLabelId  = $OldLabelId
           Document    = $Document
           Location    = $Site
           SiteLabel   = $SiteLabel
           SiteLabelId = $SiteLabelId
           Device      = $Device
           Application = $Application
           Action      = $AuditData.Operation }        
      $Report.Add($ReportLine) }
}
Cls
Write-Host "All done!" $Records.Count "Sensitivity Label records found for the last 90 days"
Write-Host "Site Labels Applied:             " $GroupLabels
Write-Host "Labels applied to new documents: " $NewDocLabels
Write-Host "Labels updated on documents:     " $LabelsChanged
Write-Host "Labeled files renamed:           " $LabelsRenamed
Write-Host "Labeled files opened (desktop):  " $OfficeFileOpens
Write-Host "Labels removed from documents:   " $LabelsRemoved
Write-Host "Mismatches detected:             " $MisMatches     
Write-Host "----------------------"
Write-Host "Report file written to" $OutputCSVFile

$Report = $Report | Sort {$_.TimeStamp -as [datetime]} 
$Report | Export-CSV -NoTypeInformation $OutputCSVFile
