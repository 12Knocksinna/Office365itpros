
# AuditRecordAuditLabel.PS1
# A script to look for audit records for the application of retention labels and report those which
# are applied by an auto-label policy. 

$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}

# Retention label to check for - change this to your preferred label name
$RetentionLabel = "Teams recordings"

# Start and end date for the audit scan. By default, we look for 14 days, but you can choose any value you like up to 365 (assuming Office 365 E5)
$StartDate = (Get-Date).AddDays(-14); $EndDate = (Get-Date) 
$OutputCSVFile = "C:\temp\TaggedTeamsRecordings.csv"

# Find the audit records
[array]$Records = (Search-UnifiedAuditLog -Operations TagApplied -StartDate $StartDate -EndDate $EndDate -Formatted -ResultSize 2000)

If (!($Records)) {Write-Host "No audit records found - exiting!"; break}

# Check the audit records for those for auto-label policies which applied the retention label we're interested in
$TaggedRecordings = [System.Collections.Generic.List[Object]]::new() 	
ForEach ($Rec in $Records) {
   $AuditData = $Rec.AuditData | ConvertFrom-Json
   If (($AuditData.DestinationLabel -eq $RetentionLabel) -and ($AuditData.UserType -eq "CustomPolicy")) { 
      $RecordingFileName = $AuditData.DestinationFileName
      $DateLoc = ($RecordingFileName.IndexOf("-202")+1)
      $RDate = $RecordingFileName.SubString($DateLoc,8)
      $TimeLoc = $DateLoc + 9
      $RTime = $RecordingFileName.SubString($TimeLoc,4)
      $RecordingDateTime = $RDate + $RTime
      [datetime]$RecordingDate = [datetime]::ParseExact($RecordingDateTime,"yyyyMMddHHmm",$null)
      [datetime]$TaggingDate = Get-Date($AuditData.CreationTime)
      $TimeToTag = ($TaggingDate - $RecordingDate)
      $TotalSeconds = $TotalSeconds + $TimeToTag.TotalSeconds
      $TimeToTagFormatted = "{0:dd}d:{0:hh}h:{0:mm}m" -f $TimeToTag
# Add the data about our record          
      $DataLine = [PSCustomObject] @{
         Workload            = $AuditData.Workload
         Recording           = $AuditData.DestinationFileName
         "Retention Label"   = $AuditData.DestinationLabel
         "Tagging Date"      = Get-Date($AuditData.CreationTime) -format g
         "Recording date"    = Get-Date($RecordingDate) -format g
         "Days to label"     = $TimeToTagFormatted
         Site                = $AuditData.SiteURL
         FullURL             = $AuditData.ObjectId }
    $TaggedRecordings.Add($DataLine) 
   } # End if
} # End ForEach

CLS
# All done
$TaggedRecordings | Export-CSV -NoTypeInformation $OutputCSVFile
Write-Host ("{0} audit records found for auto-applying the {1} retention label between {2} and {3}" -f $TaggedRecordings.Count, $RetentionLabel, $StartDate, $EndDate)
$AverageSeconds = $TotalSeconds/$TaggedRecordings.Count
$AverageTimeToTag = [timespan]::fromseconds($AverageSeconds)
$AverageTimeToTagFormatted = "{0:dd}d:{0:hh}h:{0:mm}m" -f $AverageTimeToTag
Write-Host ("Average elapsed time to auto-label recordings: {0}" -f $AverageTimeToTagFormatted)
Write-Host ("The report file is available in {0}. -f $OutputCSVFile}
