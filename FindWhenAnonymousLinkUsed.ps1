# Find out when an anonymous link is used
$StartDate = (Get-Date).AddDays(-90); $EndDate = (Get-Date)
$Records = (Search-UnifiedAuditLog -Operations AnonymousLinkUsed -StartDate $StartDate -EndDate $EndDate -ResultSize 1000)
If ($Records.Count -eq 0) {
    Write-Host "No anonymous share records found." }
Else {
    Write-Host "Processing" $Records.Count "audit records..."}
# Create output file for report
$Report = @()
# Scan each audit record to extract information
ForEach ($Rec in $Records) {
  $AuditData = ConvertFrom-Json $Rec.Auditdata
  $ReportLine = [PSCustomObject][Ordered]@{
    TimeStamp = Get-Date($AuditData.CreationTime) -format g
    User      = $AuditData.UserId
    Action    = $AuditData.Operation
    Object    = $AuditData.ObjectId
    IPAddress = $AuditData.ClientIP
    Workload  = $AuditData.Workload
    Site      = $AuditData.SiteUrl
    FileName  = $AuditData.SourceFileName 
    SortTime  = $AuditData.CreationTime }
  $Report += $ReportLine }
# Now that we have parsed the information for the link used audit records, let's track what happened to each link
  ForEach ($R in $Report) {
     Write-Host "Examining records for" $R.FileName
     $StartSearch = $R.TimeStamp; $EndSearch = (Get-Date $R.TimeStamp).AddDays(+7) # We'll search for any audit records 
     $AuditRecs = (Search-UnifiedAuditLog -StartDate $StartSearch -EndDate $EndSearch -IPAddresses $R.IPAddress -Operations FileAccessedExtended, FilePreviewed, FileModified, FileAccessed, FileDownloaded -ResultSize 100)
     Foreach ($AuditRec in $AuditRecs) {
       If ($AuditRec.UserIds -Like "*urn:spo:*") { # It's a continuation of anonymous access to a document
          $AuditData = ConvertFrom-Json $AuditRec.Auditdata
          $ReportLine = [PSCustomObject][Ordered]@{
            TimeStamp = Get-Date($AuditData.CreationTime) -format g
            User      = $AuditData.UserId
            Action    = $AuditData.Operation
            Object    = $AuditData.ObjectId
            IPAddress = $AuditData.ClientIP
            Workload  = $AuditData.Workload
            Site      = $AuditData.SiteUrl
            FileName  = $AuditData.SourceFileName 
            SortTime  = $AuditData.CreationTime }}
         $Report += $ReportLine }
}
$Report | Sort FileName, IPAddress, User, SortTime | Export-CSV -NoTypeInformation "c:\Temp\AnonymousLinksUsed.CSV"
