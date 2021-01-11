# Find out when an anonymous link is used by someone outside an Office 365 tenant to access SharePoint Online and OneDrive for Business documents
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindWhenAnonymousLinkUsed.ps1
$StartDate = (Get-Date).AddDays(-90); $EndDate = (Get-Date) #Maximum search range for audit log for E3 users
CLS; Write-Host "Searching Office 365 Audit Records to find anonymous sharing activity"
$Records = (Search-UnifiedAuditLog -Operations AnonymousLinkUsed -StartDate $StartDate -EndDate $EndDate -ResultSize 1000)
If ($Records.Count -eq 0) {
    Write-Host "No anonymous share records found." }
Else {
    Write-Host "Processing" $Records.Count "audit records..."
    $Report = @() # Create output file for report
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
  $RecNo = 0; CLS; $TotalRecs = $Report.Count
  ForEach ($R in $Report) {
     $RecNo++
     $ProgressBar = "Processing audit records for " + $R.FileName + " (" + $RecNo + " of " + $TotalRecs + ")" 
     Write-Progress -Activity "Checking Sharing Activity With Anonymous Links" -Status $ProgressBar -PercentComplete ($RecNo/$TotalRecs*100)
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
}}
$Report | Sort FileName, IPAddress, User, SortTime | Export-CSV -NoTypeInformation "c:\Temp\AnonymousLinksUsed.CSV"
Write-Host "All done. Output file is available in c:\temp\AnonymousLinksUsed.Csv"
# Output in grid, making sure that any duplicates created at the same time are ignored
$Report | Sort FileName, IPAddress, User, SortTime -Unique | Select Timestamp, Action, Filename, IPAddress, Workload, Site | Out-Gridview  
