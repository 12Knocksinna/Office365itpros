# Find out when an anonymous link is used by someone outside an Office 365 tenant to access SharePoint Online and OneDrive for Business documents
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindWhenAnonymousLinkUsed.ps1
$StartDate = (Get-Date).AddDays(-90); $EndDate = (Get-Date) #Maximum search range for audit log for E3 users
CLS; Write-Host "Searching Office 365 Audit Records to find anonymous sharing activity"
$Records = (Search-UnifiedAuditLog -Operations AnonymousLinkUsed -StartDate $StartDate -EndDate $EndDate -ResultSize 1000)
If ($Records.Count -eq 0) {
    Write-Host "No anonymous share records found." }
Else {
    Write-Host "Processing" $Records.Count "audit records..."
    $Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
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
    $Report.Add($ReportLine) }
  # Now that we have parsed the information for the link used audit records, let's track what happened to each link
  $RecNo = 0; CLS; $TotalRecs = $Report.Count
  $AuditReport = [System.Collections.Generic.List[Object]]::new() # Create output file 
  ForEach ($R in $Report) {
     $RecNo++
     $ProgressBar = "Processing audit records for " + $R.FileName + " (" + $RecNo + " of " + $TotalRecs + ")" 
     Write-Progress -Activity "Checking Sharing Activity With Anonymous Links" -Status $ProgressBar -PercentComplete ($RecNo/$TotalRecs*100)
     $StartSearch = (Get-Date $R.TimeStamp); $EndSearch = (Get-Date $R.TimeStamp).AddDays(+7) # We'll search for any audit records 
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
         $AuditReport.Add($ReportLine) }
}}
$AuditReport | Sort FileName, IPAddress, User, SortTime -Unique | Export-CSV -NoTypeInformation "c:\Temp\AnonymousLinksUsed.CSV"
Write-Host "All done. Output file is available in c:\temp\AnonymousLinksUsed.Csv"
# Output in grid, making sure that any duplicates created at the same time are ignored
$AuditReport | Sort FileName, IPAddress, User, SortTime -Unique | Select Timestamp, Action, Filename, IPAddress, Workload, Site | Out-Gridview  

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
