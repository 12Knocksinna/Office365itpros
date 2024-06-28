# ReportAuditRecsGuestDocAccess.PS1
# Example for Chapter 21 of Office 365 for IT Pros
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportAuditRecsGuestDocAccess.PS1
# Note: There might well be far more than 5000 FileAccessed audit records to be retrieved. You can either go back and fetch more or shorten the timeframe that you look.

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Loading Exchange Online Management module"
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

Write-Host "Searching for SharePoint file access audit records..."
[array]$Records = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date).AddDays(+1) `
  -Operations FileAccessed -ResultSize 5000 -Formatted -SessionCommand ReturnLargeSet
If ($Records.Count -eq 0) {
  Write-Host "No SharePoint file access records found for guest users." 
  Break
}

# Remove any duplicates and sort by date and extract the records for guest users
$Records = $Records | Sort-Object Identity -Unique | Where-Object {$_.UserIds -Like "*#EXT#*" } | Sort-Object { $_.CreationDate -as [datetime]} -Descending 

Write-Host "Processing" $Records.Count" SharePoint file access audit records..."
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
  $AuditData = ConvertFrom-Json $Rec.Auditdata
  If ($AuditData.SourceFileName -NotLike "*aspx*" -And $AuditData.SourceFileName -NotLike "*jpg*" ) {
    $ReportLine = [PSCustomObject]@{
      TimeStamp   = $Rec.CreationDate
      User        = $Rec.UserIds
      Action      = $AuditData.Operation
      Workload    = $AuditData.Workload
      URL         = $AuditData.SiteUrl
      Document    = $AuditData.SourceFileName 
    }      
    $Report.Add($ReportLine)
  }
}

$Report | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
