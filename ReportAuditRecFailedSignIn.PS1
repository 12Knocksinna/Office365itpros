# ReportAuditRecFailedSignIn.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportAuditRecFailedSignIn.PS1
# Example used in the Reporting and Auditing chapter of Office 365 for IT Pros
# Updated 28-June-2024 for the 2025 Book

$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Loading Exchange Online Management module"
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

Write-Host "Searching for failed login audit records..."
[array]$Records = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date).AddDays(+1) `
  -Operations UserLoginFailed  -SessionCommand  ReturnLargeSet -ResultSize 5000 -Formatted
If ($Records.Count -eq 0) {
  Write-Host "No audit records found for failed logins." 
  break
}

# Remove any duplicates and sort by date  
$Records = $Records | Sort-Object Identity -Unique | Sort-Object { $_.CreationDate -as [datetime]} -Descending
 
Write-Host "Processing" $Records.Count "audit records..."
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
  $AuditData = ConvertFrom-Json $Rec.Auditdata
  $ReportLine = [PSCustomObject]@{
    TimeStamp   = $Rec.CreationDate
    User        = $AuditData.UserId
    Action      = $AuditData.Operation
    Status      = $AuditData.ResultStatus
    IpAddress   = $AuditData.ActorIpAddress
    Error       = $AuditData.LogonError
    UserAgent   = $AuditData.ExtendedProperties.value[0] 
  }
  $Report.Add($ReportLine) 
}

$Report | Sort-Object User, Timestamp | Select-Object Timestamp, User, IpAddress, UserAgent | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
