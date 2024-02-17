# ReportSPOSiteStorageUsedGraph.PS1
# A Graph-based script to report SharePoint Online Site Storage usage data
#
# V1.0 27-Mar-2020
# V2.0 17-Feb-2024 Updated for the Graph SDK
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportSPOSiteStorageUsedGraph.PS1

Clear-Host

Connect-MgGraph -NoWelcome -Scopes Reports.Read.All, Directory.Read.All
$ObscureFlag = $false
$TempDownloadFile = "c:\temp\x.csv"

# Check if the tenant has obscured real names for reports - see https://office365itpros.com/2022/09/09/graph-usage-report-tips/
$DisplaySettings = Invoke-MgGraphRequest -Method Get -Uri 'https://graph.microsoft.com/beta/admin/reportSettings'
If ($DisplaySettings['displayConcealedNames'] -eq $true) { # data is obscured, so let's reset it to allow the report to run
   $ObscureFlag = $true
   Write-Host "Setting tenant data concealment for reports to False" -foregroundcolor red
   Invoke-MgGraphRequest -Method PATCH -Uri 'https://graph.microsoft.com/beta/admin/reportSettings' `
    -Body (@{"displayConcealedNames"= $false} | ConvertTo-Json) 
}

$CSVOutputFile = "c:\temp\SPOSiteConsumption.CSV"
[array]$SiteTypesToReport = "Group", "Team Channel", "Team Site"

Write-Host "Fetching SharePoint Online site data from the Graph..."
# Get SharePoint files usage data - includes redirects, so we will have to remove them
$URI = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D180')"
Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath $TempDownloadFile
[array]$Sites = Import-CSV $TempDownloadFile

$TotalSPOStorageUsed = [Math]::Round(($Sites."Storage Used (Byte)" | Measure-Object -Sum).Sum /1GB,2)
$Report = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Site in $Sites) {
  $DoNotProcessSite = $False
  If ([string]::IsNullOrEmpty($Site."Last Activity Date")) {
    $LastActiveDate = "No Activity" 
  } Else  {
    $LastActiveDate = Get-Date ($Site."Last Activity Date") -Format dd-MMM-yyyy 
  }
  
# Check for redirect sites returned by the Graph so we don't process them
If (($Site."Owner Display Name" -eq "System Account") -and ([string]::IsNullOrEmpty($Site."Owner Principal Name")))  {
    $DoNotProcessSite = $True 
}
# Check for the fundamental site because we don't want to process it either
If ($Site."Root Web Template" -notin $SiteTypesToReport) {
    $DoNotProcessSite = $True 
  }
  If ($DoNotProcessSite -eq $False) {
    $UsedGB = [Math]::Round($Site."Storage Used (Byte)"/1GB,2) 
    $PercentTenant = ([Math]::Round($Site.StorageUsageCurrent/1024,4)/$TotalSPOStorageUsed).tostring("P")  
    $ReportLine = [PSCustomObject]@{
         URL            = $Site."Site URL"
         Owner          = $Site."Owner Display Name"
         OwnerUPN       = $Site."Owner Principal Name"
         Files          = $Site."File Count"
         ActiveFiles    = $Site."Active File Count"
         LastActiveDate = $LastActiveDate
         Template       = $Site."Root Web Template"
         QuotaGB        = [Math]::Round($Site."Storage Allocated (Byte)"/1GB,0) 
         UsedGB         = $UsedGB
         PercentUsed    = ([Math]::Round(($Site."Storage Used (Byte)"/$Site."Storage Allocated (Byte)"),4).ToString("P"))
         PercentTenant = $PercentTenant 
    }
    $Report.Add($ReportLine) }
}

# Switch the tenant report obscure data setting back if necessary
If ($ObscureFlag -eq $True) {
  Write-Host "Resetting tenant data concealment for reports to True" -foregroundcolor red
  Invoke-MgGraphRequest -Method PATCH -Uri 'https://graph.microsoft.com/beta/admin/reportSettings' `
   -Body (@{"displayConcealedNames"= $true} | ConvertTo-Json) 
}

$Report | Export-CSV -NoTypeInformation $CSVOutputFile
$Report | Sort-Object {$_.UsedGB -as [decimal]}, url -Descending | Out-GridView
Write-Host ("{0} sites processed. Current SharePoint Online storage consumption is {1} GB. Report file available in {2}" -f $Report.count, $TotalSPOStorageUsed, $CSVOutputReport)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
