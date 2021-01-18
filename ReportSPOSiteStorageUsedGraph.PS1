# ReportSPOSiteStorageUsedGraph.PS1
# A Graph-based script to report SharePoint Online Site Storage usage data
#
# V1.0 27-Mar-2020
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportSPOSiteStorageUsedGraph.PS1
CLS
# Define the values applicable for the application used to connect to the Graph - You must change these values to match your app id,
# tenant id, and secret...
$AppId = "e716b32c-0edb-48be-9385-30a9cfd96155"
$TenantId = "c662313f-14fc-43a2-9a7a-d2e27f4f3478"
$AppSecret = 's_rkvIn1oZ1cNceUBvJ2or1lrrIsb*:='

# Build the request to get the OAuth 2.0 access token
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials"}

# Request token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
# Unpack Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token
$headers = @{Authorization = "Bearer $token"}

Write-Host "Fetching SharePoint Online site data from the Graph..."
# Get SharePoint files usage data - includes redirects, so we will have to remove them
$SPOFilesReportsURI = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D7')"
$Sites = (Invoke-RestMethod -Uri $SPOFilesReportsURI -Headers $Headers -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv

$TotalSPOStorageUsed = [Math]::Round(($Sites."Storage Used (Byte)" | Measure-Object -Sum).Sum /1GB,2)
$Report = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Site in $Sites) {
  $DoNotProcessSite = $False
  If ([string]::IsNullOrEmpty($Site."Last Activity Date")) {
    $LastActiveDate = "No Activity" }
  Else  {
    $LastActiveDate = Get-Date ($Site."Last Activity Date") -Format dd-MMM-yyyy }
  
# Check for redirect sites returned by the Graph so we don't process them
If (($Site."Owner Display Name" -eq "System Account") -and ([string]::IsNullOrEmpty($Site."Owner Principal Name")))  {
   $DoNotProcessSite = $True }
# Check for the fundamental site because we don't want to process it eiteher
If ($Site."Root Web Template" -eq "SharePoint Online Tenant Fundamental Site") {
   $DoNotProcessSite = $True }
  If ($DoNotProcessSite -eq $False) {
  $UsedGB = [Math]::Round($Site."Storage Used (Byte)"/1GB,2) 
  $PercentTenant = ([Math]::Round($Site.StorageUsageCurrent/1024,4)/$TotalSPOStorageUsed).tostring("P")  
  $SitesProcessed++
  $ReportLine = [PSCustomObject]@{
         URL            = $Site."Site URL"
 #       SiteName       = $Site.Title
         Owner          = $Site."Owner Display Name"
         OwnerUPN       = $Site."Owner Principal Name"
         Files          = $Site."File Count"
         ActiveFiles    = $Site."Active File Count"
         LastActiveDate = $LastActiveDate
         Template       = $Site."Root Web Template"
         QuotaGB        = [Math]::Round($Site."Storage Allocated (Byte)"/1GB,0) 
         UsedGB         = $UsedGB
         PercentUsed    = ([Math]::Round(($Site."Storage Used (Byte)"/$Site."Storage Allocated (Byte)"),4).ToString("P"))
         PercentTenant = $PercentTenant }
  $Report.Add($ReportLine) }
}

$Report | Export-CSV -NoTypeInformation c:\temp\SPOSiteConsumption.CSV
$Report | Sort {$_.UsedGB -as [decimal]}, url -Descending | Out-gridview
Write-Host $SitesProcessed "sites processed. Current SharePoint Online storage consumption is" $TotalSPOStorageUsed "GB. Report is in C:\temp\SPOSiteConsumption.CSV"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
