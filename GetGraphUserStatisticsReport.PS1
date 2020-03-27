# GetGraphUserStatisticsReport.PS1
# A sample script showing how to gather user activity information from the Graph and assemble it into one report
# V1.0 21-Mar-2020
# https://github.com/12Knocksinna/Office365itpros/blob/master/GetGraphUserStatisticsReport.PS1

CLS
# Define the values applicable for the application used to connect to the Graph
$AppId = "d716b32c-0edb-48be-9385-30a9cfd96165"
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f2478"
$AppSecret = 's_rkvIn1oZ1cNceUBvJ2or1lrrIsb*:='

# Construct URI and body needed for authentication
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials"
}

# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Unpack Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

# Base URL
$headers = @{Authorization = "Bearer $token"}
$ctype = "application/json"

Write-Host "Fetching user activity data from the Graph..."
# Get Teams Usage Data
$TeamsUserReportsURI = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D90')"
$TeamsUserData = (Invoke-RestMethod -Uri $TeamsUserReportsURI -Headers $Headers -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv 

# Get OneDrive for Business data
$OneDriveUsageURI = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D90')"
$OneDriveData = (Invoke-RestMethod -Uri $OneDriveUsageURI -Headers $Headers -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv 
 
# Get Exchange Activity Data
$EmailReportsURI = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='D90')"
$EmailData = (Invoke-RestMethod -Uri $EmailReportsURI -Headers $Headers -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv

# Get Exchange Storage Data   
$MailboxUsageReportsURI = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D90')"
$MailboxUsage = (Invoke-RestMethod -Uri $MailboxUsageReportsURI -Headers $Headers -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv

# Get SharePoint usage data
$SPOUsageReportsURI = "https://graph.microsoft.com/v1.0/reports/getSharePointActivityUserDetail(period='D90')"
$SPOUsage = (Invoke-RestMethod -Uri $SPOUsageReportsURI -Headers $Headers -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv

# Get Yammer usage data
$YammerUsageReportsURI = "https://graph.microsoft.com/v1.0/reports/getYammerActivityUserDetail(period='D90')"
$YammerUsage = (Invoke-RestMethod -Uri $YammerUsageReportsURI -Headers $Headers -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv

Write-Host "Processing activity data fetched from the Graph..."
# Create a list file to normalize and assemble the information we've collected from the Graph
$Report = [System.Collections.Generic.List[Object]]::new() 
# Process Teams Data
ForEach ($T in $TeamsUserData) {
   If ([string]::IsNullOrEmpty($T."Last Activity Date")) { 
      $TeamsLastActivity = "No activity"
      $TeamsDaysSinceActive = "N/A" }
   Else {
      $TeamsLastActivity = Get-Date($T."Last Activity Date") -format "dd-MMM-yyyy" 
      $TeamsDaysSinceActive = (New-TimeSpan($TeamsLastActivity)).Days }
   $ReportLine  = [PSCustomObject] @{          
     UPN               = $T."User Principal Name"
     LastActive        = $TeamsLastActivity  
     DaysSinceActive   = $TeamsDaysSinceActive      
     ReportDate        = Get-Date($T."Report Refresh Date") -format "dd-MMM-yyyy"  
     License           = $T."Assigned Products"
     ChannelChats      = $T."Team Chat Message Count"
     PrivateChats      = $T."Private Chat Message Count"
     Calls             = $T."Call Count"
     Meetings          = $T."Meeting Count"
     RecordType        = "Teams"}
   $Report.Add($ReportLine) } 

# Process Exchange Data
ForEach ($E in $EmailData) {
   If ([string]::IsNullOrEmpty($E."Last Activity Date")) { 
      $MailLastActivity = "No activity"
      $ExoDaysSinceActive = "N/A" }
   Else {
      $MailLastActivity = Get-Date($E."Last Activity Date") -format "dd-MMM-yyyy"
      $ExoDaysSinceActive = (New-TimeSpan($MailLastActivity)).Days }
  $ReportLine  = [PSCustomObject] @{          
     UPN                = $E."User Principal Name"
     DisplayName        = $E."Display Name"
     LastActive         = $MailLastActivity   
     DaysSinceActive    = $ExoDaysSinceActive    
     ReportDate         = Get-Date($E."Report Refresh Date") -format "dd-MMM-yyyy"  
     SendCount          = [int]$E."Send Count"
     ReadCount          = [int]$E."Read Count"
     ReceiveCount       = [int]$E."Receive Count"
     IsDeleted          = $E."Is Deleted"
     RecordType         = "Exchange Activity"}
   $Report.Add($ReportLine) } 
  
ForEach ($M in $MailboxUsage) {
   If ([string]::IsNullOrEmpty($M."Last Activity Date")) { 
      $MailLastActivity = "No activity" }
   Else {
      $MailLastActivity = Get-Date($M."Last Activity Date") -format "dd-MMM-yyyy"
      $ExoDaysSinceActive = (New-TimeSpan($MailLastActivity)).Days }
   $ReportLine  = [PSCustomObject] @{          
     UPN                = $M."User Principal Name"
     DisplayName        = $M."Display Name"
     LastActive         = $MailLastActivity 
     DaysSinceActive    = $ExoDaysSinceActive          
     ReportDate         = Get-Date($M."Report Refresh Date") -format "dd-MMM-yyyy"  
     QuotaUsed          = [Math]::Round($M."Storage Used (Byte)"/1GB,2) 
     Items              = [int]$M."Item Count"
     RecordType         = "Exchange Storage"}
   $Report.Add($ReportLine) } 

# SharePoint data
ForEach ($S in $SPOUsage) {
   If ([string]::IsNullOrEmpty($S."Last Activity Date")) { 
      $SPOLastActivity = "No activity"
      $SPODaysSinceActive = "N/A" }
   Else {
      $SPOLastActivity = Get-Date($S."Last Activity Date") -format "dd-MMM-yyyy"
      $SPODaysSinceActive = (New-TimeSpan ($SPOLastActivity)).Days }
   $ReportLine  = [PSCustomObject] @{          
     UPN              = $S."User Principal Name"
     LastActive       = $SPOLastActivity    
     DaysSinceActive  = $SPODaysSinceActive 
     ViewedEditedSPO  = [int]$S."Viewed or Edited File Count"     
     SyncedFileCount  = [int]$S."Synced File Count"
     SharedExtSPO     = [int]$S."Shared Externally File Count"
     SharedIntSPO     = [int]$S."Shared Internally File Count"
     VisitedPagesSPO  = [int]$S."Visited Page Count" 
     RecordType       = "SharePoint Usage"}
   $Report.Add($ReportLine) } 

# OneDrive for Business data
ForEach ($O in $OneDriveData) {
   $OneDriveLastActivity = $Null
   If ([string]::IsNullOrEmpty($O."Last Activity Date")) { 
      $OneDriveLastActivity = "No activity"
      $OneDriveDaysSinceActive = "N/A" }
   Else {
      $OneDriveLastActivity = Get-Date($O."Last Activity Date") -format "dd-MMM-yyyy" 
      $OneDriveDaysSinceActive = (New-TimeSpan($OneDriveLastActivity)).Days }
   $ReportLine  = [PSCustomObject] @{          
     UPN               = $O."Owner Principal Name"
     DisplayName       = $O."Owner Display Name"
     LastActive        = $OneDriveLastActivity    
     DaysSinceActive   = $OneDriveDaysSinceActive    
     OneDriveSite      = $O."Site URL"
     FileCount         = [int]$O."File Count"
     StorageUsed       = [Math]::Round($O."Storage Used (Byte)"/1GB,4) 
     Quota             = [Math]::Round($O."Storage Allocated (Byte)"/1GB,2) 
     RecordType        = "OneDrive Storage"}
   $Report.Add($ReportLine) } 

# Yammer Data
ForEach ($Y in $YammerUsage) {  
  If ([string]::IsNullOrEmpty($Y."Last Activity Date")) { 
      $YammerLastActivity = "No activity" 
      $YammerDaysSinceActive = "N/A" }
   Else {
      $YammerLastActivity = Get-Date($Y."Last Activity Date") -format "dd-MMM-yyyy" 
      $YammerDaysSinceActive = (New-TimeSpan ($YammerLastActivity)).Days }
  $ReportLine  = [PSCustomObject] @{          
     UPN             = $Y."User Principal Name"
     DisplayName     = $Y."Display Name"
     LastActive      = $YammerLastActivity      
     DaysSinceActive = $YammerDaysSinceActive   
     PostedCount     = [int]$Y."Posted Count"
     ReadCount       = [int]$Y."Read Count"
     LikedCount      = [int]$Y."Liked Count"
     RecordType      = "Yammer Usage"}
   $Report.Add($ReportLine) } 
 
# Get a list of users to process
CLS
$Users = $Report | Sort UPN -Unique | Select -ExpandProperty UPN
$ProgressDelta = 100/($Users.Count); $PercentComplete = 0; $UserNumber = 0
$StartDate = (Get-Date).AddDays(-30); $EndDate = (Get-Date) 
$OutData = [System.Collections.Generic.List[Object]]::new() # Create merged output file

# Process each user to extract Exchange, Teams, OneDrive, SharePoint, and Yammer statistics for their activity
ForEach ($U in $Users) {
  $UserNumber++
  $CurrentStatus = $U + " ["+ $UserNumber +"/" + $Users.Count + "]"
  Write-Progress -Activity "Extracting information for user" -Status $CurrentStatus -PercentComplete $PercentComplete
  $PercentComplete += $ProgressDelta
  $ExoData = $Null; $ExoActiveData = $Null; $TeamsData = $Null; $ODData = $Null; $SPOData = $Null; $YammerData = $Null
  # Process Exchange Storage Data
  $ExoData = $Report | ? {$_.UPN -eq $U -and $_.RecordType -eq "Exchange Storage"}
  $ExoActiveData = $Report | ? {$_.UPN -eq $U -and $_.RecordType -eq "Exchange Activity"}
  $TeamsData = $Report | ? {$_.UPN -eq $U -and $_.RecordType -eq "Teams"}
  $ODData = $Report | ? {$_.UPN -eq $U -and $_.RecordType -eq "OneDrive Storage"}
  $SPOData = $Report | ? {$_.UPN -eq $U -and $_.RecordType -eq "SharePoint Usage"}
  $YammerData = $Report | ? {$_.UPN -eq $U -and $_.RecordType -eq "Yammer Usage"}

If ((![string]::IsNullOrEmpty($ExoData.UPN))) {
# Build a line for the report file with the collected data for all workloads and write it to the list
  $ReportLine  = [PSCustomObject] @{          
     UPN                     = $ExoData.UPN
     DisplayName             = $ExoData.DisplayName
     EXOLastActive           = $ExoData.LastActive    
     EXODaysSinceActive      = $ExoData.DaysSinceActive     
     EXOQuotaUsed            = $ExoData.QuotaUsed
     EXOItems                = $ExoData.Items
     EXOSendCount            = $ExoActiveData.SendCount
     EXOReadCount            = $ExoActiveData.ReadCount
     EXOReceiveCount         = $ExoActiveData.ReceiveCount
     TeamsLastActive         = $TeamsData.LastActive
     TeamsDaysSinceActive    = $TeamsData.DaysSinceActive 
     TeamsChannelChat        = $TeamsData.ChannelChats
     TeamsPrivateChat        = $TeamsData.PrivateChats
     TeamsMeetings           = $TeamsData.Meetings
     TeamsCalls              = $TeamsData.Calls
     SPOLastActive           = $SPOData.LastActive
     SPODaysSinceActive      = $SPOData.DaysSinceActive 
     SPOViewedEditedFiles    = $SPOData.ViewedEditedSPO
     SPOSyncedFiles          = $SPOData.SyncedFileCount
     SPOSharedExtFiles       = $SPOData.SharedExtSPO
     SPOSharedIntFiles       = $SPOData.SharedIntSPO
     SPOVisitedPages         = $SPOData.VisitedPagesSPO
     OneDriveLastActive      = $ODData.LastActive
     OneDriveDaysSinceActive = $ODData.DaysSinceActive 
     OneDriveFiles           = $OneDriveFiles
     OneDriveStorage         = $ODData.StorageUsed
     OneDriveQuota           = $ODData.Quota
     YammerLastActive        = $YammerData.LastActive
     YammerDaysSinceActive   = $YammerData.DaysSinceActive 
     YammerPosts             = $YammerData.PostedCount
     YammerReads             = $YammerData.ReadCount
     YammerLikes             = $YammerData.LikedCount
     License                 = $TeamsData.License
     OneDriveSite            = $ODData.OneDriveSite
     IsDeleted               = $ExoActiveData.IsDeleted
     EXOReportDate           = $ExoData.ReportDate
     TeamsReportDate         = $TeamsData.ReportDate }
   $OutData.Add($ReportLine) } 
  }

Write-Host "Data processed for" $Users.Count "users"

