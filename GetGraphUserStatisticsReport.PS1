# GetGraphUserStatisticsReport.PS1
# A sample script showing how to gather user activity information from the Graph and assemble it into one report
# V1.3 2-Sep-2020
# https://github.com/12Knocksinna/Office365itpros/blob/master/GetGraphUserStatisticsReport.PS1
# See https://petri.com/graph-powershell-office365-usage for a description of what this script does
# And see https://github.com/12Knocksinna/Office365itpros/blob/master/GetGraphUserStatisticsReportV2.PS1 for a version with
# better performance.
# Note: Guest user activity is not recorded by the Graph - only tenant accounts are processed
# Needs the Reports.Read.All permission to get user login data

CLS
# Define the values applicable for the application used to connect to the Graph (change these for your tenant)
$AppId = "d716b32c-0edb-48be-9385-30a9cfd96155"
$TenantId = "c662313f-14fc-43a2-9a7a-d2e27f4f3478"
$AppSecret = 's_rkvIn1oZ1cNceUBvJ2or1lrrIsb*:='
$TargetDir = 'c:\temp'

# Construct URI and body needed for authentication
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials" }

# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Unpack Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

# Base URL
$headers = @{Authorization = "Bearer $token"}

Write-Host "Fetching Teams user activity data from the Graph..."
# The Graph returns information in CSV format. We convert it to allow the data to be more easily processed by PowerShell
# Get Teams Usage Data - the replace parameter is there to remove three odd leading characters (ï»¿) in the CSV data returned by the 
$TeamsUserReportsURI = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D90')"
$TeamsUserData = (Invoke-RestMethod -Uri $TeamsUserReportsURI -Headers $Headers -Method Get -ContentType "application/json") -Replace "...Report Refresh Date", "Report Refresh Date" | ConvertFrom-Csv 

Write-Host "Fetching OneDrive for Business user activity data from the Graph..."
# Get OneDrive for Business data
$OneDriveUsageURI = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D90')"
$OneDriveData = (Invoke-RestMethod -Uri $OneDriveUsageURI -Headers $Headers -Method Get -ContentType "application/json") -Replace "...Report Refresh Date", "Report Refresh Date" | ConvertFrom-Csv 
 
Write-Host "Fetching Exchange Online user activity data from the Graph..."
# Get Exchange Activity Data
$EmailReportsURI = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='D90')"
$EmailData = (Invoke-RestMethod -Uri $EmailReportsURI -Headers $Headers -Method Get -ContentType "application/json") -Replace "...Report Refresh Date", "Report Refresh Date" | ConvertFrom-Csv 

# Get Exchange Storage Data   
$MailboxUsageReportsURI = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D90')"
$MailboxUsage = (Invoke-RestMethod -Uri $MailboxUsageReportsURI -Headers $Headers -Method Get -ContentType "application/json") -Replace "...Report Refresh Date", "Report Refresh Date" | ConvertFrom-Csv 
Write-Host "Fetching SharePoint Online user activity data from the Graph..."
# Get SharePoint usage data
$SPOUsageReportsURI = "https://graph.microsoft.com/v1.0/reports/getSharePointActivityUserDetail(period='D90')"
$SPOUsage = (Invoke-RestMethod -Uri $SPOUsageReportsURI -Headers $Headers -Method Get -ContentType "application/json") -Replace "...Report Refresh Date", "Report Refresh Date" | ConvertFrom-Csv 

Write-Host "Fetching Yammer user activity data from the Graph..."
# Get Yammer usage data
$YammerUsageReportsURI = "https://graph.microsoft.com/v1.0/reports/getYammerActivityUserDetail(period='D90')"
$YammerUsage = (Invoke-RestMethod -Uri $YammerUsageReportsURI -Headers $Headers -Method Get -ContentType "application/json") -Replace "...Report Refresh Date", "Report Refresh Date" | ConvertFrom-Csv 

# Create hash table for user sign in data
$UserSignIns = @{}
# Get User sign in data
Write-Host "Fetching user sign-in data from the Graph..."
$URI = "https://graph.microsoft.com/beta/users?`$select=displayName,userPrincipalName, mail, id, CreatedDateTime, signInActivity, UserType&`$top=999"
$SignInData = (Invoke-RestMethod -Uri $URI -Headers $Headers -Method Get -ContentType "application/json") 
# Update the user sign in hash table
ForEach ($U in $SignInData.Value) {
   If ($U.UserType -eq "Member") {
     If ($U.SignInActivity.LastSignInDateTime) {
          $LastSignInDate = Get-Date($U.SignInActivity.LastSignInDateTime) -format g
          $UserSignIns.Add([String]$U.UserPrincipalName, $LastSignInDate) }
}}

# Do we have extra data to fetch?
$NextLink = $SignInData.'@Odata.NextLink'
# If we have a next link, go and process the remaining set of users
While ($NextLink -ne $Null) { 
   Write-Host "Still processing..."
   $SignInData = Invoke-WebRequest -Method GET -Uri $NextLink -ContentType "application/json" -Headers $Headers -UseBasicParsing
   $SignInData = $SignInData | ConvertFrom-JSon
   ForEach ($U in $SignInData.Value) {  
   If ($U.UserType -eq "Member") {
     If ($U.SignInActivity.LastSignInDateTime) {
          $LastSignInDate = Get-Date($U.SignInActivity.LastSignInDateTime) -format g
          $UserSignIns.Add([String]$U.UserPrincipalName, $LastSignInDate) }
   }}
   $NextLink = $SignInData.'@Odata.NextLink'
} # End while

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
   $ExoDaysSinceActive = $Null
   If ([string]::IsNullOrEmpty($E."Last Activity Date")) { 
      $ExoLastActivity = "No activity"
      $ExoDaysSinceActive = "N/A" }
   Else {
      $ExoLastActivity = Get-Date($E."Last Activity Date") -format "dd-MMM-yyyy"
      $ExoDaysSinceActive = (New-TimeSpan($ExoLastActivity)).Days }
  $ReportLine  = [PSCustomObject] @{          
     UPN                = $E."User Principal Name"
     DisplayName        = $E."Display Name"
     LastActive         = $ExoLastActivity   
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
      $ExoLastActivity = "No activity" }
   Else {
      $ExoLastActivity = Get-Date($M."Last Activity Date") -format "dd-MMM-yyyy"
      $ExoDaysSinceActive = (New-TimeSpan($ExoLastActivity)).Days }
   $ReportLine  = [PSCustomObject] @{          
     UPN                = $M."User Principal Name"
     DisplayName        = $M."Display Name"
     LastActive         = $ExoLastActivity 
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
[array]$Users = $Report | Sort UPN -Unique | Select -ExpandProperty UPN
$ProgressDelta = 100/($Users.Count); $PercentComplete = 0; $UserNumber = 0
$OutData = [System.Collections.Generic.List[Object]]::new() # Create merged output file

# Process each user to extract Exchange, Teams, OneDrive, SharePoint, and Yammer statistics for their activity
ForEach ($U in $Users) {
  $UserNumber++
  $CurrentStatus = $U + " ["+ $UserNumber +"/" + $Users.Count + "]"
  Write-Progress -Activity "Extracting information for user" -Status $CurrentStatus -PercentComplete $PercentComplete
  $PercentComplete += $ProgressDelta
  $ExoData = $Null; $ExoActiveData = $Null; $TeamsData = $Null; $ODData = $Null; $SPOData = $Null; $YammerData = $Null
  
  $UserData = $Report.Where({$_.UPN -eq $U})  # Extract records from list for the user

# Process Exchange Data
  $ExoData = $UserData.Where({$_.RecordType -eq "Exchange Storage"})
  $ExoActiveData = $UserData.Where({$_.RecordType -eq "Exchange Activity"})
  
  If ($ExoActiveData -eq $Null -or $ExoActiveData.LastActive -eq "No Activity") {
     $ExoLastActive       = "No Activity"
     $ExoDaysSinceActive  = "N/A" }
  Else {
     $ExoLastActive = $ExoActiveData.LastActive
     $ExoDaysSinceActive = $ExoActiveData.DaysSinceActive }

  $TeamsData = $UserData.Where({$_.RecordType -eq "Teams"})
  
  $SPOData = $UserData.Where({$_.RecordType -eq "SharePoint Usage"})
 
# Parse OneDrive for Business usage data 
  $ODData = $UserData.Where({$_.RecordType -eq "OneDrive Storage"})
  If ($ODData -eq $Null -or $ODData.LastActive -eq "No Activity") {
     $ODLastActive       = "No Activity"
     $ODDaysSinceActive  = "N/A"
     $ODFiles            = 0
     $ODStorage          = 0
     $ODQuota            = 1024 }
 Else {
     $ODLastActive = $ODData.LastActive
     $ODDaysSinceActive = $ODData.DaysSinceActive
     $ODFiles = $ODData.FileCount
     $ODStorage = $ODData.StorageUsed
     $ODQuota = $ODData.Quota  }

# Parse Yammer usage data
  $YammerData = $UserData.Where({$_.RecordType -eq "Yammer Usage"})
# Yammer isn't used everywhere, so make sure that we record zero data 
  If ($YammerData -eq $Null -or $YammerData.LastActive -eq "No Activity") {
     $YammerLastActive       = "No Activity"
     $YammerDaysSinceActive  = "N/A" 
     $YammerPosts             = 0
     $YammerReads             = 0
     $YammerLikes             = 0 }
 Else {
     $YammerLastActive = $YammerData.LastActive
     $YammerDaysSinceActive = $YammerData.DaysSinceActive
     $YammerPosts = $YammerData.PostedCount
     $YammerReads = $YammerData.ReadCount
     $YammerLikes = $YammerData.LikedCount }
   
# Fetch the sign in data if available
$LastAccountSignIn = $Null; $DaysSinceSignIn = 0
$LastAccountSignIn = $UserSignIns.Item($U)
If ($LastAccountSignIn -eq $Null) { $LastAccountSignIn = "No sign in data found"; $DaysSinceSignIn = "N/A"}
  Else { $DaysSinceSignIn = (New-TimeSpan($LastAccountSignIn)).Days }
   
# Figure out if the account is used
# Base is 2 if someuse uses the five workloads because the Graph is usually 2 days behind, but we have some N/A values for days used
  If ($ExoActiveData.DaysSinceActive -eq "N/A") {$ExoDays = 365} Else {$ExoDays = $ExoActiveData.DaysSinceActive}
  If ($TeamsData.DaysSinceActive -eq "N/A") {$TeamsDays = 365} Else {$TeamsDays = $TeamsData.DaysSinceActive}
  If ($SPOData.DaysSinceActive -eq "N/A") {$SPODays = 365} Else {$SPODays = $SPOData.DaysSinceActive}  
  If ($ODDaysSinceActive -eq "N/A") {$ODDays = 365} Else {$ODDays = $ODDaysSinceActive}  
  If ($YammerDaysSinceActive -eq "N/A") {$YammerDays = 365} Else {$YammerDays = $YammerDaysSinceActive}
   
# Average days per workload used...
  $AverageDaysSinceUse = [Math]::Round((($ExoDays + $TeamsDays + $SPODays + $ODDays + $YammerDays)/5),2)

  Switch ($AverageDaysSinceUse) { # Figure out if account is used
   ({$PSItem -le 8})                          { $AccountStatus = "Heavy usage" }
   ({$PSItem -ge 9 -and $PSItem -le 50} )     { $AccountStatus = "Moderate usage" }   
   ({$PSItem -ge 51 -and $PSItem -le 120} )   { $AccountStatus = "Poor usage" }
   ({$PSItem -ge 121 -and $PSItem -le 300 } ) { $AccountStatus = "Review account"  }
   default                                    { $AccountStatus = "Account unused" }
  } # End Switch

# And an override if someone has been active in just one workload in the last 14 days
  [int]$DaysCheck = 14 # Set this to your chosen value if you want to use a different period.
  If (($ExoDays -le $DaysCheck) -or ($TeamsDays -le $DaysCheck) -or ($SPODays -le $DaysCheck) -or ($ODDays -le $DaysCheck) -or ($YammerDays -le $DaysCheck)) {
     $AccountStatus = "Account in use"}

If ((![string]::IsNullOrEmpty($ExoData.UPN))) {
# Build a line for the report file with the collected data for all workloads and write it to the list
  $ReportLine  = [PSCustomObject] @{          
     UPN                     = $ExoData.UPN
     DisplayName             = $ExoData.DisplayName
     Status                  = $AccountStatus
     LastSignIn              = $LastAccountSignIn
     DaysSinceSignIn         = $DaysSinceSignIn 
     EXOLastActive           = $ExoLastActive  
     EXODaysSinceActive      = $ExoDays    
     EXOQuotaUsed            = $ExoData.QuotaUsed
     EXOItems                = $ExoData.Items
     EXOSendCount            = $ExoActiveData.SendCount
     EXOReadCount            = $ExoActiveData.ReadCount
     EXOReceiveCount         = $ExoActiveData.ReceiveCount
     TeamsLastActive         = $TeamsData.LastActive
     TeamsDaysSinceActive    = $TeamsDays 
     TeamsChannelChat        = $TeamsData.ChannelChats
     TeamsPrivateChat        = $TeamsData.PrivateChats
     TeamsMeetings           = $TeamsData.Meetings
     TeamsCalls              = $TeamsData.Calls
     SPOLastActive           = $SPOData.LastActive
     SPODaysSinceActive      = $SPODays 
     SPOViewedEditedFiles    = $SPOData.ViewedEditedSPO
     SPOSyncedFiles          = $SPOData.SyncedFileCount
     SPOSharedExtFiles       = $SPOData.SharedExtSPO
     SPOSharedIntFiles       = $SPOData.SharedIntSPO
     SPOVisitedPages         = $SPOData.VisitedPagesSPO
     OneDriveLastActive      = $ODLastActive
     OneDriveDaysSinceActive = $ODDays 
     OneDriveFiles           = $ODFiles
     OneDriveStorage         = $ODStorage
     OneDriveQuota           = $ODQuota
     YammerLastActive        = $YammerLastActive  
     YammerDaysSinceActive   = $YammerDays 
     YammerPosts             = $YammerPosts
     YammerReads             = $YammerReads
     YammerLikes             = $YammerLikes
     License                 = $TeamsData.License
     OneDriveSite            = $ODData.OneDriveSite
     IsDeleted               = $ExoActiveData.IsDeleted
     EXOReportDate           = $ExoData.ReportDate
     TeamsReportDate         = $TeamsData.ReportDate
     UsageFigure             = $AverageDaysSinceUse }
   $OutData.Add($ReportLine) } 
 } #End processing user data

Write-Host "Data processed for" $Users.Count "users"

$OutData | Sort {$_.ExoLastActive -as [DateTime]} -Descending | Out-GridView  
$OutData | Sort $AccountStatus | Export-CSV $TargetDir\Office365TenantUsage.csv -NoTypeInformation

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
