# AuditSPOOperationsByUser.PS1
# Script to generate a per-user CSV file of SharePoint Online document operations
# V1.0 4-May-2020
# Requires the Exchange Online PowerShell module and/or Azure Active Directory module depending on if you want to
# report on a tenant mailbox basis or tenant and guest user accounts.
# https://github.com/12Knocksinna/Office365itpros/blob/master/AuditSPOOperationsByUser.PS1

Param ([String]$CheckUsers)
CLS
# Check do we have the right modules loaded....
$Module = Get-Module 
$AzureAD = $Module | ? {$_.Name -like "*AzureAD*"}
If (!$AzureAD) { Write-Host "This script needs the Azure Active Directory PowerShell module" ; break}
$Exo = $Module | ? {$_.Name -eq "ExchangeOnlineManagement"}
If (!$Exo) { Write-Host "This script needs the Exchange Online Management PowerShell module" ; break }

# Check what our parameter is. If no parameter is passed, we check both tenant users and guests. If Tenant is passed, we only check mailbox owners.
Switch ($CheckUsers) {
   "Tenant" { # Just check tenant users
       $Type = "(Mailbox owners)"
       $Users = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Select UserPrincipalName,DisplayName }
     default { #Check everyone
       $Type = "(Tenant and guest user accounts)"
       $Users = Get-AzureADUser -All $True -Filter ("UserType eq 'Guest' or UserType eq 'Member'") | Select UserPrincipalName, DisplayName }
} #End switch

Write-Host $Users.Count "accounts will be checked." $Type

$StartDate = (Get-Date).AddDays(-90); $EndDate = (Get-Date).AddDays(+1); $DroppedRecords = 0; $StartProcessingTime = (Get-Date)  
$Operations = @('FileAccessed', 'FileDownloaded', 'FileModified', 'FileDeleted', 'FileUploaded')
# Get a unique session identifier  
$Sessionid = (New-Guid).Guid 
$i = 0; $TotalRecordsProcessed = 0
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
$Records = 1
While ($Records) { #Keep fetching until we have processed all records
  $i++
  Write-Host "Fetching audit records...( batch" $i ")"
  $Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations $Operations -ResultSize 2000 -SessionId $SessionId -SessionCommand ReturnLargeSet
  
  If ($i -eq 1) {$TotalRecordsAvailable = $Records[0].ResultCount}
  If ($Records.Count -ne 0) {
    $StatusProcessing = "Processing " + $Records.Count + " audit records of " + $TotalRecordsAvailable + ". Total processed so far: " + $TotalRecordsProcessed
    Write-Host $StatusProcessing}

  ForEach ($Rec in $Records) {
     $AuditData = $Rec.AuditData | ConvertFrom-JSon
     # Drop all the graphic files and web pages for which SharePoint Online creates audit records
     # Also drop records where the useragent isn't recorded because of system processing and any created by App@SharePoint
     If ($AuditData.SourceFileName -like "*.jpg*" -or $AuditData.SourceFileName -like "*.aspx*" -or $AuditData.SourceFileName -like "*.png*" -or [string]::IsNullOrEmpty($AuditData.UserAgent) -or $AuditData.UserId -eq "app@sharepoint" ) { $DroppedRecords++ }
        Else {
        # Make some sense of the User agent string
        Switch ($AuditData.UserAgent) { # This switch is used to interpret the UserAgent string and make it more human friendly
          {$AuditData.UserAgent -like "*Word 2014*"} {
                $UserAgent = "Microsoft Word 2014 (Mac)"
                $Version = "2014"; break}
          {$AuditData.UserAgent -like "*Edg/*"} { 
                $UserAgent = "Microsoft Edge"
                $Version = $AuditData.UserAgent.Split("/")[5] ; break}
          {$AuditData.UserAgent -like "*Microsoft Office Excel*"} { # Excel desktop
                $UserAgent = "Microsoft Excel (desktop)"
                $Version = $AuditData.UserAgent.Split("/")[1].Split("(")[0] ; break }
          {$AuditData.UserAgent -Like "*Faraday*"} { #Yammer process for uploading and removing files
                $UserAgent = "Yammer"
                $Version = $AuditData.UserAgent.Split("v")[1] ; break }
          {$AuditData.UserAgent -like "*Skydrive*"} { # OneDrive for Business
                $UserAgent = "OneDrive for Business"
                $Version = $AuditData.UserAgent.Split("nc ")[5] ; break }
          {$AuditData.UserAgent -like "*Microsoft Office PowerPoint*"} { #Word desktop
                $UserAgent = "Microsoft PowerPoint (desktop)"
                $Version = $AuditData.UserAgent.Split("/")[1].Split("(")[0]; break }
          {$AuditData.UserAgent -like "*Microsoft Office Word*"} { #Word desktop
                $UserAgent = "Microsoft Word (desktop)"
                $Version = $AuditData.UserAgent.Split("/")[1].Split("(")[0]; break }
          {$AuditData.UserAgent -like "*Teams/*"} {
                $UserAgent = "Teams" 
                $Version = $AuditData.UserAgent.Split("/")[3].Split("Ch")[0]; break}
          {$AuditData.UserAgent -like "*Gecko*" -and $AuditData.UserAgent -like "*Chrome*"} { 
                $UserAgent = "Chromium browser"
                $Version = $AuditData.UserAgent.Split("/")[3].Split("Safari")[0]}
          "OneDriveMpc-Transform_Thumbnail/1.0" {$UserAgent = "OneDrive"}
          "MSWAC" {
                $UserAgent = "Office Online"
                $Version = "N/A" }
          "MSOCS" {
                $UserAgent = "Office Online"
                $Version = "N/A" }
          "MSWACONSync" {
                $UserAgent = "OneDrive Online"
                $Version = "N/A"}
        default { 
                $UserAgent = $AuditData.UserAgent
                $Version = "N/A" }
        } # End Switch
        If ($Null -eq $UserAgent) {$UserAgent = "Not Recorded"}
        $ReportLine  = [PSCustomObject] @{
           TimeStamp       = Get-Date ($AuditData.CreationTime) -format g
           UPN             = $AuditData.UserId
           Operation       = $AuditData.Operation
           FileName        = $AuditData.SourceFileName
           Object          = $AuditData.ObjectId
           Site            = $AuditData.SiteUrl         
           ClientIP        = $AuditData.ClientIP
           UserAgent       = $UserAgent
           Version         = $Version
           FullClientInfo  = $AuditData.UserAgent}
         $Report.Add($ReportLine) }
   }
   # Update number of processed records 
   $TotalRecordsProcessed = $TotalRecordsProcessed + $Records.Count
   
}
CLS
# Because we've fetched large amounts of unsorted data using the ReturnLargeSet parameter, we should sort it to make sure that it's in the correct order
$Report = $Report | Sort {$_.TimeStamp -as [DateTime]}  -Descending 
$ProgressDelta = 100/($Users.count); $PercentComplete = 0; $UserNumber = 0; $UserReports = 0
ForEach ($U in $Users) {
    $UserNumber++
    $Status = "Checking SharePoint activities for " + $U.DisplayName +  " ["+ $UserNumber +"/" + $Users.Count + "]"
    Write-Progress -Activity "Processing audit data" -Status $Status -PercentComplete $PercentComplete
    $PercentComplete += $ProgressDelta
    $UserRecords = $Report | ? {$_.UPN -eq $U.UserPrincipalName} 
    If ($UserRecords) {  
       $UserReports++  
       Write-Host "Writing out data for" $U.DisplayName
       $FileName = "c:\Temp\AuditHistory" + $U.UserPrincipalName + ".csv"
       $UserRecords | Export-CSV -NoTypeInformation $FileName }
}
$OutputFiles = Get-ChildItem -Path c:\temp\ |?{$_.LastWriteTime -ge $StartProcessing}  
Write-Host " "
Write-Host "All done." $UserReports "audit reports for SharePoint activity are available in c:\temp\"
Write-Host "------------------------------------------------------------------------------"
$OutputFiles | ?{$_.Name -Like "*AuditHistory*"} | Ft Name     

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
