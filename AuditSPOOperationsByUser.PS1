# AuditSPOOperationsByUser.PS1
# Script to generate a per-user CSV file of SharePoint Online document operations
# V1.0 4-May-2020
# Requires the Exchange Online PowerShell module and/or Azure Active Directory module depending on if you want to
# report on a tenant mailbox basis or tenant and guest user accounts.

$StartDate = (Get-Date).AddDays(-7); $EndDate = (Get-Date).AddDays(+1); $DroppedRecords = 0; $StartProcessingTime = (Get-Date)  
$Operations = @('FileAccessed', 'FileDownloaded', 'FileUpdated', 'FileModified')
CLS
# Check only mailboxes
$Users = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Select UserPrincipalName,DisplayName
# Check guest users too
# $Users = Get-AzureADUser -All $True -Filter ("UserType eq 'Guest' or UserType eq 'Member'") | Select Userprincipalname, DisplayName     
Write-Host "Fetching audit records..."
$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations $Operations -ResultSize 5000

Write-Host "Processing" $Records.Count "Audit records..."
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
ForEach ($Rec in $Records) {
     $AuditData = $Rec.AuditData | ConvertFrom-JSon
     # Drop all the graphic files and web pages for which SharePoint Online creates audit records
     # Also drop records where the useragent isn't recorded because of system processing
     If ($AuditData.SourceFileName -like "*.jpg*" -or $AuditData.SourceFileName -like "*.aspx*" -or $AuditData.SourceFileName -like "*.png*" -or [string]::IsNullOrEmpty($AuditData.UserAgent) ) { $DroppedRecords++ }
        Else {
        # Make some sense of the User agent string
        Switch ($AuditData.UserAgent) {
          {$AuditData.UserAgent -like "*Word 2014*"} {
                $UserAgent = "Microsoft Word 2014 (Mac)"
                $Version = "2014"; break}
          {$AuditData.UserAgent -like "*Edg/*"} { 
                $UserAgent = "Microsoft Edge"
                $Version = $AuditData.UserAgent.Split("/")[5] ; break}
          {$AuditData.UserAgent -like "*Microsoft Office Excel*"} {
                $UserAgent = "Microsoft Excel (desktop)"
                $Version = $AuditData.UserAgent.Split("/")[1].Split("(")[0] ; break }
          {$AuditData.UserAgent -like "*Microsoft Office Word*"} {
                $UserAgent = "Microsoft Word (desktop)"
                $Version = $AuditData.UserAgent.Split("/")[1].Split("(")[0]; break }
          ($AuditData.UserAgent -like "*Teams/*"} {
                $UserAgent = "Teams" 
                $Version = $AuditData.UserAgent.Split("/")[3].Split("Ch")[0]; break}
          {$AuditData.UserAgent -like "*Gecko*" -and $AuditData.UserAgent -like "*Chrome*"} { 
                $UserAgent = "Chromium browser"
                $Version = $AuditData.UserAgent.Split("/")[3].Split("Safari")[0]}
          "OneDriveMpc-Transform_Thumbnail/1.0" {$UserAgent = "OneDrive"}
          "MSWAC"       {
                $UserAgent = "OneDrive Online"
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
           TimeStamp       = Get-Date($AuditData.CreationTime) -format g
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
CLS
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
Write-Host "All done." $UserReports "audit reports for SharePoint activity are available in c:\temp\"
Write-Host "------------------------------------------------------------------------------"
$OutputFiles | ?{$_.Name -Like "*AuditHistory*"} | Ft Name        
    
