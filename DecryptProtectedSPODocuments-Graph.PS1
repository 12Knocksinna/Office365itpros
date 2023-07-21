# Requires -Version 3.0
# DecryptProtectedSPODocuments-Graph.PS1
# A script to show how to use Microsoft Graph calls to retrieve a set of documents protected by sensitivity labels 
# with encryption from a selected SharePoint Online site (and optional folder) and decrypt them by running the 
# SPO cmdlet Unlock-SPOSensitivityLabelEncryptedFile
# https://github.com/12Knocksinna/Office365itpros/blob/master/DecryptProtectedSPODocuments-Graph.PS1
#
# Usage .\DecryptProtectedSPODocuments-Graph.PS1 -SearchSite MySite -SearchFolder OptionalFolder
# e.g.  .\DecryptProtectedSPODocuments-Graph.PS1 -SearchSite Projects -SearchFolder "Blog Posts"
# Make sure to fill in all the required variables before running the script
# Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Sites.ReadWrite.All (Application)
#    Sites.Read.All      (Application)

Param ([Parameter(Mandatory)]$SearchSite, [string]$SearchFolder)

If (!$SearchFolder) {
   Write-Host "We're going to search" $SearchSite }
Else {
   Write-Host "We're going search" $SearchSite "site and folder" $SearchFolder }

# These values need to be changed to reflect the registered app (in Azure AD) and tenant details
$AppId = "72bd89d6-060a-43c9-8063-c281d8f8b685"
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3478"
$AppSecret = "Dz6A1tbmc1Z-H4qsJxALy.-JR2_gj-1E~1"
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

# https://graph.microsoft.com/v1.0/sites?search=* for all sites
# https://graph.microsoft.com/v1.0/sites?search="Corporate Accounting (Billing)"

$URI = "https://graph.microsoft.com/v1.0/sites?search='$($SearchSite)'"
[array]$Site = (Invoke-RestMethod -Uri $URI -Headers $Headers -Method Get -ContentType "application/json")

If ($Site.Value.Count -eq 0) { # Nothing found
     Write-Host "No matching sites found - exiting"; break }
If ($Site.Value.Count -eq 1) { # Only one site found - go ahead
     $SiteId = $Site.Value.Id 
     $SiteName = $Site.Value.DisplayName
     Write-Host "Found site to process:" $SiteName }
Elseif ($Site.Value.Count -gt 1) { # More than one site found. Ask which to use
     CLS; Write-Host "More than one matching site was found."; [int]$i=1
     ForEach ($SiteOption in $Site.Value) {
        Write-Host $i ":" $SiteOption.DisplayName; $i++}
     [Int]$Answer = Read-Host "Enter the number of the site to use"
     If (($Answer -gt 0) -and ($Answer -le $i)) {
        [int]$Si = ($Answer-1)
        $SiteName = $Site.Value[$Si].DisplayName 
        Write-Host "OK. Selected site is" $Site.Value[$Si].DisplayName 
        $SiteId = $Site.Value[$Si].Id }
}
 
If ($SearchFolder) { # We've been asked to look in a specific folder, so find its drive id
   $Uri = "https://graph.microsoft.com/v1.0/sites/$($Siteid)/lists/Documents/Drive/root/children"
   $SiteData = (Invoke-RestMethod -Uri $URI -Headers $Headers -Method Get -ContentType "application/json")
   # Look for the target folder in the set of resources returned
   $TargetFolder = $SiteData.Value | Where-Object {$_.Name -eq $SearchFolder -and $_.Folder -ne $Null}

   If ($TargetFolder) { # We found the folder   
      $DriveId = $TargetFolder.Id }
   Else { # We didn't... so exit to let the user try again
      $DriveId = $Null
      Write-Host "Can't find the" $SearchFolder "folder" in the $SiteName "site"; break }}
Else { # Search folder isn't defined, so we look in the default folder of the document library in chosen site
      $Uri = "https://graph.microsoft.com/v1.0/sites/$($Siteid)/lists/Documents/Drive/root/"
      $SiteData = (Invoke-RestMethod -Uri $URI -Headers $Headers -Method Get -ContentType "application/json")
      $TargetFolder = $SiteData.Id }

# Retrieve files in the folder, including sensitivity label info. SharePoint returns a default of 200 files per call, so we use the nextlink to keep on fetching files until we are done
$BaseUrl = $TargetFolder.WebUrl + "/"
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 

Write-Host "Searching for files in the target site/folder"
If (!$SearchFolder) { # Search the root folder of the site
   $Uri = "https://graph.microsoft.com/v1.0/sites/$($Siteid)/lists/Documents/Drive/root/children?`$select=sensitivitylabel,weburl,name" }
Else { # Search the nominated folder
   $Uri = "https://graph.microsoft.com/v1.0/sites/$($Siteid)/lists/Documents/Drive/Items/$($DriveId)/children?`$select=sensitivitylabel,weburl,name"}

$Files = (Invoke-RestMethod -Uri $URI -Headers $Headers -Method Get -ContentType "application/json")
$FilesCount = $Files.Value.Count
If ($FilesCount -eq 0) { # No files found in that location
   Write-Host "No files can be found in that location - exiting";break}

ForEach ($File in $Files.Value) {
    If ($File.SensitivityLabel.ProtectionEnabled -eq $True) {
       $FileName = $BaseUrl + $File.Name
       $ReportLine = [PSCustomObject] @{
         File      = $File.Name
         FileURL   = $FileName
         Label     = $File.SensitivityLabel.DisplayName
         LabelGuid = $File.SensitivityLabel.Id  }               
    $Report.Add($ReportLine) } #End If
} # End For

$NextLink = $Files.'@Odata.NextLink'
While ($NextLink -ne $Null) {
   $Files = (Invoke-RestMethod -Uri $Nextlink -Headers $Headers -Method Get -ContentType "application/json")
   $FilesCount = $Files.Value.Count + $FilesCount
   ForEach ($File in $Files.Value) {
    If ($File.SensitivityLabel.ProtectionEnabled -eq $True) {
       $FileName = $BaseUrl + $File.Name
       $ReportLine = [PSCustomObject] @{
         File      = $File.Name
         FileURL   = $FileName
         Label     = $File.SensitivityLabel.DisplayName
         LabelGuid = $File.SensitivityLabel.Id  }               
    $Report.Add($ReportLine) } #End If
} # End For 
   $NextLink = $Files.'@odata.NextLink' }

# Prompt to go ahead with decryption after clearing screen.
CLS; [int]$i = 0
If ($Report.Count -eq 0) {
   Write-Host "No encrypted files found in" $SiteName "- exiting"; break }
Else {
   Write-Host $Report.Count "of" $FilesCount "files in" $SiteName "have sensitivity labels with encryption" }    

# List documents found and ask whether to proceed
$Report | Format-Table File, Label -AutoSize
$PromptTitle = 'Remove Encryption from documents'
$PromptMessage = 'Please confirm whether to go ahead and remove encryption from these files'
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&yes", 'yes?'
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&no", 'no?'
$cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&cancel", 'Exit'
$PromptOptions = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $cancel)
$PromptDecision = $host.ui.PromptForChoice($PromptTitle, $PromptMessage, $PromptOptions, 0) 

# Decision made Y or default to go ahead and remove protection from the documents, so let's do it.
If ($PromptDecision -eq 0) {
   ForEach ($F in $Report) {
     Write-Host "Removing encryption from" $F.File
     Unlock-SPOSensitivityLabelEncryptedFile -FileUrl $F.FileUrl -JustificationText "Administrator removed label"
     $i++ }
   Write-Host "All done. Encryption removed from $i files" }
Else {
   Write-Host "OK. Details of the encrypted files are in c:\temp\EncryptedDocuments.csv"
   $Report | Export-CSV -NoTypeInformation c:\temp\EncryptedDocuments.csv }

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
