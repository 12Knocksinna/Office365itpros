# AuditRecordsTeamsRecordingsUploads.PS1
# A script to show how to use audit records in the Office 365 audit log to track uploads of Teams meeting recordings
# V2.0 updated 4-June-2024 to deal with the new mechanism for uploading Teams meeting recordings to OneDrive for Business
# article at https://office365itpros.com/2021/06/29/track-creation-teams-meeting-recordings/
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/AuditRecordsTeamsRecordingsUploads.PS1

Write-Host "Connecting to SharePoint Online to fetch details of OneDrive for Business sites..."
# This part is to generate a hash table of OneDrive sites that we can use to look up for audit records
Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell 
# Make sure that you use the right URL for your tenant here
Connect-SPOService -Url https://office365itpros-admin.sharepoint.com
[array]$OneDriveSites = Get-SPOSite -IncludePersonalSite $true -Limit All `
    -Filter "Url -like '-my.sharepoint.com/personal/'" | Select-Object URL, Owner
# Build the hash table
$OneDriveHashTable = @{}
ForEach ($Site in $OneDriveSites) { 
    [string]$SiteKey = $Site.URL + "/"
    $OneDriveHashTable.Add($SiteKey, $Site.Owner) 
}

# Connect to Exchange Online, if we're not already connected
$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

$StartDate = (Get-Date).AddDays(-180)
$EndDate = Get-Date
Write-Host "Searching for Teams recordings upload audit records..."
[array]$Records = Search-UnifiedAuditLog -Operations FileUploaded, FileModified `
 -StartDate $StartDate -EndDate $EndDate -Formatted -ResultSize 5000 -SessionCommand ReturnLargeSet `
 -UserIds "app@sharepoint" -RecordType SharePointFileOperation

If (!($Records)) {
    Write-Host "No audit records found - exiting!"; break
}

# Remove duplicates and make sure that we have a set sorted by date
$Records = $Records | Sort-Object Identity -Unique | Sort-Object {$_.CreationDate -as [datetime]} -Descending

$TaggedRecordings = [System.Collections.Generic.List[Object]]::new() 	
ForEach ($Rec in $Records) {
   $AuditData = $Rec.AuditData | ConvertFrom-Json
  
   If (($AuditData.SourceFileExtension -eq "mp4") -and ($AuditData.SourceRelativeUrl -like "*/Recordings") `
     -and $AuditData.SourceFileName.Substring(0,4) -ne "~tmp") { 
        $RecordingFileName = $AuditData.SourceFileName
        $DateLoc = $RecordingFileName.IndexOf("-202")
        If ($DateLoc -eq -1) {
            $Topic = $RecordingFileName
        } Else {
            $Topic = $RecordingFileName.SubString(0,$DateLoc)
        }   
        # All uploads are performed by the app@sharepoint account, so we try to use the hash table
        # to figure out the owner of the target OneDrive for Business account
        $User = $OneDriveHashTable[$AuditData.SiteURL]
        If ($null -eq $User) {
            $User = "SharePoint app"
        }
        $CreationDate = Get-Date $Rec.CreationDate -format 'dd-MMM-yyyy HH:mm:ss'
        $DataLine = [PSCustomObject] @{
            Workload            = $AuditData.Workload
            Operation           = $Rec.Operations
            Date                = $CreationDate
            User                = $User
            Recording           = $RecordingFileName
            "Meeting title"     = $Topic
            Site                = $AuditData.SiteURL
            FullURL             = $AuditData.ObjectId
            Folder              = $AuditData.SourceRelativeURL
         }
    $TaggedRecordings.Add($DataLine) 

   } #End If
} #End For

$TaggedRecordings = $TaggedRecordings | Sort-Object {$_.Date -as [datetime]} -Descending
$CSVOutputFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\TeamsRecordings.csv"
$TaggedRecordings | Out-GridView -Title 'Teams Recordings Uploads' 
$TaggedRecordings | Export-CSV -NoTypeInformation $CSVOutputFile -Encoding utf8

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production envi