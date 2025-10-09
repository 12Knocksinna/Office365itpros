
	


Connect-MgGraph -Scopes Mail.ReadBasic.All, UserReadBasic.All, Mail.Send

function Get-MailboxSizeForUser {
param (
[string]$UserId
)

function Get-FolderSizeRecursive {
param (
[string]$UserId,
[string]$FolderId,
[string]$Indent = $null
)

$size = 0
$folder = Get-MgUserMailFolder -UserId $UserId -MailFolderId $FolderId
$folderName = $folder.DisplayName
$folderSize = $folder.AdditionalProperties[“sizeInBytes”]

Write-Host “$Indent Folder: $folderName”
Write-Host “$Indent Size: $([math]::Round($folderSize / 1MB, 2)) MB”

if ($folderSize) {
$size += $folderSize
}

if ($folder.ChildFolderCount -gt 0) {
$childFolders = Get-MgUserMailFolderChildFolder -UserId $UserId -MailFolderId $FolderId -All
foreach ($child in $childFolders) {
$size += Get-FolderSizeRecursive -UserId $UserId -FolderId $child.Id -Indent (“$Indent`t”)
}
}

return $size
}

$totalSize = 0
$rootFolders = Get-MgUserMailFolder -UserId $UserId -All
foreach ($folder in $rootFolders) {
$totalSize += Get-FolderSizeRecursive -UserId $UserId -FolderId $folder.Id
}

$totalSizeGB = [math]::Round($totalSize / 1GB, 2)
Write-Host “Total mailbox size for $UserId : $totalSizeGB GB”
return $totalSizeGB
}

$csvPath = “$env:TEMP\\MailboxSizes.csv”
“DisplayName,AG,TotalMailboxSizeGB” | Out-File -FilePath $csvPath -Encoding UTF8

# Get all users whose display name starts with “Crelan”
[array]$Users = Get-MgUser -Filter "startsWith(displayName,'Tony') and userType eq 'Member'" -All

foreach ($user in $Users) {
try {
$sizeGB = Get-MailboxSizeForUser -UserId $user.Id
“$($user.GivenName),$($user.SurName),$sizeGB” | Out-File -FilePath $csvPath -Append -Encoding UTF8
} catch {
Write-Warning “Failed to get mailbox size for $($user.DisplayName): $_”
}
}

# Define the CSV file path
$csvName = [System.IO.Path]::GetFileName($csvPath)
$csvBytes = [System.IO.File]::ReadAllBytes($csvPath)
$csvBase64 = [System.Convert]::ToBase64String($csvBytes)

# Create the email message with CSV attachment
$emailMessage = @{
Message = @{
Subject = “Mailbox Size Report”
Body = @{
ContentType = “Text”
Content = “Attached is the mailbox size report in CSV format.”
}
ToRecipients = @(
@{
EmailAddress = @{
Address = “”
}
}
)
Attachments = @(
@{
“@odata.type” = “#microsoft.graph.fileAttachment”
Name = $csvName
ContentBytes = $csvBase64
ContentType = “text/csv”
}
)
}
}
Send-MgUserMail -UserId “” -BodyParameter $emailMessage