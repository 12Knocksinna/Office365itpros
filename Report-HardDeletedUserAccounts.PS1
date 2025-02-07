# Report-HardDeletedUserAccounts.PS1
# Example of an Azure Automation runbook to report Entra ID hard-deleted user accounts

# Requires AuditLog.Read.All permission
# Requires Microsoft.Graph.Reports module for access to audit logs
# Requires Microsoft.Graph.Users.Actions module to send email
# Requires Microsoft.Graph.Files to upload a file to a SharePoint document library
# Requires Microsoft.Graph.Sites to access SharePoint Online site
# Requires ImportExcel module to create an Excel worksheet

Connect-MgGraph -Identity -NoWelcome
$ExcelOutputFile = "HardDeletedUsers.xlsx"

# Find hard-deleted user accounts
[array]$DeletedUserRecords = Get-MgAuditLogDirectoryAudit -Filter "ActivityDisplayName eq 'Hard Delete user'" -All -Sort 'ActivityDateTime'
If (!$DeletedUserRecords) {
        Write-Host "No hard deleted user records found"; break
}
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Record in $DeletedUserRecords) {
  $DataLine = [PSCustomObject][Ordered]@{
    TimeStamp           = (Get-Date $Record.ActivityDateTime -format 'dd-MMM-yyyy HH:mm:ss')
    DeletionInitiatedBy = $Record.InitiatedBy.User.UserPrincipalName
    DeletedUser         = $Record.TargetResources.UserPrincipalName.Substring(32,($Record.TargetResources.UserPrincipalName.length-32))
  }
  $Report.Add($DataLine)
}

$Report | Format-Table TimeStamp, DeletionInitiatedBy, DeletedUser -AutoSize

# Generate Excel worksheet
$Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Hard Deleted User Accounts" -Title  "Hard Deleted Accounts" `
    -TitleBold -TableName "Hard Deleted Accounts" 

# Generate CSV file
$Report | Export-Csv -Path "HardDeletedUsers.csv" -NoTypeInformation -Encoding UTF8

$SiteUri = "https://redmondassociates.sharepoint.com/sites/Office365Adoption"
$SiteId = $SiteUri.Split('//')[1].split("/")[0] + ":/sites/" + $SiteUri.Split('//')[1].split("/")[2]
$Site = Get-MgSite -SiteId $SiteId
If (!$Site) {
    Write-Output ("Unable to connect to site {0} with id {1}" -f $Uri, $SiteId) 
    Exit
}
[array]$Drives = Get-MgSiteDrive -SiteId $Site.Id
$DocumentsDrive = $Drives | Where-Object {$_.Name -eq "Documents"}
$TargetFile = "root:/General/" + $ExcelOutputFile + ":"
# This works interactively but not in an Azure Automation runbook
# $NewFile = Set-MgDriveItemContent -DriveId $DocumentsDrive.Id -DriveItemId $TargetFile -InFile $File
# So we use the URI method instead
$Uri = ("https://graph.microsoft.com/V1.0/sites/{0}/drive/items/root:/General/{1}:/content" -f $Site.Id, $ExcelOutputFile)
$NewFile = Invoke-MgGraphRequest -uri $uri -Method PUT -InputFilePath $ExcelOutputFile
If ($NewFile) { 
    Write-Output ("File {0} uploaded to {1} with size {2} MB" -f $NewFile.Name, $DocumentsDrive.Name, ([math]::Round($NewFile.Size/ 1MB, 2)))
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.


$Uri = ("https://graph.microsoft.com/V1.0/sites/{0}/drive/items/root:/General/{1}:/content" -f $Site.Id, $ExcelOutputFile)