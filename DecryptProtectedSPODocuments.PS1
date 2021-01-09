# DecryptProtectedSPODocuments.PS1
# Example of how to use the Unlock-SPOSensitivityLabelEncryptedFile to decrypt labelled SharePoint Online documents
# Uses SharePnP and SharePoint Online Management module - example used in https://petri.com/decrypt-sharepointonline-documents

$SPOCheck = Get-Module "Microsoft.Online.SharePoint.PowerShell"
If (!$SPOCheck) {
     Write-Host "Your PowerShell session is not connected to SharePoint Online."; break}
$SPOCheck = Get-Module "SharePointPnPPowerShellOnline"
If (!$SPOCheck) {
     Write-Host "Your PowerShell session is not connected to SharePoint Onnline PnP."; break}

$SiteURL = "https://redmondassociates.sharepoint.com/sites/rabilling"
$FolderURL= "/Shared Documents/2020"

# Connect to SharePoint PnP with cached credentials
Connect-PnPOnline -Url $SiteURL -Credentials $O365Cred 

$FolderItems = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderURL -ItemType File
ForEach ($Item in $FolderItems) {
     $ItemPath = $SiteUrl+$FolderUrl+"/"+$Item.Name
     Write-Host "Unlocking" $Item.Name
     Unlock-SPOSensitivityLabelEncryptedFile -FileUrl $ItemPath -JustificationText "Administrator removed label"
}

# $ItemPath Looks like this: https://redmondassociates.sharepoint.com/sites/rabilling/Shared%20Documents/2020/Document.docx
