# DecryptProtectedSPODocuments.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/DecryptProtectedSPODocuments.PS1
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

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
