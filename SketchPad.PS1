# SketchPad.PS1
# Used for bits of PowerShell I am working on or want to note

# Get list to update metadata for the new item
$ListId = (Get-MgSiteList -SiteId $Site.Id -Filter "DisplayName eq 'Documents'").Id
[array]$ListItems = Get-MgSiteListItem -SiteId $Site.Id -ListId $ListId
$ListItem = $ListItems[-1]
$Body = @{}
$Body.Add("Title", "Hard Deleted Users Report Created by Azure Automation")
$Status = Update-MgSiteListItemField -SiteId $site.Id -ListId $listId -ListItemId $listItem.Id -BodyParameter $Body
If ($Status) {
    Write-Output ("Updated document metadata for item {0} with title {1}" -f $ListItem.Id, $Params.Title)
}




# Report all OneDrive accounts
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
    -ConsistencyLevel eventual -CountVariable UsersFound -All -PageSize 500
If (!$Users) {
    Write-Host "No user accounts found"
    Break
}
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    Try {
        $OneDrive = Get-MgUserDefaultDrive -UserId $User.Id -ErrorAction Stop
    } Catch {
        Write-Host ("Unable to find OneDrive for {0}" -f $User.UserPrincipalName)
        Continue
    }   
    $ReportLine = [PSCustomObject]@{
        UserPrincipalName = $User.UserPrincipalName
        OneDriveUrl       = $OneDrive.WebUrl
        Created           = Get-Date $OneDrive.CreatedDateTime -format 'dd-MMM-yyyy HH:mm'
        Modified          = Get-Date $OneDrive.LastModifiedDateTime -format 'dd-MMM-yyyy HH:mm'
    }
    $Report.Add($ReportLine)
}

