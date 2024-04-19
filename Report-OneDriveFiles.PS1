# Report-OneDriveFiles.PS1
# A demo script to show how to generate a report of the files in a user's OneDrive account using 
# Microsoft Graph APIs. Based on the script to report files from a SharePoint Online document library 
# described in the article https://practical365.com/sharepoint-site-files-report/
# This script is https://github.com/12Knocksinna/Office365itpros/blob/master/Report-OneDriveFiles.PS1
# V1.0 19-Apr-2024

function UnpackFilesRecursively {
# Unpack set of items (files and folders)
param (
        [parameter(Mandatory = $true)]
        $Items, # Items to unpack
	
		[parameter(Mandatory = $true)]
        $SiteUri, # Base site URI
		
		[parameter(Mandatory = $true)]
        $FolderPath, # Folder path
		
        [parameter(Mandatory = $true)]
        $SiteFiles,
		
		[parameter(Mandatory = $false)]
		[bool]$IsNextLink
    )

  # Find sub-folders that we need to check for files
  $Folders = $Items.Value | Where-Object {$_.Folder.ChildCount -gt 0 }
  # And any files in the folder
  $Files = $Items.Value | Where-Object {$null -eq $_.Folder.ChildCount}
  
  $before = $SiteFiles.count
  
  # Report the files
  ForEach ($D in $Files) {
    $FileSize = FormatFileSize $D.Size
    # if we're reporting files from the top of the OneDrive account, set the value to the account name
    If ([string]::IsNullOrWhiteSpace($FolderPath)) {
        $FolderPathReport = $OneDriveName
    } Else {
        $FolderPathReport = $FolderPath
    }
    $ReportLine  = [PSCustomObject] @{   
        FileName = $D.Name
        Folder   = $FolderPathReport
        Author   = $D.createdby.user.displayname
        Created  = $D.createdDateTime
        Modified = $D.lastModifiedDateTime
        Size     = $FileSize
        Uri      = $D.WebUrl 
    }
    $SiteFiles.Add($ReportLine) 
  } # End If

  $NextLink = $Items."@odata.nextLink"
  $Uri = $Items."@odata.nextLink"
  While ($NextLink) { 
    $MoreData = Invoke-MgGraphRequest -Uri $Uri -Method Get
    UnpackFilesRecursively -Items $MoreData -SiteUri $SiteUri -FolderPath $FolderPath -SiteFiles $SiteFiles -IsNextLink $true
  
    $NextLink = $MoreData."@odata.nextLink"
    $Uri = $MoreData."@odata.nextLink" 
  } # End While
  
  $Count = $SiteFiles.count - $before
  If (-Not $IsNextLink) {
    Write-Host ("Folder: {0} count {1}" -f $FolderPath, $Count)
  }
  
  # Report the files in each sub-folder
  ForEach ($Folder in $Folders) {
    Write-Host "Folderpath:" $FolderPath
	$NewFolderPath = ("{0}/{1}" -f $FolderPath, $Folder.Name).TrimStart()
    Write-Host "New folder path" $NewFolderPath
    $Uri =('https://graph.microsoft.com/V1.0/users/{0}/drive/root:{1}:/children' -f $Account, $NewFolderPath)    
	$SubFolderData = Invoke-MgGraphRequest -Uri $Uri -Method Get
    UnpackFilesRecursively -Items $SubFolderData -SiteUri $SiteUri -FolderPath $NewFolderPath -SiteFiles $SiteFiles -IsNextLink $false
  } # End Foreach Folders
}

function FormatFileSize {
# Format File Size nicely
param (
        [parameter(Mandatory = $true)]
        $InFileSize
    ) 

 If ($InFileSize -lt 1KB) { # Format the size of a document
        $FileSize = $InFileSize.ToString() + " B" 
    } ElseIf ($InFileSize -lt 1MB) {
        $FileSize = $InFileSize / 1KB
        $FileSize = ("{0:n2}" -f $FileSize) + " KB"
    } Elseif ($InFileSize -lt 1GB) {
        $FileSize = $InFileSize / 1MB
        $FileSize = ("{0:n2}" -f $FileSize) + " MB" 
    } Elseif ($InFileSize -ge 1GB) {
        $FileSize = $InFileSize / 1GB
        $FileSize = ("{0:n2}" -f $FileSize) + " GB" 
    }
  Return $FileSize
} 

# Connect to the Microsoft Graph with the permission to read sites
Disconnect-MgGraph | Out-Null # Make sure that we sign out of existing sessions
Connect-MgGraph -Scopes Sites.Read.All -NoWelcome

$Global:Account = (Get-MgContext).Account
$Site = ("https://graph.microsoft.com/v1.0/users/{0}/drive/" -f $Account)
[array]$OneDriveInfo = Invoke-MgGraphRequest -Uri $Site -Method Get
If (!($OneDriveInfo)) { # Nothing found
    Write-Host "No matching OneDrive found - exiting"; break 
} Else {
    Write-Host ("Found OneDrive account owned by {0} to process. URL: {1}" -f $OneDriveInfo.owner.user.displayName, $OneDriveInfo.WebUrl)
    $DriveId = $OneDriveInfo.Id
    $SiteUri = $OneDriveInfo.WebUrl
    $Global:OneDriveName = $OneDriveInfo.name
}

# Create output list
$SiteFiles = [System.Collections.Generic.List[Object]]::new()

Write-Host "Reading from OneDrive account..."
$Uri = ("https://graph.microsoft.com/v1.0/drives/{0}/items/root/children" -f $DriveId)
[array]$Items = Invoke-MgGraphRequest -Uri $Uri -Method Get
If ($Items.Value.Count -eq 0) {
    Write-Host "No files found in the OneDrive for Business account"
    Break
}   
$StartPath = " "
UnpackFilesRecursively -Items $Items -SiteUri $SiteUri -FolderPath $StartPath -SiteFiles $SiteFiles

Write-Host ("Total files found in the OneDrive for Business account of {0}: {1}" -f $OneDriveInfo.owner.user.displayName, $SiteFiles.Count) -ForegroundColor Red
$SiteFiles | Out-GridView -Title ("OneDrive for Business Files for {0}" -f $OneDriveInfo.owner.user.displayName) 


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
