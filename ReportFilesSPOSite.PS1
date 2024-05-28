# ReportFilesSPOSite.PS1
# A demo script to show how to generate a report of the files in a SharePoint Online site using the 
# Microsoft Graph
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportFilesSPOSite.PS1
# V1.0 04-July-2022
# V1.1 27-July-2022 (Fix nextlink processing and allow user to pick document library to process)
# V1.2 2-Jan-2024 (Check against SDK V2.11)
# .\ReportFilesSPOSite.PS1 -SearchSite "Billing"

# A better version is available in https://github.com/12Knocksinna/Office365itpros/blob/master/Report-SPOFilesDocumentLibrary.PS1 and
# described in the article https://practical365.com/sharepoint-online-files-report/

Param ([Parameter(Mandatory)]$SearchSite)

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
    $ReportLine  = [PSCustomObject] @{   
        FileName = $D.Name
        Folder   = $FolderPath
        Author   = $D.createdby.user.displayname
        Created  = $D.createdDateTime
        Modified = $D.lastModifiedDateTime
        Size     = $FileSize
        Uri      = $D.WebUrl }
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
  if (-Not $IsNextLink) {
    Write-Host "  $FolderPath ($count)"
  }
  
  # Report the files in each sub-folder
  ForEach ($Folder in $Folders) {
	$NewFolderPath = $FolderPath + "/" + $Folder.Name
	$Uri = $SiteUri + "/" + $Folder.parentReference.path + "/" + $Folder.Name + ":/children"
	$SubFolderData = Invoke-MgGraphRequest -Uri $Uri -Method Get
    UnpackFilesRecursively -Items $SubFolderData -SiteUri $SiteUri -FolderPath $NewFolderPath -SiteFiles $SiteFiles -IsNextLink $IsNextLink
  } # End Foreach Folders
}

function FormatFileSize {
# Format File Size nicely
param (
        [parameter(Mandatory = $true)]
        $InFileSize
    ) 

 If ($InFileSize -lt 1KB) { # Format the size of a document
        $FileSize = $InFileSize.ToString() + " B" } 
      ElseIf ($InFileSize -lt 1MB) {
        $FileSize = $InFileSize / 1KB
        $FileSize = ("{0:n2}" -f $FileSize) + " KB"} 
      Elseif ($InFileSize -lt 1GB) {
        $FileSize = $InFileSize / 1MB
        $FileSize = ("{0:n2}" -f $FileSize) + " MB" }
      Elseif ($InFileSize -ge 1GB) {
        $FileSize = $InFileSize / 1GB
        $FileSize = ("{0:n2}" -f $FileSize) + " GB" }
  Return $FileSize
} 

# Connect to the Microsoft Graph with the permission to read sites
Disconnect-MgGraph | Out-Null # Make sure that we sign out of existing sessions
Connect-MgGraph -Scopes Sites.Read.All -NoWelcome

Write-Host "Looking for matching sites..."
$Uri = 'https://graph.microsoft.com/v1.0/sites?search="' + $SearchSite + '"'
[array]$Sites = Invoke-MgGraphRequest -Uri $uri -Method Get
$Sites = $Sites.Value
$SiteCount = $Sites.Count

If (!($Sites)) { # Nothing found
     Write-Host "No matching sites found - exiting"; break }
If ($SiteCount -eq 1) { # Only one site found - go ahead
     $Site = $Sites
     $SiteName = $Site.DisplayName
     Write-Host "Found site to process:" $SiteName 
} Elseif ($SiteCount -gt 1) { # More than one site found. Ask which to use
     Clear-Host; Write-Host "More than one matching site was found. We need you to select a site to report."; [int]$i=1
     Write-Host " "
     ForEach ($SiteOption in $Sites) {
        Write-Host ("{0}: {1} ({2})" -f $i, $SiteOption.DisplayName, $SiteOption.Name); $i++}
        Write-Host ""
     [Int]$Answer = Read-Host "Enter the number of the site to use"
     If (($Answer -gt 0) -and ($Answer -le $i)) {
        [int]$Si = ($Answer-1)
        $SiteName = $Sites[$Si].DisplayName 
        Write-Host "OK. Selected site is" $Sites[$Si].DisplayName 
        $Site = $Sites[$Si] }
}

If (!($Site)) { 
    Write-Host ("Can't find the {0} site - script exiting" -f $Uri) ; break 
}

# Get Drives in the site
Write-Host ("Checking for document libraries in the {0} site" -f $Site.DisplayName)
$Uri = "https://graph.microsoft.com/v1.0/sites/$($Site.Id)/drives"
[array]$Drives = Invoke-MgGraphRequest -Uri $Uri -Method Get
$Drives = $Drives.Value

If (!($Drives)) { # Nothing found
     Write-Host "No matching drives found - exiting"; break }
If ($Drives.Count -eq 1) { # Only one drive found - go ahead
     $Drive = $Drives
     $DriveName = $Drive.Name
     Write-Host "Found drive to process:" $DriveName 
} Elseif ($Drives.Count -gt 1) { # More than one drive found. Ask which to use
     Clear-Host; Write-Host "More than one drive found in site. We need you to select a drive to report."; [int]$i=1
     Write-Host " "
     ForEach ($DriveOption in $Drives) {
        Write-Host ("{0}: {1}" -f $i, $DriveOption.Name); $i++}
        Write-Host ""
     [Int]$Answer = Read-Host "Enter the number of the drive to use"
     If (($Answer -gt 0) -and ($Answer -le $i)) {
        [int]$Si = ($Answer-1)
        $DriveName = $Drives[$Si].Name 
        Write-Host "OK. Selected drive is" $Drives[$Si].Name 
        $Drive = $Drives[$Si] }
}

If (!($Drive)) { 
    Write-Host ("Can't find the {0} drive - script exiting" -f $Uri) ; break 
}


# Use the selected drive
$DocumentLibrary = $Drive

$SiteUri = "https://graph.microsoft.com/v1.0/sites/$($Site.Id)"
$Uri = "$SiteUri/drives/$($DocumentLibrary.Id)/root/children"

# Create output list
$SiteFiles = [System.Collections.Generic.List[Object]]::new()

Write-Host "Reading from document library..."

# Get Items in document library
[array]$Items = Invoke-MgGraphRequest -Uri $Uri -Method Get

UnpackFilesRecursively -Items $Items -SiteUri $SiteUri -FolderPath $DocumentLibrary.Name -SiteFiles $SiteFiles

Write-Host ("Total files found {0}" -f $SiteFiles.Count)
$SiteFiles | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
