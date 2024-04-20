# Report-SPOFilesDocumentLibrary.PS1
# A version of the script to report files from a SharePoint Online document library that uses the Microsoft Graph SDK
# instead of Graph API requests (see ReportFilesSPOSite.PS1 for that version). 

# V1.0 20-Apr-2024

# Set up the script parameters
Param (
    [Parameter(Mandatory = $true)]
    [string]$SiteName  # Name of the SharePoint Online site to process
)

function Get-DriveItem {
    # function to get properties of a drive item
    [CmdletBinding()]
    param (
        [Parameter()]
        $Drive,
        [Parameter()]
        $DriveItem
    )

    # Ignore any folder that has no child items
    If ($DriveItem.folder.childcount -eq 0) {
        Continue 
    }
    If ($DriveItem.folder.childcount -gt 0) {
        $FolderSize = FormatFileSize $DriveItem.Size
        Write-Output ("Processing folder {0} with {1} items ({2})" -f $DriveItem.Name, $DriveItem.folder.childcount, $FolderSize)
        # Recurse in the folder to get the childitems of the folder
        [array]$FolderItems = $null
        $CurrentFolderName = $DriveItem.Name
        $FolderItems = (Get-MgDriveItem -DriveId $Drive.id -DriveItemId $DriveItem.Id -ExpandProperty Children).Children
        If ($FolderItems) {
            $FolderItems | ForEach-Object { 
                Get-DriveItem -Drive $Drive -DriveItem (Get-MgDriveItem -DriveId $Drive.Id -DriveItemId $($_.Id) ) 
            }
        } 
    } Else {
        # Write-Output ("Processing file {0}" -f $DriveItem.Name) 
        If ($null -eq $CurrentFolderName) {
            $CurrentFolderName = "Root"
        }
        $ReportLine = [PSCustomObject]@{
            FileName        = $DriveItem.Name
            Folder          = $CurrentFolderName
            WebURL          = $DriveItem.WebUrl
            Size            = (FormatFileSize $DriveItem.Size)
            Created         = $DriveItem.CreatedDateTime
            LastModified    = $DriveItem.LastModifiedDateTime
        }
        $ReportData.Add($ReportLine)
    }
}

function FormatFileSize {
    # Format File Size nicely
    param (
            [parameter(Mandatory = $true)]
            $InFileSize
        ) 
    
    If ($InFileSize -lt 1KB) { # Format the size of a document
        $FileSize = $InFileSize.ToString() + " B" 
    } 
    ElseIf ($InFileSize -lt 1MB) {
        $FileSize = $InFileSize / 1KB
        $FileSize = ("{0:n2}" -f $FileSize) + " KB"
    } 
    Elseif ($InFileSize -lt 1GB) {
        $FileSize = $InFileSize / 1MB
        $FileSize = ("{0:n2}" -f $FileSize) + " MB" 
    }
    Elseif ($InFileSize -ge 1GB) {
        $FileSize = $InFileSize / 1GB
        $FileSize = ("{0:n2}" -f $FileSize) + " GB" 
    }
    Return $FileSize
    } 

# Disconnect from any previous Graph SDK session
Disconnect-MgGraph
# Connect to the Microsoft Graph
Connect-MgGraph -Scopes "Sites.Read.All" -NoWelcome

# Find the site
Write-Host "Looking for matching sites..."
[array]$Sites = Get-MgSite -Search ($SiteName)

If (!($Sites)) { # Nothing found
    Write-Host "No matching sites found - exiting"
    break 
}
If ($Sites.Count -eq 1) { # Only one site found - go ahead
    $Site = $Sites
    $SiteName = $Site.DisplayName
    Write-Host "Found site to process:" $SiteName 
} ElseIf ($Sites.Count -gt 1) { 
    # More than one site found. Ask which to use
    Clear-Host
    [int]$i = 1
    Write-Host "More than one matching site was found. We need you to select a site to report."
    Write-Host " "
    ForEach ($SiteOption in $Sites) {
        Write-Host ("{0}: {1} ({2})" -f $i, $SiteOption.DisplayName, $SiteOption.Name); $i++
    }
    Write-Host ""
    [Int]$Answer = Read-Host "Enter the number of the site to use"
    If (($Answer -gt 0) -and ($Answer -le $i)) {
        [int]$Si = ($Answer-1)
        $SiteName = $Sites[$Si].DisplayName 
        Write-Host ("OK. Selected site is {0}" -f $Sites[$Si].DisplayName) 
        $Site = $Sites[$Si] 
    }
}

If (!($Site)) { 
    Write-Host ("Can't find the {0} site - script exiting" -f $Uri) 
    break 
}

# Find the document libraries in the site
[array]$Drives = Get-MgSiteDrive -SiteId $Site.Id
If ($Drives.Count -eq 0) {
    Write-Host "No document libraries found in the site" -ForegroundColor Red
    Break
}

If (!($Drives)) { # Nothing found
    Write-Host "No matching drives found - exiting"; break 
}
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
       $Drive = $Drives[$Si] 
    }
}

If (!($Drive)) { 
   Write-Host ("Can't find the {0} drive - script exiting" -f $Uri) ; break 
}
# Create output list and CSV file
$Global:ReportData = [System.Collections.Generic.List[Object]]::new()
$CSVOutputFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + ("\Files {0}-{1} library.csv" -f $Site.displayName, $DriveName)
# Get the items in the root, including child folders
Write-Host "Fetching file information..."
[array]$DriveItems = Get-MgDriveItemChild -DriveId $Drive.id -DriveItemId "root"
ForEach ($DriveItem in $DriveItems) {
     # Retrieve items from all folders beneath the root
    Get-DriveItem -Drive $Drive -DriveItem $DriveItem
}

$ReportData | Out-GridView -Title ("Files in {0} document library for the {1} site" -f $DriveName, $SiteName)
$ReportData | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding UTF8
Write-Host ("Report data saved to {0}" -f $CSVOutputFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.