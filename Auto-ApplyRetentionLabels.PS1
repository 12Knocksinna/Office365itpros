# Auto-ApplyRetentionLabels.PS1
# An example of how to apply retention labels to files in SharePoint Online and OneDrive for Account sites
# V1.0 December 2024
# GitHub Link:

# Permissions RecordsManagement.ReadWrite.All (retention labels)


function Get-DriveItems {
    [CmdletBinding()]
    param (
        [Parameter()]
        $Drive,
        [Parameter()]
        $FolderId
    )
    # Get data for a folder and its children
    [array]$Data = Get-MgDriveItemChild -DriveId $Drive -DriveItemId $FolderId -All
    # Split the data into files and folders
    [array]$Folders = $Data | Where-Object {$_.folder.childcount -gt 0} | Sort-Object Name
    [array]$Files = $Data | Where-Object {$null -ne $_.file.mimetype} 

    # Process the files
    ForEach ($File in $Files) {   
        # Get retention label information
        Try {
            $FileInfo =  Get-MgDriveItemRetentionLabel -DriveId $Drive -DriveItemId $File.Id -ErrorAction Stop
        } Catch {
            Write-Host ("Error reading retention label data from file {0}" -f $File.Name) 
            Continue
        }
        If ([string]::IsNullOrEmpty($FileInfo.Name)) {
            # Attempt to apply the defined retention label
            #$Status = Update-MgDriveItemRetentionLabel -DriveId $Drive -DriveItemId $File.Id -BodyParameter $RetentionLabel
            #If ($Status.Name) { 
            #    Write-Host ("Retention label assigned to {0}" -f $File.Name)
            #}
        }    
        If ($File.LastModifiedDateTime) {
            $LastModifiedDateTime = Get-Date $File.LastModifiedDateTime -format 'dd-MMM-yyyy HH:mm'
        } Else {
            $LastModifiedDateTime = $null
        }
        If ($File.CreatedDateTime) {
            $FileCreatedDateTime = Get-Date $File.CreatedDateTime -format 'dd-MMM-yyyy HH:mm'
        }
    
        $ReportLine = [PSCustomObject]@{
            TimeStamp               = (Get-Date -format s)
            FileName                = $File.Name
            Folder                  = $File.parentreference.name
            Created                 = $FileCreatedDateTime
            Author                  = $File.CreatedBy.User.DisplayName
            LastModified            = $LastModifiedDateTime
            'Last modified by'      = $File.LastModifiedBy.User.DisplayName
            'Retention label'       = $RetentionLabel['Name']
        }
        $ReportData.Add($ReportLine)
    }

    # Process the folders
    ForEach ($Folder in $Folders) {
        Write-Host ("Processing folder {0}" -f $Folder.Name)
        Get-DriveItems -Drive $Drive -FolderId $Folder.Id
    }
}

# Need to run in app-only mode
Connect-MgGraph -Scopes RecordsManagement.ReadWrite.All

$User = Get-MgUser -UserId (Get-MgContext).Account


$LocationsFile = 'c:\temp\FilesLocations.csv'
If (!(Test-Path $LocationsFile)) {
    Write-Host "Locations file for unlabeled files not found - exiting"
    Break
}
[array]$AllLocations = Import-Csv -Path $LocationsFile
# Exclude all OneDrive for Business sites
[array]$Locations = $AllLocations | Where-Object {$_.Location -notlike "*my.sharepoint.com/personal*"}

# Define default retention label to apply
$Global:RetentionLabel = @{}
$RetentionLabel.Add("Name","Approved")
# Output PowerShell list for report
$Global:ReportData = [System.Collections.Generic.List[Object]]::new()

# Try to access each site. If we can gain access, look for unlabeled files
ForEach ($Location in $Locations) {
    $URI = $Location.Location
    Try {
        $Site = Get-MgSite -Search $URI -ErrorAction Stop
    } Catch {
        Write-Host ("Unable to access site {0} {1}" -f $URI, $_.Exception.Message) -ForegroundColor Red
        Continue
    }
    If (!$Site) {
        Write-Host "Site $URI not found"
        Continue
    } Else {
        Write-Host ("Processing site {0} to look for unlabeled files" -f $Site.DisplayName) -ForegroundColor Yellow
    }
    Try {
        [array]$AllDrives = Get-MgSiteDrive -SiteId $Site.Id -ErrorAction Stop
    } Catch {
        Write-Host ("Unable to access site {0} ({1}) {2}" -f $Site.DisplayName, $URI, $_.Exception.Message) -ForegroundColor Red
        Continue
    }

    [array]$Drives = $AllDrives | Where-Object {$_.Name -notlike "*Preservation Hold Library*" -and $_.Name -notlike "*Teams Wiki Data*"}
    If (!$Drives) {
        Write-Host "No drives found in site $URI"
        Continue
    }
    # Process each drive
    ForEach ($Drive in $Drives) {
        Write-Host ("Processing drive {0} in site {1}" -f $Drive.Name, $Site.DisplayName)
        Get-DriveItems -Drive $Drive.Id -FolderId "root"
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.