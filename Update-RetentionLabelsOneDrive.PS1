# Based on https://github.com/12Knocksinna/Office365itpros/blob/master/Report-OneDriveFiles.PS1, which creates a report of files in a user's OneDrive
# for Business account. This adaptation updates the assigned retention label for selected files.
# V1.0 18-Jul-2025

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Update-RetentionLabelsOneDrive.PS1

function Get-DriveItems {
    [CmdletBinding()]
    param (
        [Parameter()]
        $Drive,
        [Parameter()]
        $FolderId
    )
    # Get OneDrive items for a folder
    [array]$Data = Get-MgDriveItemChild -DriveId $Drive -DriveItemId $FolderId -All
    # Split the data into files and folders
    [array]$Folders = $Data | Where-Object {$_.folder.childcount -gt 0}
    $Global:TotalFolders = $TotalFolders + $Folders.Count
    [array]$Files = $Data | Where-Object {$null -ne $_.file.mimetype}

    ForEach ($File in $Files) {
    # Write-Output ("Processing file {0}" -f $File.Name) 
        $ReplacedLabel = $false

        # Get retention label information
        Try {
            $RetentionLabelInfo = $null; $RetentionLabelName = $null
            $RetentionlabelInfo = Get-MgDriveItemRetentionLabel -DriveId $OneDriveInfo.Id -DriveItemId $File.Id -ErrorAction Stop

            If ($RetentionLabelInfo.name -eq $OldRetentionLabelName) {
     
                $Status = Update-MgDriveItemRetentionLabel -DriveId $OneDriveInfo.Id -DriveItemId $File.Id -BodyParameter $NewRetentionLabelParameters -ErrorAction Stop
                If ($Status.Name) {
                    Write-Host ("File {0} had the {1} label and has been assigned the {2} retention label" -f $File.Name, $OldRetentionLabelName, $NewRetentionLabelName) -ForegroundColor Yellow
                    $ReplacedLabel = $true
                } Else {
                    Write-Host ("File {0} had the {1} label, but the update to the {2} label failed" -f $File.Name, $OldRetentionLabelName, $NewRetentionLabelName) -ForegroundColor Red
                }
            }
        } Catch {
            Write-Host ("Error reading retention label data from file {0}" -f $File.Name) 
        }
        
        If ($File.LastModifiedDateTime) {
            $FileLastModifiedDateTime = Get-Date $File.LastModifiedDateTime -format 'dd-MMM-yyyy HH:mm'
        } Else {
            $FileLastModifiedDateTime = $null
        }
        If ($File.CreatedDateTime) {
            [datetime]$FileCreated = $File.createdDateTime
            $AgeInDays = (New-TimeSpan $FileCreated).Days
            $FileCreatedDateTime = Get-Date $File.CreatedDateTime -format 'dd-MMM-yyyy HH:mm'
        }
        If ([string]::IsNullOrEmpty($RetentionLabelName)) {
            $RetentionLabelName = "No label"
        } Else {
            [string]$RetentionLabelName = $RetentionLabelName.Trim()
        }

        If ($ReplacedLabel -eq $true) {
            
            $ReportLine = [PSCustomObject]@{
                FileName            = $File.Name
                Folder              = $File.parentreference.name
                Author              = $File.createdby.user.displayname
                Created             = $FileCreatedDateTime
                Modified            = $FileLastModifiedDateTime
                Size                = (FormatFileSize $File.Size)
                DaysOld             = $AgeInDays
                'Retention label'   = $RetentionLabelInfo.name
                'New Retention label' = $NewRetentionLabelName
                FileType            = $FileType
                Bytes               = $File.Size
                WebURL              = $File.WebUrl
            }
            $ODBFiles.Add($ReportLine)
        }
    }
    ForEach ($Folder in $Folders) {
        Write-Host ("Processing folder {0} ({1} files/size {2})" -f $Folder.Name, $Folder.folder.childcount, (FormatFileSize $Folder.Size))
        Get-DriveItems -Drive $Drive -FolderId $Folder.Id
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
# User must have a valid license for OneDrive for Business...
Connect-MgGraph -Scopes Files.Read.All, RecordsManagement.Read.All, InformationProtectionPolicy.Read -NoWelcome

Write-Host "Getting ready to apply replacement retention labels in a OneDrive for Business account..."
# Discover if the tenant uses sensitivity labels
$Account = (Get-MgContext).Account

# Discover if the tenant uses retention labels
[array]$RetentionLabels = Get-MgSecurityLabelRetentionLabel
If ($RetentionLabels) {
    $Global:RetentionLabelsAvailable = $true
    $RetentionLabels = $RetentionLabels | Sort-Object DisplayName
    Write-Host ("Tenant has {0} retention labels available" -f $RetentionLabels.Count) -ForegroundColor Green
} Else {
    Write-Host "No retention labels available in the tenant. Exiting because without labels we can't switch anything" -ForegroundColor Red
    Break
}

# List retention label display names and prompt user to select the label to find and its replacement
Write-Host "Available Retention Labels:"
for ($i = 0; $i -lt $RetentionLabels.Count; $i++) {
    Write-Host ("[{0}] {1}" -f ($i + 1), $RetentionLabels[$i].DisplayName)
}
do {
    [int]$Selection = Read-Host "Enter the number of the retention label to replace"
    $IsValid = (($Selection -ge 1) -and ($Selection -le $RetentionLabels.Count))
    if (-not $IsValid) {
        Write-Host "Invalid selection. Please enter a number between 1 and $($RetentionLabels.Count)." -ForegroundColor Yellow
    }
} until ($IsValid)
[int]$OriginalSelection = $Selection
$ReplaceLabel = $RetentionLabels[$Selection - 1]

Write-Host ("You selected the {0} retention label" -f $ReplaceLabel.DisplayName) -ForegroundColor Green

do {
    [int]$Selection = Read-Host ("Now select the number of the retention label to apply to files that have the {0} label" -f $ReplaceLabel.DisplayName) 
    $IsValid = (($Selection -ge 1) -and ($Selection -le $RetentionLabels.Count) -and ($Selection -ne $OriginalSelection))
    if (-not $IsValid) {
        Write-Host "Invalid selection. Please enter a number between 1 and $($RetentionLabels.Count)." -ForegroundColor Yellow
    }
} until ($IsValid)
$NewLabel = $RetentionLabels[$Selection - 1]

Write-Host ("The {0} retention label will replace the {1} label" -f $NewLabel.DisplayName, $ReplaceLabel.DisplayName) -ForegroundColor Green

# Define the hash table for the new retention label parameters
$Global:NewRetentionLabelParameters = @{}
$NewRetentionLabelParameters.Add("Name",$NewLabel.DisplayName)

$Global:OldRetentionLabelName = $ReplaceLabel.DisplayName
$Global:NewRetentionLabelName = $NewLabel.DisplayName

# Find user's OneDrive for Business account
[array]$OneDriveInfo =  Get-MgUserDefaultDrive -UserId $Account

If (!($OneDriveInfo)) { # Nothing found
    Write-Host ("No matching OneDrive for Business account found for {0} - exiting" -f $Account)
    break 
} Else {
    Write-Host ("Found OneDrive account owned by {0} to process. URL: {1}" -f $OneDriveInfo.owner.user.displayName, $OneDriveInfo.WebUrl)
    $Global:OneDriveName = $OneDriveInfo.name
}

# Create output list and CSV file
$Global:ODBFiles = [System.Collections.Generic.List[Object]]::new()
$CSVOutputFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + ("\Relabeled OneDrive files for {0}.csv" -f $OneDriveInfo.owner.user.displayName)

[datetime]$StartProcessing = Get-Date
$Global:TotalFolders = 1

# Get the items in the root, including child folders
Write-Host "Fetching file information from OneDrive for Business..." -ForegroundColor Yellow
Get-DriveItems -Drive $OneDriveInfo.id -FolderId "root"

[datetime]$EndProcessing = Get-Date
$ElapsedTime = ($EndProcessing - $StartProcessing)
$FilesPerMinute = [math]::Round(($ODBFiles.Count / ($ElapsedTime.TotalSeconds / 60)), 2)

# Show what we've found with Out-GridView
$ODBFiles | Select-Object FileName, Folder, Author, Created, Modified, Size, DaysOld, 'Retention Label', 'Sensitivity Label' | `
    Out-GridView -Title ("OneDrive for Business Files for {0}" -f $OneDriveInfo.owner.user.displayName) 

Write-Host ("A total of {0} files were reassigned with the {1} retention label" -f $ODBFiles.Count, $NewLabel.DisplayName)
$ODBFiles | Select-Object FileName, Folder, Author, Created, Modified, Size, DaysOld, 'Retention Label', 'Sensitivity Label' | `
    Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding UTF8

Write-Host ("CSV file created at {0}" -f $CSVOutputFile) -ForegroundColor Green

Write-Host ("Processed {0} files in {1} folders in {2}:{3} minutes ({4} files/minute)" -f `
   $ODBFiles.Count, $TotalFolders, $ElapsedTime.Minutes, $ElapsedTime.Seconds, $FilesPerMinute)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.