# Apply-SensitivityLabels.PS1
# A version of the script to apply sensitivity labels to files from a SharePoint Online document library 
# Based on https://github.com/12Knocksinna/Office365itpros/blob/master/Report-SPOFilesDocumentLibrary.PS1
# This script is https://github.com/12Knocksinna/Office365itpros/blob/master/Apply-SensitivityLabels.PS1
# See https://practical365.com/sharepoint-online-files-report/ for more information.
# V1.0 21-May-2024

# Run the script with the name of the site to process.
# .\Apply-SensitivityLabels.PS1 -SiteName "Contoso"

# Set up the script parameter - the name of the site to find
Param (
    [Parameter(Mandatory = $true)]
    [string]$SiteName  # Name of the SharePoint Online site to process
)
function Get-DriveItems {
    [CmdletBinding()]
    param (
        [Parameter()]
        $DriveId,
        [Parameter()]
        $FolderId
    )
    # Get data for a folder and its children
    [array]$Data = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $FolderId -All
    # Split the data into files and folders
    [array]$Folders = $Data | Where-Object {$_.folder.childcount -gt 0} | Sort-Object Name
    $Global:TotalFolders = $TotalFolders + $Folders.Count
    [array]$Files = $Data | Where-Object {$null -ne $_.file.mimetype} 

    # Process the files to find each has a sensitivity label. If not, label it
    ForEach ($File in $Files) {  
        $SensitivityLabelName = $null; $SensitivityLabelInfo = $Null; $Status = $false
        $FileType = $File.Name.Split(".")[1]
        If ($FileType -in $ValidFileTypes) { 
            $Uri = ("https://graph.microsoft.com/V1.0/sites/{0}/drive/items/{1}/extractSensitivityLabels" -f $Site.Id, $File.id)
            Try {
                [array]$SensitivityLabelInfo = Invoke-MgGraphRequest -Uri $Uri -Method POST 
                } Catch {
                    Write-Host ("Error reading sensitivity label data from file {0}" -f $File.Name) -ForegroundColor Yellow
                    # PDFs often return errors when reading sensitivity label info, so I ignore them and only accept the error for Office documents
                    If ($FileType.ToUpper() -ne "PDF") {
                        $SensitivityLabelName = "Error"
                    }
                }
        }         
        If (($null -eq $SensitivityLabelInfo.labels.sensitivityLabelId) -and ($FileType -in $ValidFileTypes) -and $SensitivityLabelName -ne "Error") {
            Write-Host ("Assigning sensitivity label to {0}" -f $File.Name)
            $Uri = ("https://graph.microsoft.com/v1.0/drives/{0}/items/{1}/assignSensitivityLabel" -f $DriveId, $File.Id)
            Try {
                Invoke-MgGraphRequest -Method "POST" -Uri $Uri -Body $AssignmentPayload -OutputType PSObject
                $Status = $true # set so that we report the update
            } Catch {
                $Error[0].Exception.Message
                Write-Host ("Failed to assign sensitivity label to file {0}" -f $File.Name ) -ForegroundColor Red
            }
        }

        If ($File.LastModifiedDateTime) {
            $LastModifiedDateTime = Get-Date $File.LastModifiedDateTime -format 'dd-MMM-yyyy HH:mm'
        } Else {
            $LastModifiedDateTime = $null
        }
        If ($File.CreatedDateTime) {
            $FileCreatedDateTime = Get-Date $File.CreatedDateTime -format 'dd-MMM-yyyy HH:mm'
        }
        If ($Status -eq $true) {
            $ReportLine = [PSCustomObject]@{
                FileName                = $File.Name
                Folder                  = $File.parentreference.name
                Size                    = (FormatFileSize $File.Size)
                Created                 = $FileCreatedDateTime
                Author                  = $File.CreatedBy.User.DisplayName
                LastModified            = $LastModifiedDateTime
                'Last modified by'      = $File.LastModifiedBy.User.DisplayName
                'Sensitivity label'     = $AssignedSensitivityLabelName
            }
            $ReportData.Add($ReportLine)
        }
    }

    # Process the folders
    ForEach ($Folder in $Folders) {
        Write-Host ("Processing folder {0} ({1} files/size {2})" -f $Folder.Name, $Folder.folder.childcount, (FormatFileSize $Folder.Size))
        Get-DriveItems -Drive $DriveId -FolderId $Folder.Id
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

# Insert values for your tenant and app here
$AppId = "43685887-266e-46cf-915b"
$TenantId = "-14fc-43a2-9a7a-d2e27f4f3478"
$CertThumbprint = "F79286DB88C21491110109A0222348FACF694CBD"
# Guid for sensitivity label to apply - againm from your tenant
$Global:SensitivityLabelId = "27451a5b-5823-4853-bcd4-2204d03ab477"
$Global:AssignedSensitivityLabelName = "Internal"

# Parameters to assign a sensitivity label
$Global:AssignmentPayload = @{}
$AssignmentPayload.Add("AssignmentMethod", "Auto")
$AssignmentPayload.Add("Justification", "API application of label")
$AssignmentPayload.Add("SensitivityLabelId", $SensitivityLabelId)

# Connect to the Microsoft Graph
Connect-MgGraph -NoWelcome -AppId $AppId -TenantId $TenantId -CertificateThumbprint $CertThumbprint

Write-Host "Setting up to assign Sensitivity labels to unlabelled files..."
$Global:SensitivityLabelsAvailable = $true
[array]$Global:ValidfileTypes = "docx", "pptx", "xlsx", "pdf"

# Find the site
Write-Host "Looking for matching sites..."
[array]$Sites = Get-MgSite -Search ($SiteName)

If (!($Sites)) { # Nothing found
    Write-Host "No matching sites found - exiting"
    break 
}
If ($Sites.Count -eq 1) { # Only one site found - go ahead
    $Global:Site = $Sites[0]
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
        $Global:Site = $Sites[$Si] 
    }
}

If (!($Site)) { 
    Write-Host ("Can't find the {0} site - script exiting" -f $Uri) 
    break 
}

# Find the document libraries in the site
[array]$Drives = Get-MgSiteDrive -SiteId $Site.Id
If (!($Drives)) {
    Write-Host "No document libraries found in the site" -ForegroundColor Red
    Break
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

[datetime]$StartProcessing = Get-Date
$Global:TotalFolders = 1

# Create output list and CSV file
$Global:ReportData = [System.Collections.Generic.List[Object]]::new()
$CSVOutputFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + ("\Files {0}-{1} library.csv" -f $Site.displayName, $DriveName)
# Get the items in the root, including child folders
Write-Host "Fetching file information..."
Get-DriveItems -Drive $Drive.Id -FolderId "root"
[datetime]$EndProcessing = Get-Date
$ElapsedTime = ($EndProcessing - $StartProcessing)
$FilesPerMinute = [math]::Round(($ReportData.Count / ($ElapsedTime.TotalSeconds / 60)), 2)
Write-Host ""
Write-Host ("Processed {0} files in {1} folders in {2}:{3} minutes ({4} files/minute)" -f `
   $ReportData.Count, $TotalFolders, $ElapsedTime.Minutes, $ElapsedTime.Seconds, $FilesPerMinute)


$ReportData | Out-GridView -Title ("Files in {0} document library for the {1} site" -f $DriveName, $SiteName)
$ReportData | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding UTF8
Write-Host ("Report data saved to {0}" -f $CSVOutputFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
