# Update-SensitivityLabels.PS1
# A version of the script to apply sensitivity labels to files from a SharePoint Online document library. This version
# overwrites the sensitivity label for files that are already labelled. The script can be used to apply a new label to files
# Based on https://github.com/12Knocksinna/Office365itpros/blob/master/Report-SPOFilesDocumentLibrary.PS1
# This script is https://github.com/12Knocksinna/Office365itpros/blob/master/Update-SensitivityLabels.PS1
# See https://practical365.com/sharepoint-online-files-report/ for original article.
# V1.0 06-May-2025
# V2.0 09-Mar-2026  -   Updated to add parameter to choose input sensitivity label to apply
#                       And to make sure that the sensitivity label is only replaced if the new label has a higher priority than
#                       the existing label, to avoid accidentally applying a lower priority label over a higher priority one.

# +----------------------------------- Important -----------------------------------+
# Because this script uses a metered API to update sensitivity labels. The application used
# must be set up to use the metered APIs. See https://aka.ms/graph-metered-overview for details

# Uses the Sites.ReadWrite.All permission to assign sensitivity labels to files in a SharePoint Online document library.
# Alternatively, use the Sites.Selected permission to grant access to specific sites.
# Requires the SensitivityLabel.Read permission to read the sensitivity label information for the tenant.

# Run the script with the name of the site to process.
# .\Apply-SensitivityLabels.PS1 -SiteName "Contoso" -NewSensitivityLabelName "Confidential"

# If "Remove" is passed as the NewSensitivityLabelName parameter, the script will remove sensitivity labels from files instead of applying a new label. In this case, the script will only remove existing labels if they have a lower priority than the default label defined in the script, to avoid accidentally removing a higher priority label.

# Set up the script parameters - the name of the site to find and the name of the sensitivity label to apply
Param (
    [Parameter(Mandatory = $true)]
    [string]$SiteName,  # Name of the SharePoint Online site to process
    [Parameter(Mandatory = $true)]
    [string]$NewSensitivityLabelName # Name of the sensitivity label to apply to files (mandatory - if not provided, will use the default label defined in the script)
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

    If ($RemoveLabels -eq $True) {
        $RequestBody = @{}
        $RequestBody.Add("sensitivityLabelId", "")
    }

    # Process the files to find each has a sensitivity label. If not, label it
    ForEach ($File in $Files) {  
        $Status = $false
        $FileType = $File.Name.Split(".")[1]
        If ($FileType -in $ValidFileTypes) { 
            $OKToApplyLabel = $false; $OldSensitivityLabelName = $null
            Try {
                # Exract the existing sensitivity label information for the file
                $SensitivityLabelName = $null; $SensitivityLabelInfo = $null
                $Uri = ("https://graph.microsoft.com/V1.0/sites/{0}/drive/items/{1}/extractSensitivityLabels" -f $Site.Id, $File.Id)
                [array]$SensitivityLabelInfo = Invoke-MgGraphRequest -Uri $Uri -Method post -OutputType PsObject -ErrorAction Stop
            } Catch {
                Write-Host ("Error retrieving sensitivity label information for {0} ({1})" -f $File.Name, $_.Exception.Message) -ForegroundColor Red
                $OKToApplyLabel = $true
            }
            # Check the label information for the current file
            [array]$LabelsFound = $SensitivityLabelInfo.Labels
        }
        If ($RemoveLabels -eq $False) {    
            If ($LabelsFound.Count -eq 0) {
                # Go ahead and label if no existing labels are found.
                $OKToApplyLabel = $true
            } Else {
                # Check if the file already has the sensitivity label we want to apply. If it does, skip it.
                If ($SensitivityLabelInfo.Labels[0].sensitivitylabelId -eq $SensitivityLabelId) {
                    Write-Host ("Skipping {0} as it already has the sensitivity label we want to apply" -f $File.Name) -ForegroundColor Yellow
                    Continue
                }
                # Check the priority of the existing label against the priority of the new label we want to apply. We only want to apply the new label if it has a higher priority (lower number) than the existing label, otherwise we could end up applying a lower priority label over a higher priority one, which would not be good!
                [int]$ExistingLabelPriority = $SensitivityLabelsFiles | Where-Object {$_.Id -eq $LabelsFound[0].sensitivityLabelId} | Select-Object -ExpandProperty Priority
                If ($ReplacementLabelPriority -gt $ExistingLabelPriority) {
                    $OKToApplyLabel = $true
                } Else {
                    Write-Host ("Skipping {0} as it has an existing sensitivity label {1} with higher or equal priority than the new label applied by this process" -f $File.Name, $LabelsFound[0].Name) -ForegroundColor Yellow
                    Continue
                }
            }
        }
        # After all the checks, if we have determined that it's OK to apply the new label, go ahead and run the cmdlet    
        If ($OKToApplyLabel -eq $true -and $RemoveLabels -eq $false) {
            Write-Host ("Applying sensitivity label {0} to {1}" -f $NewSensitivityLabelName, $File.Name) -ForegroundColor Green
            Try {
                Set-MgDriveItemSensitivityLabel -DriveId $DriveId -DriveItemId $File.Id -SensitivityLabelId $SensitivityLabelId -ErrorAction Stop
                $Status = $true
            } Catch {
                Write-Host ("Error applying label {0} to {1} ({2})" -f $SensitivityLabelName, $File.Name, $_.Exception.Message) -ForegroundColor Red
                Continue
            }
        }
        # If we're removing labels, do it here
        If ($RemoveLabels -eq $true) {
            Write-Host ("Removing sensitivity label from {0}" -f $File.Name) -ForegroundColor Green
            Try {
                Set-MgDriveItemSensitivityLabel -DriveId $DriveId -DriveItemId $File.Id -SensitivityLabelId "" -ErrorAction Stop
                # Could also have done this with
                # Invoke-MgGraphRequest -Uri $Uri -Method Post -Body $RequestBody -ErrorAction Stop
                $Status = $true
            } Catch {
                Write-Host ("Error removing sensitivity labels from {0} ({1})" -f $File.Name, $_.Exception.Message) -ForegroundColor Red
                Continue
            }
        }
        
        # Collect information for the report
        If ($File.LastModifiedDateTime) {
            $LastModifiedDateTime = Get-Date $File.LastModifiedDateTime -format 'dd-MMM-yyyy HH:mm'
        } Else {
            $LastModifiedDateTime = $null
        }
        If ($File.CreatedDateTime) {
            $FileCreatedDateTime = Get-Date $File.CreatedDateTime -format 'dd-MMM-yyyy HH:mm'
        } Else {
            $FileCreatedDateTime = $null
        }

        If ($LabelsFound) { 
            $OldSensitivityLabelName = $SensitivityLabelsFiles | Where-Object {$_.Id -eq $LabelsFound[0].sensitivityLabelId} | Select-Object -ExpandProperty Name
        }

        If ($Status -eq $true) {
            $ReportLine = [PSCustomObject]@{
                FileName               = $File.Name
                Folder                 = $File.parentreference.name
                Size                   = (FormatFileSize $File.Size)
                Created                = $FileCreatedDateTime
                Author                 = $File.CreatedBy.User.DisplayName
                LastModified           = $LastModifiedDateTime
               'Last modified by'      = $File.LastModifiedBy.User.DisplayName
               'Old Label Name'        = If ($OldSensitivityLabelName) {$OldSensitivityLabelName} Else {"No existing label"}
               'Action'                = If ($RemoveLabels -eq $true) {"Removed label"} Else {"Applied label"}
                Timestamp              = Get-Date -Format 'dd-MMM-yyyy HH:mm:ss'
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
Disconnect-MgGraph -ErrorAction SilentlyContinue

# The script is designed to run in app-only mode, so you need to create a registered app, make sure that the app can use the
# metered APIs, and grant it the permissions listed above. Then insert the app details in the variables listed below to connect
# in app-only mode using a certificate thumbprint. See https://aka.ms/graph-metered-overview for more information about the metered APIs and how to set up an app to use them.
# Insert values for your tenant and app here
$AppId = "282eb7dc-9c47-4f13-9620-d8471178b8ad"
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3471"
$CertThumbprint = "6FBD36F588E12DE291100371E16219863E399818"

# Connect to the Microsoft Graph
Connect-MgGraph -NoWelcome -AppId $AppId -TenantId $TenantId -CertificateThumbprint $CertThumbprint

Write-Host "Setting up to assign Sensitivity labels to all files in the target document libraries..."
$Global:SensitivityLabelsAvailable = $true
[array]$Global:ValidfileTypes = "docx", "pptx", "xlsx", "pdf", "doc", "ppt", "xls"

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
} Else {
    $Drives = $Drives | Where-Object {$_.Name -ne 'Preservation Hold Library'}
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
   Write-Host ("Can't find the selected drice ({0}) - script exiting" -f $Uri) ; break 
}

If ($NewSensitivityLabelName.tolower() -ne "remove") {
    Write-Host ("Checking if it's possuble to apply the sensitivity label {0} to files in the {1} document library for the {2} site, replacing any existing label if the new label has a higher priority than the existing one." -f $NewSensitivityLabelName, $DriveName, $SiteName) -ForegroundColor Green
    $Global:RemoveLabels = $false
    Write-Output "Fetching sensitivity labels defined for the tenant..."
    # Find sensitivity labels for the tenant 
    [array]$SensitivityLabels = Get-MgSecurityDataSecurityAndGovernanceSensitivityLabel -All
    # Reduce the set to the sensitivity labels that can be applied to files
    $SensitivityLabelsFiles = [System.Collections.Generic.List[Object]]::new()
    ForEach ($Label in $SensitivityLabels) {
        If ($Label.additionalProperties.applicableTo -like "*file*") {
            $ReportLine = [PSCustomObject]@{
                Id          = $Label.Id
                Name        = $Label.Name
                Priority    = $Label.Priority
                HasProtection = $Label.HasProtection
            }
            $SensitivityLabelsFiles.Add($ReportLine)
        }
    }

    [int]$Global:ReplacementLabelPriority = $SensitivityLabelsFiles | Where-Object {$_.Name -eq $NewSensitivityLabelName} | Select-Object -ExpandProperty Priority
    # Have we found a sensitivity label to apply? We will know if we have found a priority value for that label...
    If ($ReplacementLabelPriority) {
        $Global:SensitivityLabelId = $SensitivityLabelsFiles | Where-Object {$_.Name -eq $NewSensitivityLabelName} | Select-Object -ExpandProperty Id
        Write-Host ("Found chosen sensitivity label to apply: {0} (priority {1})" -f $NewSensitivityLabelName, $ReplacementLabelPriority) -ForegroundColor Green
    } Else {
        Write-Host ("Can't find the sensitivity label {0} to apply - script exiting" -f $NewSensitivityLabelName) -ForegroundColor Red
        break
    }
} Else {
    $Global:RemoveLabels = $true
    Write-Output "Option Selected to remove sensitivity labels from labeled files."
}

# Start the ball rolling to process files in the target document library
[datetime]$StartProcessing = Get-Date
$Global:TotalFolders = 1

# Create output list and CSV file
$Global:ReportData = [System.Collections.Generic.List[Object]]::new()
$CSVOutputFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + ("\Files {0}-{1} library.csv" -f $Site.displayName, $DriveName)
# Get the items in the root, including child folders
Write-Host "Checking files to process sensitivity labels..."
Get-DriveItems -Drive $Drive.Id -FolderId "root"
[datetime]$EndProcessing = Get-Date
$ElapsedTime = ($EndProcessing - $StartProcessing)
$FilesPerMinute = [math]::Round(($ReportData.Count / ($ElapsedTime.TotalSeconds / 60)), 2)
Write-Host ""
Write-Host ("Processed {0} files in {1} folders in {2}:{3} minutes ({4} files/minute)" -f `
   $ReportData.Count, $TotalFolders, $ElapsedTime.Minutes, $ElapsedTime.Seconds, $FilesPerMinute)

$ReportData | Out-GridView -Title ("Files in {0} document library for the {1} site" -f $DriveName, $SiteName)

# Generate the report in either Excel worksheet or CSV format, depending on if the ImportExcel module is available
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\SharePoint Labels.xlsx"
    $ReportData | Export-Excel -Path $ExcelOutputFile -WorksheetName "SharePoint Labels" -Title ("SharePoint Labels {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "SharePointLabels" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\SharePoint Labels.CSV"
    $ReportData | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}   

If ($ExcelGenerated -eq $true) {
    Write-Host ("SharePoint labels report is available in Excel workbook {0}" -f $ExcelOutputFile)
} Else {
    Write-Host ("SharePoint labels report is available in CSV file {0}" -f $CSVOutputFile)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
