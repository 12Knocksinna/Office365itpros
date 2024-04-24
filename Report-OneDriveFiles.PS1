# Report-OneDriveFiles.PS1
# A demo script to show how to generate a report of the files in a user's OneDrive account using 
# Microsoft Graph APIs. Based on the script to report files from a SharePoint Online document library 
# described in the article https://practical365.com/sharepoint-site-files-report/
# This script is https://github.com/12Knocksinna/Office365itpros/blob/master/Report-OneDriveFiles.PS1
# V1.0 19-Apr-2024

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
        # Only check for sensitivity labels if they are available
        If ($SensitivityLabelsAvailable -eq $true) {        
            $FileType = $File.Name.Split(".")[1]
            If ($FileType -notin $ValidFileTypes) { 
                # This file isn't an Office document or PDF file
                [string]$SensitivityLabelName = "Unsupported file type"
            } Else {
                $SensitivityLabelInfo = $Null
                [string]$SensitivityLabelName = "None"
                # Use Graph API to get sensitivity label information
                $Uri = ("https://graph.microsoft.com/V1.0/users/{0}/drive/items/{1}/extractSensitivityLabels" -f $Account, $File.id)
                Try {
                    [array]$SensitivityLabelInfo = Invoke-MgGraphRequest -Uri $Uri -Method POST 
                    If ($SensitivitylabelInfo.labels.count -eq 1) {
                        [string]$SensitivityLabelName = $SensitivityLabelsHash[$SensitivityLabelInfo.labels.sensitivitylabelid]
                    }
                    If ($SensitivityLabelInfo.labels.count -gt 1) {
                        [string]$SensitivityLabelName = $SensitivityLabelsHash[$SensitivityLabelInfo[0].labels.sensitivitylabelid]     
                    }
                } Catch {
                    Write-Host ("Error reading sensitivity label data from file {0}" -f $File.Name) -ForegroundColor Red
                    [string]$SensitivityLabelName = "Error"
                }
                If ([string]::IsNullOrWhiteSpace($SensitivityLabelName)) {
                    $SensitivityLabelName = "None"
                }
                #Write-Host ("File {0} has sensitivity label {1}" -f $File.Name, $SensitivityLabelName )
            }       
        }  
        # Get retention label information
        If ($RetentionLabelsAvailable -eq $true) {
            Try {
                $RetentionLabelInfo = $Null; $RetentionLabelName = $null
                $Uri = ("https://graph.microsoft.com/v1.0/drives/{0}/items/{1}/retentionLabel" -f $OneDriveInfo.id, $File.Id)
                [array]$RetentionLabelInfo = Invoke-MgGraphRequest -Uri $Uri -Method Get
                $RetentionLabelName = $RetentionLabelInfo.name
            } Catch {
                Write-Host ("Error reading retention label data from file {0}" -f $File.Name) 
            }
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
       
        $FileExtension = ($File.Name.Split('.')[-1]).ToUpper()
        Switch ($FileExtension) {
            "DOCX"      { $FileType = "Word document" }
            "DOC"       { $FileType = "Older Word document" }
            "DOCM"      { $FileType = "Word macro-enabled document"}
            "XLSX"      { $FileType = "Excel workbook" }
            "XLSB"      { $FileType = "Excel binary workbook" }
            "XLS"       { $FileType = "Excel spreadsheet" }
            "PPTX"      { $FileType = "PowerPoint presentation" }
            "PPT"       { $FileType = "Older PowerPoint presentation" }
            "PDF"       { $FileType = "PDF document" }
            "TXT"       { $FileType = "Text file" }
            "MP4"       { $FileType = "Video file" }
            "NOTE"      { $FileType = "OneNote notebook" }
            "ONE"       { $FileType = "OneNote .ONE file" }
            "ONETOC2"   { $FileType = "OneNote notebook TOC" }    
            "WBM"       { $FileType = "WebM video file" }
            "MOV"       { $FileType = "QuickTime movie" }
            "DLL"       { $FileType = "Dynamic link library" }
            "WAV"       { $FileType = "Wave audio file" }
            "FLUID"     { $FileType = "Loop component" }
            "LOOP"      { $FileType = "Loop file" }
            "POD"       { $FileType = "Loop workspace file" }    
            "CSV"       { $FileType = "CSV file" }
            "EDGE"      { $FileType = "Edge file" }
            "VSD"       { $FileType = "Visio diagram" }
            "WEBM"      { $FileType = "WebM video file" }
            "PNG"       { $FileType = "PNG image" }
            "JPG"       { $FileType = "JPEG image" }
            "JPEG"      { $FileType = "JPEG image" }
            "TEC"       { $FileType = "Camtasia file" }
            "MSG"       { $FileType = "Outlook message" }
            "EML"       { $FileType = "Email message" }
            "PS1"       { $FileType = "PowerShell script" }
            "PST"       { $FileType = "Outlook data file" }
            "JSON"      { $FileType = "JSON file" }
            "ZIP"       { $FileType = "ZIP archive" }
            "SAZ"       { $FileType = "Trace file" }
            "CLIPCHAMP" { $FileType = "Clipchamp video" }
            "WHITEBOARD" { $FileType = "Whiteboard file" }
            "PFILE"     { $FileType = "Power Automate file" }
            "ODS"       { $FileType = "OpenDocument spreadsheet" }
            "MHT"       { $FileType = "MHTML file" }
            "HTML"      { $FileType = "HTML file" }
            "XML"       { $FileType = "XML file" }
            "XLR"       { $FileType = "Works spreadsheet" }
            "INI"       { $FileType = "Configuration file" }
            "ICO"       { $FileType = "Icon file" }
            "JS"        { $FileType = "JavaScript file" }
            "PSM1"      { $FileType = "PowerShell module" }
            "TREC"      { $FileType = "Camtasia recording" }
            "VSSX"      { $FileType = "Visio stencil" }
            "BANNER"    { $FileType = "Banner file" }
            "ARTICLES"  { $FileType = "Articles file" }
            Default     { $FileType = "Unknown" }
        }
        $ReportLine = [PSCustomObject]@{
            FileName            = $File.Name
            Folder              = $File.parentreference.name
            Author              = $File.createdby.user.displayname
            Created             = $FileCreatedDateTime
            Modified            = $FileLastModifiedDateTime
            Size                = (FormatFileSize $File.Size)
            DaysOld             = $AgeInDays
            'Retention label'   = $RetentionLabelName
            'Sensitivity label' = $SensitivityLabelName
            FileType            = $FileType
            Bytes               = $File.Size
            Extension           = $FileExtension
            WebURL              = $File.WebUrl

        }
        $ODBFiles.Add($ReportLine)
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

Write-Host "Setting up for the OneDrive for Business account files report..."
# Discover if the tenant uses sensitivity labels
$Account = (Get-MgContext).Account
[array]$SensitivityLabels = Get-MgBetaUserSecurityInformationProtectionSensitivityLabel -UserId $Account
If ($SensitivityLabels) {
    $Global:SensitivityLabelsAvailable = $true
    [array]$Global:ValidfileTypes = "docx", "pptx", "xlsx", "pdf"
    $Global:SensitivityLabelsHash = @{}
    ForEach ($Label in $SensitivityLabels) {
        $SensitivityLabelsHash.Add($Label.Id, $Label.Name)
    }
} Else {
    $Global:SensitivityLabelsAvailable = $false
}

# Discover if the tenant uses retention labels
[array]$RetentionLabels = Get-MgSecurityLabelRetentionLabel
If ($RetentionLabels) {
    $Global:RetentionLabelsAvailable = $true
} Else {
    $Global:RetentionLabelsAvailable = $false
}

$Global:Site = ("https://graph.microsoft.com/v1.0/users/{0}/drive/" -f $Account)
[array]$OneDriveInfo = Invoke-MgGraphRequest -Uri $Site -Method Get
If (!($OneDriveInfo)) { # Nothing found
    Write-Host "No matching OneDrive found - exiting"; break 
} Else {
    Write-Host ("Found OneDrive account owned by {0} to process. URL: {1}" -f $OneDriveInfo.owner.user.displayName, $OneDriveInfo.WebUrl)
    $Global:OneDriveName = $OneDriveInfo.name
}

# Create output list and CSV file
$Global:ODBFiles = [System.Collections.Generic.List[Object]]::new()
$CSVOutputFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + ("\OneDrive files for {0}.csv" -f $OneDriveInfo.owner.user.displayName)

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

Clear-Host
Write-Host "Generating analysis of OneDrive for Business files..." -ForegroundColor Yellow
# Generate some statistics
[array]$FileTypeCount = $ODBFiles | Group-Object FileType -NoElement | Sort-Object Count -Descending | `
    Select-Object Name, Count

# Analysis of types used in the OneDrive for Business account
$ReportData = [System.Collections.Generic.List[Object]]::new()
ForEach ($FT in $FileTypeCount.Name) {
    $FTItems = $ODBFiles | Where-Object {$_.FileType -eq $FT}
    $FileExtensionData = ($FTItems.Bytes | Measure-Object -AllStats)
    $FileCount = $FileExtensionData.Count
    $FileSize = FormatFileSize $FileExtensionData.Sum
    $FileAverageSize = FormatFileSize $FileExtensionData.Average
    $ReportLine = [PSCustomObject]@{
        'File Type'     = $FT
        Count           = $FileCount
        Size            = $FileSize
        'Average Size'  = $FileAverageSize
    }
    $ReportData.Add($ReportLine)
}

# Quota
$QuotaTotal = FormatFileSize $OneDriveInfo.quota.total
$QuotaUsed = FormatFileSize $OneDriveInfo.quota.used
$QuotaAvailable = FormatFileSize $OneDriveInfo.quota.remaining
$PercentUsed = ($OneDriveInfo.quota.used/$OneDriveInfo.quota.total).toString('P')

# Oldest files
[array]$OldestFiles = $ODBFiles | Sort-Object DaysOld -Descending | Select-Object -First 10 | `
    Select-Object FileName, Folder, Author, Created, Modified, Size, Uri, DaysOld

# Largest files
[array]$TopFiles = $ODBFiles | Sort-Object Bytes -Descending | Select-Object -First 10 | `
    Select-Object FileName, Folder, Author, Created, Modified, Size, Uri

Clear-Host
Write-Host "OneDrive for Business account statistics"
Write-Host "----------------------------------------"
Write-Host ("Total files found in the OneDrive for Business account of {0}: {1}" `
    -f $OneDriveInfo.owner.user.displayName, $ODBFiles.Count) -ForegroundColor Red
Write-Host ("Quota assigned {0}" -f $QuotaTotal)
Write-Host ("Quota used: {0} ({1})" -f $QuotaUsed, $PercentUsed)
Write-Host ("Quota remaining {0}" -f $QuotaAvailable)
Write-Host ""
$ReportData | Format-Table -AutoSize

Write-Host "Largest ten files in the account"
Write-Host "--------------------------------"
$TopFiles | Format-Table FileName, Created, Modified, Size -AutoSize
Write-Host ""

Write-Host "Oldest ten files in the account"
Write-Host "-------------------------------"
$OldestFiles | Format-Table FileName, Created, Modified, DaysOld, Size -AutoSize
Write-Host ""
$ODBFiles | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding UTF8
Write-Host ("Report data saved to file: {0}" -f $CSVOutputFile)
Write-Host ""

If ($RetentionLabelsAvailable) {
    $PercentRetentionLabels = (($ODBFiles | Where-Object {$_.'Retention Label' -ne "No label"}).Count /$ODBFiles.Count).toString('P')
    Write-Host ("Retention Label Usage is {0} of {1} files" -f $PercentRetentionLabels, $ODBFiles.Count)
    $ODBfiles | Group-Object 'Retention label' -NoElement | Sort-Object Count -Descending | Format-Table Name, Count -AutoSize
    Write-Host ""
}
If ($SensitivityLabelsAvailable) {
    [array]$ExcludedLabels = "None", "Unsupported file type"
    # Files that could receive sensitivity labels
    [array]$ValidFilesforLabels = $ODBFiles | Where-Object {$_.Extension -in $ValidFileTypes}
    # Files that have received a sensivity label
    [array]$ValidLabelledFiles = $ValidFilesForLabels | Where-Object {$_.'Sensitivity Label' -notin $ExcludedLabels}
    $PercentSensitivityLabels = ($ValidLabelledFiles.count/$ValidFilesForLabels.Count).toString('P')
    Write-Host ("Sensitivity Label Usage is {0} of {1} files that support labels ({2} total files)" `
        -f $PercentSensitivityLabels, $ValidFilesForLabels.Count, $ODBFiles.count)
    $ODBfiles | Group-Object 'Sensitivity label' -NoElement | Sort-Object Count -Descending | Format-Table Name, Count -AutoSize
    Write-Host ""
}

Write-Host ""
Write-Host ("Processed {0} files in {1} folders in {2}:{3} minutes ({4} files/minute)" -f `
   $ODBFiles.Count, $TotalFolders, $ElapsedTime.Minutes, $ElapsedTime.Seconds, $FilesPerMinute)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
