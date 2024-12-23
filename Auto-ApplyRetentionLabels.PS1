# Auto-ApplyRetentionLabels.PS1
# An example of how to apply retention labels to files in SharePoint Online and OneDrive for Account sites

# V1.0 December 2024
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Auto-ApplyRetentionLabels.PS1

# Requires the following permissions 
# RecordsManagement.ReadWrite.All (retention labels)
# Sites.Read.All (to access sites)
# Files.Read.All (to access files) (Sites.Read.All includes Files.Read.All)

function Get-DriveItems {
    [CmdletBinding()]
    param (
        [Parameter()]
        $Drive,
        [Parameter()]
        $FolderId
    )
    # Get data for a folder and its children, check each file to see if it has a retention label. If it has, and the
    # retention threshold is not met, apply the default retention label defined in the script
    [array]$Data = Get-MgDriveItemChild -DriveId $Drive -DriveItemId $FolderId -All
    # Split the data into files and folders
    [array]$Folders = $Data | Where-Object {$_.folder.childcount -gt 0} | Sort-Object Name
    [array]$Files = $Data | Where-Object {$null -ne $_.file.mimetype} 

    # Process the files
    ForEach ($File in $Files) {   
        # Get retention label information
        $FileExtension = $File.Name.Split(".")[1].ToUpper()
        If ($FileExtension -notin $SupportedFileTypes) {
            continue
        }
        If ($File.CreatedDateTime) {
            $FileCreatedDateTime = Get-Date $File.CreatedDateTime -format 'dd-MMM-yyyy HH:mm'
        }
        # Is the file already older than the retention period? If so, we leave it alone
        If ($File.LastModifiedDateTime) {
            [datetime]$LastModifiedDateTime = Get-Date $File.LastModifiedDateTime 
            If ($LastModifiedDateTime -lt $RetentionDuration) {
                Write-Host ("File {0} last modified {1} is older than retention period" -f $File.Name, $LastModifiedDateTime)
                $ReportLine = [PSCustomObject]@{
                    TimeStamp               = (Get-Date -format s)
                    FileName                = $File.Name
                    Folder                  = $File.parentreference.name
                    Created                 = $FileCreatedDateTime
                    Author                  = $File.CreatedBy.User.DisplayName
                    LastModified            = Get-Date $LastModifiedDateTime -format 'dd-MMM-yyyy HH:mm'
                    'Last modified by'      = $File.LastModifiedBy.User.DisplayName
                    'Retention label'       = $RetentionLabel['Name']
                    Path                    = $Site.WebUrl
                    Action                  = "Retention label not applied - modified date older than retention period"
                }
                $ReportData.Add($ReportLine)
                Continue
            }
        } Else {
            $LastModifiedDateTime = $null
        }
        Try {
            $FileInfo =  Get-MgDriveItemRetentionLabel -DriveId $Drive -DriveItemId $File.Id -ErrorAction Stop
        } Catch {
            Write-Host ("Error reading retention label data from file {0}" -f $File.Name) 
            Continue
        }
        If ([string]::IsNullOrEmpty($FileInfo.Name)) {
            # Attempt to apply the defined retention label
            If (!$PreviewMode) {
                $Status = Update-MgDriveItemRetentionLabel -DriveId $Drive -DriveItemId $File.Id -BodyParameter $RetentionLabel
                If ($Status.Name) { 
                    Write-Host ("Retention label assigned to {0}" -f $File.Name) -ForegroundColor DarkGray
                    $ReportLine = [PSCustomObject]@{
                        TimeStamp               = (Get-Date -format s)
                        FileName                = $File.Name
                        Folder                  = $File.parentreference.name
                        Created                 = $FileCreatedDateTime
                        Author                  = $File.CreatedBy.User.DisplayName
                        LastModified            = Get-Date $LastModifiedDateTime -format 'dd-MMM-yyyy HH:mm'
                        'Last modified by'      = $File.LastModifiedBy.User.DisplayName
                        'Retention label'       = $RetentionLabel['Name']
                        Path                    = $Site.WebUrl
                        Action                  = "Retention label applied"
                    }
                    $ReportData.Add($ReportLine)
                } Else {    
                    Write-Host ("Unable to assign retention label to {0}" -f $File.Name) -ForegroundColor Red
                    Continue
                }
            } Else {
                # Preview Mode, so just log the fact that we would have applied the label
                $ReportLine = [PSCustomObject]@{
                    TimeStamp               = (Get-Date -format s)
                    FileName                = $File.Name
                    Folder                  = $File.parentreference.name
                    Created                 = $FileCreatedDateTime
                    Author                  = $File.CreatedBy.User.DisplayName
                    LastModified            = Get-Date $LastModifiedDateTime -format 'dd-MMM-yyyy HH:mm'
                    'Last modified by'      = $File.LastModifiedBy.User.DisplayName
                    'Retention label'       = $RetentionLabel['Name']
                    Path                    = $Site.WebUrl
                    Action                  = "Retention label not applied (preview mode)"
                }
                $ReportData.Add($ReportLine)
            }        
        }     
    }

    # Process the folders found in the root
    ForEach ($Folder in $Folders) {
        Write-Host ("Processing folder {0}" -f $Folder.Name) -ForegroundColor Green
        Get-DriveItems -Drive $Drive -FolderId $Folder.Id
    }
}

# Start of Main Script

param (
    [switch]$Global:PreviewMode
)

if ($PreviewMode) {
    Write-Host "Running in preview mode. No changes will be made."
}
# Start of real work
Connect-MgGraph -Scopes RecordsManagement.ReadWrite.All, Sites.Read.All

$LocationsFile = 'c:\temp\Locations.csv'
If (!(Test-Path $LocationsFile)) {
    Write-Host "Locations file for unlabeled files not found - exiting"
    Break
}
[array]$AllLocations = Import-Csv -Path $LocationsFile
# Exclude all OneDrive for Business sites
[array]$Locations = $AllLocations | Where-Object {$_.Location -notlike "*my.sharepoint.com/personal*"}

# Define default retention label to apply
$Global:RetentionLabel = @{}
$RetentionLabel.Add("Name","General Purpose Information")
$Global:RetentionDuration = (Get-Date).AddYears(-3)

# Supported file types that we will apply retention labels to
$Global:SupportedFileTypes = "DOCX", "PPTX", "XLSX", "PDF"
# Output PowerShell lists for reports
$Global:ReportData = [System.Collections.Generic.List[Object]]::new()
$Global:ProblemSites = [System.Collections.Generic.List[Object]]::new()

# Let people know what we plan to do
If ($PreviewMode) {
    Write-Host "Running in preview mode. No retention labels will be applied to files."
}
# Try to access each site. If we can gain access, look for unlabeled files
ForEach ($Location in $Locations) {
    $Uri = $Location.Location; $Site = $null; $LookupUri = $null
    # Create a value that we can find with a call to Get-MgSite. The resulting value will be something like
    # office365itpros.sharepoint.com:/sites/SeniorTeam
   # $LookUpUri = $Uri.Split('//')[1].split("/")[0] + ":/sites/" + $Uri.Split('//')[1].split("/")[2]
    Try {
        $Global:Site = Get-MgSite -Search $Uri -ErrorAction Stop
    } Catch {
        # Try to find the site with the site name
        Try {
            $Global:Site = Get-MgSite -Search $Uri.Split("/sites/")[1] -ErrorAction Stop
        } Catch {
            Write-Host ("Unable to access site {0} {1}" -f $Uri, $_.Exception.Message) -ForegroundColor Red
            # Log the problem site  
            $ProblemSiteReportLine = [PSCustomObject]@{
                TimeStamp               = (Get-Date -format s)
                Site                    = $Uri
                Action                  = "Unable to access site"
            }
            $ProblemSites.Add($ProblemSiteReportLine)
            Continue
        }
    }
    If (!$Site) {
        Write-Host ("Site {0} not found" -f $Uri) -ForegroundColor Red
        Continue
    } Else {
        Write-Host ("Processing site {0} to look for unlabeled files" -f $Site.DisplayName) -ForegroundColor Yellow
    }
    Try {
        [array]$AllDrives = Get-MgSiteDrive -SiteId $Site.Id -ErrorAction Stop
    } Catch {
        Write-Host ("Unable to access drives in site {0} ({1}) {2}" -f $Site.DisplayName, $URI, $_.Exception.Message) -ForegroundColor Red
        $ProblemSiteReportLine = [PSCustomObject]@{
            TimeStamp               = (Get-Date -format s)
            Site                    = $Uri
            Action                  = "Unable to access drives"
        }
        $ProblemSites.Add($ProblemSiteReportLine)
        Continue
    }

    [array]$Drives = $AllDrives | Where-Object {$_.Name -notlike "*Preservation Hold Library*" -and $_.Name -notlike "*Teams Wiki Data*"}
    If (!$Drives) {
        Write-Host "No drives found in site $URI"
        Continue
    }
    # Process each drive
    ForEach ($Drive in $Drives) {
        Write-Host ("Processing drive {0} in site {1}" -f $Drive.Name, $Site.DisplayName) -ForegroundColor Cyan
        Get-DriveItems -Drive $Drive.Id -FolderId "root"
    }
}

Write-Host "The following sites could not be processed"
$ProblemSites | Format-Table Site, Action -AutoSize
[array]$SitesWithLabeledFiles = $ReportData | Where-Object {$_.Action -eq 'Retention label applied'} | Select-Object -ExpandProperty Path | Sort-Object -Unique

Write-Host ""
If ($PreviewMode) {
    Write-Host "Retention Labeling Run Statistics (Preview mode)"
    Write-Host "------------------------------------------------"
} Else {
    Write-Host "Retention Labeling Run Statistics"
    Write-Host "--------------------------------"
}
Write-Host ""
Write-Host ("Retention label used: {0}" -f $RetentionLabel['Name'])
Write-Host ("Retention date threshold: {0} ({1} days)" -f $RetentionDuration, (New-TimeSpan -Start $RetentionDuration).Days)
Write-Host ("Total {0} of files processed in {1} locations" -f $ReportData.Count, ($Locations.count - $ProblemSites.count))
If ($PreviewMode) {
    Write-Host ("Retention labels not applied because of preview mode: {0}" -f ($ReportData | Where-Object {$_.Action -eq "Retention label not applied (preview mode)"}).Count)
} Else {
    Write-Host ("Retention labels applied: {0}" -f ($ReportData | Where-Object {$_.Action -eq "Retention label applied"}).Count)
}
Write-Host ("Files older than the threshold to apply retention label: {0}" -f ($ReportData | Where-Object {$_.Action -eq "Retention label not applied - modified date older than retention period"}).Count)
Write-Host ""
Write-Host "Sites where labels were applied:"
$SitesWithLabeledFiles

Write-Host ""
Write-Host "Generating report..."
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Retention Labeling.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    }
    $ReportData | Export-Excel -Path $ExcelOutputFile -WorksheetName "Retention Labeling Report" -Title ("Retention Labeling Report") -TitleBold -TableName "RetentionLabels" 
   
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Retention Labeling.CSV"
    $ReportDtra | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}
 
If ($ExcelGenerated) {
    Write-Host ("An Excel report is available in {0}" -f $ExcelOutputFile)
} Else {    
    Write-Host ("A CSV report is available in {0}" -f $CSVOutputFile)
}  

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.