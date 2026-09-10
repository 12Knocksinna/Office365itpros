# Archive-EmailMessagesInSPO.PS1
# This script is designed to archive email messages in SharePoint Online (SPO) that are older than a threshold date/

# Needs to be run with an account or Azure Automation account that has access to the target SharePoint Online 
# sites. Given the number of files that might be checked, best to use an Azure Automation runbook with a managed identity that has access to the sites. The script uses Microsoft Graph PowerShell SDK to connect and retrieve data.

# V1.0 8-Sept-2026
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Archive-EmailMessagesinSPO.PS1

# Files.ReadWrite.All is needed to archive files in SharePoint Online

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

function Get-DriveItems {
    [CmdletBinding()]
    param (
        [Parameter()]
        $DriveId,
        [Parameter()]
        $FolderId
    )

    # Return details of archived files found in a document library. The function is recursive and will process 
    # all folders in the library. The function returns a list of archived files found in the library.

    # Get data for a folder and its children - Get-MgDriveItemChild doesn't return the archive status currently
    # Equivalent Graph SDK cmdlet when the SDK is updated (last checked with V2.39)
    # [array]$Data = Get-MgDriveItemChild -DriveId $Drive -DriveItemId $FolderId -All

    $FolderUri = ("https://graph.microsoft.com/v1.0/drives/{0}/items/{1}" -f $DriveId, $FolderId)

    Try {
        $FolderInfo = Invoke-MgGraphRequest -Uri $FolderUri -Method Get -OutputType PSObject -ErrorAction Stop
        $FolderName = $FolderInfo.Name
    } Catch {
        Write-Error "Failed to retrieve folder info: $_"
        Return
    }
    
    # Need to fetch files and folders from the current folder to check if there's anything to process.
    $Uri = ("https://graph.microsoft.com/V1.0/drives/{0}/items/{1}/children" -f $DriveId, $FolderId)
        Try {
            [array]$Data = Invoke-MgGraphRequest -Uri $Uri -Method Get -OutputType PsObject -ErrorAction Stop
        } Catch {
            Write-Error "Failed to retrieve drive items: $_"
            Return
        }

    If ($Data.'@odata.nextlink') {
        # More data needs to be retrieved, so get the next page of data
        $NextLink = $Data.'@odata.nextlink'
        $FolderData = $Data.Value
        While ($NextLink) {
            Try {
                [array]$NextData = Invoke-MgGraphRequest -Uri $NextLink -Method Get -OutputType PsObject 
            } Catch {
                Write-Error "Failed to retrieve next page of drive items: $_"
                Return
            }
            $FolderData += $NextData.Value
            $NextLink = $NextData.'@odata.nextlink'
        }
    } Else {
        $FolderData = $Data.Value
    }

    [array]$Folders = $FolderData | Where-Object {$_.folder.childcount -gt 0} | Sort-Object Name
    $Global:TotalFolders = $Global:TotalFolders + $Folders.Count
    $Global:TotalFiles = $Global:TotalFiles + ($FolderData.Count - $Folders.Count)
    # Extract unarchived files from the set returned
    [array]$Files = $FolderData | Where-Object {$null -ne $_.file.mimetype -and $null -eq $_.file.archivestatus} | Sort-Object Name

    # We only want to process files in folders that match the pattern for email messages folders. The pattern is 'Email Messages' or 'Email_Messages' or 'Email-Messages' or 'EmailMessages'. The regex pattern used is '^Email([ _])?Messages'.
    $ProcessFolder = $FolderName -match '^Email([ _])?Messages' 

    If (!$ProcessFolder) {
        Write-Host ("Skipping folder {0} as it does not match an email messages folder name" -f $FolderName)
        # Process the child folders 
        ForEach ($Folder in $Folders) {
            # Write-Host ("====> Checking folder {0} ({1} files/size {2})" -f $Folder.Name, $Folder.folder.childcount, (FormatFileSize $Folder.Size))
            Get-DriveItems -DriveId $DriveId -FolderId $Folder.Id
        }
    } Else {   
        If ($Files.count -gt 0) {
            Write-Host ("!!! Processing files in email folder {0} ({1} files/size {2})" -f $FolderName, $Files.Count, (FormatFileSize ($Files | Measure-Object -Property Size -Sum).Sum))
            # Process the files found in the folder.
            ForEach ($File in $Files) {  

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

                If ([datetime]$File.LastModifiedDateTime -lt $Global:ArchiveThresholdDate) {
                # File is older than the threshold date, so it will be archived
                    $Uri = ("https://graph.microsoft.com/beta/drives/{0}/items/{1}/archive" -f $DriveId, $File.Id)
                    Try {
                        Invoke-MgGraphRequest -Method POST -Uri $Uri -OutputType PSobject -ErrorAction Stop
                        Write-Host ("File {0} is now archived" -f $File.Name) -ForegroundColor Yellow
                    } Catch {
                        Write-Host ("Failed to archive file {0}: {1}" -f $File.Name, $_.Exception.Message) -ForegroundColor Red
                    }
                    $ArchivedFiles++
                    # Basic throttling
                    Start-Sleep -Milliseconds 25
                } Else {
                    # File is newer than the threshold date, so it will not be archived
                    Continue
                }   

                $YearCreated = Get-Date $FileCreatedDateTime -format 'yyyy'
                $FileType = $File.file.fileExtension
                # Report what's been done
                $ReportLine = [PSCustomObject]@{
                    Site                    = $SiteName
                    FileName                = $File.Name
                    Folder                  = $File.parentreference.name
                    Size                    = (FormatFileSize $File.Size)
                    Created                 = $FileCreatedDateTime
                    Author                  = $File.CreatedBy.User.DisplayName
                    LastModified            = $LastModifiedDateTime
                    'Last modified by'      = $File.LastModifiedBy.User.DisplayName
                    SizeInBytes             = $File.Size
                    YearCreated             = $YearCreated
                    FileType                = $FileType
                }
                $Global:Report.Add($ReportLine)
            }
        }

        # Process the child folders - but only those matching the email messages pattern
        $Folders = $Folders | Where-Object {$_.Name -match '^Email([ _])?Messages'}
        ForEach ($Folder in $Folders) {
        # Write-Host ("====> Checking folder {0} ({1} files/size {2})" -f $Folder.Name, $Folder.folder.childcount, (FormatFileSize $Folder.Size))
            Get-DriveItems -DriveId $DriveId -FolderId $Folder.Id
        }
    }
}

# Start of processing
# This is the code for app-only authentication. The permissions needed are
# Sites.ReadWrite.All and Group.Read.All. The app must be registered in Entra ID and the certificate uploaded to the app.
# registration. The script uses the certificate to authenticate to Microsoft Graph. Amend the values for $AppId
# $TenantId and $CertificateThumbprint to match your app registration.

$AppId = 'd4c18027-cb9c-4983-b2d8-8ba6a5e753e7'
$TenantId = 'a662313f-14fc-43a2-9a7a-d2e27f4f3478'
$CertificateThumbprint = '8D225839321765FE6D224D393CDE5A61A45CD846'
Connect-MgGraph -NoWelcome -AppId $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint

# Define value for input file containing Teams channels with email addresses. This file can be created using the 
# script in https://github.com/12Knocksinna/Office365itpros/blob/master/Report-TeamsChannelEmailAddresses.PS1
$InputFile = "C:\Temp\Teams Channels with Email Addresses.csv"
# Threshold for archiving files in days. Files older than this threshold will be archived.
[datetime]$Global:ArchiveThresholdDate = Get-Date (Get-Date).AddYears(-3) # 3 years ago. Change this value to set a different threshold for archiving files.
Write-Host ("Email items will be archived if dated before {0} " -f (Get-Date $Global:ArchiveThresholdDate -format 'dd-MMM-yyyy HH:mm'))

# The script can check the SPO sites for all Microsoft 365 Groups to find folders email-backed channels. It's obviously much more
# efficient to use an input file with data from a CSV file. The CSV file is generated by examining Teams for email-enabed channels.
If (-Not (Test-Path -Path $InputFile)) { 
    # Fetch all Microsoft 365 Groups in the tenant
    Try {
        [array]$Groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'unified')" -All -ErrorAction Stop
    } Catch {
        Write-Error "Failed to retrieve Microsoft 365 groups: $_"
        Return
    } 
} Else {
    # Read in the information from the input file
    [array]$Groups = Import-Csv -Path $InputFile | Select-Object Id, DisplayName
    # Remove any duplicates so we don't process a group multiple times
    $Groups = $Groups | Sort-Object Id -Unique 
}

$Groups = $Groups | Sort-Object DisplayName
Write-Output ("{0} Microsoft 365 groups found for checking..." -f $Groups.Count)

[int]$Global:TotalFiles = 0
[int]$Global:TotalFolders = 0
[int]$Global:ArchivedFiles = 0
[int]$Global:BlockedSites = 0
$Global:Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Group in $Groups) {
    Write-Output ("Processing group: {0}..." -f $Group.DisplayName)
    Try {
    # $GroupSite = Invoke-MgGraphRequest -Uri $Uri -Method Get -OutputType PsObject -ErrorAction Stop
        $GroupSite = Get-MgGroupSite -GroupId $Group.Id -SiteId 'root' -ErrorAction Stop
    } Catch {
        $BlockedSites++
        $ErrorMessage = $_.Exception.Message
        If ($ErrorMessage -match "notAllowed") {
            Write-Warning "Site is archived or otherwise blocked: $SiteId"
        } Else {
            Write-Warning ("Error {0}: {1}" -f $_.Exception.ResponseStatusCode, $_.Exception.Message)
    }
        Continue
    }

    $SiteId = $GroupSite.Id
    $Global:SiteName = $GroupSite.displayName
    
    Try {
        # Get the document libraries for the site.
        # $Drive = Get-MgSiteDefaultDrive -SiteId $SiteId -ErrorAction Stop
        [array]$Drives = Get-MgSiteDrive -SiteId $GroupSite.Id -ErrorAction Stop
    } Catch {
        Write-Error "Failed to retrieve drives for site $($GroupSite.DisplayName): $_"
        Continue
    }

    # Remove the preservation hold library from the list of drives to process, as it is not relevant for this report.
    $Drives = $Drives | Where-Object {$_.Name -ne 'Preservation Hold Library' -and $_.Name -ne 'Teams Wiki Data'}
    ForEach ($Drive in $Drives) {
        Write-Output ("==> Processing document library: {0}..." -f $Drive.Name)
    
        $DriveId = $Drive.Id
     
        Try {
            $RootId = Get-MgDriveRoot -DriveId $DriveId -ErrorAction Stop
        } Catch {
            Write-Error "Failed to retrieve root folder for drive $($Drive.Name): $_"
            Continue
        }
        $RootId = $RootId.Id
        [array]$Files = Get-DriveItems -DriveId $DriveId -FolderId $RootId
    }
}

$TotalArchivedFilesSize = ($Global:Report | Measure-Object -Property SizeInBytes -Sum).Sum

Write-Output ""
Write-Output ("Total Microsoft 365 groups processed: {0}" -f $Groups.Count)
If ($BlockedSites -gt 0) {
    Write-Output ("Unable to access Microsoft 365 groups: {0}" -f $BlockedSites)
}
Write-Output ("Total folders processed: {0}" -f $Global:TotalFolders)
Write-Output ("Total files processed: {0}" -f $Global:TotalFiles)
Write-Output ("Total email files archived: {0}" -f $Global:Report.Count)
Write-Output ("Total size of files archived: {0}" -f (FormatFileSize $TotalArchivedFilesSize))

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 