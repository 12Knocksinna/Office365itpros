# Report-ArchivedSPOFiles.PS1

# A script to report on SharePoint Online files that have been archived (moved to Microsoft 365 Archive), either by
# explict user action or by retention policy or label. The script reports on the files that have been archived, 
# including the file name, path, and the date it was archived.

# Needs to be run with an account or Azure Automation account that has access to the target SharePoint Online 
# sites. Given the number of files that might be checked, best to use an Azure Automation runbook with a managed identity that has access to the sites. The script uses Microsoft Graph PowerShell SDK to connect and retrieve data.

# V1.0 2-Sept-2026
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ArchivedSPOFiles.PS1

# Requires Sites.Read.All to read site information and GroupMember.Read.All to read group details. Also Mail.Send to send the email.

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

    $Uri = ("https://graph.microsoft.com/V1.0/drives/{0}/items/{1}/children" -f $DriveId, $FolderId)
    Try {
        [array]$Data = Invoke-MgGraphRequest -Uri $Uri -Method Get -OutputType PsObject 
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
    [array]$ArchivedFiles = $FolderData | Where-Object {$null -ne $_.file.mimetype -and $null -ne $_.file.archiveStatus}
    
    # Process the files
    ForEach ($File in $ArchivedFiles) {  

        If ($File.LastModifiedDateTime) {
            $LastModifiedDateTime = Get-Date $File.LastModifiedDateTime -format 'dd-MMM-yyyy HH:mm'
        } Else {
            $LastModifiedDateTime = $null
        }
        If ($File.CreatedDateTime) {
            $FileCreatedDateTime = Get-Date $File.CreatedDateTime -format 'dd-MMM-yyyy HH:mm'
        }

        $YearCreated = Get-Date $FileCreatedDateTime -format 'yyyy'
        $FileType = $File.file.fileExtension
        Switch ($FileType) {
            '.csv' { $FileType = 'CSV File' }
            '.docx' { $FileType = 'Word Document' }
            '.doc' { $FileType = 'Word Document (old format)' }
            '.mht' { $FileType = 'MHTML Document' }
            '.eml' { $FileType = 'Email Message' }
            '.xlsx' { $FileType = 'Excel Spreadsheet' }
            '.xls' { $FileType = 'Excel Spreadsheet (old format)' }
            '.pptx' { $FileType = 'PowerPoint Presentation' }
            '.ppt' { $FileType = 'PowerPoint Presentation (old format)' }
            '.pdf' { $FileType = 'PDF Document' }
            '.jpg' { $FileType = 'JPEG Image' }
            '.jpeg' { $FileType = 'JPEG Image' }
            '.mp4' { $FileType = 'MP4 Video' }
            default { $FileType = $File.file.mimeType }
        }
        
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
            ArchiveStatus           = $File.File.ArchiveStatus
            YearCreated             = $YearCreated
            FileType                = $FileType
        }
        $ReportData.Add($ReportLine)
    }

    # Process the child folders 
    ForEach ($Folder in $Folders) {
        Write-Host ("====> Processing folder {0} ({1} files/size {2})" -f $Folder.Name, $Folder.folder.childcount, (FormatFileSize $Folder.Size))
        Get-DriveItems -DriveId $DriveId -FolderId $Folder.Id
    }
}

# Start of processing
# This is the code for app-only authentication. The permissions needed are
# Sites.Read.All and Group.Read.All. The app must be registered in Entra ID and the certificate uploaded to the app.
# registration. The script uses the certificate to authenticate to Microsoft Graph. Amend the values for $AppId
# $TenantId and $CertificateThumbprint to match your app registration.

$AppId = 'd4c18027-cb9c-4983-b2d8-8ba6a5e753e7'
$TenantId = 'a662313f-14fc-43a2-9a7a-d2e27f4f3478'
$CertificateThumbprint = '8D225839321765FE6D224D393CDE5A61A45CD846'
Connect-MgGraph -NoWelcome -AppId $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint

$ProcessingStartTime = Get-Date

# Change the destination email address to receive the emailed report and the mailbox from which the report is sent
# to match valid recipients in your tenant.
$DestinationEmailAddress = 'Customer.Services@office365itpros.com'
$MsgFrom = 'Customer.Services@office365itpros.com'

Write-Output "Fetching Microsoft 365 groups...."
# Find the Microsoft 365 groups in the tenant. We use the groups to find the associated SharePoint Online sites.
# The script will then check each site for archived files.
Try {
    [array]$Groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'unified')" -All -ErrorAction Stop
} Catch {
    Write-Error "Failed to retrieve Microsoft 365 groups: $_"
    Return
}

$Groups = $Groups | Sort-Object DisplayName
Write-Output ("{0} Microsoft 365 groups found" -f $Groups.Count)
[int]$BlockedSites = 0

$Global:ReportData = [System.Collections.Generic.List[Object]]::new()
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
    $Drives = $Drives | Where-Object {$_.Name -ne 'Preservation Hold Library' -or $_.Name -ne 'Teams Wiki Data'}
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
        [array]$ArchivedFiles = Get-DriveItems -DriveId $DriveId -FolderId $RootId
    }
}

# Generate some statistics
$TotalArchivedFilesSize = ($Global:ReportData | Measure-Object -Property SizeInBytes -Sum).Sum
$StorageSavingsBytes = ($TotalArchivedFilesSize / 4)*3
$StorageSavings = FormatFileSize $StorageSavingsBytes

Write-Output ""
Write-Output ("Total Microsoft 365 groups found for processing: {0}" -f $Groups.Count)
If ($BlockedSites -gt 0) {
    Write-Output ("Unable to access Microsoft 365 groups: {0}" -f $BlockedSites)
}
Write-Output ("Total folders processed: {0}" -f $Global:TotalFolders)
Write-Output ("Total files processed: {0}" -f $Global:TotalFiles)
Write-Output ("Total archived files found: {0}" -f $Global:ReportData.Count)
Write-Output ("Total size of archived files found: {0}" -f (FormatFileSize $TotalArchivedFilesSize))
Write-Output ("Estimated storage savings from archived files: {0}" -f $StorageSavings)
Write-Output ""

# Analyze the years when archived files were created
Write-Output "Archived Files by Year of Creation"
Write-Output "----------------------------------"
$ArchivedFilesByYearOutput =  $Global:ReportData | Group-Object YearCreated | Select-Object @{
    Name='Year created'; Expression={$_.Name}}, @{Name='Count'; Expression={$_.Count} }
$ArchivedFilesByYearOutput 

# Analyze the file types of archived files
Write-Output ""
Write-Output "Archived Files by Type"
Write-Output "----------------------"
    
$ArchivedFilesByTypeOutput =  $Global:ReportData | Group-Object FileType | Select-Object @{
    Name='File type'; Expression={$_.Name}}, @{Name='Count'; Expression={$_.Count} }
$ArchivedFilesByTypeOutput | Format-Table -AutoSize

Write-Output ""
Write-Output "Sites holding archived files"
Write-Output "----------------------------"
$SitesWithArchivedFiles = $Global:ReportData | Group-Object Site -NoElement | Select-Object @{
    Name='Site name'; Expression={$_.Name}}, @{Name='Number of archived files'; Expression={$_.Count} }
$SitesWithArchivedFiles

$ProcessingEndTime = Get-Date
$ProcessingDuration = $ProcessingEndTime - $ProcessingStartTime
Write-Output ""
Write-Output "Processing completed in $($ProcessingDuration.ToString('hh\:mm\:ss'))"
Write-Output ""

# Build the array of a single TO recipient detailed in a hash table - change this to the appropriate recipient for your tenant
$ToRecipient = @{}
$ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
[array]$MsgTo = $ToRecipient

# Define the message subject
$MsgSubject = "File-Level Archiving Report for SharePoint Online Sites"
# Create the HTML content
$HtmlMsg = "<html><body>"
$HtmlMsg += "<h1>File-Level Archiving Report for SharePoint Online Sites</h1>"
$HtmlMsg += "<p>Total Microsoft 365 groups found for processing: <b>$($Groups.Count)</b></p>"
If ($BlockedSites -gt 0) {
    $HtmlMsg += "<p>Unable to access Microsoft 365 groups: <b>$($BlockedSites)</b>. This may be because the site is archived.</p>"
}
$HtmlMsg += "<p>The information contained in this report is calculated based on an examination of file properties found in the SharePoint Online sites belonging to the Microsoft 365 Groups scanned by the job..</p>"
$HtmlMsg += "<h2>Summary of Archived Files by Site</h2>"
$HtmlMsg += $SitesWithArchivedFiles | ConvertTo-Html -Fragment
$HtmlMsg += "<h2>Summary of Archived Files by Year of Creation</h2>"
$HtmlMsg += $ArchivedFilesByYearOutput | ConvertTo-Html -Fragment
$HtmlMsg += "<h2>Summary of Archived Files by File Type</h2>"
$HtmlMsg += $ArchivedFilesByTypeOutput | ConvertTo-Html -Fragment
$HtmlMsg += "<p>Total archived files found: $($Global:ReportData.Count)</p>"
$HtmlMsg += "<p>Total size of archived files found: $(FormatFileSize $TotalArchivedFilesSize)</p>"
$HtmlMsg += "<p>Estimated storage savings from archived files: $StorageSavings</p>"
$HtmlMsg += "<p>Report generated on:<b> $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')</b></p>"
$HtmlMsg += "<p>Processing completed in: $($ProcessingDuration.ToString('hh\:mm\:ss'))</p>"
$HtmlMsg += "</body></html>"

# Construct the message body 	
$MsgBody = @{}
$MsgBody.Add('Content', "$($HtmlMsg)")
$MsgBody.Add('ContentType','html')
# Build the parameters to submit the message
$Message = @{}
$Message.Add('subject', $MsgSubject)
$Message.Add('toRecipients', $MsgTo)
$Message.Add('body', $MsgBody)
$Message.Add("attachments", $MsgAttachments)

$EmailParameters = @{}
$EmailParameters.Add('message', $Message)
$EmailParameters.Add('saveToSentItems', $true)
$EmailParameters.Add('isDeliveryReceiptRequested', $true)

# Send the message
Try {
    Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters -ErrorAction Stop
    Write-Output ("File-Level Archiving report emailed to {0}" -f $ToRecipient.emailAddress.address)
    Write-Output "All done!"
} Catch {
    Write-Output "Unable to send email"
    Write-Output $_.Exception.Message
}

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 