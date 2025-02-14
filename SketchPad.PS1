# SketchPad.PS1
# Used for bits of PowerShell I am working on or want to note

# Get list to update metadata for the new item
$ListId = (Get-MgSiteList -SiteId $Site.Id -Filter "DisplayName eq 'Documents'").Id
[array]$ListItems = Get-MgSiteListItem -SiteId $Site.Id -ListId $ListId
$ListItem = $ListItems[-1]
$Body = @{}
$Body.Add("Title", "Hard Deleted Users Report Created by Azure Automation")
$Status = Update-MgSiteListItemField -SiteId $site.Id -ListId $listId -ListItemId $listItem.Id -BodyParameter $Body
If ($Status) {
    Write-Output ("Updated document metadata for item {0} with title {1}" -f $ListItem.Id, $Params.Title)
}




# Report all OneDrive accounts
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
    -ConsistencyLevel eventual -CountVariable UsersFound -All -PageSize 500
If (!$Users) {
    Write-Host "No user accounts found"
    Break
}
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    Try {
        $OneDrive = Get-MgUserDefaultDrive -UserId $User.Id -ErrorAction Stop
    } Catch {
        Write-Host ("Unable to find OneDrive for {0}" -f $User.UserPrincipalName)
        Continue
    }   
    $ReportLine = [PSCustomObject]@{
        UserPrincipalName = $User.UserPrincipalName
        OneDriveUrl       = $OneDrive.WebUrl
        Created           = Get-Date $OneDrive.CreatedDateTime -format 'dd-MMM-yyyy HH:mm'
        Modified          = Get-Date $OneDrive.LastModifiedDateTime -format 'dd-MMM-yyyy HH:mm'
    }
    $Report.Add($ReportLine)
}


# +---------------------- Example of loading lots of attachments to a message -----------------

$AttachmentsFolder = "c:\Temp\Attachments"
[array]$InputAttachments = Get-ChildItem -Path $AttachmentsFolder
If (!($InputAttachments)) {
    Write-Host "No attachments found in $AttachmentsFolder"
    Break
}   
$FileSizeThreshold = 146800640 # 140 MB in byes
$TotalFileSize = ($InputAttachments | Measure-Object -Sum Length).Sum
$FoundSizeMB = [math]::Round($TotalFileSize / 1MB, 2)
If ($TotalFileSize -gt $FileSizeThreshold) {
    Write-Host ("Total size of attachments is {1} MB. Maximum size for an Outlook message is 140 MB. Please remove some attachments and try again." -f $TotalFileSize, $FoundSizeMB)
    Break
}

[array]$MsgAttachments = $null
Write-Host ("Processing {0} attachments..." -f $InputAttachments.Count)
[array]$MsgAttachments = $null
ForEach ($File in $InputAttachments.Name) {
    $FullFileName = $AttachmentsFolder + "\" + $File
    $ConvertedContent = [Convert]::ToBase64String([IO.File]::ReadAllBytes($FullFileName))
    $FileExtension = [System.IO.Path]::GetExtension($FullFileName) 
    Switch ($FileExtension) {
        ".pdf" {
            $ContentType = "application/pdf"
        }
        ".docx" {
            $ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        }
        ".xlsx" {
            $ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }   
        ".pptx" {
            $ContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        }   
        ".jpg" {
            $ContentType = "image/jpeg"
        }   
        ".png" {
            $ContentType = "image/png"
        }   
        default {
            $ContentType = "application/octet-stream"
        }
        }
    $AttachmentDetails = @{
        "@odata.type" = "#microsoft.graph.fileAttachment"
        Name = $File
        ContentType =  $ContentType
        ContentBytes = $ConvertedContent
    }
    $MsgAttachments += $AttachmentDetails
}
Write-Host ("Sending a message with {0} attachments. Total size of attachments: {1} MB" -f $InputAttachments.Count, $FoundSizeMB)

$MsgSubject = "Lots of attachments"
$MsgFrom = (Get-MgContext).Account
$ToRecipients = @{}
$ToRecipients.Add("emailAddress",@{'address'='hans.geering@office365itpros.com'})
[array]$MsgTo = $ToRecipients
$HtmlBody = "Some important reading for you... Please note the large number of attachments and the size of some files!"
$MsgBody = @{}
$MsgBody.Add('Content', "$($HtmlBody)")
$MsgBody.Add('ContentType','html')

$Message = @{}
$Message.Add('subject', $MsgSubject)
$Message.Add('toRecipients', $MsgTo)
$Message.Add('body', $MsgBody)
$Message.Add('attachments', $MsgAttachments) 

$Params = @{}
$Params.Add('message', $Message)
$Params.Add('saveToSentItems', $true)
$Params.Add('isDeliveryReceiptRequested', $true)  

Try {
    Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params -ErrorAction Stop
    Write-Host "Mesaage sent"
} Catch {
    Write-Host ("Error sending message: {0}" -f $_.Exception.Message)
}
