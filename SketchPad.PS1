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

#---------------------- Example of uploading a large attachment to Outlook ==================

$TenantName = (Get-MgOrganization).DisplayName
# Define some variables used to construct the HTML content in the message body
#HTML header with styles
$htmlhead="<html><style>BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    </style>"

#Content for the message - obviously this is very customizable to reflect what you want to say to new users
$HtmlBody = $HtmlHead + "<body><h2>New User " + $User.DisplayName + "</h2>"
$HTMLBody = $HtmlBody + "<h1>Welcome to $($TenantName)</h1>
    <p><strong>Generated:</strong> $(Get-Date -Format g)</p>  
    <h2><u>We're Pleased to Have You Here</u></h2>
    <p><b>Welcome to your new Office 365 account</b></p>
    <p>You can open your account to access your email and documents by clicking <a href=http://www.portal.office.com>here</a> </p>
    <p>Have a great time and be sure to call the help desk if you need assistance. And be sure to read all the great articles about Office 365 published on Practical365.com.</p>"

# Define chunk size for uploading attachments. Must be a multiple of 320 KB ( 327680 bytes)
[Int32]$uploadChunkSize = 983040
# Define path to attachment file
$AttachmentFile = "C:\Temp\PowerShellBook.pdf"

$FileStream = New-Object System.IO.StreamReader($AttachmentFile)  
$FileSize = $FileStream.BaseStream.Length 

# Build structure for the attachment
$AttachmentDetail = @{}
$AttachmentDetail.Add("attachmentType", "file")
$AttachmentDetail.Add("name", [System.IO.Path]::GetFileName($AttachmentFile))
$AttachmentDetail.Add("size", $FileSize)

$AttachmentParams = @{}
$AttachmentParams.Add("AttachmentItem", $AttachmentDetail)

$MsgSubject = "A warm welcome to $($TenantName)"
$MsgFrom = (Get-MgContext).Account

# Create message structure
$MsgParams = @{}
$MsgParams.Add("Content", $HtmlBody)
$MsgParams.Add("ContentType", "html")

# Add the recipient using the mailbox's primary SMTP address
$EmailAddress  = @{address = $User.Mail} 
$EmailRecipient = @{EmailAddress = $EmailAddress}  
Write-Host "Sending welcome email to" $User.DisplayName

# Create a draft message in the signed-in user's mailbox
$NewMessage = New-MgUserMessage -UserId $MsgFrom -Body $MsgParams -ToRecipients $EmailRecipient -Subject $MsgSubject 

# Create an upload session
$UploadSession = New-MgUserMessageAttachmentUploadSession -UserId $MsgFrom -MessageId $NewMessage.Id -BodyParameter $AttachmentParams

# Upload the attachment file in chunks
$FileOffsetStart = 0              
$FileBuffer = [byte[]]::new($uploadChunkSize)
do {            
    $FileChunkByteCount = $fileStream.BaseStream.Read($FileBuffer, 0, $FileBuffer.Length) 
    Write-Verbose ($fileStream.BaseStream.Position)
    $FileOffsetEnd = $fileStream.BaseStream.Position - 1
    if ($FileChunkByteCount -gt 0) {
        $UploadRangeHeader = "bytes " + $FileOffsetStart + "-" + $FileOffsetEnd + "/" + $FileSize
        Write-Verbose $UploadRangeHeader                
        $FileOffsetStart = $fileStream.BaseStream.Position
        $binaryContent = New-Object System.Net.Http.ByteArrayContent -ArgumentList @($FileBuffer, 0, $FileChunkByteCount)
        $FileBuffer = [byte[]]::new($uploadChunkSize)
        $headers = @{
            'AnchorMailbox' = $MsgFrom
            'Content-Range' = $UploadRangeHeader
        }
        $Result = (Invoke-RestMethod -Method Put -Uri $UploadSession.UploadUrl -UserAgent "UploadAgent" -Headers $headers -Body $binaryContent.ReadAsByteArrayAsync().Result -ContentType "application/octet-stream") 
        Write-Verbose $Result 
}          

} while ($FileChunkByteCount -ne 0)      

# Send the message
Send-MgUserMessage -UserId $MsgFrom -MessageId $NewMessage.Id  


# +---------------------- Example of loading lots of attachments to a message -----------------

$AttachmentsFolder = "c:\Temp\Attachments"
[array]$InputAttachments = Get-ChildItem -Path $AttachmentsFolder
If (!($InputAttachments)) {
    Write-Host "No attachments found in $AttachmentsFolder"
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

$MsgSubject = "Lots of attachments"
$MsgFrom = (Get-MgContext).Account
$ToRecipients = @{}
$ToRecipients.Add("emailAddress",@{'address'='tony.redmond@office365itpros.com'})
[array]$MsgTo = $ToRecipients
$HtmlBody = "Some important reading for you.."
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

Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params