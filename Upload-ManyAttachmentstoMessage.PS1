# Upload-ManyAttachmentstoMessage.PS1
# An example of uploading and sending many attachments with a mail message
# +---------------------- Example of loading lots of attachments to a message -----------------

# V1.0 14-Feb-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Upload-ManyAttachmentsToMessage.PS1

Connect-MgGraph -Scopes "Mail.Send", "Mail.ReadWrite" -NoWelcome

# Customize where the attachment files are stored
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
