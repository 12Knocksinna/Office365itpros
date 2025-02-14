# Upload-LargeAttachmentChunks.PS1
# An example to show how to upload a large attachment in byte chunks for a draft message and then send the message

# V1.0 14-Feb-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Update-LargeAttachmentChunks.PS1

Connect-MgGraph -Scopes "Mail.Send", "Mail.ReadWrite" -NoWelcome

# Custiomize this address!
$RecipientAddress = 'Hans.Geering@Office365ITPros.com'
# Customize this path to a large file that exists on your system
$AttachmentFile = "C:\Temp\PowerShellBook.pdf"
[int32]$MaxAttachmentSize = 146800640
$FileStream = New-Object System.IO.StreamReader($AttachmentFile)  
$FileSize = $FileStream.BaseStream.Length 
# The maxiumum size of an attachment is 150 MB (157286400 bytes). There's some overhead, so we'll restrict this to 140 MB (146800640 bytes).
If ($FileSize -gt $MaxAttachmentSize) {
    Write-Host ("Attachment {0} is too large ({1} bytes). Maximum size is {2} bytes (140 MB)" -f $AttachmentFile, $FileSize, $MaxAttachmentSize)
    Break
}
# Build structure for the attachment
$AttachmentDetail = @{}
$AttachmentDetail.Add("attachmentType", "file")
$AttachmentDetail.Add("name", [System.IO.Path]::GetFileName($AttachmentFile))
$AttachmentDetail.Add("size", $FileSize)
$AttachmentParams = @{}
$AttachmentParams.Add("AttachmentItem", $AttachmentDetail)

# Create message structure
$HtmlBody = "<b>Isn't it nice to get a large attachment?<b>"
$MsgSubject = "A very large attachment for you to read"
$MsgFrom = (Get-MgContext).Account
$MsgParams = @{}
$MsgParams.Add("Content", $HtmlBody)
$MsgParams.Add("ContentType", "html")
$EmailAddress  = @{address = $RecipientAddress} 
$EmailRecipient = @{EmailAddress = $EmailAddress}  
Write-Host "Sending welcome email to Hans Geering"

# Create a draft message in the signed-in user's mailbox
$NewMessage = New-MgUserMessage -UserId $MsgFrom -Body $MsgParams -ToRecipients $EmailRecipient -Subject $MsgSubject 

# Create an upload session
$UploadSession = New-MgUserMessageAttachmentUploadSession -UserId $MsgFrom -MessageId $NewMessage.Id -BodyParameter $AttachmentParams

# Define chunk size for uploading attachments. Must be a multiple of 320 KB ( 327680 bytes)
[Int32]$uploadChunkSize = 983040

# Upload the attachment file in chunks
$FileOffsetStart = 0              
$FileBuffer = [byte[]]::new($uploadChunkSize)
do {            
    $FileChunkByteCount = $FileStream.BaseStream.Read($FileBuffer, 0, $FileBuffer.Length) 
    Write-Verbose ($FileStream.BaseStream.Position)
    $FileOffsetEnd = $FileStream.BaseStream.Position - 1
    if ($FileChunkByteCount -gt 0) {
        $UploadRangeHeader = "bytes " + $FileOffsetStart + "-" + $FileOffsetEnd + "/" + $FileSize
        Write-Verbose $UploadRangeHeader                
        $FileOffsetStart = $fileStream.BaseStream.Position
        $BinaryContent = New-Object System.Net.Http.ByteArrayContent -ArgumentList @($FileBuffer, 0, $FileChunkByteCount)
        $FileBuffer = [byte[]]::new($uploadChunkSize)
        $Headers = @{
            'AnchorMailbox' = $MsgFrom
            'Content-Range' = $UploadRangeHeader
        }
        $Result = (Invoke-RestMethod -Method Put -Uri $UploadSession.UploadUrl -UserAgent "UploadAgent" -Headers $Headers -Body $BinaryContent.ReadAsByteArrayAsync().Result -ContentType "application/octet-stream") 
        Write-Verbose $Result 
    }          
} while ($FileChunkByteCount -ne 0)      

# Send the message
Try {
    Send-MgUserMessage -UserId $MsgFrom -MessageId $NewMessage.Id -ErrorAction Stop
    Write-Host "Message sent to $RecipientAddress"
}
Catch {
    Write-Host "Error sending message: $_"
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.