# Find-JunkEmailDomains.PS1
# Find the domains that send our tenant junk email by examining the items in the Junk Email folder of each mailbox
# 1.0 11-Oct-2025

# GitHub link:
	

$Thumbprint = '0CF6CE3F3548FD73E7AC8CF20226ED447E125C71'
$TenantId =  'b662313f-14fc-43a2-9a7a-d2e27f4f3478'
$AppId = '9802440a-2c48-4e47-9eb8-f166ba99b11f'

Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $Thumbprint -NoWelcome
Connect-ExchangeOnline -ShowBanner:$false

# Get mailboxes - users and shared
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox, SharedMailbox -ResultSize Unlimited
If (!($Mbx)) {
    Write-Host "No user or shared mailboxes found" -ForegroundColor Red
    Break
} Else {
    Write-Host ("{0} user and shared mailboxes found" -f $Mbx.Count) -ForegroundColor Green
}   

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {
    Write-Host ("Processing mailbox {0}" -f $M.DisplayName)
    $JunkEmailFolder = Get-MgUserMailFolder -UserId $M.ExternalDirectoryObjectId -MailFolderId "junkemail"
    If ($null -eq $JunkEmailFolder) {
        Write-Host ("Failed to get Junk Email folder for {0}" -f $M.DisplayName) -ForegroundColor Red
        Continue
    }

    [array]$MailItems = Get-MgUserMailFolderMessage -UserId $M.ExternalDirectoryObjectId -MailFolderId $JunkEmailFolder.Id -All -PageSize 500 `
        -Property SentDateTime, Sender, Subject
   
    If (!($MailItems)) {
        Write-Host ("No items found in the Junk Email folder for {0}" -f $M.DisplayName) -ForegroundColor Yellow
        Continue
    } Else {    
        ForEach ($MailItem in $MailItems) {
        # Report each item found in the Junk Email folder
            $ReportItem = [PSCustomObject]@{
                DisplayName             = $M.DisplayName
                Id                      = $M.ExternalDirectoryObjectId
                UserPrincipalName       = $M.UserPrincipalName
                MailboxType             = $M.RecipientTypeDetails
                'Junk Mail Items'       = $MailItems.Count
                'Junk Mail Date'        = Get-Date $MailItem.SentDateTime -format 'dd-MMM-yyyy HH:mm'
                'Junk Mail Sender'      = $MailItem.Sender.emailAddress.Name
                'Junk Mail Email'       = $MailItem.Sender.emailAddress.Address
                'Junk Mail Subject'     = $MailItem.Subject
                'Junk Mail Domain'      = ($MailItem.Sender.emailAddress.Address -split '@')[1]
            }
            $Report.Add($ReportItem)
        }
    }
}


# Define the CSV file path
$csvName = [System.IO.Path]::GetFileName($csvPath)
$csvBytes = [System.IO.File]::ReadAllBytes($csvPath)
$csvBase64 = [System.Convert]::ToBase64String($csvBytes)

# Create the email message with CSV attachment
$emailMessage = @{
Message = @{
Subject = “Mailbox Size Report”
Body = @{
ContentType = “Text”
Content = “Attached is the mailbox size report in CSV format.”
}
ToRecipients = @(
@{
EmailAddress = @{
Address = “”
}
}
)
Attachments = @(
@{
“@odata.type” = “#microsoft.graph.fileAttachment”
Name = $csvName
ContentBytes = $csvBase64
ContentType = “text/csv”
}
)
}
}
Send-MgUserMail -UserId “” -BodyParameter $emailMessage