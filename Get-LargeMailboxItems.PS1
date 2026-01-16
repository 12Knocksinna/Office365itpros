<# 
.SYNOPSIS
Get-LargeMailboxItems.PS1
Script to find large mailbox items in the Inbox of a specified mailbox over a given size threshold and within a date

.PARAMETER Mailbox
Identifier for the mailbox (could be SMTP address, display name, UPN, or GUID).

.PARAMETER Threshold 
Size of the mailbox items to filter for. Pass in MB, KB, GB, or bytes

.PARAMETER Start
Start date/time in local time, e.g. 1-Jan-2026

.PARAMETER End
End date/time in local time, e.g. 16-Jan-2026

.EXAMPLE
.\Get-LargeMailboxItems.ps1 -Mailbox "Lotte.Vetler@office365itpros.com" -Threshold 10MB `
  -Start "1-Dec-2025" -End "1-Jan-2026" 

.NOTES
Requires an app configured to connect to the Microsoft Graph PowerShell SDK in app-only mode with permissions
Mail.Read to read mailbox items
MailboxFolder.Read.All to read mailbox folder names
User.Read.All to find mailboxes in the tenant.

Version 1.0 16  -Jan-2026
GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Get-LargeMailboxItems.PS1

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $Mailbox,

    [Parameter(Mandatory)]
    [int32] $Threshold,

    [Parameter(Mandatory)]
    [string] $Start,

    [Parameter(Mandatory)]
    [string] $End
)

# Change these values to match your app registration and your tenant!!!
$AppId = "d224fa41-4c2a-4380-b688-31625f90207a"
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f3478"
$CertThumbprint = "0CF6CE3F3548FD73E7AC8CF20226ED447E125C71"
Connect-MgGraph -App $AppId -TenantId $TenantId -CertificateThumbprint $CertThumbprint -NoWelcome

Try {
    $StartDate = (Get-Date $Start -ErrorAction Stop).ToString('s') + "Z" 
} Catch {
    Write-Host "$Start is not a valid date" -ForegroundColor Red
    Break
}
Try {
    $EndDate = (Get-Date $End -ErrorAction Stop).ToString('s') + "Z"
} Catch {
    Write-Host "$End is not a valid date" -ForegroundColor Red
    Break
}

# Attempt to find account for the mailbox to check that it exists
Try {
    $User = Get-MgUser -UserId $Mailbox -ErrorAction Stop
} Catch {
    Write-Host "Cannot find mailbox for $Mailbox" -ForegroundColor Red  
    Break
}

# Run server-side filter to find large items in the target folder within the date range
Try {
    [array]$LargeItems = Get-MgUserMailFolderMessage -UserId $User.Id -MailFolderId Inbox -All `
     -PageSize 500 -Property Sender,ReceivedDateTime,BodyPreview,Subject,ToRecipients `
     -Filter "(hasattachments eq true and receivedDateTime ge $startdate and receivedDateTime le $enddate) AND singleValueExtendedProperties/any(ep:ep/id eq 'Integer 0x0E08' and cast(ep/value, 'Edm.Int32') gt $Threshold)" -expand "singleValueExtendedProperties(`$filter=(id eq 'Integer 0x0e08'))" -ErrorAction Stop 
} Catch {
    Write-Host "Error retrieving mailbox items for $Mailbox. $_" -ForegroundColor Red  
    Break
}

If (!($LargeItems)) {
    Write-Host "No items larger than $Threshold bytes found in the Inbox of $Mailbox" -ForegroundColor Yellow
    Break
}

$Report = [System.Collections.Generic.List[Object]]::new()
# Client-side size check
$LargeItems | ForEach-Object {
    $Size = [int64]$_.SingleValueExtendedProperties[0].Value
    If ($Size -gt $Threshold) {
        $ReportLine = [PSCustomObject][Ordered]@{
            Subject          = $_.Subject
            From             = $_.Sender.EmailAddress.Address
            ReceivedDateTime = $_.ReceivedDateTime
            SizeMB           = [math]::Round($Size / 1MB, 2)
            Id               = $_.Id
        }
        $Report.Add($ReportLine)
    }
}
$Report = $Report | Sort-Object {$_.ReceivedDateTime -as [datetime]} -Descending

Write-Host ""
Write-Host ("{0} item(s) largher than the threshold of {1} bytes were found in {2}'s mailbox" -f $LargeItems.Count, $Threshold, $User.DisplayName) -ForegroundColor Cyan
Write-Host "------------------------------------------------------------------------------------------" -ForegroundColor Cyan

$Report | Select-Object Subject,From,ReceivedDateTime,SizeMB | Format-Table -AutoSize 
$Report | Sort-Object SizeMB -Descending | Out-GridView -Title "Large Emails in Inbox over $($Threshold / 1MB) MB"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.