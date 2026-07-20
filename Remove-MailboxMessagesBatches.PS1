# Remove-MailboxMessagesBatches.PS1
# An example of how to use JSON batching to permanently remove (send to the Purges folder) mailbox items.

# V1.0 20-Jul-2026
# GitHub link:

# https://github.com/12Knocksinna/Office365itpros/blob/master/Remove-UtilityMessagesFromMailboxes.PS1 is the original version of a script
# to remove utility messages from mailboxes. This script is a modified version of that script to illustrate how to use JSON batching to remove mailbox items.


# Connect to the Microsoft Graph
$AppThumbprint = "F79286DB88C21491110109A0222348FACF694CBD"
$AppId = "d58578ac-7cb4-4b5a-a296-f19218a03f11"
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3478"

# The app needs the following application permissions:
# User.Read.All: Read user account details
# Mail.ReadWrite: Read and write mail in all mailboxes
# Mail.Send: Send mail as a user

Connect-MgGraph -NoWelcome -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $AppThumbprint

# List the email addresses of the notification messages to search for and remove
$UM0 = "notifications@communityhub.microsoft.com"
$UM1 = "messages-noreply@linkedin.com"
$UM2 = "noreply@microsoft.com"
$UM3 = "notifications-noreply@linkedin.com"
$UM4 = "noreply@yammer.com"
$UM5 = "microsoft@powerapps.com"
$UM6 = "o365mc@microsoft.com"
$UM7 = "jobs-listings@linkedin.com"
$UM8 = "postmaster@online.zdnet.com"
$UM9 = "noreply@amazon.com"
$UM10 = "notify@mstechcommunity.microsoft.com"
$UM11 = "irishtimeseveningupdate@comms.irishtimes.com"

# Set the start and end dates for the search - starting five years ago and ending 14 days ago
[string]$StartDate = (Get-Date).AddDays(-1825).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
[string]$EndDate = (Get-Date).AddDays(-14).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

# Define the service plan IDs for Exchange Online (Plan 1) and Exchange Online (Plan 2)
$ExoServicePlan1 = "9aaf7827-d63c-4b61-89c3-182f06f82e5c"
$ExoServicePlan2 = "efb87545-963c-4e0d-99df-69c6916d9eb0" 

# Find users assigned a license that includes the Exchange Online (Plan 1) or Exchange Online (Plan 2) service plans.
# The check also looks for users with the Exchange Online service plan enabled.
Write-Host "Finding accounts with Exchange Online mailboxes..." -ForegroundColor Cyan
[array]$Users = Get-MgUser -Filter "assignedPlans/any(c:c/servicePlanId eq $ExoServicePlan1 and capabilityStatus eq 'Enabled') `
  or assignedPlans/any(c:c/servicePlanId eq $ExoServicePlan2 and capabilityStatus eq 'Enabled')" `
    -ConsistencyLevel eventual -CountVariable Test -All -PageSize 999 -Sort ('displayname') `
    -Property Id, displayName, userprincipalName, assignedLicenses, assignedPlans, department, country, CreatedDateTime

If ($Users) {
    Write-Output ("{0} users licensed for Exchange Online found" -f $Users.Count)
} Else {
    Write-Output "No mailboxes found"
    Break
}

$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($User in $Users) {
    Write-Host ("Processing mailbox {0}" -f $User.DisplayName) -ForegroundColor Yellow

    [array]$Messages = Get-MgUserMessage -UserId $User.Id `
    -Filter "(ReceivedDateTime ge $StartDate and ReceivedDateTime le $EndDate) `
        and (sender/emailAddress/address eq '$UM0' or sender/emailAddress/address eq '$UM1' or sender/emailAddress/address eq '$UM2' `
        or sender/emailAddress/address eq '$UM3' or sender/emailAddress/address eq '$UM4' or sender/emailAddress/address eq '$UM5' `
        or sender/emailAddress/address eq '$UM6' or sender/emailAddress/address eq '$UM7' or sender/emailAddress/address eq '$UM8' `
        or sender/emailAddress/address eq '$UM9' or sender/emailAddress/address eq '$UM10' or sender/emailAddress/address eq '$UM11')" `
        -Property Id, Subject, Sender, SentDateTime -All -PageSize 999

    If ($Messages) {
        Write-Host ("Found {0} messages to delete for mailbox {1}" -f $Messages.Count, $User.DisplayName) -ForegroundColor Green
    } Else {
        Write-Host ("No messages found to delete for mailbox {0}" -f $User.DisplayName) -ForegroundColor Cyan
        Continue
    }
    $BatchSize = 20
    $Batches = [Math]::Ceiling($Messages.Count/$BatchSize)

    For ($i = 0; $i -lt $Batches; $i++) {

        $Start = $i * $BatchSize
        $End = [Math]::Min($Start + $BatchSize - 1, $Messages.Count - 1)
        $CurrentBatch = $Messages[$Start..$End] | Select-Object Id, Subject, Sender, SentDateTime

        # Build the batch request body - each entry needs a unique "id", the target method/url, and the body of the DELETE operation
        $BatchRequests = [System.Collections.Generic.List[object]]::new()
        $Counter = 0

        ForEach ($Message in $CurrentBatch) {
            $Counter++
            $URI = "/users/{0}/Messages/{1}/permanentDelete" -f $User.Id, $Message.Id
            $BatchRequest = [PSCustomObject][Ordered]@{
                id      = "$counter"
                method  = "POST"
                url     = $Uri
                headers = @{ "Content-Type" = "application/json" }
            }
            $BatchRequests.Add($BatchRequest)
        }

        $BatchBody = @{ Requests = $batchRequests }

        Try {
            $Response = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/`$batch" -Body ($batchBody | ConvertTo-Json -Depth 10) -ErrorAction Stop
            # Check each individual response inside the batch for errors
            ForEach ($Result in $Response.responses) {
                $FailedStatus = $false
                If ($Result.status -ge 400) {
                    $FailedMessage = $currentBatch[[int]$result.id - 1]
                    $FailedStatus = $true
                    Write-Warning "Failed to delete $($FailedMessage.Subject): Status $($Result.status) - $($Result.body.error.message)"
                }
                If ($FailedStatus) {
                    $MessageStatus = "Message deletion failed"
                } Else {
                    $MessageStatus = "Message deleted"
                }
                $ReportLine = [PSCustomObject][Ordered]@{  
                    Timestamp   = (Get-Date -format s)  
                    Action      = $MessageStatus
                    Mailbox     = $User.DisplayName
                    SMTPAddress = $User.UserPrincipalName
                    Subject     = $Message.Subject
                    SenderEmail = $Message.Sender.EmailAddress.Address
                    SenderName  = $Message.Sender.EmailAddress.Name
                    Sent        = $Message.SentDateTime
                }
                $Report.Add($ReportLine)  
            }
            Write-Host "Batch $($i + 1) of $batches completed"
        } Catch {
            Write-Error "Batch $($i + 1) failed entirely: $_"
        }

        # Small pause to avoid hammering the throttling limits between batches
        Start-Sleep -Milliseconds 200
    }
}

Write-Host ("Message removal complete. {0} messages removed from {1} unique mailboxes" -f ($Report | Where-Object { $_.Action -eq 'Message Deleted' }).Count, ($Report | Select-Object -ExpandProperty Mailbox | Sort-Object -Unique).Count)

# Generate reports in Excel or CSV format depending on whether the ImportExcel module is available
If (Get-Module ImportExcel -ListAvailable) {
  $ExcelGenerated = $True
  Import-Module ImportExcel -ErrorAction SilentlyContinue
  $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Mailbox Deletions.xlsx"
  Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
  $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Mailbox Deletions" -Title ("Mailbox Deletions {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "MailboxDeletions"
  
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Mailbox Deletions.CSV"
    Remove-Item $CSVOutputFile -ErrorAction SilentlyContinue
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8 
}

Write-Host ("Report saved to {0}" -f (If ($ExcelGenerated) { $ExcelOutputFile } Else { $CSVOutputFile })) -ForegroundColor Green


# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.