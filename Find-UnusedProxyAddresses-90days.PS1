# Find-UnusedProxyAddresses-90days.PS1


# A script to check user and shared mailboxes to discover if any of the SMTP proxy addresses assigned to the mailbox    
# have not been used to send or receive email in the last 10 days. The script uses historical message trace data downloaded from Exchange Online.
# 
# V1.0         24 September 2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-UnusedProxyAddresses-90days.PS1

# The first part of this script is code taken from Analyze-MTHistoricalLogs.PS1 to read the historical message trace data downloaded from Exchange Online.
# The second part determines which proxy addresses assigned to user and shared mailboxes have not been used to send or receive email in the last 90 days.

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If (!($Modules -contains "ExchangeOnlineManagement")) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    # Get accepted domains
    [array]$Domains = Get-AcceptedDomain | Select-Object -ExpandProperty DomainName
}

# Folder where the historical message tracking logs downloaded from Exchange Online are stored. Change this to the folder you use.
$DataFolder = "c:\temp\MtData\"
# Find the set of downloaded historical message trace data files to process
[array]$DataFiles = Get-ChildItem -Path $DataFolder | Select-Object -ExpandProperty Name

If (!($DataFiles)) {
    Write-Host "No historical message tracking logs to analyze - exiting"
    Break
}

Write-Host ("Preparing to process {0} historical message trace data files..." -f $DataFiles.count)
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report
# Define a set of bad outcomes from message trace data entries that we don't need to report on. These include
# messages intercepted as spam, quarantined messages, and failed messages.
[array]$BadOutcomes = "Receive, Fail", "Receive, Deliver, Quarantined", "Receive, Deliver, FilteredAsSpam"

Write-Host "Analyzing historical message trace data..."
ForEach ($File in $DataFiles) {
    $MtDataFile = $DataFolder + $File
    [array]$MtData = Import-CSV -Path $MtDataFile -Encoding unicode
        ForEach ($Line in $MtData) {
            If (!([string]::IsNullOrEmpty($Line.origin_timestamp_utc))){
                [array]$RecipientStatus = $Line.Recipient_Status.split(";")
                # array of individual recipients for a message
                $RecipientInfo = [System.Collections.Generic.List[Object]]::new()
                ForEach ($RecipientDetail in $RecipientStatus) {
                    $Recipient = $RecipientDetail.Split("##")[0]
                    $RecipientOutcome = $RecipientDetail.Split("##")[1]
                    $RecipientLine = [PSCustomObject]@{ 
                        Recipient = $Recipient
                        Outcome   = $RecipientOutcome
                    }
                    $RecipientInfo.Add($RecipientLine)
                }    
                $SenderDomain  = $Line.Sender_address.Split("@")[1]
                If ($SenderDomain -in $Domains) {
                    $Direction = "Originating" 
                } Else {
                    $Direction = "Incoming"
                }
                # Only report on messages with a good outcome
                If (!($RecipientOutcome -in $BadOutcomes)) {
                    $ReportLine = [PSCustomObject]@{ 
                        Timestamp        = $Line.origin_timestamp_utc
                        Sender           = $Line.sender_address
                        Subject          = $Line.message_subject
                        Recipient        = $RecipientInfo.Recipient
                        RecipientDomain  = $RecipientInfo.Recipient.Split("@")[1]
                        RecipientStatus  = $Line.Recipient_Status
                        RecipientInfo    = $RecipientInfo
                        Outcome          = $RecipientOutcome 
                        Bytes            = $Line.total_bytes
                        Message_id       = $Line.message_id
                        Sender_Domain    = $SenderDomain
                        Client_IP        = $Line.original_client_ip
                        Direction        = $Direction
                    }
                      $Report.Add($ReportLine) 
                }  
            }
        }
}

$Report = $Report | Sort-Object Sender, @{Expression = { $_.Timestamp -as [datetime] }; Descending = $true}

# Create an array of the latest messages sent or received by each email address in the tenant
[array]$SenderAddresses = $Report | Where-Object {$_.Sender_domain -in $Domains} | Sort-Object {$_.Timestamp -as [datetime]} -Descending | Group-Object Sender |  ForEach-Object { $_.Group | Select-Object -First 1 }
[array]$RecipientAddresses = $Report | Where-Object {$_.RecipientDomain -in $Domains} | Sort-Object {$_.Timestamp -as [datetime]} -Descending | Group-Object Recipient |  ForEach-Object { $_.Group | Select-Object -First 1 }

Write-Host ("{0} unique sender addresses found in the last 90 days" -f $SenderAddresses.count)
Write-Host ("{0} unique recipient addresses found in the last 90 days" -f $RecipientAddresses.count)
Write-Host "Now checking for unused proxy addresses assigned to mailboxes..."
[array]$Mailboxes = Get-ExoMailbox -ResultSize Unlimited -Properties EmailAddresses -RecipientTypeDetails UserMailbox, SharedMailbox | Sort-Object DisplayName

$UnusedProxyReport = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Mbx in $Mailboxes) {
    $MbxAddresses = $Mbx.EmailAddresses | Where-Object {$_ -like "smtp:*"} 
    ForEach ($Address in $MbxAddresses) {
        $SmtpAddress = $Address.Substring(5)  # Remove the smtp: prefix
        If ($Address.SubString(0,4) -ceq "SMTP") {
            $AddressType = "Primary"
        } Else {
            $AddressType = "Proxy" 
        }
        If ($SmtpAddress -notin $SenderAddresses.Sender -and $SmtpAddress -notlike "*.onmicrosoft.com") {
            $ReportLine = [PSCustomObject]@{
                DisplayName         = $Mbx.DisplayName
                PrimarySmtpAddress  = $Mbx.PrimarySmtpAddress
                Alias               = $Mbx.Alias
                MailboxType         = $Mbx.RecipientTypeDetails
                UnusedProxyAddress  = $SmtpAddress
                AddressType         = $AddressType
                CheckType           = "Sender"
            }
            $UnusedProxyReport.Add($ReportLine)
        }
        If ($SmtpAddress -notin $RecipientAddresses.Recipient  -and $SmtpAddress -notlike "*.onmicrosoft.com") {
            $ReportLine = [PSCustomObject]@{
                DisplayName         = $Mbx.DisplayName
                PrimarySmtpAddress  = $Mbx.PrimarySmtpAddress
                Alias               = $Mbx.Alias
                MailboxType         = $Mbx.RecipientTypeDetails
                UnusedProxyAddress  = $SmtpAddress
                AddressType         = $AddressType
                CheckType           = "Recipient"
            }
            $UnusedProxyReport.Add($ReportLine)
        }
    }
}

# Show what we've found
$UnusedProxyReport | Select-Object DisplayName, UnusedProxyAddress, AddressType, CheckType, MailboxType, PrimarySmtpAddress | Out-GridView -Title "Unused Proxy Addresses"

# And generate an output file
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\UnusedProxyAddresses.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    } 
    $UnusedProxyReport | Export-Excel -Path $ExcelOutputFile -WorksheetName "Unused Proxy Addresses" -Title ("Unused Proxy Addresses Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "UnusedProxyAddresses" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\UnusedProxyAddresses.CSV"
    $UnusedProxyReport | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}

If ($ExcelGenerated) {
    Write-Host ("Excel worksheet output written to {0}" -f $ExcelOutputFile)
} Else {
    Write-Host ("CSV output file written to {0}" -f $CSVOutputFile)
}   

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.