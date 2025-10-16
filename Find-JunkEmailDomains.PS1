# Find-JunkEmailDomains.PS1
# Find the domains that send our tenant junk email by examining the items in the Junk Email folder of each mailbox. The script
# works by extracting the domain from each item found in user and shared mailboxes. You can then use the output to create
# a report of the domains that send the most junk email to your tenant.
# V1.0 11-Oct-2025
# V1.16-Oct-2025 Added option to delete items found in the Junk Email folder after processing

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-JunkEmailDomains.PS1

param(
    # If set to $true, delete all items found in the Junk Email folder after processing
    [Parameter(Mandatory)]
    [bool]$DeleteItemsNow = $false
)

Function Get-PublicSuffix {
    [CmdletBinding()]
    param(
        # Domain name (e.g., "11.pen-and-sword.uk", "www.bbc.co.uk")
        [Parameter(Mandatory, ValueFromPipeline)]
        [string]$Domain,

        # Return registrable domain (eTLD+1) instead of just the public suffix
        [switch]$Registrable,

        # Force re-download of the Public Suffix List
        [switch]$RefreshSuffixList
    )

    begin {
        # Cache PSL in script scope for the session
        if (-not $script:__PSL -or $RefreshSuffixList) {
            $uri = 'https://publicsuffix.org/list/public_suffix_list.dat'
            $raw = Invoke-WebRequest -Uri $uri -UseBasicParsing | Select-Object -ExpandProperty Content -ErrorAction Stop
            $rules = $raw -split "`n" | ForEach-Object {
                $t = $_.Trim()
                if ($t -and -not $t.StartsWith('//')) { $t }
            }
            $script:__PSL = [PSCustomObject]@{
                Exceptions = $rules | Where-Object { $_.StartsWith('!') } | ForEach-Object { $_.Substring(1) }
                Normals    = $rules | Where-Object { -not $_.StartsWith('!') }
            }
        }
        $idn = [System.Globalization.IdnMapping]::new()
    }

    process {
        if ([string]::IsNullOrWhiteSpace($Domain)) { return $null }

        # Normalize domain
        $trimmed = $Domain.Trim().TrimEnd('.').ToLowerInvariant()
        if ($trimmed -match '^(?:\d{1,3}\.){3}\d{1,3}$') { return $null } # IP address check

        $labels = ($trimmed -split '\.') | ForEach-Object { try { $idn.GetAscii($_) } catch { $_ } }
        if ($labels.Count -eq 0) { return $null }

        # Helper: rule match
        function Test-RuleMatch($ruleLabels, $domainLabels) {
            $ri = $ruleLabels.Count - 1
            $di = $domainLabels.Count - 1
            while ($ri -ge 0 -and $di -ge 0) {
                $rule = $ruleLabels[$ri]
                $dom  = $domainLabels[$di]
                if ($rule -eq '*') { $ri--; $di--; continue }
                if ($rule -ieq $dom) { $ri--; $di--; continue }
                return $false
            }
            return ($ri -lt 0)
        }

        # PSL matching
        $bestLen = 1
        $matchedByExcept = $false

        foreach ($ex in $script:__PSL.Exceptions) {
            $exLabels = $ex -split '\.'
            if (Test-RuleMatch $exLabels $labels) {
                $len = [Math]::Max($exLabels.Count - 1, 1)
                if ($len -gt $bestLen) { $bestLen = $len; $matchedByExcept = $true }
            }
        }

        if (-not $matchedByExcept) {
            foreach ($nr in $script:__PSL.Normals) {
                $nrLabels = $nr -split '\.'
                if (Test-RuleMatch $nrLabels $labels) {
                    $len = $nrLabels.Count
                    if ($len -gt $bestLen) { $bestLen = $len }
                }
            }
        }

        $suffix = ($labels[($labels.Count - $bestLen)..($labels.Count - 1)] -join '.')
        if ($Registrable) {
            if ($labels.Count -gt $bestLen) {
                return ($labels[($labels.Count - ($bestLen + 1))..($labels.Count - 1)] -join '.')
            } else {
                return $null
            }
        } else {
            return $suffix
        }
    }
}
	
# The script can only run in app-only mode, so define the settings to connect interactively using a certificate. If using an app, make sure
# that the values for the appid, tenantid, and certificate thumbprint variables below match your app registration.
# Alternatively, you can run this code in Azure Automation and use a managed identity to authenticate
$Thumbprint = '0CF6CE3F3548FD73E7AC8CF20226ED447E125C71'
$TenantId =  'a662313f-14fc-43a2-9a7a-d2e27f4f3478'
$AppId = '9802440a-2c48-4e47-9eb8-f166ba99b11f'

# For interactive use, the signed-in user must be an Exchange administrator
Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $Thumbprint -NoWelcome
Connect-ExchangeOnline -ShowBanner:$false

# Change these variables to select the mailbox the message will come from and the destination SMTP address
$DestinationEmailAddress = "SomeAdminUser@office365itpros.com"
$MsgFrom = 'noreply@office365itpros.com'

# Get mailboxes - users and shared
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox, SharedMailbox -ResultSize Unlimited | Sort-Object DisplayName
If (!($Mbx)) {
    Write-Host "No user or shared mailboxes found" -ForegroundColor Red
    Break
} Else {
    Write-Host ("{0} user and shared mailboxes found" -f $Mbx.Count) -ForegroundColor Green
}   

# Find the domains for junk email senders
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
    } Else  {    
        Write-Host ("{0} items found in the Junk Email folder for {1}" -f $MailItems.Count, $M.DisplayName) -ForegroundColor Green
        ForEach ($MailItem in $MailItems) {
        $EmailDomain = ($MailItem.Sender.emailAddress.Address -split '@')[1]
        # Extract the root domain unless it's a .onmicrosoft.com service domain
        If ($EmailDomain -Notlike "*.onmicrosoft.com") {
            $EmailDomain = Get-PublicSuffix -Domain $EmailDomain -Registrable 
        }
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
                'Junk Mail Domain'      = $EmailDomain
            }
            $Report.Add($ReportItem)
        }
        If ($DeleteItemsNow) {
            Write-Host ("Deleting {0} spammy items from the Junk Email folder for {1}" -f $MailItems.Count, $M.DisplayName) -ForegroundColor Yellow
            ForEach ($MailItem in $MailItems) {
                Try {
                    Remove-MgUserMailFolderMessage -UserId $M.ExternalDirectoryObjectId -MailFolderId $JunkEmailFolder.Id -MessageId $MailItem.Id -ErrorAction Stop
                } Catch {
                    Write-Host ("Failed to delete item {0} from the Junk Email folder for {1}" -f $MailItem.Id, $M.DisplayName) -ForegroundColor Red
                }
            }
        }
    }
}

# Extract the set of domains
[array]$JunkEmailDomains = $Report | Group-Object 'Junk Mail Domain' -NoElement | Select-Object -ExpandProperty Name
# Remove consumer domains like gmail.com, outlook.com, yahoo.com, etc.
[array]$ConsumerDomains = @('gmail.com','outlook.com','yahoo.com','hotmail.com','live.com','aol.com','icloud.com','protonmail.com','zoho.com','gmx.com','msn.com')
$JunkEmailDomains = $JunkEmailDomains | Where-Object {$_ -and ($_ -notin $ConsumerDomains)} | Sort-Object -Unique

Write-Host "Generating report..."
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Junk Email Domains.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    }
    $JunkEmailDomains | Export-Excel -Path $ExcelOutputFile -WorksheetName "JunkEmailDomains" -Title ("Junk Email Domains {0}" -f (Get-Date -format 'dd-MMM-yyyy')) `
        -TitleBold -TableName "JunkEmailDomains"
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Junk Email Domains.CSV"
    $JunkEmailDomains | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}
 
If ($ExcelGenerated) {
    Write-Host ("An Excel report of Junk Email Domains is available in {0}" -f $ExcelOutputFile)
    $OutputFile = $ExcelOutputFile
} Else {    
    Write-Host ("A CSV report of Junk Email Domains is available in {0}" -f $CSVOutputFile)
    $OutputFile = $CSVOutputFile
}

# Encode the output file so that it can be added as an attachment to an email
$EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($OutputFile))
# Create the attachments array
$MsgAttachments = @(
    @{
        '@odata.type' = '#microsoft.graph.fileAttachment'
        Name = (Split-Path $OutputFile -Leaf)
        ContentBytes = $EncodedAttachmentFile
        ContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
)

# Create a string of domains to block for inclusion in the email body
[string]$JunkEmailDomainsOutput = $JunkEmailDomains -join "', '"

$TransportRuleName = "'Quarantine Traffic from Junk Email Domains'"

# Define the message recipient (see earlier)
$ToRecipient = @{}
$ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
[array]$MsgTo = $ToRecipient
# Define the message subject
$MsgSubject = "Important: Junk Email Domains Report"
# Create the HTML content
$HtmlMsg = "</body></html><p>The output file for the <b>Junk Email Domains Report</b> are attached to this message. Please review the information at your convenience and consider creating a transport rule to block these domains</p>"
$HtmlMsg = $HtmlMsg + "<p>You can use PowerShell commands like this to create a transport rule to block these domains:<p></p>"
$HtmlMsg = $HtmlMsg + "<p>New-TransportRule -Name $TransportRuleName -SenderDomainIs '$JunkEmailDomainsOutput' -Mode Enforce  -Quarantine 1 -SenderAddressLocation HeaderOrEnvelope -Comments 'Blocks messages from domains that send junk email'</p>"

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
    Write-Output ("Junk Email Domains report emailed to {0}" -f $ToRecipient.emailAddress.address)
} Catch {
    Write-Output "Unable to send email"
    Write-Output $_.Exception.Message
}

Write-Output "All done"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
