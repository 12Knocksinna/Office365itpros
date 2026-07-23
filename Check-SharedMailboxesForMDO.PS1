# Check-SharedMailboxesForMDO.PS1
# A script to check the Entra ID accounts used for shared mailboxes to figure out if they should be licensed for Microsoft Defender for Office 365 (MDO).

# Basically, a shared mailbox needs an MDO license if it receives benefit of MDO protection. In other words, the shared mailbox receives email from
# external domains that don't belong to the tenant. If a shared mailbox only receives email from internal users, it doesn't need an MDO license.
# See https://office365itpros.com/2025/10/31/mdo-p2-licensing/ for more information.

# V1.0 13-Nov-2025
# V2.0 23-Jul-2026 Updated to fetch all message trace data for the last 90 days in 10-day chunks. This means that the check
#                  for external email traffic is more accurate.
#                  Also included a check for Teams meetings organized by shared mailboxes becaause they might need 
#                  a OneDrive license to retain meeting recordings. See https://office365itpros.com/2026/07/16/schedule-teams-meeting-shared-mbx/

# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Check-SharedMailboxesForMDO.PS1

# This ia an example of running a script in app-only mode for both the Microsoft Graph PowerShell SDK and Exchange Online management module.
# An X.509 certificate loaded into a registered app is used to authenticate for both connections. The app must have these permissions:

# Calendar.Read for Microsoft Graph to read the shared mailbox calendar to check for Teams meetings organized by the mailbox.
# Mail.Send to send the report email from a service account to the recipient of your choice.
# Domain.Read.All to read the list of verified domains for the tenant.

function Get-TeamsMeetings {
    # Function to check events in a shared mailbox calendar to detect if any meetings are Teams meetings organized by
    # the shared mailbox. If so, the shared mailbox might need a OneDrive license to retain meeting recordings.
    param (
        [string]$UserId
    )

    $StartDate = (Get-Date).AddDays(-90).toString('yyyy-MM-dd')
    $EndDate = (Get-Date).AddDays(7).toString('yyyy-MM-dd')
    $DefaultUserCalendar = Get-MgUserDefaultCalendar -UserId $UserId
    [array]$Events = Get-MgUserCalendarView -UserId $UserId -CalendarId $DefaultUserCalendar.Id -EndDateTime $EndDate -StartDateTime $StartDate

    [int]$TeamsMeetingsFound = 0

    ForEach ($Event in $Events) {
    # Check if the event is a meeting organized by the shared mailbox
        If ((($Event.organizer.emailaddress.address.trim() -eq $M.primarySMTPAddress.trim()) -and $Event.OnlineMeetingProvider -eq 'teamsForBusiness')) {
            $TeamsMeetingsFound++
        }
    }
    return $TeamsMeetingsFound
}
function Get-DateRangeData {
    param (
        [datetime]$StartDate,
        [datetime]$EndDate
    )
    
    # Function to fetch message trace data for expanded events over a 10-day period.
    [array]$Messages = $null
    Write-Host ("Collecting message trace data between {0} and {1}..." -f $StartDate, $EndDate)
    [int]$BatchSizeForMessages = 2000
    Try {
        # The warning action is suppressed here because we don't want to see warnings when more data is available
        [array]$MessagePage = Get-MessageTraceV2 -StartDate $StartDate -EndDate $EndDate `
        	-ResultSize $BatchSizeForMessages -Status "Delivered" -ErrorAction Stop -WarningAction SilentlyContinue
        $Messages += $MessagePage
    } Catch {
        Write-Host ("Error fetching message trace data: {0}" -f $_.Exception.Message)
        Break
    }
    If ($MessagePage.count -eq $BatchSizeForMessages) {
        Do {
            Write-Host ("Fetched {0} messages so far for this chunk..." -f $Messages.count)
            $LastMessageFetched = $MessagePage[-1]
            $LastMessageFetchedDate = $LastMessageFetched.Received.ToString("O")
            $LastMessageFetchedRecipient = $LastMessageFetched.RecipientAddress
            # Fetch the next page of messages
            [array]$MessagePage = Get-MessageTraceV2 -StartDate $StartDate -EndDate $LastMessageFetchedDate `
                -StartingRecipientAddress $LastMessageFetchedRecipient -ResultSize $BatchSizeForMessages -Status "Delivered" -ErrorAction Stop -WarningAction SilentlyContinue
            If ($MessagePage) {
                $Messages += $MessagePage
            }
        } While ($MessagePage.count -eq $BatchSizeForMessages)
}

    # Remove Exchange Online public folder hierarchy synchronization messages
    $Messages = $Messages | Where-Object {$_.Subject -NotLike "*HierarchySync*"}
    return $Messages
}

# --- Start of processing --

# Define the bits needed to connect in app-only mode. Make sure to update these for your tenant.
$AppId = "9725b779-7b69-41b9-9137-394a5801e651"
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3478"
$Thumbprint = "8CD4AD070C7447BA767EAB5DA659A02E6411BB80"

# Connect to the Microsoft Graph 
Write-Output "Connecting to Microsoft Graph..."
Connect-MgGraph -AppId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint 

# Get the tenant information and the list of verified domains for the tenant
Write-Host "Fetching tenant information and verified domains..."
$TenantInfo = Get-MgOrganization
$DefaultDomain = $TenantInfo.VerifiedDomains | Where-Object {$_.IsDefault -eq $true} | Select-Object -ExpandProperty Name    

# Find the set of verified domains for the tenant that we can use to check against incming email
[array]$Domains = Get-MgDomain -All 
If ($Domains.Count -eq 0) {
    Write-Output "No verified domains found in the tenant. Exiting."
    Break
} Else {
    [array]$VerifiedDomains = $Domains | Where-Object {$_.IsVerified -eq $true} | Select-Object -ExpandProperty Id    
}

# Connect to Exchange Online in app-only mode - before this can happen, the app must be assigned the manage Exchange as app permission and
# be assigned the Exchange administrator role for the tenant
Write-Output "Connecting to Exchange Online..."
Connect-ExchangeOnline -CertificateThumbprint $Thumbprint -AppId $AppId -Organization $DefaultDomain -ShowBanner:$false
 
# Define the MDO service plan IDs for P1 and P2 along with the Exchange Online Plan 2 plan. 
# The MDO Plan 1 is only used for Microsoft 365 SME tenants.
$MDOPlan1 = "f20fedf3-f3c3-43c3-8267-2bfdd51c0939"
$MDOPlan2 = "8e0c0a52-6a6c-4d40-8370-dd62790dcd70"
$EXOPlan2 = "efb87545-963c-4e0d-99df-69c6916d9eb0"

Write-Output "Checking shared mailboxes for MDO license requirements..."
# Get the shared mailboxes
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
If ($Mbx) {
    Write-Output ("Fetched {0} shared mailboxes from the tenant." -f $Mbx.Count)
} Else {  
    Write-Output "No shared mailboxes found in the tenant. Exiting."
    Break
}

# Determine 1-day chunks to fetch message trace data for the tenant
$StartDate = (Get-Date).AddDays(-90).ToUniversalTime()
$EndDate = (Get-Date).ToUniversalTime()

$iterations = 90
[int]$i = 0

[array]$DateRanges = $null

For ($i=0; $i -lt $iterations; $i++) {
    $RangeStart = $StartDate.AddDays($i + 1)
    $RangeEnd = $StartDate.AddDays(($i + 1) +1 )
    If ($RangeEnd -gt $EndDate) {
        $RangeEnd = $EndDate
    }
    $DateRanges += @{Start=$RangeStart; End=$RangeEnd}
}

# Fetch the message trace data
[array]$Data = $null
[array]$Messages = $null

# Fetch each of the 1-day chunks of message trace data and add it to the Messages array
Write-Output ("Fetching message trace data for {0} days in 1-day chunks..." -f $DateRanges.Count)
ForEach ($DateRange in $DateRanges) {
    $Data = Get-DateRangeData -StartDate $DateRange.Start -EndDate $DateRange.End
    # Process the data for this date range, for example by adding it to a report or exporting it to a file
    Write-Host ("Fetched {0} messages delivered between {1} and {2}" -f $Data.count, $DateRange.Start, $DateRange.End)
    $Messages+= $Data
}

# Make sure that the messages are sorted in descending order by received date, so that we can easily find the last message sent by a user.
$Messages = $Messages | Sort-Object {$_.Received -as [datetime]} -Descending

# Create an array of unique sender addresses from the message trace data, which will be used to check against the mailboxes in the tenant
[array]$MessageTable = $Messages | Sort-Object SenderAddress -Unique

$Report = [System.Collections.Generic.List[Object]]::new()  
Write-Host ("Checking shared mailboxes for MDO Plan 2. If running in an SME tenant, modify the code to check for MDO Plan 1 {0} instead." -f $MDOPlan1) -ForegroundColor Yellow
ForEach ($M in $Mbx) {
    $NeedsMDO = $false
    Write-Output ("Checking shared mailbox {0}" -f $M.DisplayName)
    $NeedsMDO = $false
    $HasMDO = $false
    $CheckForOneDriveLicense = $false

    $InboxTrafficCheck = $MessageTable | Where-Object {$_.RecipientAddress -eq $M.PrimarySmtpAddress}
    #[array]$InboxTraffic = Get-MessageTraceV2 -RecipientAddress $M.PrimarySmtpAddress -StartDate (Get-Date).AddDays(-10) -EndDate (Get-Date) -ResultSize 5000
    
    If ($InboxTrafficCheck) {
        # Get set of messasges delivered to the shared mailbox.
        [array]$InboxTraffic = $Messages | Where-Object {$_.RecipientAddress -eq $M.PrimarySmtpAddress} 
    } Else {
        $InboxTraffic = @()
        Write-Output ("No message trace data found for shared mailbox {0} over the last 90 days." -f $M.UserPrincipalName)
        Continue
    }
       
    # Check if any of the messages delivered to the shared mailbox are from external domains
    [array]$ExternalTrafficDomains = $null
    [int]$NumberOfExternalMessages = 0
    ForEach ($Message in $InboxTraffic) {
        $SenderDomain = $Message.SenderAddress.Split("@")[1].ToLower()
        If ($VerifiedDomains -notcontains $SenderDomain) {
            $NeedsMDO = $true
            $NumberOfExternalMessages++
            If ($ExternalTrafficDomains -notcontains $SenderDomain) {
                $ExternalTrafficDomains += $SenderDomain
            }
        }
    }
    If ($NeedsMDO -and $ExternalTrafficDomains) {
        [array]$AssignedPlans = Get-MgUser -UserId $M.ExternalDirectoryObjectId -Property AssignedPlans | Select-Object -ExpandProperty assignedPlans | Where-Object {$_.CapabilityStatus -eq 'Enabled'} | Select-Object -ExpandProperty ServicePlanId
        # If running in an SME tenant, use MDO Plan 1   
        If ($MDOPlan2 -notin $AssignedPlans) {
            $HasMDO = $false
        } Else {
            $HasMDO = $true
        }   
        If ($HasMDO) {
            Write-Output ("Shared mailbox {0} has an MDO Plan 2 license." -f $M.DisplayName)
        } Else {
            Write-Output ("Shared mailbox {0} does not have an MDO Plan 2 license and needs one because the mailbox has received external email from {1}" -f $M.DisplayName, ($ExternalTrafficDomains -join ", "))
        }
        If ($ExoPlan2 -in $AssignedPlans) {
            Write-Output ("Shared mailbox {0} has an Exchange Online Plan 2 license." -f $M.DisplayName)
            $EXOPlan2Licensed = $true
        } Else {
            Write-Output ("Shared mailbox {0} does not have an Exchange Online Plan 2 license." -f $M.DisplayName)
            $EXOPlan2Licensed = $false
        }

        # Check if any Teams Meetings have been organized by the shared mailbox over the last 90 days. If so, the shared mailbox might need a OneDrive license to retain meeting recordings.
        $TeamsMeetingsFound = Get-TeamsMeetings -UserId $M.ExternalDirectoryObjectId
        If ($TeamsMeetingsFound -gt 0) {
            Write-Output ("Shared mailbox {0} has organized {1} Teams meetings over the last 90 days. It might need a OneDrive license to retain meeting recordings." -f $M.DisplayName, $TeamsMeetingsFound)
            $CheckForOneDriveLicense = $true        
        } Else {
            Write-Output ("Shared mailbox {0} has not organized any Teams meetings over the last 90 days." -f $M.DisplayName)
        }   

        $ReportLine = [PSCustomObject][Ordered]@{ 
            Mailbox                         = $M.DisplayName
            PrimarySmtpAddress              = $M.PrimarySmtpAddress  
            NeedsMDO                        = $NeedsMDO
            HasMDO                          = $HasMDO
            HasExoPlan2License              = $EXOPlan2Licensed
            'Teams Meetings Organized'      = $TeamsMeetingsFound
            'Check for OneDrive License'    = $CheckForOneDriveLicense
            'Count of external email'       = $NumberOfExternalMessages
            'External Domains'              = ($ExternalTrafficDomains -join ", ")
        }
        $Report.Add($ReportLine)

    }  
}

# Send email with the results - define your own email address here and the address to receive the report.
$MsgFrom = "PowerShell.Reports@office365itpros.com"
$DestinationEmailAddress = "User.Administration@office365itpros.com"
# Build the array of a single TO recipient detailed in a hash table - change this to the appropriate recipient for your tenant
$ToRecipient = @{}
$ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
[array]$MsgTo = $ToRecipient
# Define the message subject
$MsgSubject = "Important: Shared Mailboxes that potentially need Microsoft Defender for Office 365 Licenses"
# Create the HTML content
$HtmlMsg = ("</body></html><p>The following <b>{0} Shared Mailboxes need Microsoft Defender for Office 365 Licenses</b> based on external email traffic received over the last 90 days.</p>" -f $Report.Count)
$HtmlMsg += "<p>The report also checked shared mailboxes to detect online Teams meetings organized by the mailbox. If this situation exists, you should assign a license containing an OneDrive for Business service plan to the mailbox to ensure meeting recordings are retained. </p>"
$HtmlMsg += "<p>Please review the information at your convenience</p>"
# Construct the message body 	
$HtmlBody = $Report | ConvertTo-Html -As Table -PostContent "</body></html>" -Fragment
$HtmlMsg = $HtmlMsg + $HtmlBody
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
Write-Output ""
Write-Output "Sending email..."
Write-Output ""

# Send the message
Try {
    Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters -ErrorAction Stop
    Write-Output ("Shared Mailboxes potential licensing report emailed to {0}" -f $ToRecipient.emailAddress.address)
} Catch {
    Write-Output "Unable to send email"
    Write-Output $_.Exception.Message
}

Write-Output "All done"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.