# Check-SharedMailboxesForMDO.PS1
# A script to check the Entra ID accounts used for shared mailboxes to figure out if they should be licensed for Microsoft Defender for Office 365 (MDO).

# Basically, a shared mailbox needs an MDO license if it receives benefit of MDO protection. In other words, the shared mailbox receives email from
# external domains that don't belong to the tenant. If a shared mailbox only receives email from internal users, it doesn't need an MDO license.
# See https://office365itpros.com/2025/10/31/mdo-p2-licensing/ for more information.

# V1,0 13-Nov-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Check-SharedMailboxesForMDO.PS1

# This ia an example of running a script in app-only mode for both the Microsoft Graph PowerShell SDK and Exchange Online management module.
# An X.509 certificate loaded into a registered app is used to authenticate for both connections.

# Define the bits needed to connect in app-onl;y mode
$AppId = "9725b779-7b69-41b9-9137-394a5801e651"
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f3478"
$Thumbprint = "8CD4AD070C7447BA767EAB5DA659A02E6411BB80"

# Connect to the Microsoft Graph 
Write-Output "Connecting to Microsoft Graph..."
Connect-MgGraph -AppId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint 

# Get the tenant information and the list of verified domains for the tenant
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
 
# Define the MDO service plan IDs for P1 and P2 along with the Exchange Online Plan 2 plan. The MDO Plan 1 is onbly used for Microsoft 365 SME tenants.
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

$Report = [System.Collections.Generic.List[Object]]::new()  
ForEach ($M in $Mbx) {
    $NeedsMDO = $false
    Write-Output ("Checking shared mailbox {0}" -f $M.DisplayName)
    $DeliveredTraffic = $null
    $NeedsMDO = $false
    $HasMDO = $false

    [array]$InboxTraffic = Get-MessageTraceV2 -RecipientAddress $M.PrimarySmtpAddress -StartDate (Get-Date).AddDays(-10) -EndDate (Get-Date) -ResultSize 5000
    If ($InboxTraffic.Count -eq 0) {
        Write-Output ("No message trace data found for shared mailbox {0} over the last 10 days." -f $M.UserPrincipalName)
        Continue
    }
    $DeliveredTraffic = $InboxTraffic | Where-Object {$_.Status -eq "Delivered"}
    [array]$ExternalTrafficDomains = $null
    [int]$NumberOfExternalMessages = 0
    ForEach ($Message in $DeliveredTraffic) {
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
        
        If ($MDOPlan2 -notin $AssignedPlans) {
            $HasMDO = $false
        } Else {
            $HasMDO = $true
        }   
        If ($HasMDO) {
            Write-Output ("Shared mailbox {0} has an MDO license." -f $M.DisplayName)
        } Else {
            Write-Output ("Shared mailbox {0} does not have an MDO license and needs one because the mailbox has received external email from {1}" -f $M.DisplayName, ($ExternalTrafficDomains -join ", "))
        }
        If ($ExoPlan2 -in $AssignedPlans) {
            Write-Output ("Shared mailbox {0} has an Exchange Online Plan 2 license." -f $M.DisplayName)
            $EXOPlan2Licensed = $true
        } Else {
            Write-Output ("Shared mailbox {0} does not have an Exchange Online Plan 2 license." -f $M.DisplayName)
            $EXOPlan2Licensed = $false
        }
        $ReportLine = [PSCustomObject][Ordered]@{ 
            Mailbox                     = $M.DisplayName
            PrimarySmtpAddress          = $M.PrimarySmtpAddress  
            NeedsMDO                    = $NeedsMDO
            HasMDO                      = $HasMDO
            HasExoPlan2License          = $EXOPlan2Licensed
            'Count of external email'   = $NumberOfExternalMessages
            'External Domains'          = ($ExternalTrafficDomains -join ", ")
        }
        $Report.Add($ReportLine)

    }  
    
}

# Send email with the results
$MsgFrom = "PowerShell.Reports@office365itpros.com"
$DestinationEmailAddress = "User.Administration@office365itpros.com"
# Build the array of a single TO recipient detailed in a hash table - change this to the appropriate recipient for your tenant
$ToRecipient = @{}
$ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
[array]$MsgTo = $ToRecipient
# Define the message subject
$MsgSubject = "Important: Shared Mailboxes that need Microsoft Defender for Office 365 Licenses"
# Create the HTML content
$HtmlMsg = ("</body></html><p>The following <b>{0} Shared Mailboxes need Microsoft Defender for Office 365 Licenses</b> based on their email traffic over the last 10 days. Please review the information at your convenience</p>" -f $Report.Count)
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
    Write-Output ("Shared Mailboxes without MDO2 license report emailed to {0}" -f $ToRecipient.emailAddress.address)
} Catch {
    Write-Output "Unable to send email"
    Write-Output $_.Exception.Message
}

Write-Output "All done"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.