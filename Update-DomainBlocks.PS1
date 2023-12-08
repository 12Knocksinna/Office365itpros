# Update-DomainBlocks.PS1
# Example of how to use PowerShell to read tenant-level block instructions for inbound email and 
# apply them through the tenant allow/block lists and through a transport rule (mail flow rule)
# to block all inbound email from selected top-level domains.
# V1.0 1-Dec-2023
# https://github.com/12Knocksinna/Office365itpros/blob/master/Update-DomainBlocks.PS1

Function Add-MessageRecipients {
    # Function to build an addressee list to send email   
    [cmdletbinding()]
        Param(
        [array]$ListOfAddresses )
        ForEach ($SMTPAddress in $ListOfAddresses) {
            @{ emailAddress = @{address = $SMTPAddress}}    
        }
} 

# Connect to the Graph SDK endpoint (tested with SDK V2.10)
Connect-MgGraph -Scopes Sites.Read.All, Sites.Manage.All, Mail.Send -NoWelcome
# To run in Azure Automation with a managed identity use
# Connect-MgGraph -NoWelcome -Identity

Connect-ExchangeOnline -SkipLoadingCmdletHelp
# to run in Azure Automation
# $ServiceDomain = (Get-MgOrganization).VerifiedDomains | Where-Object {$_.IsInitial -eq $True} | Select-Object -ExpandProperty Name
# Connect-ExchangeOnline -ManagedIdentity -TenantName $ServiceDomain -SkipLoadingCmdletHelp

# Transport rules to use for blocking top-level domains (TLDs) and individual domains
$TransportRuleName = "Block Inbound messages from selected top-level domains"
$TransportRuleName2 = "Block Email From Selected Domains"
$ExpirationDate = (Get-Date).AddDays(90).ToUniversalTime()

# Recipient for the email sent at the end of the script - define the addresses you want to use here
$EmailRecipient = "Peter.Bedson@o365maestro.onmicrosoft.com"
# Message will be sent from the account running the script. If used with the Mail.Send permission in an Azure
# Automation runbook, the sender can be any mailbox in the organization
$MsgFrom = (Get-MgContext).Account
$NewRuleCreated = $false; $NewRuleCreated2 = $false

# SharePoint Site that stores the list holding data about blocks to apply
$Site = Get-MgSite -Search 'Office 365 Adoption' | Select-Object -First 1
# List in the site with the block information
$List = Get-MgSiteList -SiteId $Site.Id -Filter "displayName eq 'Domain Blocks'"

# Read items from the list and build a PowerShell list object containing blocks to apply
Write-Output ("Fetching information about tenant blocks from SharePoint Online list {0}" -f $List.displayName)
[array]$Data = Get-MgSiteListItem -ListId $List.Id -SiteId $Site.Id  -ExpandProperty Fields
[array]$Items = $Data.Fields.additionalProperties

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Item in $Items) {
    $ItemtoBlock    = $Item.Title
    $BlockType      = $Item.DomainType
    $Notes          = $Item.Value
    $DataLine = [PSCustomObject][Ordered]@{
        ItemToBlock = $ItemtoBlock
        BlockType   = $BlockType
        Notes       = $Notes 
    }
    $Report.Add($DataLine)
}

# Now process the blocks in the list
Write-Output "Processing block instructions"
# List to capture details of what the script does
$BlockReport = [System.Collections.Generic.List[Object]]::new()
# Create arrays for the two kinds of block processed by the script
[array]$SenderBlocks = $Report | Where-Object {$_.BlockType -eq "Sender"} | Sort-Object ItemToBlock
[array]$URLBlocks = $Report | Where-Object {$_.BlockType -eq "URL"} | Sort-Object ItemToBlock
# Fetch the current set of blocks for each type
[array]$CurrentSenderBlocks = Get-TenantAllowBlockListItems -ListType Sender | Select-Object Identity, Value
[array]$CurrentURLBlocks = Get-TenantAllowBlockListItems -ListType URL | Select-Object -ExpandProperty Value
# array to hold the TLDs to block with the transport rule
[array]$TLDBlocksforTransportRule = $Null
[array]$IndividualDomainBlocks = $Null

# Process sender blocks
ForEach ($BlockSender in $SenderBlocks) {
    Write-Output ("Processing block for {0}" -f $BlockSender.ItemToBlock)
    # Check if a block already exists for this sender. If so, remove it
    $CheckSender = $CurrentSenderBlocks | Where-Object {$_.Value -eq $BlockSender.ItemToBlock}
    If ($CheckSender) {
        Write-Output ("Removing current block for {0}" -f $BlockSender.ItemToBlock)
        $Status = Remove-TenantAllowBlockListItems -ListType Sender -Ids $CheckSender.Identity -ErrorAction SilentlyContinue
    }
    # Add the new sender block list entry
    $Status = (New-TenantAllowBlockListItems -ListType Sender -Entries $BlockSender.ItemToBlock -Block `
        -ExpirationDate $ExpirationDate -ErrorAction SilentlyContinue)
    If ($Status) {
        Write-Output ("Block successfully applied for {0}" -f $BlockSender.ItemToBlock)
        $BlockData = [PSCustomObject][Ordered]@{
            Timestamp   = Get-Date -Format s
            Block       = $BlockSender.ItemToBlock
            BlockType   = 'Sender'
        }
        $BlockReport.Add($BlockData)
    } Else {
        Write-Output "Error occurred adding block"
    }
}

# Process URL blocks
ForEach ($BlockURL in $URLBlocks) {
    Write-Output ("Processing block for {0}" -f $BlockURL.ItemToBlock)
    If (($BlockURL.ItemToBlock.Substring(0,1)) -eq ".") {
        $URLToBlock = ("*{0}/*" -f $BlockURL.ItemToBlock)
        $TLDBlock = ("\{0}$" -f $BlockURL.ItemToBlock)
        $TLDBlocksforTransportRule += $TLDBlock
    } Else {
        $URLToBlock = $BlockURL.ItemToBlock
        $IndividualDomainBlocks += $BlockURL.ItemToBlock
    }
    # If URL is already blocked, remove the current block
    If ($URLToBlock -in $CurrentURLBlocks) {
        Write-Output ("Removing current block for {0}" -f $URLToBlock)
        $Status = Remove-TenantAllowBlockListItems -ListType URL -Entries $URLToBlock -ErrorAction SilentlyContinue
    }
    # Add the block for 90 days
    $Status = (New-TenantAllowBlockListItems -ListType URL -Entries $URLToBlock -Block `
        -ExpirationDate $ExpirationDate -ErrorAction SilentlyContinue)
    If ($Status) {
        Write-Output ("Block successfully applied for {0}" -f $URLToBlock)
        $BlockData = [PSCustomObject][Ordered]@{
            Timestamp   = Get-Date -Format s
            Block       = $URLToBlock
            BlockType   = 'URL'
        }
        $BlockReport.Add($BlockData)
    } Else {
        Write-Output "Error occurred adding block"
    }
}

If ($TLDBlocksforTransportRule) {
# Now to update the transport rule to block TLDs if any TLDs are to be blocked. 
# First, check if a rule exists. If it doesn't, create it
    [array]$CheckTransportRule = Get-TransportRule -Identity $TransportRuleName -ErrorAction SilentlyContinue
    $Comments = ("Rule updated automatically on {0} to process TLDs: {1}" -f (Get-Date -format 'dd-MMM-yy HH:mm'), ($TLDBlocksforTransportRule -join ", "))
    If (!($CheckTransportRule)) {
    # Transport rule not present, so create new rule
        $NewRule = New-TransportRule -Name $TransportRuleName -Enabled $True `
            -FromAddressMatchesPatterns $TLDBlocksforTransportRule -SenderAddressLocation 'Header' `
            -Comments $Comments -Quarantine $true
        If ($NewRule) {
            Write-Output "Transport rule created to block email from specified TLDs" 
            $NewRuleCreated = $true
        }
    } Else {
    # We have a transport rule, so update it
        Write-Output "Updating transport rule for TLD blocks..."
        Set-TransportRule -Identity $TransportRuleName -FromAddressMatchesPatterns $TLDBlocksforTransportRule `
            -ErrorAction SilentlyContinue -Comments $Comments
        $BlockedTLDs = Get-TransportRule -Identity $TransportRuleName | Select-Object -ExpandProperty FromAddressMatchesPatterns
        If (!(Compare-Object -ReferenceObject $TLDBlocksForTransportRule -DifferenceObject $BlockedTLDs)) {
            Write-Output ("Transport rule updated to block email from these TLDs {0}:" -f ($TLDBlocksforTransportRule -join ", "))
        }    
    }
}

# And check if we have to block individual domains
If ($IndividualDomainBlocks) {
    [array]$CheckTransportRule = Get-TransportRule -Identity $TransportRuleName2 -ErrorAction SilentlyContinue
    $Comments = ("Rule updated automatically on {0} to process domains: {1}" -f (Get-Date -format 'dd-MMM-yy HH:mm'), ($IndividualDomainBlocks -join ", "))
    If (!($CheckTransportRule)) {
    # Transport rule not present, so create new rule
        $NewRule = New-TransportRule -Name $TransportRuleName2 -Enabled $True `
            -SenderDomainIs $IndividualDomainBlocks `
            -Comments $Comments -Quarantine $true
        If ($NewRule) {
            Write-Output "Transport rule created to block email for individual domains" 
            $NewRuleCreated2 = $true
        }
    } Else {
    # We have a transport rule, so update it
        Write-Output "Updating transport rule to block specific email domains..."
        Set-TransportRule -Identity $TransportRuleName2 -SenderDomainIs $IndividualDomainBlocks -Comments $Comments `
            -ErrorAction SilentlyContinue 
        $BlockedDomains = Get-TransportRule -Identity $TransportRuleName2 | Select-Object -ExpandProperty SenderDomainIs
        If (!(Compare-Object -ReferenceObject $IndividualDomainBlocks -DifferenceObject $BlockedDomains)) {
            Write-Output ("Transport rule updated to block email from these domains: {0}" -f ($IndividualDomainBlocks -join ", "))
        }    
    }
}

$ToRecipientList   = @( $EmailRecipient )
[array]$MsgToRecipients = Add-MessageRecipients -ListOfAddresses $ToRecipientList
$MsgSubject = "Tenant blocks for email"
$HtmlHead = "<h2>Updates to tenant block list</h2><p>The following tenant blocks have been applied.</p>"
$HtmlBody = $BlockReport | ConvertTo-Html -Fragment 
$HtmlMsg = "</body></html><p>" + $HtmlHead + $Htmlbody + "<p>"
If ($NewRuleCreated) {
    $HtmlMsg = $HtmlMsg + ("<p><b>New transport rule created to block top-level domains: {0}</b></p>" -f $TransportRuleName)
} Else {
    $HtmlMsg = $HtmlMsg + ("<p>Transport rule <b>{0}</b> updated to block top-level domains</p>" -f $TransportRuleName)
}
If ($NewRuleCreated2) {
    $HtmlMsg = $HtmlMsg + ("<p><b>New transport rule created to block email from specific domains: {0}</b></p>" -f $TransportRuleName2)
} Else {
    $HtmlMsg = $HtmlMsg + ("<p>Transport rule <b>{0}</b> updated to block email from specific domains</p>" -f $TransportRuleName2)
}
# Construct the message body
$MsgBody = @{
  Content = "$($HtmlMsg)"
  ContentType = 'html'  
}

$Message =  @{subject           = $MsgSubject}
$Message += @{toRecipients      = $MsgToRecipients}  
$Message += @{body              = $MsgBody}
$Params   = @{'message'         = $Message}
$Params  += @{'saveToSentItems' = $True}
$Params  += @{'isDeliveryReceiptRequested' = $True}

# And send the message using the parameters that we've filled in
Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params
Write-Output ("Message containing update about tenant blocks sent to {0}!" -f $EmailRecipient)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
