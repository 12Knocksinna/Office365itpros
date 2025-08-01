# New-RetentionPolicySharedMailboxes.PS1
# A script to add shared mailboxes to a dedicated Microsoft 365 retention policy.

# V1.0 2-August-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/New-RetentionPolicySharedMailboxes.PS1
# Similar to https://github.com/12Knocksinna/Office365itpros/blob/master/Switch-LitigationHoldsForEDiscovery.PS1 which switches litigation holds to eDiscovery holds
# as explained in https://office365itpros.com/2025/04/24/litigation-hold-ediscovery/

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -Notin $Modules) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -showBanner:$false -UserPrincipalName (Get-MgContext).Account
}

# Connect to the Security & Compliance Center PowerShell
Write-Host "Connecting to Security & Compliance Center PowerShell..." -ForegroundColor Yellow
Connect-IPPSSession -ShowBanner:$false -UserPrincipalName (Get-MgContext).Account

# Find shared mailboxes
Write-Host "Finding shared mailboxes..." -ForegroundColor Yellow
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
If ($Mbx) {
    # Only take mailboxes with unlimited hold duration
    Write-Host ("{0} Shared mailboxes found" -f $Mbx.Count) -ForegroundColor Green
} Else {
    Write-Host "No shared mailboxes found" -ForegroundColor Green
    Break
}

# Create file containing mailbox data for the holds
$MbxReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {

    $ReportLine = [PSCustomObject][Ordered]@{   
        UserPrincipalName = $M.UserPrincipalName
        DisplayName = $M.DisplayName
        ObjectId = $M.ExternalDirectoryObjectId
    }
    $MbxReport.Add($ReportLine)
}

# Can't have more than 1,000 mailboxes as locations for a retention policy: https://learn.microsoft.com/en-us/microsoft-365/compliance/retention-limits?view=o365-worldwide#locations-for-retention-policies
If ($MbxReport.Count -gt 1000) { 
    Write-Host ("There are {0} shared mailboxes. Retention policy locations are limited to 1000 mailboxes. Please split the mailboxes into smaller groups and run the script again." -f $MailboxesToHold.Count) -ForegroundColor Red
    Break
}

[array]$MailboxesToHold = $MbxReport.UserPrincipalName | Sort-Object -Unique

Write-Host "Creating Microsoft 365 retention policy for shared mailboxes..." -ForegroundColor Yellow
$NewPolicy = New-RetentionCompliancePolicy -Name "Shared Mailboxes Retention Policy" -ExchangeLocation $MailboxesToHold `
    -Comment ("Retention policy for shared mailboxes created by New-RetentionPolicySharedMailboxes.PS1 script on {0}" -f (Get-Date).ToString("dd-MMM-yyyy")) 
If ($NewPolicy) {
    Write-Host ("Retention policy {0} created" -f $NewPolicy.Name) -ForegroundColor Green
    $NewPolicyRule = New-RetentionComplianceRule -Name TwoYearSharedMailboxRetention -Policy "Shared Mailboxes Retention Policy" -RetentionDuration 730 `
        -Comment "Created by New-RetentionPolicySharedMailboxes.PS1 script"
    If ($NewPolicyRule) {
        Write-Host ("Retention rule {0} created" -f $NewPolicyRule.Name) -ForegroundColor Green
    } Else {
        Write-Host "Failed to create retention rule" -ForegroundColor Red
        Break
    }
} Else {
    Write-Host "Failed to create retention policy" -ForegroundColor Red
    Break
}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.