# Switch-LitigationHoldsForEDiscovery.ps1
# A script to move all mailboes with litigation hold enabled to an eDiscovery case. Once the eDiscovery hold is active, the
# litigation holds can be removed

# V1.0 20-Apr-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Switch-LitigationHoldsForEDiscovery.PS1

# Check permissions available to the signed-in account and disconnect from the Graph if the requisite permissions are not available

[string[]]$RequiredScopes = @("Sites.Read.All","User.ReadBasic.All", "Reports.Read.All", "ReportSettings.ReadWrite.All")    
Connect-MgGraph -NoWelcome -Scopes $RequiredScopes

[string[]]$CurrentScopes = (Get-MgContext).Scopes

$CheckScopes =[object[]][Linq.Enumerable]::Intersect($RequiredScopes,$CurrentScopes)
If ($CheckScopes.Count -ne 4) { # Must have all 4 required scopes
    Write-Host ("To run this script, you need to connect to Microsoft Graph with the following scopes: {0}" -f ($RequiredScopes -join ", ")) -ForegroundColor Red
    Disconnect-Graph
    Break
}

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -Notin $Modules) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -showBanner:$false -UserPrincipalName (Get-MgContext).Account
}

Connect-IPPSSession -ShowBanner:$false -UserPrincipalName (Get-MgContext).Account

# Find mailboxes with litigation hold enabled
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -Properties LitigationHoldEnabled, LitigationHoldDuration `
    -Filter {LitigationHoldEnabled -eq $True} -ResultSize Unlimited
If ($Mbx) {
    # Only take mailboxes with unlimited hold duration
    $Mbx = $Mbx | Where-Object { $_.LitigationHoldDuration -eq "Unlimited" }
    Write-Host ("Litigation hold is enabled for {0} mailboxes" -f $Mbx.Count) -ForegroundColor Green
} Else {
    Write-Host "No mailboxes with litigation hold enabled" -ForegroundColor Green
    Break
}

# Handle data obfuscation setting for the tenant because we need real rather than obfuscated names for the OneDrive accounts
$ObfuscationChanged = $false
If ((Get-MgAdminReportSetting).DisplayConcealedNames -eq $True) {
   $Parameters = @{ displayConcealedNames = $False }
   Update-MgAdminReportSetting -BodyParameter $Parameters
   $ObfuscationChanged = $true
}

Write-Host "Finding OneDrive usage information..." -ForegroundColor Yellow
$Uri = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D90')"
$Data = Invoke-MgGraphRequest -uri $Uri -method Get  -OutputFilePath data.csv
[array]$OneDriveData = Import-CSV data.csv | Select-Object 'Site Id', 'Owner Principal Name'
$OneDriveHash = @{}
ForEach ($OneDrive in $OneDriveData) {
    $OneDriveHash.Add($OneDrive.'Owner Principal Name'.toLower(), [string]$OneDrive.'Site Id')
}

# Create file containing mailbox and OneDrive data for the holds
$MbxReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {

    $OneDriveURL = $null

    # Get OneDrive URL for the mailbox
    $SiteId = $OneDriveHash[$M.UserPrincipalName.ToLower()]
    If ($SiteId) {
        $OneDrive = Get-MgSite -SiteId $SiteId -ErrorAction SilentlyContinue
        If ($OneDrive) { 
            $OneDriveURL = $OneDrive.WebUrl + "/"
        }
    } Else {
        $OneDriveURL = "No OneDrive URL found"
    }

    $ReportLine = [PSCustomObject][Ordered]@{   
        UserPrincipalName = $M.UserPrincipalName
        DisplayName = $M.DisplayName
        LitigationHoldEnabled = $M.LitigationHoldEnabled
        OneDriveUrl = $OneDriveURL
    }
    $MbxReport.Add($ReportLine)
}

# Create arrays of the mailbioxes and OneDrive accounts to place on hold
[string[]]$MailboxesToHold = $MbxReport.UserPrincipalName

# Can't have more than 1,000 mailboxes on hold for a single case: https://learn.microsoft.com/en-us/purview/ediscovery-create-holds?WT.mc_id=M365-MVP-9501#ediscovery-hold-limits
If ($MailboxesToHold.Count -gt 1000) { 
    Write-Host ("There are {0} mailboxes to place on hold. eDiscovery standard cases are limited to 1000 mailboxes. Please split the mailboxes into smaller groups and run the script again." -f $MailboxesToHold.Count) -ForegroundColor Red
}

[string[]]$OneDriveToHold = $MbxReport.OneDriveUrl | Sort-Object -Unique

Write-Host "Creating standard eDiscovery case to replace litigation holds..." -ForegroundColor Yellow
$Case = New-ComplianceCase -Description "eDiscovery Case for Mailbox Litigation Holds" -Name "eDiscovery Case to Replace Mailbox Litigation Holds"

If ($Case) {
    Write-Host ("eDiscovery case created {0} - now applying holds to {1} mailboxes and {2} OneDrive accounts" -f $Case.Name, $MailboxesToHold.Count, $OneDriveToHold.Count) -ForegroundColor Green
    $HoldPolicy = New-CaseHoldPolicy -Name "Hold Policy for Mailboxes on Litigation Hold" -Case $Case.Name -ExchangeLocation $MailboxesToHold `
        -SharePointLocation $OneDriveToHold -Comment "Hold policy for mailboxes and OneDrive accounts to replace litigation holds"

    $HoldRule = New-CaseHoldRule -Name "Hold Policy Litigation Rule 1" -Policy $HoldPolicy.Name -ContentMatchQuery ""
    # Set-CaseHoldPolicy -Identity "eDiscovery Replace Litigtion Hold Policy" -AddSharePointLocation $OneDriveToHold 
} Else {
    Write-Host "An error occurred creating the eDiscovery case" -ForegroundColor Red
    Break
}

If ($HoldRule) {}
    Write-Host ("Hold policy {0} created with rule {1}" -f $HoldPolicy.Name, $HoldRule.Name) -ForegroundColor Green
    Write-Host "The litigation hold replacement eDiscovery case is now active. After 3 hours it should be safe to remove the litigation holds from the mailboxes." -ForegroundColor Green
} Else {
    Write-Host "An error occurred creating the eDiscovery hold policy and rule" -ForegroundColor Red
}

# Reset tenant obfuscation settings to True
If ($ObfuscationChanged) {
    If ((Get-MgAdminReportSetting).DisplayConcealedNames -eq $False) {
       $Parameters = @{ displayConcealedNames = $True }
       Update-MgAdminReportSetting -BodyParameter $Parameters
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.