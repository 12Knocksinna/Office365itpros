# Report-MailboxAuditConfiguration.PS1
# A script to report the audit configuration for all user and shared mailboxes
# V1.0 22-May-2024
# GitHub link:

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement -notin $Modules") {
    Write-Host "Hold on while we connect to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited `
    -Properties DefaultAuditSet, AuditOwner, AuditDelegate, AuditAdmin, AuditEnabled |`
    Sort-Object RecipientTypeDetails -Descending
Write-Host ("{0} mailboxes found" -f $Mbx.Count)
# Define the default audit set
[array]$DefaultAuditSet = "Admin", "Delegate", "Owner"
[string]$DefaultAuditSetReport = $DefaultAuditSet -join ", "
# Define the auditable actions that we want to check for
[array]$CriticalAuditActions = "MailItemsAccessed", "Send", "searchqueryinitiatedexchange"
# Create output CSV file in the downloads folder
$CSVOutputFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + ("\MailboxAuditingReport.csv" -f $Site.displayName, $DriveName)

Write-Host "Checking mailboxes..."
$Report = [System.Collections.Generic.List[Object]]::new()
[int]$CustomAuditConfigurations = 0
ForEach ($M in $Mbx) {
    # Check if the mailbox uses a custom audit configuration
    $CustomAuditSet = $null; $UsesDefaultAuditSet = $true; [array]$MissingActions = $null
    If ($null -ne (Compare-Object -ReferenceObject $DefaultAuditSet -DifferenceObject $M.DefaultAuditSet)) {
       $CustomAuditSet = $M.DefaultAuditSet -join ","
       $UsesDefaultAuditSet = $false
       $CustomAuditConfigurations++
    } 
    ForEach ($Action In $CriticalAuditActions) {
        If ($Action -notin $M.AuditOwner) {
            $MissingActions += $Action
        }
    }
    # create formatted strings for the actions audited in each set
    $AuditOwner = $M.AuditOwner -join ", "
    $AuditDelegate = $M.AuditDelegate -join ", "
    $AuditAdmin = $M.AuditAdmin -join ", "
    # if the mailbox uses a custom configuration, report that, otherwise report the default set
    If ($null -ne $CustomAuditSet) {
        $MailboxAuditSet = $CustomAuditSet
    } Else {
        $MailboxAuditSet = $DefaultAuditSetReport
    }
    # This set is to make sure that mailbox auditing is enabled and audit events are being ingested by
    # the audit log
    If ($M.AuditEnabled -eq $true) {
        Set-Mailbox -Identity $M.UserPrincipalName -AuditEnabled $true -WarningAction SilentlyContinue
    }
    # Generate the report line
    $ReportLine = [PSCustomObject]@{
        UserPrincipalName        = $M.UserPrincipalName
        RecipientTypeDetails     = $M.RecipientTypeDetails
        AuditEnabled             = $M.AuditEnabled
        DefaultAuditSet          = $MailboxAuditSet
        'Uses default audit set' = $UsesDefaultAuditSet
        'Missing actions'        = $MissingActions -join ', '   
        AuditOwner               = $AuditOwner
        AuditDelegate            = $AuditDelegate
        AuditAdmin               = $AuditAdmin
    }
    $Report.Add($ReportLine)
}

[array]$NotEnabledforAuditing = $Report | Where-Object { $_.AuditEnabled -eq $false }
Write-Host ("All done. Of the {0} mailboxes, {1} are audit enabled, and {2} use custom audit configurations" -f $Report.Count, ($Mbx.count - $NotEnabledforAuditing.count), $CustomAuditConfigurations)
$Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
Write-Host ("Output data is available in {0}" -f $CSVOutputFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.