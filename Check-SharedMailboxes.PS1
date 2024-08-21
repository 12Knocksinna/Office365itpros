# Check-SharedMailboxes.PS1
# Check if people are signing into shared mailboxes. If they are, check if the accounts for the mailboxes are licensed for Exchange Online (Plan 1 or Plan 2).
#
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Check-SharedMailboxes.PS1

# Connect to the Microsoft Graph and Exchange Online
Connect-MgGraph -Scopes "AuditLog.Read.All", "User.Read.All" -NoWelcome

$Modules = Get-Module | Select-Object -ExpandProperty Name
If ('ExchangeOnlineManagement' -notin $Modules) {
    Write-Output "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}
Write-Output "Finding shared mailboxes..."
$Mbx = Get-ExoMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Sort-Object DisplayName
If ($Mbx) {
    Write-Output ("{0} shared mailboxes found" -f $Mbx.Count)
} Else {
    Write-Output "No shared mailboxes found"
    Break
}
# Define the service plan IDs for Exchange Online (Plan 1) and Exchange Online (Plan 2)
$ExoServicePlan1 = "9aaf7827-d63c-4b61-89c3-182f06f82e5c"
$ExoServicePlan2 = "efb87545-963c-4e0d-99df-69c6916d9eb0" 
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($M in $Mbx) {
    $ExoPlan1Found = $false; $ExoPlan2Found = $false; $LogsFound = "No"
    Write-Output ("Checking sign-in records for {0}" -f $M.DisplayName)
    $UserId = $M.ExternalDirectoryObjectId
    [array]$Logs = Get-MgAuditLogSignIn -Filter "userid eq '$UserId'" -All -Top 1
    If ($Logs) {
        $LogsFound = "Yes"
        Write-Host ("Sign-in records found for shared mailbox {0}" -f $M.DisplayName) -ForegroundColor Red
        # Check if the shared mailbox is licensed
        $User = Get-MgUser -UserId $M.ExternalDirectoryObjectId -Property UserPrincipalName, AccountEnabled, Id, DisplayName, assignedPlans
        [array]$ExoPlans = $User.AssignedPlans | Where-Object {$_.Service -eq 'exchange' -and $_.capabilityStatus -eq 'Enabled'}
        If ($ExoServicePlan1 -in $ExoPlans.ServicePlanId) {
            $ExoPlan1Found = $true

        } ElseIf ($ExoServicePlan2 -in $ExoPlans.ServicePlanId) {
            $ExoPlan2Found = $true
        }
    
        If ($ExoPlan1Found -eq $true) {
            Write-Output ("Shared mailbox {0} has Exchange Online (Plan 1) license" -f $M.DisplayName)
        } ElseIf ($ExoPlan2Found -eq $true) {
            Write-Output ("Shared mailbox {0} has Exchange Online (Plan 2) license" -f $M.DisplayName)
        }  Else {
            Write-Host ("Shared mailbox {0} has no Exchange Online license" -f $M.DisplayName) -ForegroundColor Yellow
        }  
    } 

    $ReportLine = [PSCustomObject] @{ 
        DisplayName                 = $M.DisplayName
        ExternalDirectoryObjectId   = $M.ExternalDirectoryObjectId
        'Sign in Record Found'      = $LogsFound
        'Exchange Online Plan 1'    = $ExoPlan1Found
        'Exchange Online Plan 2'    = $ExoPlan2Found
    }
    $Report.Add($ReportLine)
}  

$Report | Out-GridView -Title "Shared Mailbox Sign-In Records and Licensing Status" 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.