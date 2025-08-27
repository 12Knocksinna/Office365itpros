# Report-LastAppAccessedByUsers.PS1
# Use the Entra ID sign-in audit log and the Microsoft 365 audit log to report the last Microsoft 365 app
# accessed by users

# V1.0 27-Aug-2025


If ([Environment]::UserInteractive) { 
    # We're running interactively...
    Write-Host "Script running interactively... connecting to the Graph" -ForegroundColor Yellow
    Connect-MgGraph -NoWelcome
    [array]$Modules = Get-Module | Select-Object -ExpandProperty Name
    If ("ExchangeOnlineManagement" -Notin $Modules) {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -ShowBanner:$false 
    }
} Else { 
    # We're not, so likely in Azure Automation
    Write-Host "Running the script to identify the last app accessed by Users" 
    Connect-MgGraph -Identity
    $Tenant = Get-MgOrganization
    # Connect with a managed identity
    Connect-ExchangeOnline -Identity -Organization $Tenant.Id
}

# Check that we have the right permissions
[string[]]$CurrentScopes = (Get-MgContext).Scopes
[string[]]$RequiredScopes = @('AuditLog.Read.All','User.Read.All', 'Reports.Read.All')

$CheckScopes =[object[]][Linq.Enumerable]::Intersect($RequiredScopes,$CurrentScopes)
If ($CheckScopes.Count -ne 3) { 
    Write-Host ("To run this script, you need to connect to Microsoft Graph with the following scopes: {0}" -f $RequiredScopes) -ForegroundColor Red
    Disconnect-Graph
    Break
}

# Get licensed users for the tenant
[array]$Users = Get-MgUser -All -PageSize 500 -Filter "usertype eq 'member' and assignedLicenses/`$count ne 0" -Sort displayName -ConsistencyLevel Eventual -CountVariable Count

If ($Users) {
    Write-Host "Found $($Users.Count) user accounts to process."
    $Report = [System.Collections.Generic.List[Object]]::new()
} Else {
    Write-Host "No user accounts found to process."
    Break
}
# Loop through the users to see what we can find
[datetime]$StartProcessing = Get-Date
ForEach ($User in $Users) {
    Write-Host ("Checking sign-in records and audit records for {0}" -f $User.DisplayName)
    $UPN = $User.UserPrincipalName.tolower()
    $LastAppSignIn = $null; $LastAppSignInDate = $null; $AuditLogApp = $null; $AuditLogAction = $null; $AuditLogTimeStamp = $null
    # Check Entra ID sign-in logs
    [array]$Logs = Get-MgBetaAuditLogSignIn -Filter "userPrincipalName eq '$UPN'" -Top 1
    If ($Logs) {
        $LastAppSignIn = $Logs | Select-Object -ExpandProperty AppDisplayName
        $LastAppSignInDate = Get-Date ($Logs | Select-Object -ExpandProperty CreatedDateTime) -format 'dd-MMM-yyyy HH:mm:ss'
    }
    # Check Microsoft 365 audit log
    [array]$AuditLogs = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date) -UserIds $UPN -ResultSize 1 -Formatted  
    If ($AuditLogs) {
        $AuditLogApp = $AuditLogs.RecordType
        $AuditLogAction = $AuditLogs.Operations
        $AuditLogTimeStamp = Get-Date ($AuditLogs.CreationDate) -format 'dd-MMM-yyyy HH:mm:ss'
    }

    # Modify the audit app to make it more readable - Audit log record types are defined at https://learn.microsoft.com/en-us/graph/api/resources/auditlog?view=graph-rest-1.0
    Switch ($AuditLogApp) {
        "Exchang*" {            
            $AuditLogApp = "Exchange Online"
        }
        "AzureActiveDirectory*" {
            $AuditLogApp = "Entra ID"
        }
        "SharePoint*" {
            $AuditLogApp = "SharePoint Online"
        }
        "CopilotInteraction" {
            $AuditLogApp = "Microsoft 365 Copilot"
        }   
        "MicrosoftTeams" {
            $AuditLogApp = "Microsoft Teams"
        }
        "MIP*" {
            $AuditLogApp = "Microsoft Information Protection"
        }
        "Compliance*" {
            $AuditLogApp = "Microsoft Purview Compliance"
        }
        "Planner*" {
            $AuditLogApp = "Microsoft Planner"
        }
        "Power*" {
            $AuditLogApp = "Microsoft Power Platform"
        }
        "Project*" {
            $AuditLogApp = "Microsoft Project"
        }
    }
    
    If ($Logs.count -gt 0 -or $AuditLogs.count -gt 0) {
        $Report.Add([PSCustomObject][ordered]@{
            UserPrincipalName       = $UPN
            LastAppSignIn           = $LastAppSignIn
            LastAppSignInDate       = $LastAppSignInDate
            AuditLogApp             = $AuditLogApp
            AuditLogAction          = $AuditLogAction
            AuditLogTimeStamp       = $AuditLogTimeStamp
        })
    }
}
[datetime]$EndProcessing = Get-Date
$TimeRequired = $EndProcessing - $StartProcessing
Write-Host ("Time taken to process {0} users: {1}" -f $users.count, $TimeRequired.totalSeconds) -ForegroundColor Green