# Report-LastAppAccessedByUsers.PS1
# Use the Entra ID sign-in audit log and the Microsoft 365 audit log to report the last Microsoft 365 app
# accessed by users

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-LastAppAccessedByUsers.PS1

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

# Get licensed member accounts for the tenant, excluding utility accounts
[array]$Users = Get-MgUser -All -PageSize 500 -Filter "usertype eq 'member' and accountenabled eq true and employeeType ne 'Utility' and assignedLicenses/`$count ne 0" `
    -Sort displayName -ConsistencyLevel Eventual -CountVariable Count -Property DisplayName, UserPrincipalName, id, employeeType, userType, accountEnabled

# $Users = $Users | Where-Object {$_.DisplayName -NotMatch "admin|test|service|breakglass|break glass|monitor|audit|helpdesk|it dept|itdept|it team|itteam|exchange admin|exchangeadmin|sharepoint admin|sharepointadmin|onedrive admin|onedriveadmin|teams admin|teamsadmin"}

If ($Users) {
    Write-Host "Found $($Users.Count) user accounts to process."
    $Report = [System.Collections.Generic.List[Object]]::new()
} Else {
    Write-Host "No user accounts found to process."
    Break
}

# Loop through the users to see what we can find
[datetime]$StartProcessing = Get-Date
[int]$i = 0
ForEach ($User in $Users) {
    $i++
    Write-Host ("Checking sign-in records and audit records for {0} {1}/{2}" -f $User.DisplayName, $i, $Users.Count)
    $LastAppSignIn = $null; $LastAppSignInDate = $null; $AuditLogTimeStamp = $null
    # Check Entra ID sign-in logs
    $UPN = $User.UserPrincipalName.tolower()
    [array]$Logs = Get-MgAuditLogSignIn -Filter "userPrincipalName eq '$UPN'" -Top 3
    If ($Logs) {
        [array]$AppNames = $Logs.AppDisplayName | Sort-Object -Unique
        $LastAppSignIn = $AppNames -join "; "
        $LastAppSignInDate = Get-Date ($Logs[0].CreatedDateTime) -format 'dd-MMM-yyyy HH:mm:ss'
    }
    # Check Microsoft 365 audit log
    [array]$AuditLogs = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date) -UserIds $UPN -ResultSize 3 -Formatted  
    If ($AuditLogs) {
        [array]$AuditLogApps = $AuditLogs.RecordType | Sort-Object -Unique
        [array]$AuditLogActions = $AuditLogs.Operations | Sort-Object -Unique
        $AuditLogTimeStamp = Get-Date ($AuditLogs[0].CreationDate) -format 'dd-MMM-yyyy HH:mm:ss'
    }

    [array]$AuditLogAppsPrint = $null
    ForEach ($App in $AuditLogApps) {

        # Modify the record type contained in the audit records to make it more understandable - Audit log record types are defined at https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-schema?WT.mc_id=M365-MVP-9501#auditlogrecordtype
        Switch -Wildcard ($App) {
            "Exchang*" {            
                $AuditLogAppsPrint = "Exchange Online"
            }
            "AzureActiveDirectory*" {
                $AuditLogAppsPrint += "Entra ID"
            }
            "SharePoint*" {
                $AuditLogAppsPrint += "SharePoint Online"
            }
            "CopilotInteraction" {
                $AuditLogAppsPrint += "Microsoft 365 Copilot"
            }   
            "MicrosoftTeams" {
                $AuditLogAppsPrint += "Microsoft Teams"
            }
            "MIP*" {
                $AuditLogAppsPrint += "Microsoft Information Protection"
            }
            "Compliance*" {
                $AuditLogAppsPrint += "Microsoft Purview Compliance"
            }
            "Planner*" {
                $AuditLogAppsPrint += "Microsoft Planner"
            }
            "Power*" {
                $AuditLogAppsPrint += "Microsoft Power Platform"
            }
            "Project*" {
                $AuditLogAppsPrint += "Microsoft Project"
            }
        }
    }

    If ($Logs.count -gt 0 -or $AuditLogs.count -gt 0) {
        $Report.Add([PSCustomObject][ordered]@{
            UserPrincipalName       = $UPN
            Name                    = $User.DisplayName
            LastAppSignIn           = $LastAppSignIn
            LastAppSignInDate       = $LastAppSignInDate
            AuditLogApp             = $AuditLogAppsPrint -join ", "
            AuditLogAction          = $AuditLogActions -join ", "
            AuditLogTimeStamp       = $AuditLogTimeStamp
        })
    }
}
[datetime]$EndProcessing = Get-Date
$TimeRequired = $EndProcessing - $StartProcessing
$Minutes = [math]::Floor($TimeRequired.TotalSeconds / 60)
$Seconds = [math]::Round($TimeRequired.TotalSeconds % 60, 2)
Write-Host ("Total processing time for {0} users: {1}m {2}s" -f $users.count, $Minutes, $Seconds) -ForegroundColor Cyan
Write-Host ("Average required per user {0} seconds" -f [math]::Round($TimeRequired.TotalSeconds / $users.count, 2))

# Create a nice HTML report
# Generate sortable HTML table with type-aware sorting - use number as the type for numeric values, date for dates, and string for text
$HtmlHeader = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Last App Access by User Accounts</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; background: #f4f6f8; color: #222; }
h1 { background: #0078d4; color: #fff; padding: 16px; border-radius: 6px 6px 0 0; margin-bottom: 20px; }
table { width: 100%; background: #fff; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-collapse: collapse; }
th, td { padding: 12px; text-align: left; }
th { background: #e5eaf1; cursor: pointer; position: relative; }
th:hover { background: #d0e7fa; }
th::after { content: '↕'; position: absolute; right: 8px; opacity: 0.5; }
tr:nth-child(even) { background: #f0f4fa; }
tr:hover { background: #d0e7fa; }
</style>
<script>
function parseValue(val, type) {
    if(type === 'number') return parseFloat(val.replace(/,/g,'')) || 0;
    if(type === 'date') return new Date(val);
    return val.toLowerCase();
}
function sortTable(n, type) {
    var table = document.getElementById('UserAppStats');
    var rows = Array.from(table.rows).slice(1);
    var dir = table.getAttribute('data-sortdir'+n) === 'asc' ? 'desc' : 'asc';
    rows.sort(function(a, b) {
        var x = parseValue(a.cells[n].innerText, type);
        var y = parseValue(b.cells[n].innerText, type);
        if(x < y) return dir === 'asc' ? -1 : 1;
        if(x > y) return dir === 'asc' ? 1 : -1;
        return 0;
    });
    rows.forEach(function(row) { table.tBodies[0].appendChild(row); });
    table.setAttribute('data-sortdir'+n, dir);
}
</script>
</head>
<body>
<h1>Last App Access by Users</h1>
<table id="UserAppStats">
<thead>
<tr>
<th onclick="sortTable(0,'string')">UserPrincipalName</th>
<th onclick="sortTable(1,'string')">Name</th>
<th onclick="sortTable(2,'string')">LastAppSignIn</th>
<th onclick="sortTable(3,'string')">LastAppSignInDate</th>
<th onclick="sortTable(4,'string')">AuditLogApp</th>
<th onclick="sortTable(5,'string')">AuditLogAction</th>
<th onclick="sortTable(6,'date')">AuditLogTimeStamp</th>
</tr>
</thead>
<tbody>
"@

$HtmlRows = foreach ($Row in $Report ) {
    "<tr><td>$($row.UserPrincipalName)</td><td>$($row.Name)</td><td>$($row.LastAppSignIn)</td><td>$($row.LastAppSignInDate)</td><td>$($row.AuditLogApp)</td><td>$($row.AuditLogAction)</td><td>$($row.AuditLogTimeStamp)</td></tr>"
}

$HtmlFooter = @"
</tbody>
</table>
</body>
</html>
"@

#Generate the full HTML content and save it to a file
$HTMLReportFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\LastAppAccessbyUsers.html"
$HTMLFile = $HtmlHeader + ($HtmlRows -join "`n") + $HtmlFooter
$HTMLFile | Out-File -FilePath $HTMLReportFile -Encoding utf8

If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\LastAppAccessbyUsers.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    }
    $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Last App Access byUsers" -Title ("Last App Access by Users Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "LastAppAccessByUsers" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\LastAppAccessbyUsers.CSV"
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}
 
If ($ExcelGenerated) {
    Write-Host ("An Excel report about the last app access by users is available in {0}" -f $ExcelOutputFile)
} Else {
    Write-Host ("A CSV report about the last app access by users is available in {0}" -f $CSVOutputFile)
}

Write-Host "Data for last app access by users is also available in a HTML format at $HTMLReportFile" 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 