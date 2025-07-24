# Report-IndividualUserSessions.PS1
# An example script to report the events for an individual user session based on the session identifier assigned by Entra ID.
# Microsoft Blog: https://techcommunity.microsoft.com/blog/microsoft-entra-blog/strengthen-identity-threat-detection-and-response-with-linkable-token-identifier/4434615

# V1.0 22-Jul-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-IndividualUserSessions.PS1

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Error "The ExchangeOnlineManagement module is not loaded. Connecting now."
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
}

Connect-MgGraph -Scopes AuditLog.Read.All, User.Read.All -NoWelcome

$StartDate = (Get-Date).AddDays(-30) # Sign-in records only valid for 30 days
$UserId = Read-Host "Enter the user ID to report sessions for " # Prompt for the user ID

$UserId = $UserId.ToLower()
Try {
    $User = Get-MgUser -UserId $UserId -ErrorAction Stop
} Catch {
    Write-Error "User not found"
    Break
}

Write-Host ("Checking for sign-in sessions for user {0} ({1})" -f $User.DisplayName, $UserId)
[array]$Logs = Get-MgBetaAuditLogSignIn -Filter "userPrincipalName eq '$UserId'" -All
[array]$Sessions = $Logs | Group-Object SessionId -NoElement | Select-Object -ExpandProperty Name
# Remove the blank session ID for old records
$Sessions = $Sessions | Where-Object {([string]::IsNullOrEmpty($_)) -eq $false }

If ($Sessions.Count -eq 0) {
    Write-Host "No sessions found for user $UserId"
    Break
} Else {
    Write-Host ("Found {0} sessions for user {1}" -f $Sessions.Count, $UserId)
}

# Define HTML style 
$HtmlStyle = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; background: #f4f6f8; color: #222; }
h1 { background: #0078d4; color: #fff; padding: 16px; border-radius: 6px 6px 0 0; margin-bottom: 0; }
table { border-collapse: collapse; width: 100%; background: #fff; border-radius: 0 0 6px 6px; overflow: hidden; }
th, td { padding: 10px 12px; text-align: left; }
th { background: #e5eaf1; color: #222; }
tr { background: #fff; color: #222; }
tr:nth-child(even) { background: #f0f4fa; color: #222; }
tr:hover { background: #d0e7fa; color: #222; }
.caption { font-size: 14px; color: #555; margin-bottom: 12px; }
</style>
"@

[string]$HTMLTable = $null
$FullReport = [System.Collections.Generic.List[Object]]::new()

ForEach ($Session in $Sessions) {
    Write-Host ("Searching for audit records for {0} in session {1}" -f $UserId, $Session)
    [array]$Records = Search-UnifiedAuditLog -Formatted -StartDate $StartDate -EndDate (Get-Date) -UserIds $UserId -freeText $Session -SessionCommand ReturnLargeSet -ResultSize 5000
    If ($Records.Count -eq 0) {
        Write-Host "No audit records found for session $Session"
        Continue
    } Else {
        $Records = $Records | Sort-Object Identity -Unique
        $Records = $Records | Sort-Object {$_.CreationDate -as [datetime]} 
        Write-Host ("Found {0} audit records for session {1}" -f $Records.Count, $Session)
        $Report = [System.Collections.Generic.List[Object]]::new()
    }
    ForEach ($Record in $Records) {
            $AuditData = $Record.AuditData | ConvertFrom-Json
           
            $ReportLine = [PSCustomObject]@{
                SessionId     = $Session
                UserId        = $AuditData.UserId
                Activity      = $AuditData.Operation
                Status        = $AuditData.ResultStatus
                Timestamp     = (Get-Date $Record.CreationDate -format 'dd-MMM-yyyy HH:mm:ss')
                RecordType    = $Record.RecordType
                ClientIP      = $AuditData.ClientIP
                Id            = $AuditData.Id
            }
            $Report.Add($ReportLine)
    }
    # Create the HTML content for the session records
    $HTMLSession = $Report | ConvertTo-HTML -Fragment -As Table
    $HTMLSessionHeading = "<h2>Audit Events found for Session ID: $Session</h2><p>"
    $HTMLTable = $HTMLTable + $HTMLSessionHeading + $HTMLSession
    $FullReport.AddRange($Report)
}


$HtmlReport = @"
<html><head>
$HtmlStyle
<title>Audit Records found for $UserId</title>
</head><body><p>Report generated: <b>$((Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss 'UTC'"))</b></p>
<h1>Audit Records Found for $UserId</h1>
<p><p>Please check these audit events to validate that the account is not being used incorrectly.</p>
$HtmlTable
</body>
</html>
"@

Write-Host "Generating HTML report for audit records found for user $UserId"
$OutputHTMLFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\AuditEventsForSessions.html"
$HtmlReport | Out-File -FilePath $OutputHTMLFile -Encoding utf8
Write-Host "A HTML report containing the audit records is available in $OutputHTMLFile"

Write-Host "Generating an Excel or CSV report for audit records found for user $UserId"
# Generate reports
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    $ExcelTitle = ("Audit events for {0}" -f $UserId)
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $OutputXLSXFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\AuditEventsForSessions.xlsx"
    If (Test-Path $OutputXLSXFile) {
        Remove-Item $OutputXLSXFile -ErrorAction SilentlyContinue
    }
    $FullReport | Export-Excel -Path $OutputXLSXFile -WorksheetName "Audit Events" -Title $ExcelTitle -TitleBold -TableName "AuditEvents" 
} Else {
    $OutputCSVFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\AuditEventsForSessions.csv"
    $FullReport | Export-Csv -Path $OutputCSVFile -NoTypeInformation -Encoding Utf8
}
  
If ($ExcelGenerated) {
    Write-Host ("An Excel worksheet containing the report data is available in {0}" -f $OutputXLSXFile)
} Else {
    Write-Host ("A CSV file containing the report data is available in {0}" -f $OutputCSVFile)
}


# Unsupported record types: PlannerTask, Discovery, AzureActiveDirectoryStsLogon
# unsupported operations: Add member to group, MoveToDeletedItems
# Session Id is only available after a successful authentication, so some records will not have a session ID.

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.