# Find-CopilotAuditRecords.PS1
# An example of how to find and report audit records generated for Copilot for Microsoft 365 interactions
# (not from mobile devices)

# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $ModulesLoaded) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
} 

Write-Host "Searching for Copilot audit records..."
[array]$Records = search-UnifiedAuditLog -StartDate (Get-Date).Adddays(-90) -EndDate (Get-Date).AddDays(1) -Formatted `
    -ResultSize 5000 -SessionCommand ReturnLargeSet -Operations CopilotInteraction
If (!($Records)) {
    Write-Host "No Copilot audit records found - exiting"
    Break
} Else {
    # Remove any duplicate records and make sure that everything is sorted in date order
    $Records = $Records | Sort-Object Identity -Unique 
    $Records = $Records | Sort-Object {$_.CreationDate -as [datetime]}
    Write-Host ("{0} Copilot audit records found. Now analyzing the content" -f $Records.count)
}

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
    $AuditData = $Rec.AuditData | ConvertFrom-Json
    $CopilotApp = $null;  $Context = $null; $CopilotLocation = $null
   
    Switch ($Auditdata.copiloteventdata.contexts.type) {
        "xlsx" {
            $CopilotApp = "Excel"
        }
        "docx" {
            $CopilotApp = "Word"
        }
        "pptx" {
            $CopilotApp = "PowerPoint"
        }
    }

    If ($Auditdata.copiloteventdata.contexts.id -like "*https://teams.microsoft.com/*") {
        $CopilotApp = "Teams"
    } ElseIf ($AuditData.CopiloteventData.AppHost -eq "bizchat") {
        $CopilotApp = "Copilot for Microsoft 365 Chat"
    }

    If ($Auditdata.copiloteventdata.contexts.id) {
        $Context = $Auditdata.copiloteventdata.contexts.id
    } ElseIf ($Auditdata.copiloteventdata.threadid) {
        $Context = $Auditdata.copiloteventdata.threadid
        $CopilotApp = "Teams"
    }

    If ($Auditdata.copiloteventdata.contexts.id -like "*-my.sharepoint.com*") {
        $CopilotLocation = "OneDrive for Business"
    } ElseIf ($Auditdata.copiloteventdata.contexts.id -like "*https://teams.microsoft.com/*") {
        $CopilotLocation = "Teams"
    } 

    $ReportLine = [PSCustomObject][Ordered]@{
        TimeStamp       = (Get-Date $Rec.CreationDate -format "dd-MMM-yyyy HH:mm")
        User            = $Rec.UserIds
        App             = $CopilotApp
        Location        = $CopilotLocation
        Context         = $Context   
    }
    $Report.Add($ReportLine)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.