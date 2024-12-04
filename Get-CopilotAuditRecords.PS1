

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ($Modules -notcontains "ExchangeOnlineManagement") {     
    Write-Host "Connecting to Microsoft Teams..."
    Connect-ExchangeOnline -ShowBanner:$False
}

[array]$Records = Search-UnifiedAuditLog –StartDate (Get-Date).AddDays(-180) –EndDate (Get-Date).AddDays(1) –ResultSize 5000 –SessionCommand ReturnLargeSet -Formatted -Operations 'CopilotInteraction'
$Records = $Records | Sort-object Identity -Unique | Sort-Object {$_.CreationDate -as [DateTime]}
$Records | Group-Object userids -NoElement | Sort-Object Count -Descending | Format-Table Name, Count

$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Rec in $Records) {
    $MsgCount = 0
    $AuditData = $Rec.AuditData | ConvertFrom-Json
    Switch ($Auditdata.CopilotEventData.AppHost) {
        "Outlook" {
            $MsgCount = $AuditData.CopilotEventData.AccessedResources.Count
        }
        "Word" {
            $MsgCount = $Auditdata.CopilotEventData.Contexts.Id.count
        }
    }
    $ReportLine = [pscustomobject]@{
        User            = $Rec.UserIds
        CreationDate    = $Rec.CreationDate
        'Copilot App'   = $Auditdata.CopilotEventData.AppHost
        Context         = $AuditData.CopilotEventData.Contexts.Id
        FileType        = $AuditData.CopilotEventData.Contexts.Type
        Prompt          = $AuditData.CopilotEventData.Messages.IsPrompt
        ThreadId        = $AuditData.CopilotEventData.Messages.Id
        ItemCount       = $MsgCount
    }
    $Report.Add($ReportLine)
}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.