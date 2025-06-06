# Find Copilot interactions for a user between two dates using the aiInteractionHistory API
# V1.0 4-Dec-2024
# V1.1 24-May-2025 added $Top parameter to the API call to fetch 100 records at a time, which seems to be the current maximum

# Requires an app with User.Read.All and AiEnterpriseInteraction.Read.All permissions
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-CopilotInteractions-Graph.PS1

# disconnect any previous Graph session so that we can run in app-only mode using the AiEnterpriseInteraction.Read.All permission
# will also need the User.Read.All permission to run Get-MgUser. These values will be different for your tenant...
Disconnect-MgGraph | Out-Null
$AppId = '9107c646-b5a3-4a64-a06d-8509df0f00a4'
$Thumbprint = '32C9529B1FFD08BCD483A5D98807E47A472C5318'
$TenantId = 'a662313f-14fc-43a2-9a7a-d2e27f4f3478'
Connect-MgGraph -AppId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome

$StartDate = (Get-Date).AddDays(-30).toString('yyyy-MM-ddT00:00:00Z')
$EndDate = (Get-Date).AddDays(1).toString('yyyy-MM-ddT00:00:00Z')
$StartDateForReport = Get-Date $StartDate -format 'dd-MMM-yyyy'
$EndDateForReport = Get-Date $EndDate -format 'dd-MMM-yyyy'

$UserPrincipalName = Read-Host "Enter the user principal name for the user to search for" 
$User = Get-MgUser -UserId $UserPrincipalName.trim() -ErrorAction SilentlyContinue
If (!$User) {
    Write-Host ("User {0} not found" -f $UserPrincipalName)
    Break
}
# Has the account got a Copilot license?
[array]$UserLicenses = Get-MgUserLicenseDetail -UserId $User.Id | Select-Object -ExpandProperty SkuId
If ("639dec6b-bb19-468b-871c-c5c441c4b0cb" -notin $UserLicenses) {
    Write-Host ("User {0} does not have a Copilot license, so we can't check their Copilot interactions" -f $User.DisplayName)
    Break
}

$Uri = ("https://graph.microsoft.com/beta/copilot/users/{0}/interactionHistory/getAllEnterpriseInteractions?`$top=100&`$filter=createdDateTime gt {1} and createdDateTime lt {2}" `
    -f $User.Id, $StartDate, $EndDate)
Write-Host ("Searching for Copilot interactions for {0} between {1} and {2}" -f $User.DisplayName, $StartDateForReport, $EndDateForReport)
[array]$CopilotData = $null
# Get the first set of records
[array]$Data = Invoke-MgGraphRequest -Uri $Uri -Method GET
$CopilotData = $Data.Value
If (!($CopilotData)) {
    Write-Host ("No Copilot interactions found for {0} between {1} and {2}" -f $User.DisplayName, $StartDateForReport, $EndDateForReport)
    Break
}

$Nextlink = $Data.'@odata.nextLink'
While ($null -ne $Nextlink) {
    Write-Host ("Fetching more records - currently at {0}" -f $CopilotData.count)
    [array]$Data = Invoke-MgGraphRequest -Uri $Nextlink -Method Get
    $CopilotData += $Data.Value
    $Nextlink = $Data.'@odata.nextLink'
}
# Remove any null records
$CopilotData = $CopilotData | Where-Object { $_ -ne $null }
$CopilotData = $CopilotData | Sort-Object {$_.createdDateTime -as [datetime]}

Write-Host ("{0} Copilot interactions for {1} between {2} and {3} have been retrieved" -f $CopilotData.count, $User.DisplayName, $StartDateForReport, $EndDateForReport)
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Record in $CopilotData) {

    If ($Record.createdDateTime) {
        $Timestamp = Get-Date $Record.createdDateTime -format 'dd-MMM-yyyy HH:mm:ss'
    } else {
        $Timestamp = $null
    }
    Switch ($Record.interactionType) {
        "userPrompt" {
            $AppName = $User.displayname
            $AppId = $record.from.user.id
        }
        "aiResponse" {
            $AppName = $Record.from.application.displayname
            $AppId = $Record.from.application.id
        }
        Default {
            $AppName = $Record.interfrom.application.displayname
            $AppId = $Record.from.application.id
        }
    }

    If ($Record.body.content.length -gt 100) {
        $Body = $Record.body.content.ToString().Substring(0,100)
    } else {
        $Body = $Record.body.content.ToString()
    }
    
    $AutoGeneratedFlag = $False
    # This section checks for some of the fingerprints that indicate that the interaction is automatic rather than user-generated
    Switch ($AppName) {
        "Copilot in Outlook" {
            $AppName = "Outlook"
            If ($Body -like "*VisualTheming*" -or $Body -like "*data:image;base64*") {
                $AutoGeneratedFlag = $True
            }
        }
        "Copilot in Word" {
            If ($Body -like "*[AutoGenerated]*") {
                $AutoGeneratedFlag = $True  
            }
        }
    }

    $ReportLine = [pscustomobject]@{
        User            = $User.UserPrincipalName
        Timestamp       = $Timestamp
        'Copilot App'   = $AppName
        AppId           = $AppId
        Contexts        = ($Record.contexts.displayName -join ", ")
        InteractionType = $Record.interactionType
        ThreadId        = $Record.sessionid
        Body            = $Body
        Attachments     = ($Record.attachments.name -join ", ")
        Mentions        = ($Record.mentions.name -join ", ")
        Links           = ($Record.Links.LinkUrl -join ", ")
        AutoGenerated   = $AutoGeneratedFlag
    }
    $Report.Add($ReportLine)
}

$Report | Out-GridView -Title 'Copilot Interactions for $User'

# Some basic computations
$NumberOfAutomaticInteractions = $Report | Where-Object { $_.AutoGenerated -eq $True } | Measure-Object | Select-Object -ExpandProperty Count
$UserInteractions = $Report | Where-Object {$_.InteractionType -eq "userPrompt"} | Measure-Object | Select-Object -ExpandProperty Count
$CopilotResponses = $Report.Count - ($UserInteractions + $NumberOfAutomaticInteractions)
$PercentCopilotResponses = ($CopilotResponses/$Report.Count).ToString("P")
$PercentAutomaticInteractions = ($NumberOfAutomaticInteractions/$Report.Count).ToString("P")
$PerecentUserInteractions = ($UserInteractions/$Report.Count).ToString("P")

Write-Host ""
Write-Host ("Copilot interactions for {0} betweeen {1} and {2}" -f $User.DisplayName, $StartDateForReport, $EndDateForReport)
$Report | Group-Object 'Copilot App' | Select-Object Name, Count | Sort-Object Count -Descending | Format-Table

Write-Host ("{0} of the {1} interactions are automatic ({2})" -f $NumberOfAutomaticInteractions, $Report.Count, $PercentAutomaticInteractions)
Write-Host ("{0} of the interactions are user prompts ({1})" -f $UserInteractions, $PerecentUserInteractions)
Write-Host ("{0} of the interactions are Copilot responses to user prompts" -f $PercentCopilotResponses)

# Generate reports
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    $ExcelTitle = ("Copilot interactions for {0} between {1} and {2}" -f $User.DisplayName, $StartDateForReport, $EndDateForReport)
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $OutputXLSXFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\CopilotInteractions.xlsx"
    If (Test-Path $OutputXLSXFile) {
        Remove-Item $OutputXLSXFile -ErrorAction SilentlyContinue
    }
    $Report | Export-Excel -Path $OutputXLSXFile -WorksheetName "Copilot Interactions" -Title $ExcelTitle -TitleBold -TableName "CopilotInteractions" 
} Else {
    $OutputCSVFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\CopilotInteractions.csv"
    $Report | Export-Csv -Path $OutputCSVFile -NoTypeInformation -Encoding Utf8
}
  
If ($ExcelGenerated) {
    Write-Host ("An Excel worksheet containing the report data is available in {0}" -f $OutputXLSXFile)
} Else {
    Write-Host ("A CSV file containing the report data is available in {0}" -f $OutputCSVFile)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.