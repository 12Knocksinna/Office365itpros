# Report-CopilotUsageWithUserDetail.PS1
# Simple script to show how to combine the usage data for Microsoft 365 Copilot with user details.

# V1.0 1-May-2025
# V1.1 09-Oct-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-CopllotUsageWithUserDetail.PS1

# Connect to the Graph
Connect-MgGraph -Scopes "Reports.Read.All", "User.Read.All", "ReportSettings.ReadWrite.All" 

$ConcealedNames = $false
Write-Host "Checking tenant settings for usage data obfuscation..."
If ((Get-MgAdminReportSetting).DisplayConcealedNames -eq $true) {
    $Parameters = @{ displayConcealedNames = $false }
    Write-Host "Switching tenant settings to allow access to unobfuscated usage data..."
    Update-MgAdminReportSetting -BodyParameter $Parameters
    $ConcealedNames = $true
}

Write-Host "Finding Microsoft 365 Copilot usage data..." -ForegroundColor Green
$Uri = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='D180')"
[array]$SearchRecords = Invoke-GraphRequest -Uri $Uri -Method Get -OutputType PSObject
If (!($SearchRecords)) {
    Write-Host "No usage data found for Microsoft 365 Copilot"
    Break
}

# Store the fetched usage data in an array
[array]$CopilotUsageData = $SearchRecords.Value

# Check do we have more usage data records to fetch and fetch more if a nextlink is available
$NextLink = $SearchRecords.'@Odata.NextLink'
While ($null -ne $NextLink) {
    $SearchRecords = $null
    [array]$SearchRecords = Invoke-MgGraphRequest -Uri $NextLink -Method GET -OutputType PsObject
    $CopilotUsageData += $SearchRecords.Value
    Write-Host ("{0} usage data records fetched so far..." -f $UsageData.count)
    $NextLink = $SearchRecords.'@odata.NextLink' 
}

If ($CopilotUsageData) {
    Write-Host ("{0} Microsoft 365 Copilot usage records fetched" -f $CopilotUsageData.Count)
    # Get the date of the usage data
    [datetime]$ReportRefreshDate =  $CopilotUsageData[0].'reportRefreshDate'
} Else {
    Write-Host "No Microsoft 365 Copilot usage data found"
    Break
}

# Find users with Microsoft 365 Copilot full and trial licenses
Write-Host "Finding users with a Microsoft 365 Copilot license..." -ForegroundColor Green
[array]$UserData = Get-MgUser -ConsistencyLevel eventual -CountVariable Var -Filter "assignedLicenses/any(s:s/skuId eq 639dec6b-bb19-468b-871c-c5c441c4b0cb) `
    or assignedLicenses/any(s:s/skuid eq ea2d19f9-23bc-4f7f-9c12-a677b26a8e2a)" `
    -Property Id, DisplayName, Department, JobTitle, UserPrincipalName, Mail -PageSize 500 -All
If ($UserData.Count -eq 0) {
    Write-Host "No users found with a Microsoft 365 Copilot license" -ForegroundColor Red
    break
} Else {
    Write-Host ("{0} users found with a Microsoft 365 Copilot license" -f $UserData.Count) -ForegroundColor Green
}

Write-Host "Creating report..." -ForegroundColor Green
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($User in $UserData) {
    $CopilotInfo = $CopilotUsageData | Where-Object { $_.userPrincipalName -eq $User.userPrincipalName }
    $OneNoteDate = $null; $OutlookDate = $null; $LastActivityDate = $null; $ExcelDate = $null; $WordDate = $null; $TeamsDate = $null; $BizChatDate = $null; $LoopDate = $null; $PowerPointDate = $null
    If ($CopilotInfo) {
        If ($CopilotInfo.oneNoteCopilotLastActivityDate) {
            $OneNoteDate = Get-Date ($CopilotInfo.oneNoteCopilotLastActivityDate) -format 'dd-MMM-yyyy'
        }
        If ($CopilotInfo.outlookCopilotLastActivityDate) {
            $OutlookDate = Get-Date ($CopilotInfo.outlookCopilotLastActivityDate) -format 'dd-MMM-yyyy'
        }
        If ($CopilotInfo.excelCopilotLastActivityDate) {
            $ExcelDate = Get-Date ($CopilotInfo.excelCopilotLastActivityDate) -format 'dd-MMM-yyyy'
        }
        If ($CopilotInfo.lastActivityDate) {
            $LastActivityDate = Get-Date ($CopilotInfo.lastActivityDate) -format 'dd-MMM-yyyy'
        } 
        If ($CopilotInfo.microsoftTeamsCopilotLastActivityDate) {
            $TeamsDate = Get-Date ($CopilotInfo.microsoftTeamsCopilotLastActivityDate) -format 'dd-MMM-yyyy'
        }
        If ($CopilotInfo.copilotChatLastActivityDate) {
            $BizChatDate = Get-Date ($CopilotInfo.copilotChatLastActivityDate) -format 'dd-MMM-yyyy'
        }
        If ($CopilotInfo.wordCopilotLastActivityDate) {
            $WordDate = Get-Date ($CopilotInfo.wordCopilotLastActivityDate) -format 'dd-MMM-yyyy'
        }
        If ($CopilotInfo.powerPointCopilotLastActivityDate) {
            $PowerPointDate = Get-Date ($CopilotInfo.powerPointCopilotLastActivityDate) -format 'dd-MMM-yyyy'
        }
        If ($CopilotInfo.loopCopilotLastActivityDate) {
            $LoopDate = Get-Date ($CopilotInfo.loopCopilotLastActivityDate) -format 'dd-MMM-yyyy'
        }
        
        $ReportItem = [PSCustomObject]@{
            UserPrincipalName   = $User.userPrincipalName
            DisplayName         = $User.displayName
            JobTitle            = $User.jobTitle
            Mail                = $User.mail
            Department          = $User.department
            OneNote             = $OneNoteDate
            PowerPoint          = $PowerPointDate
            Outlook             = $OutlookDate
            Excel               = $ExcelDate
            Word                = $WordDate
            Teams               = $TeamsDate
            BizChat             = $BizChatDate
            Loop                = $LoopDate
            LastUsedDate        = $LastActivityDate
        }
        $Report.Add($ReportItem)
    }
}

Write-Host ("All done. {0} users have Microsoft 365 Copilot licenses and usage data was found for {1} users." -f $UserData.Count, `
    ($Report | Where-Object {$_.LastUsedDate -ne $null}).count) -ForegroundColor Green
$Report | Sort-Object DisplayName | Out-GridView -Title ("Microsoft 365 Copilot Usage Report using data refreshed on {0}" -f (Get-Date $ReportRefreshDate -format 'dd-MMM-yyyy')) 

# Reset tenant obfuscation settings to True if we switched the setting earlier
If ((Get-MgAdminReportSetting).DisplayConcealedNames -eq $false -and $ConcealedNames -eq $true) {
    Write-Host "Resetting tenant settings to obfuscate usage data..."
    $Parameters = @{ displayConcealedNames = $True }
    Update-MgAdminReportSetting -BodyParameter $Parameters
}
 
Write-Host ""
Write-Host "Here are the departments where people who don't use Copilot work:" -ForegroundColor Green
$GroupedReport = $Report | Where-Object { $_.LastUsedDate -eq $null } |
    Group-Object -Property Department | ForEach-Object {
        [PSCustomObject]@{
            Department = $_.Name
            UserCount  = $_.Group.Count
        }
    }

$GroupedReport | Sort-Object -Property Department | Format-Table -AutoSize

Write-Host "Generating report..."
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Copilot Usage.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    }
    $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Copilot Usage Report" -Title ("Copilot Usage Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) `
        -TitleBold -TableName "CopilotUsage"
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Copilot Usage.CSV"
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}
 
If ($ExcelGenerated) {
    Write-Host ("An Excel report of Copilot usage with account data is available in {0}" -f $ExcelOutputFile)
} Else {    
    Write-Host ("A CSV report of Copilot usage with account data is available in {0}" -f $CSVOutputFile)
}  

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.