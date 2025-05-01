# Report-CopilotUsageWithUserDetail.PS1
# Simple script to show how to combine the usage data for Microsoft 365 Copilot with user details.

# V1.0 1-May-2025
# GitHub link:

# Connect to the Graph
Connect-MgGraph -Scopes "Reports.Read.All", "User.Read.All" -ErrorAction Stop -NoWrite

Write-Host "Finding Microsoft 365 Copilot usage data..." -ForegroundColor Green
$Uri = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='D90')"
[array]$SearchRecords = Invoke-GraphRequest -Uri $Uri -Method Get
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
    [array]$SearchRecords = Invoke-MgGraphRequest -Uri $NextLink -Method GET 
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

Write-Host "Finding users with a Microsoft 365 Copilot license..." -ForegroundColor Green
[array]$UserData = Get-MgUser -All -Filter "assignedLicenses/any(s:s/skuId eq 639dec6b-bb19-468b-871c-c5c441c4b0cb)" `
    -Property Id, DisplayName, Department, JobTitle, UserPrincipalName, Mail 
If ($UserData.Count -eq 0) {
    Write-Host "No users found with a Microsoft 365 Copilot license" -ForegroundColor Red
    break
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

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.