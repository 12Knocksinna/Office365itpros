# Find-UnderusedCopilotLicenses.PS1
# a script to check users with Microsfot 365 Copilot licenses who might not be using the features as they should
# V1.0 5-Nov-2024
# GitHub link:


# Connect to Microsoft Graph
If (!(Get-MgContext).Account) {
    Write-Host "Connecting to Microsoft Graph..."
    Connect-MgGraph -NoWelcome -Scopes Reports.Read.All, ReportSettings.ReadWrite.All, User.Read.All
}   

Write-Host "Scanning for user accounts with Microsoft 365 Copilot licenses..."
[array]$Users = Get-MgUser -Filter "usertype eq 'Member' and assignedLicenses/any(s:s/skuId eq 639dec6b-bb19-468b-871c-c5c441c4b0cb)" `
    -ConsistencyLevel Eventual -CountVariable Licenses -All -Sort 'displayName' `
    -Property Id, displayName, signInActivity, userPrincipalName -PageSize 999

If (!$Users) {
    Write-Host "No users with Microsoft 365 Copilot licenses found"
    Break
} Else {
    Write-Host ("{0} users with Microsoft 365 Copilot licenses found" -f $Users.Count)
}

# Make sure that we can fetch usage data that isn't obfuscated
If ((Get-MgAdminReportSetting).DisplayConcealedNames -eq $true) {
    $Parameters = @{ displayConcealedNames = $false }
    Update-MgAdminReportSetting -BodyParameter $Parameters
    $ConcealedNames = $false
}
 
# Define the score that marks a user as underusing Microsoft 365 Copilot
[double]$MicrosoftCopilotScore = 30

# Fetch usage data for Copilot
Write-Host "Fetching Microsoft 365 Copilot usage data..."
$Uri = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='D90')"
[array]$UsageData = Invoke-GraphRequest -Uri $Uri -Method Get | Select-Object -ExpandProperty Value
If (!($UsageData)) {
    Write-Host "No usage data found for Microsoft 365 Copilot"
    Break
}

$CopilotReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    $LastSignIn = $null; $ScoreApps = 7
    [array]$UserData = $UsageData | Where-Object {$_.UserPrincipalName -eq $User.UserPrincipalName}
    If (!($UserData)) {
        # can't assess a user if we don't have usage data
        Write-Host ("No Microsoft 365 Copilot usage data found for {0}" -f $User.DisplayName)
        Continue
    }
    If ($User.SignInActivity.LastSuccessfulSignInDateTime) {
        $LastSignIn = $User.SignInActivity.LastSuccessfulSignInDateTime 
    } Else {
        $LastSignIn = $User.SignInactivity.LastSignInDateTime
    }
    If ($null -eq $LastSignIn) {
        $LastSignIn = "Never"
        $DaysSinceSignIn = "N/A"
    } Else {
        # Is it more than 30 days since a sign-in?
        $LastSignIn = Get-Date $LastSignIn -format 'dd-MMM-yyyy HH:mm:ss'
        $DaysSinceSignIn = (New-TimeSpan ($LastSignIn)).Days
    }
    # Check dates of use for the various Copilot features
    # OneNote
    If (-not ([string]::IsNullOrEmpty($UserData.oneNoteCopilotLastActivityDate))) {
        $OneNoteDate = Get-Date $UserData.oneNoteCopilotLastActivityDate -format 'dd-MMM-yyyy'
        $OneNoteDays = (New-TimeSpan $OneNoteDate).Days
    } Else {
        $OneNoteDate = 'Not used'
        $OneNoteDays = 0
        $ScoreApps = $ScoreApps -1
    }
    #Teams
    If (-not ([string]::IsNullOrEmpty($UserData.microsoftTeamsCopilotLastActivityDate))) {
        $TeamsDate = Get-Date $UserData.microsoftTeamsCopilotLastActivityDate -format 'dd-MMM-yyyy'
        $TeamsDays = (New-TimeSpan $TeamsDate).Days
    } Else {
        $TeamsDate = 'Not used'
        $TeamsDays = 0
        $ScoreApps = $ScoreApps -1
    }
    #Outlook
    If (-not ([string]::IsNullOrEmpty($UserData.outlookCopilotLastActivityDate))) {
        $OutlookDate = Get-Date $UserData.outlookCopilotLastActivityDate -format 'dd-MMM-yyyy'
        $OutlookDays = (New-TimeSpan $OutlookDate).Days
    } Else {
        $OutlookDate = 'Not used'
        $OutlookDays = 0
        $ScoreApps = $ScoreApps -1
    }
    # Word
    If (-not ([string]::IsNullOrEmpty($UserData.wordCopilotLastActivityDate))) {
        $WordDate = Get-Date $UserData.wordCopilotLastActivityDate -format 'dd-MMM-yyyy'
        $WordDays = (New-TimeSpan $WordDate).Days
    } Else {
        $WordDate = 'Not used'
        $WordDays = 0
        $ScoreApps = $ScoreApps -1
    }
    # Microsoft 365 Chat
    If (-not ([string]::IsNullOrEmpty($UserData.copilotChatLastActivityDate))) {
        $ChatDate = Get-Date $UserData.copilotChatLastActivityDate -format 'dd-MMM-yyyy'
        $ChatDays = (New-TimeSpan $ChatDate).Days
    } Else {
        $ChatDate = 'Not used'
        $ChatDays = 0
        $ScoreApps = $ScoreApps -1
    }
    # Excel
    If (-not ([string]::IsNullOrEmpty($UserData.excelCopilotLastActivityDate))) {
        $ExcelDate = Get-Date $UserData.excelCopilotLastActivityDate -format 'dd-MMM-yyyy'
        $ExcelDays = (New-TimeSpan $ExcelDate).Days
    } Else {
        $ExcelDate = 'Not used'
        $ExcelDays = 0
        $ScoreApps = $ScoreApps -1
    }
    # PowerPoint
    If (-not ([string]::IsNullOrEmpty($UserData.powerPointCopilotLastActivityDate))) {
        $PowerPointDate = Get-Date $UserData.powerPointCopilotLastActivityDate -format 'dd-MMM-yyyy'
        $PowerPointDays = (New-TimeSpan $PowerPointDate).Days
    } Else {
        $PowerPointDate = 'Not used'
        $PowerPointDays = 0
        $ScoreApps = $ScoreApps -1
    }
    
    # Compute a score for the user
    $Score = $OutlookDays + $TeamsDays + $OneNoteDays + $ExcelDays + $WordDays + $ChatDays + $PowerPointDays
    [double]$UserScore = ($Score / $ScoreApps)

    $ReportLine = [PSCustomObject][Ordered]@{ 
        UserPrincipalName       = $User.UserPrincipalName
        User                    = $User.DisplayName
        'Last sign in'          = $LastSignIn
        'Days since sign in'    = $DaysSinceSignIn
        'Copilot data from'     = Get-Date $UserData.reportRefreshDate -format 'dd-MMM-yyyy'
        'Copilot in Teams'      = $TeamsDate
        'Days since Teams'      = $TeamsDays
        'Copilot in Outlook'    = $OutlookDate
        'Days since Outlook'    = $OutlookDays
        'Copilot in Word'       = $WordDate
        'Days since Word'       = $WordDays
        'Copilot in Chat'       = $ChatDate
        'Days since Chat'       = $ChatDays
        'Copilot in Excel'      = $ExcelDate
        'Days since Excel'      = $ExcelDays
        'Copilot in PowerPoint' = $PowerPointDate
        'Days since PowerPoint' = $PowerPointDays
        'Copilot in OneNote'    = $OneNoteDate
        'Days since OneNote'    = $OneNoteDays
        'Number active apps'    = $ScoreApps
        'Overall Score'         = $UserScore
    }
    $CopilotReport.Add($ReportLine)
}

# Extract the set of users who should be considered as underusing Copilot
[array]$UnderusedCopilot = $CopilotReport | Where-Object {$_.'Overall Score' -gt $MicrosoftCopilotScore}
# If there are no underused Copilot users, say so - and if we have, give the administrator the chance to remove the licenses
If (!($UnderusedCopilot)) {
    Write-Host "No users found to be underusing an assigned Microsoft 365 Copilot license"
} Else {
    [int]$RemovedLicenses = 0
    Write-Host ("The folllowing {0} users are underusing their assigned Microsoft 365 Copilot license" -f $UnderusedCopilot.Count)
    $UnderusedCopilot | Sort-Object {$_.'Overall Score' -as [double]} | Select-Object User, UserPrincipalName, 'Number active apps', 'Overall Score' | Format-Table -AutoSize
    [string]$Decision = Read-Host "Do you want to remove the licenses from these users"
    If ($Decision.Substring(0,1).toUpper() -eq "Y") {
        ForEach ($User in $UnderusedCopilot) {
            Write-Host ("Removing Microsoft 365 Copilot license from {0}" -f $User.User)
            Try {
                Set-MgUserLicense -UserId $User.UserPrincipalName -AddLicenses @{} -RemoveLicenses "639dec6b-bb19-468b-871c-c5c441c4b0cb" -ErrorAction Stop | Out-Null
                $RemovedLicenses++
            } Catch {
                Write-Host ("Failed to remove Microsoft 365 Copilot license from {0}: {1}" -f $User.User, $_.Exception.Message)
            }
        }
        Write-Host ("{0} licenses removed" -f $RemovedLicenses)
    } Else {
        Write-Host "No licenses removed"
    }
}


# Reset tenant obfuscation settings to True if we switched the setting earlier
If ((Get-MgAdminReportSetting).DisplayConcealedNames -eq $false -and $ConcealedNames) {
    $Parameters = @{ displayConcealedNames = $True }
    Update-MgAdminReportSetting -BodyParameter $Parameters
 }
 
# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.