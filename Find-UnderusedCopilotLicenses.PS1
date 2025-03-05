# Find-UnderusedCopilotLicenses.PS1
# A script to check users with Microsfot 365 Copilot licenses who might not be using the features as they should
# And if we find any underused licenses, we can give them to someone else...
# V1.0 5-Nov-2024
# V1.1 1-Feb-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-UnderusedCopilotLicenses.PS1

# Connect to Microsoft Graph
If (!(Get-MgContext).Account) {
    Write-Host "Connecting to Microsoft Graph..."
    Connect-MgGraph -NoWelcome -Scopes Reports.Read.All, ReportSettings.ReadWrite.All, User.ReadWrite.All
    # Reports.Read.All is needed to fetch usage data
    # ReportSettings.ReadWrite.All is needed to change the tenant settings to allow access to unobfuscated usage data
    # User.ReadWrite.All is needed to read license data for user accounts and to remove licenses from accounts. Also to read sign-in data for users.
}   

# Define the score that marks a user as underusing Microsoft 365 Copilot
[double]$MicrosoftCopilotScore = 30

# Sku Id for the Microsoft 365 Copilot license
[guid]$CopilotSKUId = "639dec6b-bb19-468b-871c-c5c441c4b0cb"

Write-Host "Scanning for user accounts with Microsoft 365 Copilot licenses..."
[array]$Users = Get-MgUser -Filter "usertype eq 'Member' and assignedLicenses/any(s:s/skuId eq $CopilotSkuId)" `
    -ConsistencyLevel Eventual -CountVariable Licenses -All -Sort 'displayName' `
    -Property Id, displayName, signInActivity, userPrincipalName -PageSize 999

If (!$Users) {
    Write-Host "No users with Microsoft 365 Copilot licenses found"
    Break
} Else {
    Write-Host ("{0} users with Microsoft 365 Copilot licenses found" -f $Users.Count)
}

$ConcealedNames = $false
# Make sure that we can fetch usage data that isn't obfuscated
Write-Host "Checking tenant settings for usage data obfuscation..."
If ((Get-MgAdminReportSetting).DisplayConcealedNames -eq $true) {
    $Parameters = @{ displayConcealedNames = $false }
    Write-Host "Switching tenant settings to allow access to unobfuscated usage data..."
    Update-MgAdminReportSetting -BodyParameter $Parameters
    $ConcealedNames = $true
}

# Fetch usage data for Copilot
Write-Host "Fetching Microsoft 365 Copilot usage data..."
$Uri = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='D90')"
[array]$SearchRecords = Invoke-GraphRequest -Uri $Uri -Method Get 
If (!($SearchRecords)) {
    Write-Host "No usage data found for Microsoft 365 Copilot"
    Break
}

# Store the fetched usage data in an array
[array]$UsageData = $SearchRecords.value

# Check do we have more usage data records to fetch and fetch more if a nextlink is available
$NextLink = $SearchRecords.'@Odata.NextLink'
While ($null -ne $NextLink) {
    $SearchRecords = $null
    [array]$SearchRecords = Invoke-MgGraphRequest -Uri $NextLink -Method GET 
    $UsageData += $SearchRecords.value
    Write-Host ("{0} usage data records fetched so far..." -f $UsageData.count)
    $NextLink = $SearchRecords.'@odata.NextLink' 
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
    If ($ScoreApps -gt 0) { 
        [double]$UserScore = ($Score / $ScoreApps)
    } Else {
        [double]$UserScore = 0
    }

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
[array]$UnderusedCopilot = $CopilotReport | Where-Object {$_.'Overall Score' -gt $MicrosoftCopilotScore -or $_.'Overall Score' -eq 0}
# If there are no underused Copilot users, say so - and if we have, give the administrator the chance to remove the licenses
If (!($UnderusedCopilot)) {
    Write-Host "No users found to be underusing an assigned Microsoft 365 Copilot license"
} Else {
    Clear-Host
    $LicenseReport = [System.Collections.Generic.List[Object]]::new()
    Write-Host ("The following {0} users are underusing their assigned Microsoft 365 Copilot license" -f $UnderusedCopilot.Count)
    $UnderusedCopilot | Sort-Object {$_.'Overall Score' -as [double]} | Select-Object User, UserPrincipalName, 'Number active apps', 'Overall Score' | Format-Table -AutoSize
    [string]$Decision = Read-Host "Do you want to remove the Microsoft 365 Copilot licenses from these users"
    If ($Decision.Substring(0,1).toUpper() -eq "Y") {
        ForEach ($User in $UnderusedCopilot) {
            # Check that the user still has a Copilot license...      
            $UserLicenseData = $User = Get-MgUser -Userid $User.UserPrincipalName -Property Id, displayName, userPrincipalName, assignedLicenses, licenseAssignmentStates
            If ($CopilotSKUId -notin $UserLicenseData.assignedLicenses.skuId) {
                Write-Host ("The {0} account does not have a Microsoft 365 Copilot license" -f $UserLicenseData.displayName)
                Continue
            }
            # Direct assigned license or group-assigned license?
            [array]$CopilotLicense = $User.LicenseAssignmentStates | Where-Object {$_.skuId -eq $CopilotSkuId}
            If ($null -eq $CopilotLicense[0].assignedByGroup) {
                # Process the removal of a direct-assigned license
                Try {
                    Write-Host ("Removing direct-assigned Microsoft 365 Copilot license from {0}" -f $UserLicenseData.displayName) -ForegroundColor Yellow
                    Set-MgUserLicense -UserId $UserLicenseData.Id -AddLicenses @{} -RemoveLicenses @($CopilotSKUId) -ErrorAction Stop | Out-Null
                    $LicenseReportLine = = [PSCustomObject][Ordered]@{ 
                        UserPrincipalName   = $UserLicenseData.UserPrincipalName
                        User                = $UserLicenseData.displayName
                        Action              = "Removed direct assigned Copilot license"
                        SkuId               = $CopilotSKUId
                        Timestamp           = Get-Date -format s
                    }
                    $LicenseReport.Add($LicenseReportLine)
                } Catch {
                    Write-Host ("Failed to remove Microsoft 365 Copilot license from {0}: {1}" -f $UserLicenseData.displayName, $_.Exception.Message) -ForegroundColor Red
                }
            } Else {
                # Process the removal of a group-assigned license
                Write-Host ("Removing group-assigned Microsoft 365 Copilot license from {0}" -f $UserLicenseData.displayName) -ForegroundColor Yellow
                $GroupId = $CopilotLicense[0].assignedByGroup
                Try {
                    Remove-MgGroupMemberDirectoryObjectByRef -DirectoryObjectId $UserLicenseData.Id -GroupId $GroupId -ErrorAction Stop
                    $LicenseReportLine = [PSCustomObject][Ordered]@{ 
                        UserPrincipalName   = $UserLicenseData.UserPrincipalName
                        User                = $UserLicenseData.displayName
                        Action              = ("Removed group assigned Copilot license from {0}" -f $GroupId)
                        SkuId               = $CopilotSKUId
                        Timestamp           = Get-Date -format s
                    }
                    $LicenseReport.Add($LicenseReportLine)
                } Catch {
                    Write-Host ("Failed to remove Microsoft 365 Copilot license for {0} from group {1}: {2}" -f $UserLicenseData.displayName, $GroupId, $_.Exception.Message) -ForegroundColor Red
                }
            }
            
        }
        Write-Host ("{0} Microsoft 365 Copilot licenses removed" -f $LicenseReport.Count)
    } Else {
        Write-Host "No Microsoft 365 Copilot licenses removed"
    }
}

If ($LicenseReport) {
    Write-Host ""
    Write-Host "License removal report"
    $LicenseReport | Select-Object Timestamp, User, UserPrincipalName, Action | Sort-Object Timestamp | Format-Table -AutoSize
}

Write-Host "Generating report..."
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Underused Copilot Licenses.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    }
    $UnderusedCopilot | Export-Excel -Path $ExcelOutputFile -WorksheetName "Copilot License Report" -Title ("Underused Copilot License Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "UnderusedCopilot" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Underused Copilot License.CSV"
    $UnderusedCopilot | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}
 
If ($ExcelGenerated) {
    Write-Host ("An Excel report of underused Microsoft 365 Copilot licenses is available in {0}" -f $ExcelOutputFile)
} Else {    
    Write-Host ("A CSV report of underused Microsoft 365 Copilot licenses is available in {0}" -f $CSVOutputFile)
}  

# Reset tenant obfuscation settings to True if we switched the setting earlier
If ((Get-MgAdminReportSetting).DisplayConcealedNames -eq $false -and $ConcealedNames -eq $true) {
    Write-Host "Resetting tenant settings to obfuscate usage data..."
    $Parameters = @{ displayConcealedNames = $True }
    Update-MgAdminReportSetting -BodyParameter $Parameters
}
 
# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.