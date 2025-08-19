# Find-UnderusedCopilotLicenseswithAudit.PS1
# A script to check users with Microsfot 365 Copilot licenses who might not be using the features as they should
# And if we find any underused licenses, we can give them to someone else...
# This is a version of Find-UnderusedCopilotLicenses.PS1 that uses the Microsoft Graph to fetch audit data to add 
# more information to what we know about Copilot interactions in different apps.
# V1.0 1-Mar-2025
# V1.1 5-Mar-2025   Use Graph API requests to submit and fetch user data because of issue with cmdlets
#                   See https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3199
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-UnderusedCopilotLicensesWithAudit.PS1

# Connect to Microsoft Graph
If (!(Get-MgContext).Account) {
    Write-Host "Connecting to Microsoft Graph..."
    Connect-MgGraph -NoWelcome -Scopes Reports.Read.All, ReportSettings.ReadWrite.All, User.ReadWrite.All, AuditLog.Read.All
    # Reports.Read.All is needed to fetch usage data
    # AuditLog.Read.All needed to fetch audit data (users must also hold a suitable Compliance role like Audit Manager or Audit Reader role)
    # ReportSettings.ReadWrite.All is needed to change the tenant settings to allow access to unobfuscated usage data
    # User.ReadWrite.All is needed to read license data for user accounts and to remove licenses from accounts. Also to read sign-in data for users.
}   
function Get-UserScore {
    param (
        [int]$Score,
        [int]$ScoreApps,
        [int]$TotalInteractions
    )
    if ($ScoreApps -gt 0) { 
        [double]$UserScore = (($Score / $ScoreApps) - ($TotalInteractions / 10))
    } else {
        [double]$UserScore = 0
    }
    return $UserScore
}

Disconnect-MgGraph # Remove any existing session
Connect-MgGraph -Scopes AuditLog.Read.All, User.ReadWrite.All, Reports.Read.All, ReportSettings.ReadWrite.All

# Define the score that marks a user as underusing Microsoft 365 Copilot
[double]$MicrosoftCopilotScore = 30

# Sku Id for the Microsoft 365 Copilot license
[guid]$CopilotSKUId = "639dec6b-bb19-468b-871c-c5c441c4b0cb"

Write-Host "Scanning for user accounts with Microsoft 365 Copilot licenses..."
[array]$Users = Get-MgUser -Filter "assignedLicenses/any(s:s/skuId eq $CopilotSkuId)" `
    -All -Sort 'displayName' -Property Id, displayName, signInActivity, userPrincipalName -PageSize 500

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

If ($UsageData) {
    Write-Host ("{0} Microsoft 365 Copilot usage records fetched" -f $UsageData.Count)
    # Get the date of the usage data
    [datetime]$ReportRefreshDate =  $UsageData[0].'reportRefreshDate'
} Else {
    Write-Host "No Microsoft 365 Copilot usage data found"
    Break
}

Write-Host "Fetching audit data for Copilot interactions over the last 30 days..."
Write-Host "This might take some time, depending on the amount of data to be processed."
Write-Host "An audit job is submitted to run in the background. When complete, we will fetch the audit records."
Write-Host "Please wait..."

# Set the parameters for the audit query. We're looking for audit records for roughly the same period as covered
# by the usage data
Set-MgRequestContext -MaxRetry 10 -RetryDelay 15| Out-Null
$AuditJobName = ("Copilot Interactions audit job created at {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm'))
$EndDate = (Get-Date $ReportRefreshDate).AddHours(23)
$StartDate = (Get-Date $EndDate).AddDays(-30)
$AuditQueryStart = (Get-Date $StartDate -format s)
$AuditQueryEnd = (Get-Date $EndDate -format s)
[array]$AuditQueryOperations = "CopilotInteraction"
$AuditQueryParameters = @{}
#$AuditQueryParameters.Add("@odata.type","#microsoft.graph.security.auditLogQuery")
$AuditQueryParameters.Add("displayName", $AuditJobName)
$AuditQueryParameters.Add("OperationFilters", $AuditQueryOperations)
$AuditQueryParameters.Add("filterStartDateTime", $AuditQueryStart)
$AuditQueryParameters.Add("filterEndDateTime", $AuditQueryEnd)

# Submit the audit query
# Too many problems with this beta cmdlet so we'll use the Graph API instead
# $AuditJob = New-MgBetaSecurityAuditLogQuery -BodyParameter $AuditQueryParameters
$Uri = "https://graph.microsoft.com/beta/security/auditLog/queries"
$AuditJob = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $AuditQueryParameters

# Check the audit query status every 20 seconds until it completes
[int]$i = 1
[int]$SleepSeconds = 20
$SearchFinished = $false; [int]$SecondsElapsed = 20
Write-Host "Checking audit query status..."
Start-Sleep -Seconds 30
# This cmdlet is not working...
#$AuditQueryStatus = Get-MgBetaSecurityAuditLogQuery -AuditLogQueryId $AuditJob.Id
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}" -f $AuditJob.id)
$AuditQueryStatus = Invoke-MgGraphRequest -Uri $Uri -Method Get

While ($SearchFinished -eq $false) {
    $i++
    Write-Host ("Waiting for audit search to complete. Check {0} after {1} seconds. Current state {2}" -f $i, $SecondsElapsed, $AuditQueryStatus.status)
    If ($AuditQueryStatus.status -eq 'succeeded') {
        $SearchFinished = $true
    } Else {
        Start-Sleep -Seconds $SleepSeconds
        $SecondsElapsed = $SecondsElapsed + $SleepSeconds
        # $AuditQueryStatus = Get-MgBetaSecurityAuditLogQuery -AuditLogQueryId $AuditJob.Id
        $AuditQueryStatus = Invoke-MgGraphRequest -Uri $Uri -Method Get
    }
}

# Fetch the audit records returned by the query
# This cmdlet isn't working either
# [array]$AuditRecords = Get-MgBetaSecurityAuditLogQueryRecord -AuditLogQueryId $AuditJob.Id -All -PageSize 999
$AuditRecords = [System.Collections.Generic.List[string]]::new()
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}/records?`$top=999" -f $AuditJob.Id)
[array]$AuditSearchRecords = Invoke-MgGraphRequest -Uri $Uri -Method GET
[array]$AuditRecords = $AuditSearchRecords.value

$NextLink = $AuditSearchRecords.'@Odata.NextLink'
While ($null -ne $NextLink) {
    $AuditSearchRecords = $null
    [array]$AuditSearchRecords = Invoke-MgGraphRequest -Uri $NextLink -Method GET 
    $AuditRecords += $AuditSearchRecords.value
    Write-Host ("{0} audit records fetched so far..." -f $AuditRecords.count)
    $NextLink = $AuditSearchRecords.'@odata.NextLink' 
}

Write-Host ("Audit query {0} returned {1} records" -f $AuditJobName, $AuditRecords.Count)
$AuditRecords = $AuditRecords | Sort-Object CreatedDateTime -Descending

$CopilotUserAuditReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $AuditRecords) {
    $AuditData = $Rec.AuditData
    $CopilotApp = 'Copilot for Microsoft 365'; $Context = $null; $CopilotLocation = $null

    $CopilotApp = $AuditData.copiloteventdata.apphost
    If ($null -eq $CopilotApp) {  
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
            "TeamsMeeting" {
                $CopilotApp = "Teams"
                $CopilotLocation = "Teams meeting"
            }
            "StreamVideo" {
                $CopilotApp = "Stream"
                $CopilotLocation = "Stream video player"
            }
            Default {
                $CopilotApp = "Copilot for Microsoft 365"
                $CopilotLocation = "Unknown"
            }
        }
    }

    If ($Auditdata.copiloteventdata.contexts.id -like "*https://teams.microsoft.com/*") {
        $CopilotApp = "Teams"
    } ElseIf ($AuditData.CopiloteventData.AppHost -eq "bizchat" -or $AuditData.CopiloteventData.AppHost -eq "Office") {
        $CopilotApp = "Copilot for Microsoft 365 Chat"
        $CopilotLocation = "Chat"
    }

    If ($Auditdata.CopilotEventData.contexts.id) {
        $Context = $Auditdata.CopilotEventData.contexts.id
    } ElseIf ($Auditdata.CopilotEventData.threadid) {
        $Context = $Auditdata.CopilotEventData.threadid
       # $CopilotApp = "Teams"
    }

    If ($Auditdata.copiloteventdata.contexts.id -like "*/sites/*") {
        $CopilotLocation = "SharePoint Online"
    } ElseIf ($Auditdata.copiloteventdata.contexts.id -like "*https://teams.microsoft.com/*") {
        $CopilotLocation = "Teams"
        If ($Auditdata.copiloteventdata.contexts.id -like "*ctx=channel*") {
            $CopilotLocation = "Teams Channel"
        } Else {
            $CopilotLocation = "Teams Chat"
        }
    } ElseIf ($Auditdata.copiloteventdata.contexts.id -like "*/personal/*") {
        $CopilotLocation = "OneDrive for Business"
    } 
    If ($CopilotApp -eq "Outlook") {
        $CopilotLocation = "Mailbox"
    }
    If ($CopilotApp -eq 'Loop') {
        $CopilotLocation = "SharePoint Embedded"
    }
       
    # Make sure that we report the resources used by Copilot and the action (like read) used to access the resource
    [array]$AccessedResources = $AuditData.copiloteventdata.accessedResources.name | Sort-Object -Unique
    [string]$AccessedResources = $AccessedResources -join ", "
    [array]$AccessedResourceLocations = $AuditData.copiloteventdata.accessedResources.id | Sort-Object -Unique
    [string]$AccessedResourceLocations = $AccessedResourceLocations -join ", "
    [array]$AccessedResourceActions = $AuditData.copiloteventdata.accessedResources.action | Sort-Object -Unique
    [string]$AccessedResourceActions = $AccessedResourceActions -join ", "
    $TimeStamp = Get-Date $Rec.CreatedDateTime -format "dd-MMM-yyyy HH:mm:ss"

    $ReportLine = [PSCustomObject][Ordered]@{
        TimeStamp                       = $TimeStamp
        User                            = $Rec.userPrincipalName
        UserId                          = $Rec.userId
        App                             = $CopilotApp
        'Resource Location'             = $CopilotLocation 
        'App context'                   = $Context   
    }
    $CopilotUserAuditReport.Add($ReportLine)
}

# Find the set of users who actually used Copilot over the last 30 days
[array]$UsersOfCopilot = $CopilotUserAuditReport | Sort-Object UserId -Unique | Select-Object User, UserId
$CopilotUserAuditData =  [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $UsersOfCopilot) {
    $CopilotChatInteractions = $CopilotUserAuditReport | Where-Object {($_.UserId -eq $User.UserId) -and ($_.App -eq 'Copilot for Microsoft 365 Chat')} | Measure-Object | Select-Object -ExpandProperty Count
    $CopilotExcelInteractions = $CopilotUserAuditReport | Where-Object {($_.UserId -eq $User.UserId) -and ($_.App -eq 'Excel')} | Measure-Object | Select-Object -ExpandProperty Count
    $CopilotLoopInteractions = $CopilotUserAuditReport | Where-Object {($_.UserId -eq $User.UserId) -and ($_.App -eq 'Loop')} | Measure-Object | Select-Object -ExpandProperty Count
    $CopilotOneNoteInteractions = $CopilotUserAuditReport | Where-Object {($_.UserId -eq $User.UserId) -and ($_.App -eq 'OneNote')} | Measure-Object | Select-Object -ExpandProperty Count
    $CopilotOutlookInteractions = $CopilotUserAuditReport | Where-Object {($_.UserId -eq $User.UserId) -and ($_.App -eq 'Outlook')} | Measure-Object | Select-Object -ExpandProperty Count
    $CopilotPowerPointInteractions = $CopilotUserAuditReport | Where-Object {($_.UserId -eq $User.UserId) -and ($_.App -eq 'PowerPoint')} | Measure-Object | Select-Object -ExpandProperty Count
    $CopilotWordInteractions = $CopilotUserAuditReport | Where-Object {($_.UserId -eq $User.UserId) -and ($_.App -eq 'Word')} | Measure-Object | Select-Object -ExpandProperty Count
    $TotalInteractions = $CopilotUserAuditReport | Where-Object {$_.UserId -eq $User.UserId} | Measure-Object | Select-Object -ExpandProperty Count
    $OtherInteractions = $TotalInteractions - ($CopilotChatInteractions + $CopilotExcelInteractions + $CopilotLoopInteractions + $CopilotOneNoteInteractions + $CopilotOutlookInteractions + $CopilotPowerPointInteractions + $CopilotWordInteractions)

    $CopilotUserAuditReportLine = [PSCustomObject][Ordered]@{
        User                    = $User.User
        UserId                  = $User.UserId
        ChatInteractions        = $CopilotChatInteractions
        ExcelInteractions       = $CopilotExcelInteractions
        LoopInteractions        = $CopilotLoopInteractions
        OutlookInteractions     = $CopilotOutlookInteractions
        PowerPointInteractions  = $CopilotPowerPointInteractions
        WordInteractions        = $CopilotWordInteractions
        OtherInteractions       = $OtherInteractions
        TotalInteractions       = $TotalInteractions
    }
    $CopilotUserAuditData.Add($CopilotUserAuditReportLine)
}

Write-Host "Generating Copilot license usage assessment report..."
$CopilotReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    $LastSignIn = $null; $ScoreApps = 7; $UserScore = 0
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
        $OneNoteDays = (New-TimeSpan -Start $OneNoteDate -End $ReportRefreshDate).Days
    } Else {
        $OneNoteDate = 'Not used'
        $OneNoteDays = 0
        $ScoreApps = $ScoreApps -1
    }
    #Teams
    If (-not ([string]::IsNullOrEmpty($UserData.microsoftTeamsCopilotLastActivityDate))) {
        $TeamsDate = Get-Date $UserData.microsoftTeamsCopilotLastActivityDate -format 'dd-MMM-yyyy'
        $TeamsDays = (New-TimeSpan -Start $TeamsDate -End $ReportRefreshDate).Days
    } Else {
        $TeamsDate = 'Not used'
        $TeamsDays = 0
        $ScoreApps = $ScoreApps -1
    }
    #Outlook
    If (-not ([string]::IsNullOrEmpty($UserData.outlookCopilotLastActivityDate))) {
        $OutlookDate = Get-Date $UserData.outlookCopilotLastActivityDate -format 'dd-MMM-yyyy'
        $OutlookDays = (New-TimeSpan -Start $OutlookDate -End $ReportRefreshDate).Days
    } Else {
        $OutlookDate = 'Not used'
        $OutlookDays = 0
        $ScoreApps = $ScoreApps -1
    }
    # Word
    If (-not ([string]::IsNullOrEmpty($UserData.wordCopilotLastActivityDate))) {
        $WordDate = Get-Date $UserData.wordCopilotLastActivityDate -format 'dd-MMM-yyyy'
        $WordDays = (New-TimeSpan -Start $WordDate -End $ReportRefreshDate).Days
    } Else {
        $WordDate = 'Not used'
        $WordDays = 0
        $ScoreApps = $ScoreApps -1
    }
    # Microsoft 365 Chat
    If (-not ([string]::IsNullOrEmpty($UserData.copilotChatLastActivityDate))) {
        $ChatDate = Get-Date $UserData.copilotChatLastActivityDate -format 'dd-MMM-yyyy'
        $ChatDays = (New-TimeSpan -Start $ChatDate -End $ReportRefreshDate).Days
    } Else {
        $ChatDate = 'Not used'
        $ChatDays = 0
        $ScoreApps = $ScoreApps -1
    }
    # Excel
    If (-not ([string]::IsNullOrEmpty($UserData.excelCopilotLastActivityDate))) {
        $ExcelDate = Get-Date $UserData.excelCopilotLastActivityDate -format 'dd-MMM-yyyy'
        $ExcelDays = (New-TimeSpan -Start $ExcelDate -End $ReportRefreshDate).Days
    } Else {
        $ExcelDate = 'Not used'
        $ExcelDays = 0
        $ScoreApps = $ScoreApps -1
    }
    # PowerPoint
    If (-not ([string]::IsNullOrEmpty($UserData.powerPointCopilotLastActivityDate))) {
        $PowerPointDate = Get-Date $UserData.powerPointCopilotLastActivityDate -format 'dd-MMM-yyyy'
        $PowerPointDays = (New-TimeSpan -Start $PowerPointDate -End $ReportRefreshDate).Days
    } Else {
        $PowerPointDate = 'Not used'
        $PowerPointDays = 0
        $ScoreApps = $ScoreApps -1
    }
    
    # Retrieve audit data if available
    [array]$UserAuditData = $CopilotUserAuditData | Where-Object {$_.User -eq $User.UserPrincipalName}
    If ($UserAuditData) {
        Write-Host ("Copilot audit data found for {0}" -f $User.DisplayName)
    }

    # Compute a score for the user
    $Score = $OutlookDays + $TeamsDays + $OneNoteDays + $ExcelDays + $WordDays + $ChatDays + $PowerPointDays
    $UserScore = Get-UserScore -Score $Score -ScoreApps $ScoreApps -TotalInteractions $UserAuditData.TotalInteractions

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
        'Chat interactions'     = $UserAuditData.ChatInteractions
        'Excel interactions'    = $UserAuditData.ExcelInteractions
        'Loop interactions'     = $UserAuditData.LoopInteractions
        'OneNote interactions'  = $UserAuditData.OneNoteInteractions
        'Outlook interactions'  = $UserAuditData.OutlookInteractions
        'PowerPoint interactions'= $UserAuditData.PowerPointInteractions
        'Word interactions'     = $UserAuditData.WordInteractions
        'Total interactions'    = $UserAuditData.TotalInteractions
        'Overall Score'         = "{0:N2}" -f $UserScore
    }
    $CopilotReport.Add($ReportLine)
}

# Extract the set of users who should be considered as underusing Copilot
[array]$UnderusedCopilot = $CopilotReport | Where-Object {
    $OS = [double]$_.'Overall Score'
    $OS -gt $MicrosoftCopilotScore -or $_.'Overall Score' -as [double] -eq 0}

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
