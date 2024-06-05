# Find-CandidateCopilotUsers.PS1
# A script showing how to use the Microsoft Graph PowerShell SDK to identify users who might be suitable to 
# receive Copilot for Microsoft 365 licenses
# V1.0 15-Feb-2024
# https://github.com/12Knocksinna/Office365itpros/blob/master/Find-CandidateCopilotUsers.PS1
# See https://practical365.com/copilot-for-microsoft-365-licenses-decision/ for an article explaining how to use
# this script.

Connect-MgGraph -NoWelcome -Scopes Reports.Read.All, ReportSettings.ReadWrite.All, User.Read.All

$TempDownloadFile = "c:\temp\x.csv"
$ObscureFlag = $false
$CSVOutputFile = "C:\temp\CopilotUserAnalysis.CSV"

$Uri = "https://graph.microsoft.com/beta/admin/reportSettings"
# Check if the tenant has obscured real names for reports - see https://office365itpros.com/2022/09/09/graph-usage-report-tips/
$DisplaySettings = Invoke-MgGraphRequest -Method Get -Uri $Uri
If ($DisplaySettings['displayConcealedNames'] -eq $true) { # data is obscured, so let's reset it to allow the report to run
   $ObscureFlag = $true
   Write-Host "Setting tenant data concealment for reports to False" -foregroundcolor red
   Invoke-MgGraphRequest -Method PATCH -Uri $Uri -Body (@{"displayConcealedNames"= $false} | ConvertTo-Json) 
}

# This command finds user accounts with an eligible Copilot base license. It specifies the SKU identifers for
# Office 365 E3, Office 365 E5, Microsoft 365 E3, amd Microsoft 365 E5. There are other variants of these SKUs
# for government and academic use, so it's important to pass the SKU identifiers in use within the tenant
Write-Host "Finding user accounts to check..."
[array]$Users = Get-MgUser -Filter "assignedLicenses/any(s:s/skuId eq 6fd2c87f-b296-42f0-b197-1e91e994b900) `
    or assignedLicenses/any(s:s/skuid eq c7df2760-2c81-4ef7-b578-5b5392b571df) `
    or assignedLicenses/any(s:s/skuid eq 05e9a617-0261-4cee-bb44-138d3ef5d965) `
    or assignedLicenses/any(s:s/skuid eq 06ebc4ee-1bb5-47dd-8120-11324bc54e06)" `
    -ConsistencyLevel Eventual -CountVariable Licenses -All -Sort 'displayName' `
    -Property Id, displayName, signInActivity, userPrincipalName -PageSize 999

Write-Host "Fetching usage data for Teams, Exchange, and OneDrive for Business..."
# Get Teams user activity detail for the last 30 days
$Uri = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D30')"
Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath $TempDownloadFile
[array]$TeamsUserData = Import-CSV $TempDownloadFile
# Get Email activity data
$Uri = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='D30')"
Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath $TempDownloadFile
[array]$EmailUserData = Import-CSV $TempDownloadFile
# Get OneDrive data 
$Uri = "https://graph.microsoft.com/v1.0/reports/getOneDriveActivityUserDetail(period='D30')"
Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath $TempDownloadFile
[array]$OneDriveUserData = Import-CSV $TempDownloadFile
# Get Apps detail
$Uri = "https://graph.microsoft.com/v1.0/reports/getM365AppUserDetail(period='D30')"
Invoke-mgGraphRequest -Uri $Uri -Method GET -OutputFilePath $TempDownloadFile
[array]$AppsUserData = Import-CSV $TempDownloadFile

Write-Host "Analyzing information..."
$CopilotReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    Write-Host ("Checking activity for {0}..." -f $User.displayName)
    $UserTeamsData = $TeamsUserData | Where-Object 'User Principal Name' -match $User.UserPrincipalName
    $UserOneDriveData = $OneDriveUserData | Where-Object 'User Principal Name' -match $User.UserPrincipalName
    $UserEmailData = $EmailUserData | Where-Object 'User Principal Name' -match $User.UserPrincipalName
    $UserAppsData = $AppsUserData | Where-Object 'User Principal Name' -match $User.UserPrincipalName

    $LastSignInDate = $null
    $DaysSinceLastSignIn = $null
    If ($User.signInActivity.LastSignInDateTime) {
        $LastSignInDate = Get-Date $User.signInActivity.LastSignInDateTime -format 'dd-MMM-yyyy'
        $DaysSinceLastSignIn = (New-TimeSpan $User.signInActivity.LastSignInDateTime).Days
    }
    [int]$Fails = 0; [int]$Points = 0
    # Test 1 - Has the user signed in withim the last 15 days
    If ($DaysSinceLastSignIn -le 15) {
        $Test1 = "Pass"
        $Points = 25
    } Else {
        $Test1 = "Fail"
        $Fails++
    }
    # Test 2 - 20 or more Teams chat and attending at least 5 Teams meetings
    If (([int]$UserTeamsData.'Team Chat Message Count' -ge 20) -and ([int]$UserTeamsData.'Meetings Attended Count' -ge 5)) {
        $Test2 = "Pass"
        $Points = $Points + 20
    } Else {
        $Test2 = "Fail"
        $Fails++
    }
    # Test 3 Sends at least 3 messages daily (22 work days over 30) and receives 6 messages daily
    If (([int]$UserEmailData.'Send Count' -ge 66) -and ([int]$UserEmailData.'Receive Count' -ge 132)) {
        $Test3 = "Pass" 
        $Points = $Points + 20
    } Else {
        $Test3 = "Fail"
        $Fails++
    }
    # Test 4 - OneDrive - views or edits at least 10 files
    If ([int]$UserOneDriveData.'Viewed Or Edited File Count' -ge 10) {
        $Test4 = "Pass"
        $Points = $Points + 20
    } Else {
        $Test4 = "Fail"
        $Fails++
    }
    # Test 5 - Apps - user must use Outlook, Word, and Excel
    If ($UserAppsData.Outlook -eq "Yes" -and $UserAppsData.Word -eq "Yes" -and $UserAppsData.Excel -eq "Yes") {
        $Test5 = "Pass"
        $Points = $Points + 15
    } Else {
        $Test5 = "Fail"
        $Fails++
    }

    If ($Points -ge 85) {
        $CopilotApproved = "Approved"
    } Else {
        $CopilotApproved = "Not eligible"
    }

    $ReportLine = [PSCustomObject][Ordered]@{ 
        User                = $User.displayName
        UPN                 = $User.UserPrincipalName
        'Last Signin'       = $LastSignInDate
        'Days since signin' = $DaysSinceLastSignIn
        Points              = $Points
        'Test 1: Sign in'   = $Test1
        'Test 2: Teams'     = $Test2
        'Test 3: Email'     = $Test3
        'Test 4: OneDrive'  = $Test4
        'Test 5: Apps'      = $Test5
        'Copilot approved'  = $CopilotApproved
    }
    $CopilotReport.Add($ReportLine)
}

[array]$CopilotRecommendations = $CopilotReport | Where-Object {$_.'Copilot Approved' -eq 'Approved'} | Select-Object User, UPN
$CopilotReport | Export-CSV -NoTypeInformation $CSVOutputFile
Clear-Host
Write-Host ""
Write-Host ("Based on analysis of user activity and apps, there are {0} users recommended" -f $CopilotRecommendations.count)
Write-Host ("to receive Copilot for Microsoft 365 licenses. Details are available in this file: {0}" -f $CSVOutputFile)
Write-Host ""
$CopilotRecommendations

# Switch the tenant report obscure data setting back if necessary
If ($ObscureFlag -eq $True) {
    Write-Host "Resetting tenant data concealment for reports to True" -foregroundcolor red
    Invoke-MgGraphRequest -Method PATCH -Uri 'https://graph.microsoft.com/beta/admin/reportSettings' `
     -Body (@{"displayConcealedNames"= $true} | ConvertTo-Json) 
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.

