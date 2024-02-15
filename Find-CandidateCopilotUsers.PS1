# Find-CandidateCopilotUsers.PS1
# A script showing how to use the Microsoft Graph PowerShell SDK to identify users who might be suitable to 
# receive Copilot for Microsoft 365 licenses
# V1.0 15-Feb-2024

Connect-MgGraph -NoWelcome -Scopes Reports.Read.All, Directory.Read.All

$TempDownloadFile = "c:\temp\x.csv"
$ObscureFlag = $false

# Check if the tenant has obscured real names for reports - see https://office365itpros.com/2022/09/09/graph-usage-report-tips/
$DisplaySettings = Invoke-MgGraphRequest -Method Get -Uri 'https://graph.microsoft.com/beta/admin/reportSettings'
If ($DisplaySettings['displayConcealedNames'] -eq $true) { # data is obscured, so let's reset it to allow the report to run
   $ObscureFlag = $true
   Write-Host "Setting tenant data concealment for reports to False" -foregroundcolor red
   Invoke-MgGraphRequest -Method PATCH -Uri 'https://graph.microsoft.com/beta/admin/reportSettings' `
    -Body (@{"displayConcealedNames"= $false} | ConvertTo-Json) 
}

Write-Host "Finding user accounts to check..."
[array]$Users = Get-MgUser -All -Sort 'displayName' `
    -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -consistencyLevel eventual -CountVariable UsersFound `
    -Property Id, displayName, signInActivity, userPrincipalName

Write-Host "Fetching usage data for Teams, Exchange, and OneDrive for Business..."
# Get Teams user activity detail for the last 180 days
$Uri = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='D180')"
Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath $TempDownloadFile
[array]$TeamsUserData = Import-CSV $TempDownloadFile
# Get Email activity data
$Uri = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='D180')"
Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath $TempDownloadFile
[array]$EmailUserData = Import-CSV $TempDownloadFile
# Get OneDrive data 
$Uri = "https://graph.microsoft.com/v1.0/reports/getOneDriveActivityUserDetail(period='D180')"
Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath $TempDownloadFile
[array]$OneDriveUserData = Import-CSV $TempDownloadFile

$CopilotReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    Write-Host ("Checking activity for {0}..." -f $User.displayName)
    $UserTeamData = $TeamsUserData | Where-Object 'User Principal Name' -match $User.UserPrincipalName
    $UserOneDriveData = $OneDriveUserData | Where-Object 'User Principal Name' -match $User.UserPrincipalName
    $UserEmailData = $EmailUserData | Where-Object 'User Principal Name' -match $User.UserPrincipalName

    $LastSignInDate = $null
    $DaysSinceLastSignIn = $null
    If ($User.signInActivity.LastSignInDateTime) {
        $LastSignInDate = Get-Date $User.signInActivity.LastSignInDateTime -format 'dd-MMM-yyyy'
        $DaysSinceLastSignIn = (New-TimeSpan $User.signInActivity.LastSignInDateTime).Days
    }

    $ReportLine = [PSCustomObject][Ordered]@{ 
        User                = $User.displayName
        UPN                 = $User.UserPrincipalName
        'Last Signin'       = $LastSignInDate
        'Days since signin' = $DaysSinceLastSignIn

    }
    $CopilotReport.Add($ReportLine)
}


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

