# Get-GraphUserStatistics.ps1
# Example of using the Microsoft Graph PowerShell SDK to get usage report data for Entra ID accounts from multiple Microsoft 365 workloads.
# Upgrade from the original script (GetGraphUserStatisticsReport.PS1) that uses Graph API requests...

# Version 2.1 8-Oct-2025    SDK cmdlets used, email of report added, lots more error handling
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Get-GraphUserStatistics.PS1

# Note: Guest user activity is not recorded by the Graph - only tenant accounts are processed
# Original article https://office365itpros.com/2020/09/14/office-365-user-activity-report/ (2020) and then https://office365itpros.com/2022/08/08/microsoft-365-user-activity-2022/ (2022)

# Flag to let script code know if we're running interactively or within Azure Automation. Make sure that the following modules are installed:
# Microsoft.Graph.Authentication
# Microsoft.Graph.Users
# Microsoft.Graph.Users.Actions
# Microsoft.Graph.Reports
# Microsoft.Graph.Identity.DirectoryManagement

$Interactive = $false
$ReportName = "Microsoft 365 User Activity Report V2.1"

# Determine if we're interactive or not
If ([Environment]::UserInteractive) { 
    # We're running interactively...
    Clear-Host
    Write-Host "Script running interactively... connecting to the Graph" -ForegroundColor Yellow
    Connect-MgGraph -NoWelcome -Scopes User.Read.All, Reports.Read.All, ReportSettings.ReadWrite.All, Mail.Send, Organization.Read.All
    $Interactive = $true
    # Email address to use when sending email from interactive session
    $MsgFrom = (Get-MgContext).Account    
} Else { 
    # We're not, so likely in Azure Automation
    Write-Output "Executing the runbook to create the last activity report for user member accounts..." 
    Connect-MgGraph -Identity -NoWelcome
    # Email address to use when sending email from Azure Automation
    $MsgFrom = "no-reply@office365itpros.com"
}

# Set the home domain and base OneDrive domain to use when creating OneDrive URLs
$HomeDomain = (Get-MgOrganization).VerifiedDomains | Where-Object { $_.IsDefault -eq $true } | Select-Object -ExpandProperty Name
$BaseOneDriveDomain = "https://"+$HomeDomain.Split(".")[0]+"-my.sharepoint.com/personal/"

# Check that we have the right permissions - in Azure Automation, we assume that the automation account has the right permissions
If ($Interactive) {
    [string[]]$CurrentScopes = (Get-MgContext).Scopes
    [string[]]$RequiredScopes = @('Reports.Read.All','ReportSettings.ReadWrite.All','User.Read.All','Mail.Send','Organization.Read.All')

    $CheckScopes =[object[]][Linq.Enumerable]::Intersect($RequiredScopes,$CurrentScopes)
    If ($CheckScopes.Count -ne 5) { 
        Write-Host ("To run this script, you need to connect to Microsoft Graph with the following scopes: {0}" -f $RequiredScopes) -ForegroundColor Red
        Disconnect-Graph
        Break
    }
}

$StartTime1 = Get-Date
# Define email address to receive the emailed copy of the report
$DestinationEmailAddress = "someuser@office365itpros.com"

# Check if report data is obfuscated in the tenant. If it is, turn obfuscation off so that we get real data for analysis
$ObfuscationChanged = $false
If ((Get-MgAdminReportSetting).DisplayConcealedNames -eq $True) {
   $Parameters = @{ displayConcealedNames = $False }
   Update-MgAdminReportSetting -BodyParameter $Parameters
   $ObfuscationChanged = $true
}

# Create temporary files to hold usage data downloaded from the Graph
$TeamsDataFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\TeamsData.csv"
If (Test-Path $TeamsDataFile) {
    Remove-Item $TeamsDataFile -ErrorAction SilentlyContinue
}
$OneDriveDataFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\OneDriveData.csv"
If (Test-Path $OneDriveDataFile) {
    Remove-Item $OneDriveDataFile -ErrorAction SilentlyContinue
}
$EmailDataFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\EmailData.csv"
If (Test-Path $EmailDataFile) {
    Remove-Item $EmailDataFile -ErrorAction SilentlyContinue
}
$EmailStorageDataFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\EmailStorageData.csv"
If (Test-Path $EmailStorageDataFile) {
    Remove-Item $EmailStorageDataFile -ErrorAction SilentlyContinue
}
$SPODataFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\SPOData.csv"
If (Test-Path $SPODataFile) {   
    Remove-Item $SPODataFile -ErrorAction SilentlyContinue
}
$YammerDataFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\YammerData.csv"
If (Test-Path $YammerDataFile) {    
    Remove-Item $YammerDataFile -ErrorAction SilentlyContinue
}

# Download the usage data and store each set in an array
# Fetch Teams user activity data
Write-Output "Fetching Teams user activity data for the last 180 days..."
[array]$TeamsUserData = $null
$TeamsUserReportsURI = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D180')"
Try {
    Invoke-MgGraphRequest -Uri $TeamsUserReportsURI -Method Get -OutputFilePath $TeamsDataFile
    [array]$TeamsUserData = Import-Csv -Path $TeamsDataFile 
} Catch {
    Write-Output "Error fetching Teams user activity data: $_" 
}   

# Fetch OneDrive user activity data
Write-Output "Fetching OneDrive user activity data for the last 180 days..."
[array]$OneDriveData = $null
$OneDriveUsageUri = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D180')"
Try {   
    Invoke-MgGraphRequest -Uri $OneDriveUsageURI -Method Get -OutputFilePath $OneDriveDataFile
    [array]$OneDriveData = Import-Csv -Path $OneDriveDataFile 
} Catch {
    Write-Output "Error fetching OneDrive user activity data: $_" 
}

# Fetch Exchange Online data
Write-Output "Fetching Exchange email activity data for the last 180 days..."
[array]$EmailData = $null
$EmailReportsUri = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='D180')"
Try {   
    Invoke-MgGraphRequest -Uri $EmailReportsURI -Method Get -OutputFilePath $EmailDataFile
    [array]$EmailData = Import-Csv -Path $EmailDataFile 
} Catch {
    Write-Output "Error fetching Exchange email activity data: $_" 
}

# Fetch mailbox usage data
Write-Output "Fetching Exchange mailbox usage data for the last 180 days..."
$MailboxUsageReportsUri = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D180')"
[array]$MailboxData = $null
Try {   
    Invoke-MgGraphRequest -Uri $MailboxUsageReportsURI -Method Get -OutputFilePath $EmailStorageDataFile
    [array]$MailboxData = Import-Csv -Path $EmailStorageDataFile 
} Catch {
    Write-Output "Error fetching Exchange mailbox usage data: $_"  
}

# Get SharePoint usage data
Write-Host "Fetching SharePoint Online user activity data from the Graph..."
$SPOUsageReportsUri = "https://graph.microsoft.com/v1.0/reports/getSharePointActivityUserDetail(period='D180')"
[array]$SPOData = $null
Try {
    Invoke-MgGraphRequest -Uri $SPOUsageReportsUri -Method Get -OutputFilePath $SPODataFile
    [array]$SPOData = Import-Csv -Path $SPODataFile
} Catch {
    Write-Output "Error fetching SharePoint Online user activity data: $_" 
}

Write-Host "Fetching Viva Engage user activity data from the Graph..."
# Get Viva Engage usage data
$YammerUsageReportsUri = "https://graph.microsoft.com/v1.0/reports/getYammerActivityUserDetail(period='D180')"
[array]$YammerData = $null
Try {
    Invoke-MgGraphRequest -Uri $YammerUsageReportsUri -Method Get -OutputFilePath $YammerDataFile
    [array]$YammerData = Import-Csv -Path $YammerDataFile
} Catch {
    Write-Output "Error fetching Viva Engage user activity data: $_" 
}

# Create the hash table for the usage data fetched for each account
$DataTable = @{}
# Get User sign in data
Write-Output "Fetching user account data from the Graph..."
Try {
    [array]$Users = Get-MgUser -All -PageSize 500 -filter "userType eq 'Member'" `
        -Property DisplayName, UserPrincipalName, signInActivity, Mail, Id, CreatedDateTime, AccountEnabled
} Catch {
    Write-Output "Error fetching user sign-in data: $_" 
}   

# If necessary, reset tenant obfuscation settings to True
If ($ObfuscationChanged) {
    If ((Get-MgAdminReportSetting).DisplayConcealedNames -eq $False) {
       $Parameters = @{ displayConcealedNames = $True }
       Update-MgAdminReportSetting -BodyParameter $Parameters
    }
}

$StartTime2 = Get-Date
Write-Output "Processing activity data fetched from the Graph..."
# Process Teams Data
ForEach ($T in $TeamsUserData) {
    If ([string]::IsNullOrEmpty($T."Last Activity Date")) { 
        $TeamsLastActivity = "No activity"
        $TeamsDaysSinceActive = "N/A" 
    } Else {
        $TeamsLastActivity = Get-Date($T."Last Activity Date") -format "dd-MMM-yyyy" 
        $TeamsDaysSinceActive = (New-TimeSpan($TeamsLastActivity)).Days 
    }
    $ReportLine  = [PSCustomObject] @{          
        TeamsUPN               = $T."User Principal Name"
        TeamsLastActive        = $TeamsLastActivity  
        TeamsDaysSinceActive   = $TeamsDaysSinceActive      
        TeamsReportDate        = Get-Date($T."Report Refresh Date") -format "dd-MMM-yyyy"  
        TeamsLicense           = $T."Assigned Products"
        TeamsChannelChats      = $T."Team Chat Message Count"
        TeamsPrivateChats      = $T."Private Chat Message Count"
        TeamsCalls             = $T."Call Count"
        TeamsMeetings          = $T."Meeting Count"
        TeamsRecordType        = "Teams"}
    $DataTable[$T."User Principal Name"] = $ReportLine
} 

# Process Exchange Data
ForEach ($E in $EmailData) {
    $ExoDaysSinceActive = $Null
    If ([string]::IsNullOrEmpty($E."Last Activity Date")) { 
        $ExoLastActivity = "No activity"
        $ExoDaysSinceActive = "N/A" 
    } Else {
        $ExoLastActivity = Get-Date($E."Last Activity Date") -format "dd-MMM-yyyy"
        $ExoDaysSinceActive = (New-TimeSpan($ExoLastActivity)).Days 
    }
    $ReportLine  = [PSCustomObject] @{          
        ExoUPN                = $E."User Principal Name"
        ExoDisplayName        = $E."Display Name"
        ExoLastActive         = $ExoLastActivity   
        ExoDaysSinceActive    = $ExoDaysSinceActive    
        ExoReportDate         = Get-Date($E."Report Refresh Date") -format "dd-MMM-yyyy"  
        ExoSendCount          = [int]$E."Send Count"
        ExoReadCount          = [int]$E."Read Count"
        ExoReceiveCount       = [int]$E."Receive Count"
        ExoIsDeleted          = $E."Is Deleted"
        ExoRecordType         = "Exchange Activity"}
    [Array]$ExistingData = $DataTable[$E."User Principal Name"] 
    [Array]$NewData = $ExistingData + $ReportLine
    $DataTable[$E."User Principal Name"] = $NewData 
} 
  
ForEach ($M in $MailboxData) {
    If ([string]::IsNullOrEmpty($M."Last Activity Date")) { 
        $ExoLastActivity = "No activity" 
    } Else {
        $ExoLastActivity = Get-Date($M."Last Activity Date") -format "dd-MMM-yyyy"
        $ExoDaysSinceActive = (New-TimeSpan($ExoLastActivity)).Days 
    }
    $ReportLine  = [PSCustomObject] @{          
        MbxUPN                = $M."User Principal Name"
        MbxDisplayName        = $M."Display Name"
        MbxLastActive         = $ExoLastActivity 
        MbxDaysSinceActive    = $ExoDaysSinceActive          
        MbxReportDate         = Get-Date($M."Report Refresh Date") -format "dd-MMM-yyyy"  
        MbxQuotaUsed          = [Math]::Round($M."Storage Used (Byte)"/1GB,2) 
        MbxQuotaPercentUsed   = ($M.'Storage Used (Byte)'/$M.'Prohibit Send/Receive Quota (Byte)').toString('P')
        MbxItems              = [int]$M."Item Count"
        MbxRecordType         = "Exchange Storage"
    }
    [Array]$ExistingData = $DataTable[$M."User Principal Name"] 
    [Array]$NewData = $ExistingData + $ReportLine
    $DataTable[$M."User Principal Name"] = $NewData 
} 

# SharePoint data
ForEach ($S in $SPOData) {
    If ([string]::IsNullOrEmpty($S."Last Activity Date")) { 
        $SPOLastActivity = "No activity"
        $SPODaysSinceActive = "N/A" 
    } Else {
        $SPOLastActivity = Get-Date($S."Last Activity Date") -format "dd-MMM-yyyy"
        $SPODaysSinceActive = (New-TimeSpan ($SPOLastActivity)).Days 
    }
    $ReportLine  = [PSCustomObject] @{          
        SPOUPN              = $S."User Principal Name"
        SPOLastActive       = $SPOLastActivity    
        SPODaysSinceActive  = $SPODaysSinceActive 
        SPOViewedEdited     = [int]$S."Viewed or Edited File Count"     
        SPOSyncedFileCount  = [int]$S."Synced File Count"
        SPOSharedExt        = [int]$S."Shared Externally File Count"
        SPOSharedInt        = [int]$S."Shared Internally File Count"
        SPOVisitedPages     = [int]$S."Visited Page Count" 
        SPORecordType       = "SharePoint Usage"
    }
    [Array]$ExistingData = $DataTable[$S."User Principal Name"] 
    [Array]$NewData = $ExistingData + $ReportLine
    $DataTable[$S."User Principal Name"] = $NewData 
}  

# OneDrive for Business data
ForEach ($O in $OneDriveData) {
    $OneDriveLastActivity = $Null
    If ([string]::IsNullOrEmpty($O."Last Activity Date")) { 
        $OneDriveLastActivity = "No activity"
        $OneDriveDaysSinceActive = "N/A" 
    } Else {
        $OneDriveLastActivity = Get-Date($O."Last Activity Date") -format "dd-MMM-yyyy" 
        $OneDriveDaysSinceActive = (New-TimeSpan($OneDriveLastActivity)).Days 
    }
    If ($Interactive) {
        $ODSite = $O."Owner Principal Name" -replace "@","_" -replace "\.","_"
        $ODSite = $BaseOneDriveDomain + $ODSite
    } Else {
        # Use the Graph to find the OneDrive site URL and return in the same format as the interactive version
        # This code is overkill, and it works fine in app-only mode but the Get-MgUserDefaultDrive cmdlet fails in an Azure Automation runbook. The fix
        # is to simply use the $BaseOneDriveDomain + the UPN with substitutions as done in the interactive code above.
        $User = $O."Owner Principal Name".trim()
        $UserId = Get-MgUser -UserId $User | Select-Object -ExpandProperty Id
        $ODSite = ((Get-MgUserDefaultDrive -UserId $UserId).WebUrl).Split("/Documents")[0]
    }
    $ReportLine  = [PSCustomObject] @{          
        ODUPN               = $O."Owner Principal Name"
        ODDisplayName       = $O."Owner Display Name"
        ODLastActive        = $OneDriveLastActivity    
        ODDaysSinceActive   = $OneDriveDaysSinceActive    
        ODSite              = $ODSite
        ODFileCount         = [int]$O."File Count"
        ODStorageUsed       = [Math]::Round($O."Storage Used (Byte)"/1GB,4) 
        ODQuota             = [Math]::Round($O."Storage Allocated (Byte)"/1GB,2) 
        ODRecordType        = "OneDrive Storage"
    }
    [Array]$ExistingData = $DataTable[$O."Owner Principal Name"] 
    [Array]$NewData = $ExistingData + $ReportLine
    $DataTable[$O."Owner Principal Name"] = $NewData 
}

# Yammer Data
ForEach ($Y in $YammerUsage) {  
    If ([string]::IsNullOrEmpty($Y."Last Activity Date")) { 
        $YammerLastActivity = "No activity" 
        $YammerDaysSinceActive = "N/A" 
    } Else {
        $YammerLastActivity = Get-Date($Y."Last Activity Date") -format "dd-MMM-yyyy" 
        $YammerDaysSinceActive = (New-TimeSpan ($YammerLastActivity)).Days 
    }
    $ReportLine  = [PSCustomObject] @{          
        YUPN             = $Y."User Principal Name"
        YDisplayName     = $Y."Display Name"
        YLastActive      = $YammerLastActivity      
        YDaysSinceActive = $YammerDaysSinceActive   
        YPostedCount     = [int]$Y."Posted Count"
        YReadCount       = [int]$Y."Read Count"
        YLikedCount      = [int]$Y."Liked Count"
        YRecordType      = "Yammer Usage"
    }
    [Array]$ExistingData = $DataTable[$Y."User Principal Name"] 
    [Array]$NewData = $ExistingData + $ReportLine
    $DataTable[$Y."User Principal Name"] = $NewData 
}

$StartTime3 = Get-Date
# Set up progress bar
If ($Interactive) {
    $ProgressDelta = 100/($Users.Count); $PercentComplete = 0; $UserNumber = 0
}

$OutData = [System.Collections.Generic.List[Object]]::new() # Create merged output file

# Process each user to extract Exchange, Teams, OneDrive, SharePoint, and Yammer (Viva Engage) statistics for their activity
ForEach ($User in $Users) {
    $U = $User.UserPrincipalName
    $UserNumber++
    If ($Interactive) {
        $CurrentStatus = $U + " ["+ $UserNumber +"/" + $Users.Count + "]"
        Write-Progress -Activity "Extracting information for user" -Status $CurrentStatus -PercentComplete $PercentComplete
        $PercentComplete += $ProgressDelta
    }
   
    $UserData = $DataTable[$U]  # Extract data for the user - everything is in a single keyed access to the hash table

    # Process Exchange Data
    [string]$ExoUPN = (Out-String -InputObject $UserData.ExoUPN).Trim()
    [string]$ExoLastActive = (Out-String -InputObject $UserData.ExoLastActive).Trim()
    If ([string]::IsNullOrEmpty($ExoUPN) -or $ExoLastActive -eq "No Activity") {
        $ExoDaysSinceActive  = "N/A"
        $ExoLastActive = "No Activity" 
    } Else {
        [string]$ExoLastActive = (Out-String -InputObject $UserData.ExoLastActive).Trim()
        [string]$ExoDaysSinceActive = (Out-String -InputObject $UserData.ExoDaysSinceActive).Trim() 
    }
 
    # Parse OneDrive for Business usage data 
    [string]$ODUPN = (Out-String -InputObject $UserData.ODUPN).Trim()
    [string]$ODLastActive = (Out-String -InputObject $UserData.ODLastActive).Trim()  # Possibility of a second OneDrive account for some users.
    If (($ODLastActive -Like "*No Activity*") -or ([string]::IsNullOrEmpty($ODLastActive))) {
        $ODLastActive = "No Activity"
    } # this is a hack until I figure out a better way to handle the situation
    If ($null -eq [string]::IsNullOrEmpty($ODUPN) -or $ODLastActive -eq "No Activity") {
        [string]$ODDaysSinceActive  = "N/A"
        [string]$ODLastActive = "No Activity"
        $ODFiles            = 0
        $ODStorage          = 0
        $ODQuota            = 1024 
    } Else {
        [string]$ODDaysSinceActive = (Out-String -InputObject $UserData.ODDaysSinceActive).Trim()
        [string]$ODLastActive = (Out-String -InputObject $UserData.ODLastActive).Trim()
        [string]$ODFiles = (Out-String -InputObject $UserData.ODFileCount).Trim()
        [string]$ODStorage = (Out-String -InputObject $UserData.ODStorageUsed).Trim()
        [string]$ODQuota = (Out-String -InputObject $UserData.ODQuota).Trim()  
    }

    # Parse Yammer usage data; Yammer isn't used everywhere, so make sure that we record zero data 
    [string]$YUPN = (Out-String -InputObject $UserData.YUPN).Trim()
    [string]$YammerLastActive = (Out-String -InputObject $UserData.YLastActive).Trim()
    If (([string]::IsNullOrEmpty($YUPN) -or ($YammerLastActive -eq "No Activity"))) { 
        [string]$YammerLastActive = "No Activity"  
        [string]$YammerDaysSinceActive  = "N/A" 
        $YammerPosts             = 0
        $YammerReads             = 0
        $YammerLikes             = 0 
    } Else {
        $YammerDaysSinceActive = (Out-String -InputObject $UserData.YDaysSinceActive).Trim()
        $YammerPosts = (Out-String -InputObject $UserData.YPostedCount).Trim()
        $YammerReads = (Out-String -InputObject $UserData.YReadCount).Trim()
        $YammerLikes = (Out-String -InputObject $UserData.YLikedCount).Trim() 
    }
  
    If ($UserData.TeamsDaysSinceActive -gt 0) {
        [string]$TeamsDaysSinceActive = (Out-String -InputObject $UserData.TeamsDaysSinceActive).Trim()
        [string]$TeamsLastActive = (Out-String -InputObject $UserData.TeamsLastActive).Trim() 
    } Else { 
        [string]$TeamsDaysSinceActive = "N/A"
        [string]$TeamsLastActive = "No Activity" 
    }
 
    If ($UserData.SPODaysSinceActive -gt 0) {
        [string]$SPODaysSinceActive = (Out-String -InputObject $UserData.SPODaysSinceActive).Trim()
        [string]$SPOLastActive = (Out-String -InputObject $UserData.SPOLastActive).Trim() 
    } Else { 
        [string]$SPODaysSinceActive = "N/A"
        [string]$SPOLastActive = "No Activity" 
    }
 
    # Fetch the sign in data if available
    $LastAccountSignIn = $Null; $DaysSinceSignIn = 0
    If ($null -eq $User.SignInActivity.LastSuccessfulSignInDateTime) { 
        $LastAccountSignIn = "No sign in data found"; $DaysSinceSignIn = "N/A"
    } Else { 
        $LastAccountSignIn = Get-Date($User.SignInActivity.LastSuccessfulSignInDateTime) -format 'dd-MMM-yyyy HH:mm'
        $DaysSinceSignIn = (New-TimeSpan($User.SignInActivity.LastSuccessfulSignInDateTime)).Days 
    }
   
    # Figure out if the account is used
    [int]$ExoDays = 365; [int]$TeamsDays = 365; [int]$SPODays = 365; [int]$ODDays = 365; [int]$YammerDays = 365

    # Base is 2 if someuse uses the five workloads because the Graph is usually 2 days behind, but we have some N/A values for days used
    If ($ExoDaysSinceActive -ne "N/A") {$ExoDays = $ExoDaysSinceActive -as [int]}
    If ($TeamsDaysSinceActive -eq "N/A") {$TeamsDays = 365} Else {$TeamsDays = $TeamsDaysSinceActive -as [int]}
    If ($SPODaysSinceActive -eq "N/A") {$SPODays = 365} Else {$SPODays = $SPODaysSinceActive -as [int]}  
    If ($ODDaysSinceActive -eq "N/A") {$ODDays = 365} Else {$ODDays = $ODDaysSinceActive -as [int]} 
    If ($YammerDaysSinceActive -eq "N/A") {$YammerDays = 365} Else {$YammerDays = $YammerDaysSinceActive -as [int]}
   
    # Average days per workload used...
    $AverageDaysSinceUse = [Math]::Round((($ExoDays + $TeamsDays + $SPODays + $ODDays + $YammerDays)/5),2)
    # Add icons to account status
    Switch ($AverageDaysSinceUse) { # Figure out if account is used
        ({$PSItem -le 12})                         { $AccountStatus = "🟩💪 Heavy usage" }
        ({$PSItem -ge 13 -and $PSItem -le 50} )    { $AccountStatus = "🟨🙂 Moderate usage" }   
        ({$PSItem -ge 51 -and $PSItem -le 120} )   { $AccountStatus = "🟧😐 Poor usage" }
        ({$PSItem -ge 121 -and $PSItem -le 300 } ) { $AccountStatus = "🟦🧐 Review account"  }
        default                                    { $AccountStatus = "🟥🚫 Account unused" }
    } # End Switch

    # And an override if someone has been active in just one workload in the last 14 days
    [int]$DaysCheck = 14 # Set this to your chosen value if you want to use a different period.
    If (($ExoDays -le $DaysCheck) -or ($TeamsDays -le $DaysCheck) -or ($SPODays -le $DaysCheck) -or ($ODDays -le $DaysCheck) -or ($YammerDays -le $DaysCheck)) {
        $AccountStatus = "🟢💪 Account in use"
    }

    If ((![string]::IsNullOrEmpty($ExoUPN))) {
    # Build a line for the report file with the collected data for all workloads and write it to the list
    $OutLine  = [PSCustomObject] @{          
        UPN                                     = $U
        DisplayName                             = (Out-String -InputObject $UserData.ExoDisplayName).Trim()
        'Active Status'                         = $AccountStatus
        'Last successful sign in'               = $LastAccountSignIn
        'Days since sign in'                    = $DaysSinceSignIn 
        'Last active in EXO'                    = $ExoLastActive  
        'Days since active in EXO'              = $ExoDaysSinceActive  
        'EXO Quota used'                        = (Out-String -InputObject $UserData.MbxQuotaUsed).Trim() + " GB"
        'EXO Quota % used'                      = (Out-String -InputObject $UserData.MbxQuotaPercentUsed).Trim()
        'Items in EXO mailbox'                  = (Out-String -InputObject $UserData.MbxItems).Trim()
        'EXO messages sent'                     = (Out-String -InputObject $UserData.ExoSendCount).Trim()
        'EXO messages read'                     = (Out-String -InputObject $UserData.ExoReadCount).Trim()
        'EXO messages received'                 = (Out-String -InputObject $UserData.ExoReceiveCount).Trim()
        'Last active in Teams'                  = $TeamsLastActive
        'Days since active in Teams'            = $TeamsDays 
        'Number of channel messages'            = (Out-String -InputObject $UserData.TeamsChannelChats).Trim()
        'Number of chat messages'               = (Out-String -InputObject $UserData.TeamsPrivateChats).Trim()
        'Number of meetings'                    = (Out-String -InputObject $UserData.TeamsMeetings).Trim()
        'Number of Teams calls'                 = (Out-String -InputObject $UserData.TeamsCalls).Trim()
        'Last active in SPO'                    = $SPOLastActive
        'Days since last active in SPO'         = $SPODays 
        'Number of SPO files edited'            = (Out-String -InputObject $UserData.SPOViewedEdited).Trim()
        'Number of SPO files synced'            = (Out-String -InputObject $UserData.SPOSyncedFileCount).Trim()
        'Number of SPO files shared (Ext)'      = (Out-String -InputObject $UserData.SPOSharedExt).Trim()
        'Number of SPO files shared (int)'      = (Out-String -InputObject $UserData.SPOSharedInt).Trim()
        'Number of SPO pages visited'           = (Out-String -InputObject $UserData.SPOVisitedPages).Trim()
        'Last active in OneDrive'               = $ODLastActive
        'Days since last active in OneDrive'    = $ODDaysSinceActive
        'Number of OneDrive files'              = $ODFiles
        'OneDrive storage used'                 = $ODStorage + " GB"
        'OneDrive quota assigned'               = $ODQuota + " GB"
        'OneDrive quota used'                   = ($ODStorage/$ODQuota).ToString("P")
        'Last active in Viva Engage'            = $YammerLastActive
        'Days since last active in Viva Engage' = $YammerDaysSinceActive
        'Number of Viva Engage posts'           = $YammerPosts
        'Number of Viva Engage reads'           = $YammerReads
        'Number of Viva Engage reactions'       = $YammerLikes
        'Assigned licenses'                     = (Out-String -InputObject $UserData.TeamsLicense).Trim()
        'OneDrive site URL'                     = (Out-String -InputObject $UserData.ODSite).Trim()
        'EXO data from'                         = (Out-String -InputObject $UserData.ExoReportDate).Trim()
        'Teams data from'                       = (Out-String -InputObject $UserData.TeamsReportDate).Trim()
        'Average days since use'                = $AverageDaysSinceUse
        'Account created date'                  = Get-Date $User.CreatedDateTime -Format 'dd-MMM-yyyy HH:mm'
        'Account enabled'                       = $User.AccountEnabled
        }
    $OutData.Add($OutLine)
    }
} #End processing user data

#Clear-Host
$StartTime4 = Get-Date
$GraphTime = $StartTime2 - $StartTime1
$PrepTime = $StartTime3 - $StartTime2
$ReportTime = $StartTime4 - $StartTime3
$ScriptTime = $StartTime4 - $StartTime1
$AccountsPerMinute = [math]::Round(($Outdata.count/($ScriptTime.TotalSeconds/60)),2)
$GraphElapsed = "{0}:{1:D2}" -f $GraphTime.Minutes, $GraphTime.Seconds
$PrepElapsed = "{0}:{1:D2}" -f $Preptime.Minutes, $Preptime.Seconds
$ReportElapsed = "{0}:{1:D2}" -f $ReportTime.Minutes, $ReportTime.Seconds
$ScriptElapsed = "{0}:{1:D2}" -f $ScriptTime.Minutes, $ScriptTime.Seconds


Write-Output " "
Write-Output ("Statistics for {0}" -f $ReportName) 
Write-Output "------------------------------------------------------"
Write-Output ("Time to fetch data from Microsoft Graph: {0}" -f $GraphElapsed)
Write-Output ("Time to prepare date for processing:     {0}" -f $PrepElapsed)
Write-Output ("Time to create report from data:         {0}" -f $ReportElapsed)
Write-Output ("Total time for script:                   {0}" -f $ScriptElapsed)
Write-Output ("Total accounts processed:                {0}" -f $Outdata.count)
Write-Output ("Account processing rate per minute:     {0}" -f $AccountsPerMinute)   
Write-Output " "

If ($Interactive) {
    $OutData | Sort-Object {$_.ExoLastActive -as [DateTime]} -Descending | Out-GridView -Title "User Activity Summary"
}

$OutData = $OutData | Sort-Object UPN 
# And generate an output file
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $true
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\AccountActivitySummary.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    } 
    $OutData | Export-Excel -Path $ExcelOutputFile -WorksheetName "Inactive Guests" -Title ("Inactive Guests Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "InactiveGuests" 
    $AttachmentFile = $ExcelOutputFile
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveGuests.CSV"
    $OutData | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
    $AttachmentFile = $CSVOutputFile
}

If ($ExcelGenerated) {
    Write-Output ("Excel worksheet output written to {0}" -f $ExcelOutputFile)
} Else {
    Write-Output ("CSV output file written to {0}" -f $CSVOutputFile)
}   

# Prepare to send the email
# Encount the output file to an email
$EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($AttachmentFile))

$MsgAttachments = @(
    @{
        '@odata.type' = '#microsoft.graph.fileAttachment'
        Name = (Split-Path $AttachmentFile -Leaf)
        ContentBytes = $EncodedAttachmentFile
        ContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
)

# Build the array of a single TO recipient detailed in a hash table - change this to the appropriate recipient for your tenant
$ToRecipient = @{}
$ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
[array]$MsgTo = $ToRecipient
# Define the message subject
$MsgSubject = "Important: User Activity Summary Report"
# Create the HTML content
$HtmlMsg = "</body></html><p>The output file for the <b>User Activity Summary Report</b> is attached to this message. Please review the information at your convenience</p>"
# Construct the message body 	
$MsgBody = @{}
$MsgBody.Add('Content', "$($HtmlMsg)")
$MsgBody.Add('ContentType','html')
# Build the parameters to submit the message
$Message = @{}
$Message.Add('subject', $MsgSubject)
$Message.Add('toRecipients', $MsgTo)
$Message.Add('body', $MsgBody)
$Message.Add("attachments", $MsgAttachments)

$EmailParameters = @{}
$EmailParameters.Add('message', $Message)
$EmailParameters.Add('saveToSentItems', $true)
$EmailParameters.Add('isDeliveryReceiptRequested', $true)

# Send the message
Try {
    Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters -ErrorAction Stop
    Write-Output ("Inactive guest account report emailed to {0}" -f $ToRecipient.emailAddress.address)
} Catch {
    Write-Output "Unable to send email"
    Write-Output $_.Exception.Message
}

Write-Output "All done"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.