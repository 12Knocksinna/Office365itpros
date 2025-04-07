# Report-FileSharingEvents.ps1
# Report audit events for file sharing events from SharePoint Online and OneDrive for Business whenn sharing occurs with tenant or guest
# accounts. Excludes sharing for SharePoint Embedded and apps that use SharePoint Embedded.

# V1.0 6-Apr-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-FileSharingAuditEvents.PS1

Connect-MgGraph -Scopes "AuditLog.Read.All","User.ReadBasic.All" -ErrorAction Stop

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -Notin $Modules) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -showBanner:$false -UserPrincipalName (Get-MgContext).Account
}

Write-Host "Checking for sensitivity labels..." -ForegroundColor Yellow
# Connect to compliance endpoint to retrieve details of sensitivity labels
Connect-IPPSSession -ShowBanner:$false -UserPrincipalName (Get-MgContext).Account

[array]$Labels = Get-Label -ErrorAction SilentlyContinue
If ($Labels) {
    $LabelsHash = @{}
    ForEach ($Label in $Labels) {
        $LabelsHash.Add([string]$Label.ImmutableId, $Label.DisplayName)
    }
}

Write-Host "Setting up background audit query to find file sharing events..."$AuditJobName = ("Audit job created at {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm'))
$EndDate = (Get-Date).AddHours(1)
$StartDate = (Get-Date $EndDate).AddDays(-180)
$AuditQueryStart = (Get-Date $StartDate -format s)
$AuditQueryEnd = (Get-Date $EndDate -format s)
[array]$AuditOperationFilters = "SharingSet"

$AuditQueryParameters = @{}
$AuditQueryParameters.Add("@odata.type","#microsoft.graph.security.auditLogQuery")
$AuditQueryParameters.Add("displayName", $AuditJobName)
$AuditQueryParameters.Add("OperationFilters", $AuditOperationFilters)
$AuditQueryParameters.Add("filterStartDateTime", $AuditQueryStart)
$AuditQueryParameters.Add("filterEndDateTime", $AuditQueryEnd)

$Uri = "https://graph.microsoft.com/beta/security/auditLog/queries"
$AuditJob = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $AuditQueryParameters

[int]$i = 1
[int]$SleepSeconds = 20
$SearchFinished = $false; [int]$SecondsElapsed = 20
Write-Host "Checking audit query status..."
Start-Sleep -Seconds 30
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
        $AuditQueryStatus = Invoke-MgGraphRequest -Uri $Uri -Method Get
    }
}

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
$AuditRecords = $AuditRecords | Sort-Object {$_.createdDateTime -as [datetime]} -Descending

# Hash table to store user details so we don't have to call the Graph API for each user
$UserDetails = @{}
# Defined the type of target user or group to filter out unwanted records
[array]$TargetTypes = "Member", "Guest"

$AuditReport =  [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $AuditRecords) {

    # Ignore anything in SharePoint Embedded like Loop or Outlook Newsletters
    If ($Rec.AuditData.ObjectId -like "*CSP_*" -or $Rec.AuditData.ObjectId -like "*contentstorage*") { Continue }
    # No interest in sharing done by the app@sharepoint account because these actions are for files like Teams meeting recordings
    If ($Rec.UserPrincipalName -eq 'app@sharepoint') { Continue }
    # No interest in SharePoint system group assigned to allow users access to a single item in a list or library
    If ($Rec.AuditData.TargetUserOrGroupName -like 'Limited Access System Group*') { Continue }
    # And focus solely on sharing links created by users
    If ($Rec.AuditData.TargetUserOrGroupType -notin $TargetTypes) { Continue }

    If ($Rec.AuditData.UserAgent -like "*Teams*") {
        $UserAgent = "Teams"
    } Else {
        $UserAgent = "SharePoint"
    }
# Find what permissions are in the sharing link     
    $Matches = [regex]::Matches($Rec.AuditData.EventData, "<(.*?)>(.*?)</\1>")
    $SharingLinkData = @{}
    ForEach ($Match in $Matches) {
        $SharingLinkData[$Match.Groups[1].Value] = $Match.Groups[2].Value
    }

    # Get the user name from the user principal name. If the user principal name is not in the hash table, get it from the Graph API
    # and add it to the hash table for future use.  
    $UserName = $null; $UPN = $null; $User = $null
    If ($Rec.userPrincipalName -like "*@*") {
        $UserName = $UserDetails[$Rec.userPrincipalName]
        $UPN = $Rec.userPrincipalName

        If ($null -eq $UserName) {
            Write-Host ("Checking user {0}..." -f $Rec.userPrincipalName)
            $User = Get-MgUser -UserId $Rec.userPrincipalName -ErrorAction SilentlyContinue
            If ($null -ne $User) {
                $UserDetails.Add($User.UserPrincipalName, $User.displayName)
                $UserName = $User.displayName
                $UPN = $User.UserPrincipalName
            } Else {
                Write-Host ("User {0} is unknown in the directory" -f $Rec.userPrincipalName)
                $UserName = $Rec.userPrincipalName
                $UPN = $Rec.userPrincipalName
            }
        }
    }

    $TargetUserName = $null; $TargetUPN = $null
    If ($Rec.Auditdata.targetUserorGroupName -like "*@*") {
        $TargetUserOrGroupName = $Rec.AuditData.targetUserorGroupName
        $TargetUserName = $UserDetails[$TargetUserorGroupName]
        $TargetUPN = $Rec.AuditData.targetUserorGroupName

        If ($null -eq $TargetUserName) {
            $TargetUser = Get-MgUser -UserId $TargetUserorGroupName -ErrorAction SilentlyContinue
            If ($null -ne $TargetUser) {
                $UserDetails.Add($TargetUser.UserPrincipalName, $TargetUser.displayName)
                $TargetUserName = $TargetUser.displayName
                $TargetUPN = $TargetUser.UserPrincipalName
            } Else {
                Write-Host ("User {0} is unknown in the directory" -f $TargetUserOrGroupName)
                $TargetUserName = $TargetUserOrGroupName
                $TargetUPN = $TargetUserOrGroupName
            }
        }
    } Else {
        $TargetUserName = $Rec.AuditData.targetUserorGroupName
        $TargetUPN = $Rec.AuditData.targetUserorGroupName
    }

    If ($TargetUPN -like "*#EXT#*") {
        $TargetDomain = $TargetUPN.split("_")[1].Split("#")[0]
    } Else {
        $TargetDomain = $TargetUPN.split("@")[1]
    }

    If ($Rec.AuditData.sensitivityLabelId) {
        $LabelName = $LabelsHash[$Rec.AuditData.sensitivityLabelId]
    } else {
        $LabelName = $null
    }

    If ($Rec.AuditData.SiteSensitivityLabelId) {
        $SiteLabelName = $LabelsHash[$Rec.AuditData.SiteSensitivityLabelId]
    } else {
        $SiteLabelName = $null
    }

    $ReportLine = [PSCustomObject][Ordered]@{ 
        CreatedDateTime         = Get-Date $Rec.createdDateTime -format 'dd-MMM-yyyy HH:mm:ss'
        User                    = $UserName
        UPN                     = $UPN
        Operation               = $Rec.operation
        Permission              = $SharingLinkData['PermissionsGranted']
        FileName                = $Rec.AuditData.SourceFileName
        TargetUserName          = $TargetUserName
        TargetUPN               = $TargetUPN
        TargetType              = $Rec.AuditData.TargetUserOrGroupType
        SiteUrl                 = $Rec.AuditData.SiteUrl
        SensitivityLabelId      = $Rec.AuditData.sensitivityLabelId
        SensitivityLabel        = $LabelName
        SiteSensitiviityLabel   = $Rec.AuditData.SiteSensitivityLabelId
        SiteSensitivityLabel    = $SiteLabelName
        UserAgent               = $UserAgent
        Object                  = $Rec.AuditData.ObjectId
        TargetDomain            = $TargetDomain

    }
    $AuditReport.Add($ReportLine)
}

$AuditReport | Out-GridView -Title "Audit Report for File Sharing Events with Users and Guests"

Write-Host "Generating report..."
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\File Sharing Events.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    }
    $AuditReport | Export-Excel -Path $ExcelOutputFile -WorksheetName "File Sharing Ecvents" `
        -Title ("File Sharing Events {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "FileSharingEvents"
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\File Sharing Ecvents.CSV"
    $AuditReport | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}
 
If ($ExcelGenerated) {
    Write-Host ("An Excel report of calendar items is available in {0}" -f $ExcelOutputFile)
} Else {    
    Write-Host ("A CSV report of calendar items is available in {0}" -f $CSVOutputFile)
}  

Write-Host ""
Write-Host "Here's a snapshot of the domains users are sharing files with..."
$AuditReport | Group-Object TargetDomain -NoElement | Sort-Object Count -Descending | Format-Table Name, Count

Write-Host "All done..."

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.