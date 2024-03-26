# Find-LoopAppUsers.PS1
# Example script to show how to find which licensed users are using the Loop App
# V1.0 17-Mar-2024
# https://github.com/12Knocksinna/Office365itpros/blob/master/Find-LoopAppUsers.PS1

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Connecting to Exchange Online" -ForegroundColor Red
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}
# To read user information from Entra ID
Connect-MgGraph -NoWelcome -Scopes User.Read.All

$CSVOutPutFile = "c:\temp\LoopUserInformation.csv"

# This command finds user accounts with a license containing the Loop app. It specifies the SKU identifers for
# Microsoft 365 Business Standard, Microsoft 365 Business Premium, Microsoft 365 E3, amd Microsoft 365 E5. 
# There are other variants of these SKUs # for government and academic use, so it's important to pass the SKU 
# identifiers in use within your tenant. The service plan for Loop is MICROSOFT_LOOP (c4b8c31a-fb44-4c65-9837-a21f55fcabda)
Write-Host "Finding user accounts to check..."
[array]$LoopLicensedUsers = Get-MgUser -Filter "assignedLicenses/any(s:s/skuId eq cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46) `
    or assignedLicenses/any(s:s/skuid eq f245ecc8-75af-4f8e-b61f-27d8114de5f3) `
    or assignedLicenses/any(s:s/skuid eq 05e9a617-0261-4cee-bb44-138d3ef5d965) `
    or assignedLicenses/any(s:s/skuid eq 06ebc4ee-1bb5-47dd-8120-11324bc54e06)" `
    -ConsistencyLevel Eventual -CountVariable Licenses -All -Sort 'displayName' `
    -Property Id, displayName, signInActivity, userPrincipalName, department, jobtitle, country

If ($LoopLicensedUsers.count -eq 0) {
    Write-Host "No users can be found eligble to use the Loop app - exiting" -ForegroundColor Red
    break
}

# Now find licensed user accounts so that we can detect who might be using the Loop app without a supported license
# (OK until 30 June 2024)
[Array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'"  `
  -ConsistencyLevel eventual -CountVariable Records -All `
  -Property id, displayName, userPrincipalName, country, department, assignedlicenses, `
  licenseAssignmentStates, createdDateTime, jobTitle, signInActivity, companyName |  `
  Sort-Object DisplayName

Write-Host ("{0} user accounts have eligible licenses to use the Loop app out of {1} licensed accounts" -f $LoopLicensedUsers.count, $Users.count)

# Figure out licensed user accounts that can't use the Loop app (but might)
[array]$UsersNotLicensedForLoop = $Users | Where-Object {$_.id -notin $LoopLicensedUsers.id}

# Set parameters for the unified audit log search
$StartDate = (Get-Date).AddDays(-15)
$EndDate = (Get-Date).AddDays(1)
# Define the set of SharePoint file audit records we are interested in
[array]$Operations = "FileModified", "FileModifiedExtended", "FileUploaded", "FileAccessed"
Write-Host "Searching the unified audit log for file operations performed using the Loop app..."
[array]$Records = Search-UnifiedAuditLog -Operations $Operations -StartDate $StartDate -EndDate $EndDate -Formatted `
 -SessionCommand ReturnLargeSet -ResultSize 5000

If (!($Records)) {
    Write-Host "No audit records can be found - exiting"
    Break
} Else {
    # Remove any duplicate records and make sure that everything is sorted in date order
    $Records = $Records | Sort-Object Identity -Unique 
    $Records = $Records | Sort-Object {$_.CreationDate -as [datetime]}
}

Write-Host ("Analyzing {0} audit records to identify users of the Loop app..." -f $Records.count)
$LoopRecords = [System.Collections.Generic.List[Object]]::new()

ForEach ($Rec in $Records) {
   $AuditData = $Rec.AuditData | ConvertFrom-JSON
   If ($Auditdata.Applicationid -eq 'a187e399-0c36-4b98-8f04-1edc167a0996' -and $AuditData.Operation -ne 'UserLoggedIn') {
    $ReportLine = [PSCustomObject][Ordered]@{
        UserPrincipalName   = $Rec.UserIds
        Timestamp           = $Rec.CreationDate
        fileName            = $AuditData.SourceFileName
        Operation           = $Rec.operations
        ObjectId            = $AuditData.ObjectId
    } 
    $LoopRecords.Add($ReportLine)      
   }
}

$LoopActiveUsers = $LoopRecords | Sort-Object UserPrincipalName -Unique

# First run through identifies user accounts who have a license that allows them to use the Loop app
# and are either active or not
$LicensedLoopUsersActivity = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $LoopLicensedUsers) {
    $AuditRecord = $null
# Fetch user information from the data we have already retrieved
    $AuditRecord = $LoopActiveUsers | Where-Object userPrincipalName -match $User.userPrincipalName
    If ($null -eq $AuditRecord) {
        Write-Host ("Can't find usage of the Loop app by {0}" -f $User.displayname)
        $LoopReportLine = [PSCustomObject][Ordered]@{
            User            = $User.UserPrincipalName
            Name            = $User.displayName
            'Job title'     = $User.jobTitle
            Department      = $User.department
            Country         = $User.country
            Licensed        = "True"
            Active          = "False"
            'Last Active'   = "N/A"
        }
        $LicensedLoopUsersActivity.Add($LoopReportLine)
    } Else {
        Write-Host ("User {0} last used the Loop app on {1}" -f $User.displayName, $AuditRecord.TimeStamp)
        $LoopReportLine = [PSCustomObject][Ordered]@{
            User            = $User.UserPrincipalName
            Name            = $User.displayName
            'Job title'     = $User.jobTitle
            Department      = $User.department
            Country         = $User.country
            Licensed        = "True"
            Active          = "True"
            'Last Active'   = $AuditRecord.TimeStamp
        }
        $LicensedLoopUsersActivity.Add($LoopReportLine)
   }
}

# Now check for unlicensed use of the Loop app
ForEach ($User in $UsersNotLicensedForLoop) {
    $AuditRecord = $null
    # Fetch user information from the data we have already retrieved
        $AuditRecord = $LoopActiveUsers | Where-Object userPrincipalName -match $User.userPrincipalName
        If ($AuditRecord) {
            Write-Host ("Unlicensed user {0} last used the Loop app on {1}" -f $User.displayName, $AuditRecord.TimeStamp) -ForegroundColor Red
            $LoopReportLine = [PSCustomObject][Ordered]@{
                User            = $User.UserPrincipalName
                Name            = $User.displayName
                'Job title'     = $User.jobTitle
                Department      = $User.department
                Country         = $User.country
                Licensed        = "False"
                Active          = "True"
                'Last Active'   = $AuditRecord.TimeStamp
            }
            $LicensedLoopUsersActivity.Add($LoopReportLine)
       }
}

$LicensedLoopUsersActivity | Out-GridView
$LicensedLoopUsersActivity | Export-CSV -Path $CSVOutPutFile -NoTypeInformation -Encoding utf8
Write-Host ("Loop user activity information is available in {0}" -f $CSVOutPutFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.