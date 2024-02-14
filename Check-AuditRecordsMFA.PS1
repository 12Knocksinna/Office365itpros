# Check-AuditRecordsMFA.PS1
# A script to demonstrate how to use Entra ID sign-in audit records to check MFA usage by user accounts
# V1.0 13-Feb-2024
# https://github.com/12Knocksinna/Office365itpros/blob/master/Check-AuditRecordsMFA.PS1

Connect-MgGraph -Scopes Directory.Read.All, AuditLog.Read.All -NoWelcome

$CSVOutputFile = "c:\temp\CheckAuditRecordsMFA.csv"
$TenantId = (Get-MgOrganization).Id
$StartDate = (Get-Date).AddDays(-15)
$StartDateS = (Get-Date $StartDate -Format s) + "Z"

Write-Host "Looking for sign-in records..."
[array]$AuditRecords = Get-MgBetaAuditLogSignIn -Top 5000 `
  -Filter "(CreatedDateTime ge $StartDateS) and (signInEventTypes/any(t:t eq 'interactiveuser')) and (usertype eq 'Member')"
If (!$AuditRecords) {
    Write-Host "No sign-in records found - exiting"
    Break
}

# Eliminate any member sign-ins from other tenants
$AuditRecords = $AuditRecords | Where-Object HomeTenantId -match $TenantId

Write-Host "Finding user accounts to check..."
[array]$Users = Get-MgUser -All -Sort 'displayName' `
    -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -consistencyLevel eventual -CountVariable UsersFound `
    -Property Id, displayName, signInActivity, userPrincipalName

Write-Host ("Checking {0} sign-in audit records for {1} user accounts..." -f $AuditRecords.count, $Users.count)
[int]$MFAUsers = 0
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    $Authentication = "No sign-in records found"
    $Status = $null; $MFARecordDateTime = $null; $MFAMethodsUsed = $null; $MFAStatus = $null
    $UserLastSignInDate = $null
    [array]$UserAuditRecords = $AuditRecords | Where-Object {$_.UserId -eq $User.Id} | `
        Sort-Object {$_.CreatedDateTIme -as [datetime]} 
    
    If ($UserAuditRecords) {
        $MFAFlag = $false
        If ("multifactorauthentication" -in $UserAuditRecords.AuthenticationRequirement) {
            # The set of sign-in records contain at least one MFA record, so we extract details
            $MFAUsers++
            $Authentication = "MFA"
            ForEach ($Record in $UserAuditRecords) {
                $Status = $Record.Status.AdditionalDetails
                $MFARecordDateTime = $Record.CreatedDateTIme 
                If ($Status -eq 'MFA completed in Azure AD') {
                    # Found a record that specifies the methods used, so capture that for the report
                    $MFAStatus = "MFA Performed"
                    $MFAMethodsUsed =  $Record.AuthenticationDetails.AuthenticationMethod -join ", "
                    $MFAFlag = $true
                } ElseIf ($MFAFlag -eq $false) {
                    # Otherwise capture details for use of an existing claim
                    $MFAStatus = "Existing claim in the token used"
                    $MFAMethodsUsed = 'Existing claim'                  
                }
            }
        } Else {
            # No MFA sign-in records exist for the user, so they use single-factor
            $Authentication = "Single factor"
        }
    }
    $UserLastSignInDate = $User.SignInActivity.LastSignInDateTime
    $ReportLine = [PSCustomObject][Ordered]@{ 
        User            = $User.Id
        Name            = $User.DisplayName
        UPN             = $User.UserPrincipalName
        LastSignIn      = $UserLastSignInDate
        Authentication  = $Authentication
        'MFA timestamp' = $MFARecordDateTime
        'MFA status'    = $MFAStatus
        'MFA methods'   = $MFAMethodsUsed
    }
    $Report.Add($ReportLine)
}

$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation $CSVOutputFile
[float]$MFACheck = ($MFAUsers/$Users.Count)*100
$PercentMFAUsers = ($MFAUsers/$Users.Count).toString('P')
Write-Host ("{0} users of {1} found to use MFA" -f $MFAUsers, $Users.count)
If ($MFACheck -gt 38) {
    Write-Host ("Congratulations - Tenant MFA usage percentage {0} is better than the 38% rate for Entra ID" -f $PercentMFAUsers)
} Else {
    Write-Host ("Oh dear! Tenant MFA usage percentage {0} is worse than Entra ID norm (38%)" -f $PercentMFAUsers)
}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
