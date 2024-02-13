# Check-AuditRecordsMFA.PS1
# A script to demonstrate how to use Entra ID sign-in audit records to check MFA usage by user accounts
# V1.0 13-Feb-2024
# https://github.com/12Knocksinna/Office365itpros/blob/master/Check-AuditRecordsMFA.PS1

Connect-MgGraph -Scopes Directory.Read.All, AuditLog.Read.All

$CSVOutputFile = "c:\temp\CheckAuditRecordsMFA.csv"
$TenantId = (Get-MgOrganization).Id
$StartDate = (Get-Date).AddDays(-60)
$StartDateS = (Get-Date $StartDate -Format s) + "Z"

Write-Host "Looking for audit sign-in records..."
[array]$AuditRecords = Get-MgBetaAuditLogSignIn -Top 5000 `
  -Filter "(CreatedDateTime ge $StartDateS) and (signInEventTypes/any(t:t eq 'interactiveuser')) and (usertype eq 'Member')"
If (!$AuditRecords) {
    Write-Host "No audit records found - exiting"
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
    $Authentication = "No signin records found"
    $Status = $null; $MFARecordDateTime = $null; $MFAMethodsUsed = $null; $MFAStatus = $null
    $UserLastSignInDate = $null
    [array]$UserAuditRecords = $AuditRecords | Where-Object {$_.UserId -eq $User.Id} | `
        Sort-Object {$_.CreatedDateTIme -as [datetime]} 
    
    If ($UserAuditRecords) {
        $MFAFlag = $false
        If ("multifactorauthentication" -in $UserAuditRecords.AuthenticationRequirement) {
            $MFAUsers++
            $Authentication = "MFA"
        } Else {
            $Authentication = "Single factor"
        }
        ForEach ($Record in $UserAuditRecords) {
            $Status = $Record.Status.AdditionalDetails
            If ($Record.AuthenticationRequirement -eq "multifactorauthentication") {
                $MFARecordDateTime = $Record.CreatedDateTIme 
                If ($MFAFlag -eq $false) {
                    $MFAStatus = "Existing claim in the token used"
                    $MFAMethodsUsed = 'Existing claim'
                }       
            }
            If ($Status -eq 'MFA completed in Azure AD') {
                $MFAStatus = "MFA Performed"
                $MFAMethodsUsed =  $Record.AuthenticationDetails.AuthenticationMethod -join ", "
                $MFAFlag = $true
            }
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

$Report | Export-CSV -NoTypeInformation $CSVOutputFile
$PercentMFAUsers = ($MFAUsers/$Users.Count).toString('P')

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
