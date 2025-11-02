# Report-DailySignIns.PS1
# An example of how to analyze and report daily sugn-ins recorded in the Entra ID sign-in audit log

# GitHub Link:
# V1.0 1-Oct-2025

Connect-MgGraph -Scopes "AuditLog.Read.All", "User.Read.All"

# Get tenant identifier
$TenantData = Get-MgOrganization
$TenantId = $TenantData.Id

# Set dates up to find what happened yesterday
$Today = Get-Date
$StartDate=$Today.AddDays(-30).ToString('yyyy-MM-ddT00:00:00Z')
$EndDate=$Today.ToString('yyyy-MM-ddT00:00:00Z')

Write-Output "Retrieving interactive sign-in records from $StartDate to $EndDate"
[array]$SignInRecords = Get-MgBetaAuditLogSignIn -All -Filter "createdDateTime ge $StartDate and createdDateTime lt $EndDate" -PageSize 999
If ($SignInRecords.Count -eq 0) {
    Write-Output "No sign-in records found"
    Break
}   Else {
    Write-Output ("Found {0} sign-in records to process..." -f $SignInRecords.Count)
    $Report = [System.Collections.Generic.List[Object]]::new()  
}

[array]$AppSignIns = $SignInRecords | Where-Object { $_.AppDisplayName -ne $null -and $_.Status.ErrorCode -eq 0 } | Group-Object -Property AppDisplayName -NoElement 
[array]$UserSignIns = $SignInRecords | Where-Object { $_.UserDisplayName -ne $null -and $_.Status.ErrorCode -eq 0} | Group-Object -Property UserDisplayName -NoElement 
# Some failed sign-ins have a null userPrincipalName, notably when a tennat account fails with an attempt using B2B authentication with another tenant
[array]$FailedSignIns = $SignInRecords | Where-Object { $_.Status.ErrorCode -ne 0 -and -not [string]::IsNullOrWhiteSpace($_.userPrincipalName) } | Group-Object -Property UserDisplayName -NoElement 
[array]$SingleFactorSignIns = $SignInRecords | Where-Object { $_.AuthenticationRequirement -eq "singleFactorAuthentication" -and $_.Status.ErrorCode -eq 0 } 
[array]$MfaSignIns = $SignInRecords | Where-Object { $_.ConditionalAccessStatus -eq "Success" -and $_.Status.ErrorCode -eq 0 } 
[array]$IncomingGuestMemberSignIns = $SignInRecords | Where-Object { $_.UserType -eq "Guest" -and $_.Status.ErrorCode -eq 0 -and $_.ResourceTenantId -eq $TenantId }
[array]$OutboundGuestMemberSignIns = $SignInRecords | Where-Object { $_.UserType -eq "Guest" -and $_.Status.ErrorCode -eq 0 -and $_.ResourceTenantId -ne $TenantId }

[array]$MfaUsers = $MfaSignIns | Group-Object userPrincipalName -NoElement  | Select-Object -ExpandProperty Name
[array]$SingleFactorUsers = $SingleFactorSignIns | Group-Object userPrincipalName -NoElement  | Select-Object -ExpandProperty Name
[array]$SingleFactorApps = $SingleFactorSignIns | Group-Object AppDisplayName -NoElement | Select-Object -ExpandProperty Name
[int]$CASuccess = $SignInRecords | Where-Object { $_.ConditionalAccessStatus -eq "Success" -and $_.Status.ErrorCode -eq 0 } | Measure-Object | Select-Object -ExpandProperty Count
[array]$CAFailureEvents = $SignInRecords | Where-Object { $_.ConditionalAccessStatus -eq "Failure" }
[int]$CANotApplied = $SignInRecords | Where-Object { $_.ConditionalAccessStatus -eq "NotApplied" -and $_.Status.ErrorCode -eq 0 } | Measure-Object | Select-Object -ExpandProperty Count
[int]$CAFailures = $CAFailureEvents.count
[array]$CAFailureUsers = $CAFailureEvents | Group-Object userPrincipalName -NoElement  | Select-Object -ExpandProperty Name

Write-Output ""
Write-Output ("Daily Sign-In Report for {0}" -f (Get-Date $EndDate -format 'dd MMMM yyyy'))
Write-Output ""
Write-Output "Top ten Applications signed into during the day"
Write-Output "-----------------------------------------------" 
$AppSignIns | Sort-Object Count -Descending | Select-Object -First 10 | Format-Table Name, Count -AutoSize
Write-Output ""
Write-Output "Top five users with successful sign-ins during the day"
Write-Output "------------------------------------------------------"
$UserSignIns | Sort-Object Count -Descending | Select-Object -First 5 | Format-Table Name, Count -AutoSize
Write-Output ""
Write-Output "Top five users with failed sign-ins during the day"
Write-Output "--------------------------------------------------"
$FailedSignIns | Sort-Object Count -Descending | Select-Object -First 5 | Format-Table Name, Count -AutoSize
Write-Output ""
Write-Output ("Number of users who signed in using single-factor authentication:  {0}" -f ($SingleFactorUsers.count))
Write-Output ("Number of users who signed in using multi-factor authentication:   {0}" -f ($MfaUsers.count))
Write-Output ("Accounts signed in using single-factor authentication:             {0}" -f ($SingleFactorUsers -join ", "))
Write-Output ("Apps accessed using single-factor authentication:                  {0}" -f ($SingleFactorApps -join ",  "))  
Write-Output ("Incoming guest user sign-ins:                                      {0}" -f ($IncomingGuestMemberSignIns.count))
Write-Output ("Outbound guest user sign-ins:                                      {0}" -f ($OutboundGuestMemberSignIns.count))
Write-Output ("Number of sign-ins where Conditional Access succeeded:             {0}" -f $CASuccess)
Write-Output ("Number of sign-ins where Conditional Access was not applied:       {0}" -f $CANotApplied)
Write-Output ("Number of sign-ins where Conditional Access failed:                {0}" -f $CAFailures) 
Write-Output ("Users with sign-ins where Conditional Access failed:               {0}" -f ($CAFailureUsers -join ", "))
Write-Output ""

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.