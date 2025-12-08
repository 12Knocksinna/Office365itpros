# Get-TenantUserCount.PS1
# A simple script to count the number of users in a Microsoft 365 tenant
# Provides counts for all users, member users, guest users, and licensed users
# V1.0 8-Dec-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Get-TenantUserCount.PS1

# Check if we're already connected to Microsoft Graph
If (!(Get-MgContext)) {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes User.Read.All, Organization.Read.All -NoWelcome
} Else {
    Write-Host "Already connected to Microsoft Graph" -ForegroundColor Green
}

# Get the organization name
$OrgName = (Get-MgOrganization).DisplayName
Write-Host "`nUser Count Report for: $OrgName" -ForegroundColor Cyan
Write-Host ("Report generated: {0}" -f (Get-Date -Format "dd-MMM-yyyy HH:mm:ss")) -ForegroundColor Cyan
Write-Host "=" * 60

# Count all users
Write-Host "`nCounting all users..." -ForegroundColor Yellow
$AllUsersCount = 0
Get-MgUser -ConsistencyLevel eventual -Count AllUsersCount | Out-Null
Write-Host ("Total users in tenant: {0}" -f $AllUsersCount) -ForegroundColor Green

# Count member users (not guests)
Write-Host "`nCounting member users..." -ForegroundColor Yellow
$MemberUsersCount = 0
Get-MgUser -Filter "userType eq 'Member'" -ConsistencyLevel eventual -Count MemberUsersCount | Out-Null
Write-Host ("Member users: {0}" -f $MemberUsersCount) -ForegroundColor Green

# Count guest users
Write-Host "`nCounting guest users..." -ForegroundColor Yellow
$GuestUsersCount = 0
Get-MgUser -Filter "userType eq 'Guest'" -ConsistencyLevel eventual -Count GuestUsersCount | Out-Null
Write-Host ("Guest users: {0}" -f $GuestUsersCount) -ForegroundColor Green

# Count licensed users (users with at least one license assigned)
Write-Host "`nCounting licensed users..." -ForegroundColor Yellow
$LicensedUsersCount = 0
Get-MgUser -Filter "assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -Count LicensedUsersCount | Out-Null
Write-Host ("Licensed users: {0}" -f $LicensedUsersCount) -ForegroundColor Green

# Display summary
Write-Host "`n" 
Write-Host "=" * 60
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "=" * 60
Write-Host ("Total Users:          {0,10}" -f $AllUsersCount) -ForegroundColor White
Write-Host ("  Member Users:       {0,10}" -f $MemberUsersCount) -ForegroundColor White
Write-Host ("  Guest Users:        {0,10}" -f $GuestUsersCount) -ForegroundColor White
Write-Host ("Licensed Users:       {0,10}" -f $LicensedUsersCount) -ForegroundColor White
Write-Host ("Unlicensed Users:     {0,10}" -f ($AllUsersCount - $LicensedUsersCount)) -ForegroundColor White
Write-Host "=" * 60

# Create a summary object for output
$Summary = [PSCustomObject][Ordered]@{
    'Organization'       = $OrgName
    'Report Date'        = Get-Date -Format "dd-MMM-yyyy HH:mm:ss"
    'Total Users'        = $AllUsersCount
    'Member Users'       = $MemberUsersCount
    'Guest Users'        = $GuestUsersCount
    'Licensed Users'     = $LicensedUsersCount
    'Unlicensed Users'   = ($AllUsersCount - $LicensedUsersCount)
}

# Output the summary object
$Summary

Write-Host "`nScript completed successfully!" -ForegroundColor Green
