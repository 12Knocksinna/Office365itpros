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
Write-Host "----------------------------------------------------------"

# Count all users
Write-Host "`nCounting all users..." -ForegroundColor Yellow
Try {
    [array]$AllUsers = Get-MgUser -ConsistencyLevel eventual -CountVariable AllUsersCount `
      -All -PageSize 500 -ErrorAction Stop -Property Id, UserType, displayName, userPrincipalName
} Catch {
    Write-Host "Error counting all users: $_" -ForegroundColor Red
    Break
}
Write-Host ("Total users in tenant: {0}" -f $AllUsers.Count) -ForegroundColor Green

# Count member users (not guests)
Write-Host "`nCounting member users..." -ForegroundColor Yellow
$MemberUsersCount = $AllUsers | Where-Object { $_.UserType -eq 'Member' } | Measure-Object | Select-Object -ExpandProperty Count
Write-Host ("Member users: {0}" -f $MemberUsersCount) -ForegroundColor Green

# Count guest users
Write-Host "`nCounting guest users..." -ForegroundColor Yellow
$GuestUsersCount = $AllUsers | Where-Object { $_.UserType -eq 'Guest' } | Measure-Object | Select-Object -ExpandProperty Count
Write-Host ("Guest users: {0}" -f $GuestUsersCount) -ForegroundColor Green

# Count licensed users (users with at least one license assigned)
Write-Host "`nCounting licensed users..." -ForegroundColor Yellow
Get-MgUser -Filter "assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -CountVariable LicensedUsersCount | Out-Null
Write-Host ("Licensed users: {0}" -f $LicensedUsersCount) -ForegroundColor Green

# Display summary
Write-Host "`n"
Write-Host "----------------------------------------------------------"
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "----------------------------------------------------------"
Write-Host ("Total Users:          {0,10}" -f $AllUsersCount) -ForegroundColor White
Write-Host ("  Member Users:       {0,10}" -f $MemberUsersCount) -ForegroundColor White
Write-Host ("  Guest Users:        {0,10}" -f $GuestUsersCount) -ForegroundColor White
Write-Host ("Licensed Users:       {0,10}" -f $LicensedUsersCount) -ForegroundColor White
Write-Host ("Unlicensed Users:     {0,10}" -f ($AllUsersCount - $LicensedUsersCount)) -ForegroundColor White
Write-Host "----------------------------------------------------------"

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

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
