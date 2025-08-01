# Report-IOSDevicesAuthenticator.ps1
# Report the set of users that have iOS devices with the Authenticator app installed. The method used is to scan for users
# whose preferred secondary authentication method is set to "push" and then check the mobile devices associated with those users to fimd
# if any devices run iOS.

# V1.0 9-July-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-iOSDevicesAuthenticator.PS1

# Connect to the Microsoft Graph and Exchange Online PowerShell modules
Connect-MgGraph -Scopes UserAuthenticationMethod.Read.All, User.Read.All -NoWelcome
[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ($Modules -notcontains 'ExchangeOnlineManagement') {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
}

Write-Host "Finding accounts to process..." -ForegroundColor Yellow
# Find Entra ID accounts that have a license assigned and are of type "Member"
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual `
    -CountVariable Records -All -PageSize 500 -Property Id, displayName, UserPrincipalName, department, country

$Report = [System.Collections.Generic.List[Object]]::new()

# Check each user for their authentication methods - we create a complete report in case the information is useful to administrators
ForEach ($User in $Users) {
    Write-Host "Processing user: $($User.displayName)" -ForegroundColor Cyan
    Try {
        $Uri = ("https://graph.microsoft.com/beta/users/{0}/authentication/signInPreferences" -f $User.Id)
        $AuthData = Invoke-MgGraphRequest -Uri $Uri -Method Get -ErrorAction Stop 
    } Catch {
        Write-Host "Failed to retrieve sign-in preferences for user $($User.displayName): $($_.Exception.Message)" -ForegroundColor Red
        Continue
    }
    $ReportLine = [PSCustomObject]@{
        User                            = $User.displayName
        UPN                             = $User.userPrincipalName
        Department                      = $User.department
        Country                         = $User.country
        Id                              = $User.Id
        'System preferred MFA enabled'  = $AuthData.isSystemPreferredAuthenticationMethodEnabled
        'System preferred MFA method'   = $AuthData.systemPreferredAuthenticationMethod
        'Secondary auth method'         = $AuthData.userPreferredMethodForSecondaryAuthentication 
    }
    $Report.Add($ReportLine)
}

# Filter the report to find users with the secondary authentication method set to "push"
[array]$StrongMethodUsers = $Report | Where-Object {$_.'Secondary auth method' -eq 'push'}
$ReportIOSDevices = [System.Collections.Generic.List[Object]]::new() 

# Check each user to find out what devices they use. If they have an iOS device, we check if it's been active in the last 30 days
ForEach ($User in $StrongMethodUsers) {
    Write-Host "Processing user with strong method: $($User.User)" -ForegroundColor Cyan
    Try {
        [array]$Devices = Get-MobileDevice -Mailbox $User.Id -ErrorAction Stop
    } Catch {
        Write-Host "Failed to retrieve mobile devices details for user $($User.User): $($_.Exception.Message)" -ForegroundColor Red
        Continue
    }

    ForEach ($Device in $Devices) {
       
        If (($Device.DeviceOS.SubString(0,3)) -ne "iOS") {
            Write-Host "Device doesn't run IOS: $($Device.DeviceId)" -ForegroundColor Yellow
            Continue
        }
        $DaysSinceLastSync = $Null; $SyncStatus = "OK"
        $DeviceStats = Get-ExoMobileDeviceStatistics -Identity $Device.DistinguishedName
        If ($Device.FirstSyncTime) {
            $DaysSinceFirstSync = (New-TimeSpan $Device.FirstSyncTime).Days 
        }
        If (!([string]::IsNullOrWhiteSpace($DeviceStats.LastSuccessSync))) {
            $DaysSinceLastSync = (New-TimeSpan $DeviceStats.LastSuccessSync).Days 
        } Else {
            $DaysSinceLastSync = $DaysSinceFirstSync
        }
        If ($DaysSinceLastSync -gt 30)  {
            $SyncStatus = ("Warning: {0} days since last sync" -f $DaysSinceLastSync) 
        }
        If ($Null -eq $DaysSinceLastSync) {
            $SyncStatus = "Never synched" 
            $DeviceStatus = "Unknown" 
        } Else {
            $DeviceStatus =  $DeviceStats.Status 
        }

        # Only report devices that have synced in the last 30 days
        If ($DaysSinceLastSync -le 30) {

            $ReportIOSLine = [PSCustomObject]@{
                DeviceId            = $Device.DeviceId
                DeviceOS            = $Device.DeviceOS
                Model               = $Device.DeviceModel
                UA                  = $Device.DeviceUserAgent
                User                = $Device.UserDisplayName
                UPN                 = $User.UPN
                FirstSync           = $Device.FirstSyncTime
                DaysSinceFirstSync  = $DaysSinceFirstSync
                LastSync            = $DeviceStats.LastSuccessSync
                DaysSinceLastSync   = $DaysSinceLastSync
                SyncStatus          = $SyncStatus
                Status              = $DeviceStatus
                Policy              = $DeviceStats.DevicePolicyApplied
                State               = $DeviceStats.DeviceAccessState
                LastPolicy          = $DeviceStats.LastPolicyUpdateTime
                DeviceDN            = $Device.DistinguishedName }
            $ReportIOSDevices.Add($ReportIOSLine)
        }
    } #End Devices

}

Write-Host ""
Write-Host "Users of iOS devices that are actively in use"
Write-Host "---------------------------------------------"
$ReportIOSDevices | Sort-Object User | Select-Object User, UPN, DeviceOS | Format-Table -AutoSize

$ReportIOSDevices | Export-Csv -Path "C:\Temp\ReportIOSDevices.csv" -NoTypeInformation
Write-Host "Report saved to C:\Temp\ReportIOSDevices.csv" -ForegroundColor Green
<<<<<<< HEAD

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
=======
>>>>>>> 9cd04fa912f87a627029497bc909aecc6d0c7c6a
