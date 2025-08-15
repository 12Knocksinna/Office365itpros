# Remove-ObsoleteMobileDevices.PS1
# An example script to show how to remove obsolete mobile device partnerships from user mailboxes

# V1.0 14-Aug-2025
# GitHub link:

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ($Modules -notcontains "ExchangeOnlineManagement") {
    Connect-ExchangeOnline -ShowBanner:$false
    Continue
} Else {
    # Make sure that a connection exists...
    Try {
        Get-ExoMailbox -ResultSize 1 -ErrorAction Stop | Out-Null
    } Catch {
        Write-Host "The Exchange Online management module is loaded but not connected - connecting now..."
        Connect-ExchangeOnline -ShowBanner:$false 
    }
}

Write-Host "Finding user mailboxes..."
[array]$Mbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
If (!($Mbx)) {
    Write-Host "No user mailboxes found" -ForegroundColor Red
    Break
} Else {
    Write-Host ("Checking mobile devices for {0} mailboxes" -f $Mbx.count)
}

$DaysObsoleteThreshold = 365
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {
    [array]$Devices = Get-MobileDevice -Mailbox $M.ExternalDirectoryObjectId  
    If ($Devices) {
        Write-Host ("Found {0} mobile devices for mailbox {1}" -f $Devices.count, $M.DisplayName)
        ForEach ($Device in $Devices) {
            $LastSyncTime = $null
            $ObsoleteDevice = $false
            Try {
                $DeviceStats = Get-MobileDeviceStatistics -Id $Device.Identity -ErrorAction Stop
            } Catch {
                Write-Host ("Failed to get statistics for device {0} - {1}" -f $Device.DeviceId, $_.Exception.Message) -ForegroundColor Yellow
                Continue
            }

            If ([string]::IsNullOrEmpty($DeviceStats.LastSuccessSync)) {
                $LastSyncTime = "Never"
                $DaysSinceLastSync = 'N/A'
            } Else {
                $LastSyncTime = Get-Date $DeviceStats.LastSuccessSync -format 'dd-MMM-yyyy HH:mm'
                $DaysSinceLastSync = (New-TimeSpan -Start $DeviceStats.LastSuccessSync -End (Get-Date)).Days    
            }

            If ($DaysSinceLastSync -gt $DaysObsoleteThreshold -or $DaysSinceLastSync -eq 'N/A') {
                $ObsoleteDevice = $true
            }

            $Report.Add([PSCustomObject]@{
                User                = $M.DisplayName
                Email               = $M.PrimarySmtpAddress
                'Mobile Device Id'  = $Device.DeviceId
                FirstSyncTime       = Get-Date $Device.FirstSyncTime -format 'dd-MMM-yyyy HH:mm'
                LastSyncTime        = $LastSyncTime
                DaysSinceLastSync   = $DaysSinceLastSync
                DeviceType          = $Device.DeviceType
                DeviceOS            = $Device.DeviceOS
                UserAgent           = $DeviceStats.DeviceUserAgent
                ObsoleteDevice      = $ObsoleteDevice
                Identity            = $Device.Identity
            })
        }

    }
}

[array]$ObsoleteDevices = $Report | Where-Object { $_.ObsoleteDevice -eq $true }
[array]$ObsoleteMailboxes = ($ObsoleteDevices | Select-Object -ExpandProperty User | Get-Unique)
Write-Host ""
Write-Host ("Mobile devices checked for {0} mailboxes." -f $Mbx.Count)
Write-Host ("Found {0} obsolete mobile devices ({1:P2} of total)" -f $ObsoleteDevices.Count, ($ObsoleteDevices.Count / $Report.Count)) 
Write-Host ("Obsolete devices were found with partnerships with {0} mailboxes ({1})" -f $ObsoleteMailboxes.Count, ($ObsoleteMailboxes -join ", "))

$ObsoleteDevices | Out-GridView -Title "Obsolete Mobile Devices Report"

$GoAhead = Read-Host "Do you want to remove the obsolete mobile devices? (Y/N)"
If ($GoAhead.tolower() -eq "y") {
    ForEach ($Device in $ObsoleteDevices) {
        Try {
            #Remove-MobileDevice -Identity $Device.Identity -Confirm:$false -ErrorAction Stop
            Write-Host ("Removed obsolete mobile device {0} for user {1}" -f $Device.'Mobile Device Id', $Device.User) -ForegroundColor Green
        } Catch {
            Write-Host ("Failed to remove mobile device {0} for user {1} - {2}" -f $Device.'Mobile Device Id', $Device.User, $_.Exception.Message) -ForegroundColor Red
        }
    }
}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.