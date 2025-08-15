# Remove-ObsoleteMobileDevices.PS1
# An example script to show how to remove obsolete mobile device partnerships from user mailboxes. Users can't do this any longer
# because Microsoft has removed the options to manage mobile devices from OWA and the new Outlook: https://office365itpros.com/2025/08/15/mobile-device-management-owa/

# This script fetches all user mailboxes, finds the device partnerships that exist for each mailbox, and checks the date of the last
# successful synchronization for each device. If a device hasn't synchronized in more than a year (adjustable), it's deemed to be obsolete
# and be a candidate for removal. The administrator can use the set of obsolete devices found by the script to remove the partnerships for
# the devices.

# V1.0 14-Aug-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Remove-ObsoleteMobileDevices.PS1

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
                $DeviceStats = Get-ExoMobileDeviceStatistics -Id $Device.Identity -ErrorAction Stop
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

# Extract the obsolete devices from the full set
[array]$ObsoleteDevices = $Report | Where-Object { $_.ObsoleteDevice -eq $true }
[array]$ObsoleteMailboxes = ($ObsoleteDevices | Select-Object -ExpandProperty User | Get-Unique)
Write-Host ""
Write-Host ("Mobile devices checked for {0} mailboxes." -f $Mbx.Count)
Write-Host ("Found {0} obsolete mobile devices ({1:P2} of total)" -f $ObsoleteDevices.Count, ($ObsoleteDevices.Count / $Report.Count)) 
Write-Host ("Obsolete devices were found with partnerships with {0} mailboxes ({1})" -f $ObsoleteMailboxes.Count, ($ObsoleteMailboxes -join ", "))

$ObsoleteDevices | Out-GridView -Title "Obsolete Mobile Devices Report"

Write-Host "Generating an Excel or CSV report containing details of obsolete mobile devices"
# Generate reports
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    $ExcelTitle = ("Obsolete mobile devices report for {0}" -f (Get-Date -Format 'dd-MMM-yyyy'))
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $OutputXLSXFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\ObsoleteMobileDevices.xlsx"
    If (Test-Path $OutputXLSXFile) {
        Remove-Item $OutputXLSXFile -ErrorAction SilentlyContinue
    }
    $ObsoleteDevices | Export-Excel -Path $OutputXLSXFile -WorksheetName "Obsolete Mobile Devices" -Title $ExcelTitle -TitleBold -TableName "ObsoleteDevices" -AutoSize
} Else {
    $OutputCSVFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\ObsoleteMobileDevices.csv"
    $ObsoleteDevices | Export-Csv -Path $OutputCSVFile -NoTypeInformation -Encoding Utf8
}
  
If ($ExcelGenerated) {
    Write-Host ("An Excel worksheet containing the report data is available in {0}" -f $OutputXLSXFile)
} Else {
    Write-Host ("A CSV file containing the report data is available in {0}" -f $OutputCSVFile)
}

# Ask the user if they want to remove the obsolete devices
Write-Host ""
Write-Host "IMPORTANT: You should review the obsolete mobile device data before removing any device partnerships." -ForegroundColor Red
Write-Host ""
$GoAhead = Read-Host "Do you want to remove the obsolete mobile devices? (Y/N)"
If ($GoAhead.tolower() -eq "y") {
    ForEach ($Device in $ObsoleteDevices) {
        Try {
            Remove-MobileDevice -Identity $Device.Identity -Confirm:$false -ErrorAction Stop
            Write-Host ("Removed obsolete mobile device {0} for user {1}" -f $Device.'Mobile Device Id', $Device.User) -ForegroundColor Green
        } Catch {
            Write-Host ("Failed to remove mobile device {0} for user {1} - {2}" -f $Device.'Mobile Device Id', $Device.User, $_.Exception.Message) -ForegroundColor Red
        }
    }
} Else {
    Write-Host "No devices removed" -ForegroundColor Yellow
    Write-Host "You can edit the output file containing obsolete mobile device data and make changes there before using the file as input to remove devices."
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.