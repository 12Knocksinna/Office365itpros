# Fix-ProxyEmailAddresses.PS1
# Fix up primary and proxy email addresses for mailboxes, groups, and distribution lists
# after a domain has been removed from a Microsoft 365 tenant

# V1.0 14-Mar-2025
# GitHub link:

[array]$Modules = Get-Module -ListAvailable | Where-Object { $_.Name -eq "ExchangeOnlineManagement" }
If (!$Modules) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
}

[array]$Domains = Get-AcceptedDomain 
$PrimaryDomain = $Domains | Where-Object { $_.Default -eq $true } | Select-Object -ExpandProperty DomainName
[array]$Domains = $Domains | Select-Object -ExpandProperty DomainName


Write-Host "Checking mailboxes..."
[array]$Mailboxes = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox, discoveryMailbox
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Mailbox in $Mailboxes) {
    $ExternalAddresses = $Mailbox.EmailAddresses | Where-Object { $_ -like "SMTP:*" -and ($_.Split(':')[1].Split('@')[1] -notin $Domains) }
    If ($ExternalAddresses) {
        $ReportLine = [PSCustomObject][Ordered]@{
            DisplayName             = $Mailbox.DisplayName
            PrimarySmtpAddress      = $Mailbox.PrimarySmtpAddress
            EmailAddresses          = $ExternalAddresses -join ", "
            Type                    = "mailbox"
            Identity                = $Mailbox.Alias
        }
        $Report.Add($ReportLine)
    }
}

Write-Host "Checking Microsoft 365 Groups..."
[array]$Groups = Get-UnifiedGroup -ResultSize Unlimited

ForEach ($Group in $Groups) {
    $ExternalAddresses = $Group.EmailAddresses | Where-Object { $_ -like "SMTP:*" -and ($_.Split(':')[1].Split('@')[1] -notin $Domains) }
    If ($ExternalAddresses) {
        $ReportLine = [PSCustomObject][Ordered]@{
            DisplayName             = $Group.DisplayName
            PrimarySmtpAddress      = $Group.PrimarySmtpAddress
            EmailAddresses  = $ExternalAddresses -join ", "
            Type                    = "group"
            Identity                = $Group.Alias
        }
        $Report.Add($ReportLine)
    }
}

Write-Host "Checking Distribution Lists..."
[array]$DLs = Get-DistributionGroup -ResultSize Unlimited

ForEach ($DL in $DLs) {
    $ExternalAddresses = $DL.EmailAddresses | Where-Object { $_ -like "SMTP:*" -and ($_.Split(':')[1].Split('@')[1] -notin $Domains) }
    If ($ExternalAddresses) {
        $ReportLine = [PSCustomObject][Ordered]@{
            DisplayName             = $DL.DisplayName
            PrimarySmtpAddress      = $DL.PrimarySmtpAddress
            EmailAddresses          = $ExternalAddresses -join ", "
            Type                    = "dl"
            Identity                = $DL.Alias
        }
        $Report.Add($ReportLine)
    }
}

Write-Host "Checking Dynamic distribution groups..."
[array]$DDLs = Get-DynamicDistributionGroup -ResultSize Unlimited

ForEach ($DDL in $DDLs) {
    $ExternalAddresses = $DDL.EmailAddresses | Where-Object { $_ -like "SMTP:*" -and ($_.Split(':')[1].Split('@')[1] -notin $Domains) }
    If ($ExternalAddresses) {
        $ReportLine = [PSCustomObject][Ordered]@{
            DisplayName             = $DDL.DisplayName
            PrimarySmtpAddress      = $DDL.PrimarySmtpAddress
            EmailAddresses          = $ExternalAddresses -join ", "
            Type                    = "ddl"
            Identity                = $DDL.Alias
        }
        $Report.Add($ReportLine)
    }
}

Write-Host ("{0} mailboxes, {1} groups, {2} distribution lists, and {3} dynamic distribution lists checked" -f $Mailboxes.Count, $Groups.Count, $DLs.Count, $DDLs.Count)
Write-Host ("Problems found in {0} objects" -f $Report.Count)

$Report | Format-Table -AutoSize

ForEach ($Object in $Report) {

    $UpdatePrimary = $false
    $NewPrimarySmtpAddress = $null

    # Check if primary SMTP address needs to be updated
    If ($Object.PrimarySmtpAddress.Split('@')[1] -notin $Domains) {
        Write-Host ("Primary SMTP address must be updated from {0}" -f $Object.PrimarySmtpAddress)
        $NewPrimarySmtpAddress = ("{0}@{1}" -f $Object.PrimarySmtpAddress.Split('@')[0], $PrimaryDomain)
        $UpdatePrimary = $true
    }

    If ($UpdatePrimary) {
        Write-Host ("Setting new primary SMTP address {0}" -f $NewPrimarySmtpAddress)
        Switch ($Object.Type) {
            "mailbox" {
                Set-Mailbox -Identity $Object.Identity -EmailAddresses @{Remove=$Object.PrimarySmtpAddress; Add=$NewPrimarySmtpAddress} -ErrorAction SilentlyContinue
                Set-Mailbox -Identity $Object.Identity -WindowsEmailAddress $NewPrimarySmtpAddress -ErrorAction SilentlyContinue
            }   
            "group" {
                Set-UnifiedGroup -Identity $Object.Identity -PrimarySmtpAddress $NewPrimarySmtpAddress -ErrorAction SilentlyContinue
            }
            "dl" {
                Set-DistributionGroup -Identity $Object.Identity -PrimarySmtpAddress $NewPrimarySmtpAddress -ErrorAction SilentlyContinue
            }
            "ddl" {
                Set-DynamicDistributionGroup -Identity $Object.Identity -PrimarySmtpAddress $NewPrimarySmtpAddress -ErrorAction SilentlyContinue
            }
        }
    }

    [array]$EmailAddresses = $Object.EmailAddresses -split ", "
    ForEach ($Address in $EmailAddresses) {
        If ($Address.Split('@')[1] -notin $Domains) {
            $AddressToRemove = $Address.Split(':')[1]
            Write-Host ("Removing address {0} from {1}" -f $Address, $Object.DisplayName)
            Switch ($Object.Type) {
            "mailbox" {
                Set-Mailbox -Identity $Object.Identity -EmailAddresses @{Remove=$AddressToRemove} -ErrorAction SilentlyContinue   
            }
            "group" {
                Set-UnifiedGroup -Identity $Object.Identity -EmailAddresses @{Remove=$AddressToRemove} -ErrorAction SilentlyContinue 
            }
             "dl" {
                Set-DistributionGroup -Identity $Object.Identity -EmailAddresses @{Remove=$AddressToRemove} -ErrorAction SilentlyContinue 
            }
             "ddl" {
                Set-DynamicDistributionGroup -Identity $Object.Identity -EmailAddresses @{Remove=$AddressToRemove} -ErrorAction SilentlyContinue 
            }
            }          
        }
    }

}
Write-Host "All done"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization.  Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.