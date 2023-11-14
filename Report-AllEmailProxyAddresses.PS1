# A script to report all the proxy addresses assigned to mail-enabled objects in the organization
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-AllEmailProxyAddresses.PS1

If ($Null -eq (Get-ConnectionInformation)) {
    Connect-ExchangeOnline
}

$CSVOutputFile = "C:\Temp\EmailProxyAddresses.csv"
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report
# First, let's do mailboxes

Write-Host "Fetching details of user, shared, equipment, and room mailboxes..."
[array]$Mbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox
Write-Host ("Processing details for {0} mailboxes..." -f $Mbx.count)
ForEach ($M in $Mbx) {
    ForEach ($Address in $M.EMailAddresses) {
        $AddressType = $Address.Split(":")[0]
        $AddressProxy = $Address.Split(":")[1]
        If ($AddressType -eq 'smtp') {
            $ReportLine = [PSCustomObject]@{ 
                ProxyAddress = $AddressProxy
                Name         = $M.DisplayName
                UPN          = $M.userPrincipalName
                ObjectId     = $M.ExternalDirectoryObjectId
                Type         = $M.RecipientTypeDetails
            }
            $Report.Add($ReportLine)
        }
    }
}

# Now let's do group mailboxes
Write-Host "Fetching details of group mailboxes..."
[array]$GroupMailboxes = Get-UnifiedGroup -ResultSize Unlimited

Write-Host ("Processing details of {0} group mailboxes" -f $GroupMailboxes.count)
ForEach ($Gmbx in $GroupMailboxes) {
    ForEach ($Address in $Gmbx.EMailAddresses) {
        $AddressType = $Address.Split(":")[0]
        $AddressProxy = $Address.Split(":")[1]
        If ($AddressType -eq 'smtp') {
            $ReportLine = [PSCustomObject]@{ 
                ProxyAddress = $AddressProxy
                Name         = $Gmbx.DisplayName
                UPN          = $Gmbx.primarySmtpAddress
                ObjectId     = $Gmbx.ExternalDirectoryObjectId
                Type         = "Group Mailbox"
            }
            $Report.Add($ReportLine)
        }
    }
}

# Mail-enabled public folders
Write-Host "Fetching details of mail-enabled public folders..."
[array]$PFs = Get-MailPublicFolder -ResultSize Unlimited

If ($PFs) {
    Write-Host ("Processing details for {0} mail-enabled public folders..." -f $PFs.Count)
    ForEach ($PF in $PFs) {
        ForEach ($Address in $PF.EMailAddresses) {
            $AddressType = $Address.Split(":")[0]
            $AddressProxy = $Address.Split(":")[1]
            If ($AddressType -eq 'smtp') {
                $ReportLine = [PSCustomObject]@{ 
                    ProxyAddress = $AddressProxy
                    Name         = $PF.DisplayName
                    UPN          = $PF.primarySmtpAddress
                    ObjectId     =  "N/A"
                    Type         = "Public folder"
                }
                $Report.Add($ReportLine)
            }
        }
    }
}

# Distribution lists

Write-Host "Fetching details of distribution lists..."
[array]$DLs = Get-DistributionGroup -ResultSize Unlimited

If ($DLs) {
    Write-Host ("Processing details for {0} distribution lists..." -f $DLs.Count)
    ForEach ($DL in $DLs) {
        ForEach ($Address in $DL.EMailAddresses) {
            $AddressType = $Address.Split(":")[0]
            $AddressProxy = $Address.Split(":")[1]
            If ($AddressType -eq 'smtp') {
                $ReportLine = [PSCustomObject]@{ 
                    ProxyAddress = $AddressProxy
                    Name         = $DL.DisplayName
                    UPN          = $DL.primarySmtpAddress
                    ObjectId     = $DL.ExternalDirectoryObjectId
                    Type         = "Distribution list"
                }
                $Report.Add($ReportLine)
            }
        }
    }
}

# Dynamic distribution lists
Write-Host "Fetching details of dynamic distribution lists..."
[array]$DDLs = Get-DynamicDistributionGroup -ResultSize Unlimited

If ($DDLs) {
    Write-Host ("Processing details for {0} dynamic distribution lists..." -f $DLs.Count)
    ForEach ($DDL in $DDLs) {
        ForEach ($Address in $DDL.EMailAddresses) {
            $AddressType = $Address.Split(":")[0]
            $AddressProxy = $Address.Split(":")[1]
            If ($AddressType -eq 'smtp') {
                $ReportLine = [PSCustomObject]@{ 
                    ProxyAddress = $AddressProxy
                    Name         = $DDL.DisplayName
                    UPN          = $DDL.primarySmtpAddress
                    ObjectId     = 'N/A'
                    Type         = "Dynamic Distribution list"
                }
                $Report.Add($ReportLine)
            }
        }
    }
}

$Report = $Report | Sort-Object UPN
$Report | Export-CSV -NoTypeInformation $CSVOutputFile
Write-Host ("Output file is available in {0}" -f $CSVOutputFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
