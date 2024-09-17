# Report-RecoverableItemsPS.PS1
# Report items stored in the Exchange Online Recoverable Items folder using the Get-RecoverableItems cmdlet
# Other version which uses the Microsoft Graph PowerShell SDK is described in https://office365itpros.com/2024/09/17/report-recoverable-items/

# 17-Sep-2024
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-RecoverableItemsPS.PS1

# Connect to Exchange Online, if we're not already connected
$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

[datetime]$StartDate = (Get-Date).AddDays(-365)
$StartDate = Get-Date ($StartDate) -format "dd-MMM-yyyy hh:mm"
[datetime]$EndDate = Get-Date -format "dd-MMM-yyyy hh:mm"

Write-Host "Scanning for mailboxes..."
[array]$NoReportFolders = @("Audits", "Calendar Logging")
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Sort-Object DisplayName
If (!($Mbx)) {
    Write-Host "No mailboxes found - exiting!"; break
}  

Write-Host ("Processing {0} mailboxes..." -f $Mbx.Count)
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {
    Write-Host ("Processing mailbox {0}" -f $M.UserPrincipalName) -ForegroundColor Yellow
    [array]$Items = Get-RecoverableItems -Identity $M.UserPrincipalName -SourceFolder RecoverableItems `
     -FilterStartTime $StartDate.toString() -FilterEndTime $EndDate.toString() -ResultSize Unlimited

    # If some items are returned, report them
    Write-Host ("Found {0} items" -f $Items.Count)
    ForEach ($Item in $Items) {
        $DateOfRecord = [datetime]::ParseExact($Item.LastModifiedTime, "MM/dd/yyyy HH:mm:ss", $null);
        $ReportLine = [PSCustomObject][Ordered]@{
            Mailbox              = $M.UserPrincipalName
            Subject              = $Item.Subject
            'Last Modified Time' = Get-Date ($DateOfRecord) -format 'dd-MMM-yyyy HH:mm'
            LastParent           = $Item.LastParentPath
            ItemClass            = $Item.ItemClass
        }
        $Report.Add($ReportLine)   
    }
}

Write-Host ("Details of {0} items from Recoverable Items reported from {1} mailboxes" -f $Report.count, $Mbx.count)
$Report | Out-GridView -Title ("Items found in Recoverable Items folder from {0}" -f $StartDate)
$Report | Export-CSV -Encoding utf8 c:\temp\RecoverableItemsFiles.csv
Write-Host "Output CSV file available in c:\temp\RecoverableItemsFiles.csv"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.