# Apply-ExcludeAllHoldstoInactiveMailboxes.PS1
# GitHub Link:  
# V1.0 11-Dec-2025

# Clean up inactive mailboxes by removing (excluding) them from all holds

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {   
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
}

[array]$InactiveMailboxes = Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited | Sort-Object WhenCreated -Descending
If ($InactiveMailboxes) {
    [int]$ComplianceTagged = $InactiveMailboxes | Where-Object { $_.ComplianceTagHoldApplied -ne $null } | Measure-Object | Select-Object -ExpandProperty Count
    [int]$InPlaceHolds = $InactiveMailboxes | Where-Object { $_.InPlaceHolds.Count -gt 0 } | Measure-Object | Select-Object -ExpandProperty Count
    [int]$DelayHoldApplied = $InactiveMailboxes | Where-Object { $_.$DelayHoldApplied -eq $true } | Measure-Object | Select-Object -ExpandProperty Count
    [int]$DelayReleaseHoldApplied = $InactiveMailboxes | Where-Object { $_.$DelayReleaseHoldApplied -eq $true } | Measure-Object | Select-Object -ExpandProperty Count
} Else {
    Write-Host "No inactive mailboxes found"
    Break
}

Write-Host ""
Write-Host ("Found {0} inactive mailboxes" -f $InactiveMailboxes.Count)
Write-Host "Of these:"  
Write-Host (" - {0} have compliance tags applied" -f $ComplianceTagged)
Write-Host (" - {0} have in-place holds noted" -f $InPlaceHolds)            
Write-Host (" - {0} have delay hold applied" -f $DelayHoldApplied)
Write-Host (" - {0} have delay release hold applied" -f $DelayReleaseHoldApplied)
Write-Host
Write-Host "Processing inactive mailboxes to exclude from all holds..."

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Mailbox in $InactiveMailboxes) {
    Write-Host ("Processing inactive mailbox {0} ({1}) created on {2}" -f $Mailbox.DisplayName, $Mailbox.PrimarySmtpAddress, $WhenCreated)
    $ReportLine = [PSCustomObject][Ordered]@{
        Timestamp                    = Get-Date -Format 'dd-MMM-yyyy HH:mm:ss'
        DisplayName                  = $Mailbox.DisplayName
        PrimarySmtpAddress           = $Mailbox.PrimarySmtpAddress
        WhenCreated                  = Get-Date $Mailbox.WhenCreated -Format 'dd-MMM-yyyy HH:mm'
        WhenSoftDeleted              = If ($Mailbox.WhenSoftDeleted) { Get-Date $Mailbox.WhenSoftDeleted -Format 'dd-MMM-yyyy HH:mm' } Else { 'N/A' }
        ComplianceTagHoldApplied     = $Mailbox.ComplianceTagHoldApplied
        InPlaceHoldsCount            = $Mailbox.InPlaceHolds.Count
        DelayHoldApplied             = $Mailbox.DelayHoldApplied
        DelayReleaseHoldApplied      = $Mailbox.DelayReleaseHoldApplied
        IsInactiveMailbox            = $Mailbox.IsInactiveMailbox
        IsSoftDeletedByRemoveMailbox = $Mailbox.IsSoftDeletedByRemove 
    }
  
    Try {
        Set-Mailbox -Identity $Mailbox.DistinguishedName -ExcludeFromAllHolds -InactiveMailbox -ErrorAction Stop
        $ReportLine | Add-Member -NotePropertyName 'ExcludedFromAllHolds' -NotePropertyValue 'Succeeded'
    } Catch {
        Write-Host ("Failed to exclude mailbox {0} from all holds: {1}" -f $Mailbox.DisplayName, $_.Exception.Message) -ForegroundColor Red
        $ReportLine | Add-Member -NotePropertyName 'ExcludedFromAllHolds' -NotePropertyValue 'Failed'
    }
    $Report.Add($ReportLine)  
}

$Report | Out-GridView

# Export the report as an Excel worksheet or CSV file
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $true
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\ExcludeAllHolds.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    } 
    $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Exclude All Holds" -Title ("Exclude All Holds Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "ExcludeAllHolds" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\ExcludeAllHolds.CSV"
    $Report  | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}

If ($ExcelGenerated) {
    Write-Output ("Excel worksheet output written to {0}" -f $ExcelOutputFile)
} Else {
    Write-Output ("CSV output file written to {0}" -f $CSVOutputFile)
} 

Write-Output "All done - enjoy the information about inactive mailboxes that have been excluded from all holds!"


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.