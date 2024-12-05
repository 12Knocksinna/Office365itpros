# Find-InactiveEmailUsers.PS1
# Find inactive email users based on the message trace information, which means we can only go back ten days...

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-InactiveEmailUsers.PS1

# V1.0 26-Nov-2024

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ($Modules -notcontains "ExchangeOnlineManagement") {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$False
}

$EndDate = Get-Date
$StartDate = $EndDate.AddDays(-10)

[array]$Messages = $Null 

$Page = 1 
Write-Host "Collecting message trace data for the last 10 days"
Do 
{ 
    [array]$CurrentMessages = (Get-MessageTrace -Status Delivered -PageSize 5000 -Page $Page `
       -StartDate $StartDate -EndDate $EndDate | Select-Object SenderAddress, Subject, Received)
    $Page++ 
    $Messages += $CurrentMessages 
}  Until ($Null -eq $CurrentMessages) 

$Messages = $Messages | Sort-Object {$_.Received -as [datetime]} -Descending
[array]$MessageTable = ($Messages | Sort-Object SenderAddress -Unique)

Write-Host "Looking for mailboxes..."
[array]$Mbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails 'UserMailbox' | Sort-Object DisplayName
Write-Host ("Processing {0} mailboxes to check activity..." -f $Mbx.count)
$Report = [System.Collections.Generic.List[Object]]::new()
[int]$ActiveStatusCount = 0
ForEach ($M in $Mbx) {
    $LastActiveDate = $null
    If ($MessageTable -Match $M.PrimarySMTPAddress) {
        $ActiveStatus = "Active"; $ActiveStatusCount++
        $LastActiveDate = $Messages | Where-Object {$_.SenderAddress -eq $M.PrimarySMTPAddress} | Select-Object -First 1 | Select-Object -ExpandProperty Received
        Write-Host ("{0} is active - message found on {1}" -f $M.DisplayName, $LastActiveDate) -Foregroundcolor Yellow 
    } Else {
        $ActiveStatus = "Inactive"
        Write-Host ("{0} is inactive" -f $M.DisplayName) -Foregroundcolor Red 
    }
    If ($null -ne $LastActiveDate) {
        $LastActiveDate = Get-Date ($LastActiveDate) -format 'dd-MMM-yyyy HH:mm:ss'
    } Else {
        $LastActiveDate = "N/A"
    }

    $Reportline = [pscustomobject]@{
        Name                = $M.DisplayName 
        UPN                 = $M.UserPrincipalName
        PrimarySMTPAddress  = $M.PrimarySMTPAddress
        Active              = $ActiveStatus 
        'Last sent message' = $LastActiveDate
    }
    $Report.Add($ReportLine)
    $Text = ("Mailbox state checked on {0} and determined as {1}. Last message addressed on {2}" `
         -f (Get-Date -format g), $ActiveStatus, $LastActiveDate )
      # This line updates the mailbox with details of the assessment. Comment it out if you don't want to do this
    Set-Mailbox -Identity $M.Alias -CustomAttribute15 $Text
}

Write-Host ""
Write-Host ("Total mail users checked: {0}" -f $Mbx.count)
Write-Host ("Active mail users:        {0}" -f $ActiveStatusCount)
Write-Host ("Inactive mail users:      {0}" -f ($Mbx.count - $ActiveStatusCount))
Write-Host ""

# Generate reports
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $OutputXLSXFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveMailUsers.xlsx"
    $Report | Export-Excel -Path $OutputXLSXFile -WorksheetName "Inactive Mail Users Report" -Title ("Inactive Mail Users Report{0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "InactiveMailUsers" 
} Else {
    $OutputCSVFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveMailUsers.csv"
    $Report | Export-Csv -Path $OutputCSVFile -NoTypeInformation -Encoding Utf8
}
  
If ($ExcelGenerated) {
  Write-Host ("An Excel report is available in {0}" -f $OutputXLSXFile)
} Else {
  Write-Host ("A CSV report is available in {0}" -f $OutputCSVFile)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.