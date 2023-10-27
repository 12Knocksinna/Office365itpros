# FindInactiveDls
# Find inactive distribution lists based on the message trace informnation, which means we can only go back 7 days...
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindInactiveDLs.PS1

# Updated 27-Oct-2023

$EndDate = Get-Date
$StartDate = $EndDate.AddDays(-10)
$CSVFile = "c:\Temp\ListofDLs.csv"
[array]$Messages = $Null 
$Report = [System.Collections.Generic.List[Object]]::new()
$Page = 1 
Write-Host "Collecting message trace data for the last 10 days"
Do 
{ 
    [array]$CurrentMessages = (Get-MessageTrace -Status Expanded -PageSize 5000 -Page $Page `
       -StartDate $StartDate -EndDate $EndDate | Select-Object Received, RecipientAddress)
    $Page++ 
    $Messages += $CurrentMessages 
} 
Until ($Null -eq $CurrentMessages) 

$MessageTable = @{}
$Messagetable = ($Messages | Sort-Object RecipientAddress -Unique | Select-Object RecipientAddress, Received)
[array]$DLs = Get-DistributionGroup -ResultSize Unlimited
Write-Host ("Processing {0} distribution lists..." -f $DLs.count)
ForEach ($DL in $DLs) {
      If ($MessageTable -Match $DL.PrimarySMTPAddress) {
         $ActiveStatus = "Active"
         Write-Host ("{0} is active" -f $DL.DisplayName) -Foregroundcolor Yellow 
      } Else {
         $ActiveStatus = "Inactive"
         Write-Host ("{0} is inactive" -f $DL.DisplayName) -Foregroundcolor Red 
      }
      $Reportline = [pscustomobject]@{
         Name = $DL.DisplayName 
         Active = $ActiveStatus 
      }
      $Report.Add($ReportLine)
      $Text = ("DL state checked on {0} and determined as {1}" -f (Get-Date -format g), $ActiveStatus )
      # This line updates the DL with details of the assessment. Comment it out if you don't want to do this
      Set-DistributionGroup -Identity $DL.Alias -CustomAttribute15 $Text
}

$Report | Export-CSV $CSVFile -NoTypeInformation

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
