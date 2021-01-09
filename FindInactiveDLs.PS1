# FindInactiveDls
# Find inactive distribution lists based on the message trace informnation, which means we can only go back 7 days...
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindInactiveDLs.PS1
$EndDate = Get-Date
$StartDate = $EndDate.AddDays(-10)
$Messages = $null 
$Page = 1 
Write-Host "Collecting message trace data for the last 10 days"
Do 
{ 
    $CurrMessages = (Get-MessageTrace -Status Expanded -PageSize 5000 -Page $Page -StartDate $StartDate -EndDate $EndDate | Select Received, RecipientAddress)
    $Page++ 
    $Messages += $CurrMessages 
} 
Until ($CurrMessages -eq $Null) 

$MessageTable = @{}
$Messagetable = ($Messages | Sort RecipientAddress -Unique | Select RecipientAddress, Received)
$DLs = Get-DistributionGroup -ResultSize Unlimited
Write-Host "Processing" $DLs.Count "distribution lists..."
$Results = ForEach ($DL in $DLs) {
       If ($MessageTable -Match $DL.PrimarySMTPAddress) {
          [pscustomobject]@{Name = $DL.DisplayName ; Active = "Yes"}
          Write-Host $DL.DisplayName "is active" -Foregroundcolor Yellow }
       Else {
          [pscustomobject]@{Name = $DL.DisplayName ; Active = "No"}
          Write-Host $DL.DisplayName "inactive" -Foregroundcolor Red }
}

$Results | Export-CSV c:\Temp\ListofDLs.csv -NoTypeInformation

