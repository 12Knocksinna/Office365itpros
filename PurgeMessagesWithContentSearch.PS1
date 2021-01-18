# PurgeMessagesWithContentSearch.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/PurgeMessagesWithContentSearch.PS1
#
# A script to purge messages from Exchange Online using an Office 365 content search and a purge action applied to 
# the results of that search.
# V1.0 March 2020
# ---------------------------------------
# Some information to identify the messages we want to purge
$Sender = "badactor@badboys.com"
$Subject = "Special Offer for you"
# Date range for the search - make this as precise as possible
$StartDate = "10-Mar-2020"
$EndDate = "13-Mar-2020"
$Start = (Get-Date $StartDate).ToString('yyyy-MM-dd')   
$End = (Get-Date $EndDate).ToString('yyyy-MM-dd')
$ContentQuery = '(c:c)(received=' + $Start + '..' + $End +')(senderauthor=' + $Sender + ')(subjecttitle="' + $Subject + '")'

CLS
If (Get-ComplianceSearch -Identity "Remove Offensive Information") {
   Write-Host "Cleaning up old search"
   Try {
      $Status = Remove-ComplianceSearch -Identity "Remove Offensive Information" -Confirm:$False  } 
   Catch {
       Write-Host "We can't clean up the old search" ; break }}

New-ComplianceSearch -Name "Remove Offensive Information" -ContentMatchQuery $ContentQuery -ExchangeLocation All -AllowNotFoundExchangeLocationsEnabled $True | Out-Null
                                                                            
Write-Host "Starting Search..."
Start-ComplianceSearch -Identity "Remove Offensive Information" | Out-Null
$Seconds = 0
While ((Get-ComplianceSearch -Identity "Remove Offensive Information").Status -ne "Completed") {
  $Seconds++
  Write-Host "Still searching... (" $Seconds ")"
   Sleep -Seconds 1 }
$ItemsFound = (Get-ComplianceSearch -Identity "Remove Offensive Information").Items

If ($ItemsFound -gt 0) {
   $Stats = Get-ComplianceSearch -Identity "Remove Offensive Information" | Select -Expand SearchStatistics | Convertfrom-JSON
   $Data = $Stats.ExchangeBinding.Sources |?{$_.ContentItems -gt 0}
   Write-Host ""
   Write-Host "Total Items found matching query:" $ItemsFound 
   Write-Host ""
   Write-Host "Items found in the following mailboxes"
   Write-Host "--------------------------------------"
   Foreach ($D in $Data)  {Write-Host $D.Name "has" $D.ContentItems "items of size" $D.ContentSize }
   Write-Host " "
   $Iterations = 0; $ItemsProcessed = 0
   While ($ItemsProcessed -lt $ItemsFound) {
       $Iterations++
       Write-Host "Deleting items... (" $Iterations ")"
       New-ComplianceSearchAction -SearchName "Remove Offensive Information" -Purge -PurgeType HardDelete -Confirm:$False | Out-Null
       While ((Get-ComplianceSearchAction -Identity "Remove Offensive Information_Purge").Status -ne "Completed") { # Let the search action complete
           Sleep -Seconds 2 }
       $ItemsProcessed = $ItemsProcessed + 10 # Can remove a maximum of 10 items per mailbox
       # Remove the search action so we can recreate it
       Remove-ComplianceSearchAction -Identity "Remove Offensive Information_Purge" -Confirm:$False  }}
  Else {
       Write-Host "No items found" }

Write-Host "All done!"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
