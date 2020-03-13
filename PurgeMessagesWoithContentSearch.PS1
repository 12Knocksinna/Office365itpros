# PurgeMessagesWoithContentSearch.PS1
#
# A script to purge messages from Exchange Online using an Office 365 content search and a purge action applied to 
# the results of that search.
# V1.0 March 2020
# ---------------------------------------
# Some information to identify the messages we want to purge
$Sender = "admin@rolysbistro.ie"
$Subject = "Special Offer from Roly's Bistro"

CLS
If (Get-ComplianceSearch -Identity "Remove Offensive Information") {
   Write-Host "Cleaning up old search"
   Try {
      Remove-ComplianceSearch -Identity "Remove Offensive Information" -Confirm:$False -ErrorAction SilentlyContinue } 
   Catch {
       Write-Host "We can't clean up the old search" ; break }}

New-ComplianceSearch -Name "Remove Offensive Information" -ContentMatchQuery "(Received:6-Mar-2020..14-Mar-2020) AND (From:$Sender) AND (Subject:$Subject)" -ExchangeLocation All -AllowNotFoundExchangeLocationsEnabled $True | Out-Null
                                                                            
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
