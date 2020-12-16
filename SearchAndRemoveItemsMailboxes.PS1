#
# SearchAndRemoveItemsMailboxes.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/SearchAndRemoveItemsMailboxes.PS1
# A script to use the Search-Mailbox cmdlet to remove calendar items from user mailboxes.
# Requires the Exchange Online management module

CLS
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script" ; break}

# Set up the search query
$StartDate = "1-Jan-2019"
$EndDate = "1-Dec-2019"
$Query = "Received:$($StartDate)..$($EndDate) kind:meetings" + ' AND (-subject:"Thread Id")'

# Find the mailboxes to process - this example uses a check against the custom attribute 12. You could also read in user details from a CSV file
$Users = Get-ExoMailbox -Filter {CustomAttribute12 -eq "Search"} -Properties CustomAttribute12 -RecipientTypeDetails UserMailbox -ResultSize Unlimited
If (!$Users) {Write-Host "No matching users found - exiting" ; break}

$UserReport = [System.Collections.Generic.List[Object]]::new() # Create output file 
ForEach ($User in $Users)  {
    $Status = (Search-Mailbox -Identity $User.UserPrincipalName-SearchQuery $Query -EstimateResultOnly -DoNotIncludeArchive -SearchDumpster:$False)
    If ($Status) {
      $ReportLine = [PSCustomObject] @{
           UserPrincipalName  = $User.UserPrincipalName
           DisplayName        = $User.DisplayName
           ItemsFound         = $Status.ResultItemsCount
           ItemsSize          = $Status.ResultItemsSize
           SearchType         = "Estimate"
           SearchTime         = Get-Date }                
     $UserReport.Add($ReportLine) } #End if
} # End For

# Filter the users where we have found some items
$ProcessUsers = $UserReport | ? {$_.ItemsFound -ne "0"}
CLS
$ProcessUsers | Format-Table DisplayName, UserPrincipalName, ItemsFound, ItemsSize -AutoSize
$PromptTitle = 'Remove items from mailboxes'
$PromptMessage = 'Please confirm whether to proceed to remove found items from mailboxes'
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&yes", 'yes?'
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&no", 'no?'
$cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&cancel", 'Exit'
$PromptOptions = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $cancel)
$PromptDecision = $host.ui.PromptForChoice($PromptTitle, $PromptMessage, $PromptOptions, 0) 

$i = 0
Switch ($PromptDecision) {
  "0" { 
       ForEach ($User in $ProcessUsers) {
         Write-Host "Removing items from the mailbox of" $User.DisplayName
         $Status = (Search-Mailbox -Identity $User.UserPrincipalName -SearchQuery $Query -DeleteContent -DoNotIncludeArchive -SearchDumpster:$False -Confirm:$False -Force)
         If ($Status) {  # Add record to capture what we did
             Write-Host "Mailbox for" $User.DisplayName "processed to remove" $Status.ResultItemsCount "items"
             $ReportLine = [PSCustomObject] @{
                 UserPrincipalName  = $User.UserPrincipalName
                 DisplayName        = $User.DisplayName
                 ItemsFound         = $Status.ResultItemsCount
                 ItemsSize          = $Status.ResultItemsSize
                 SearchType         = "Removal"
                 SearchTime         = Get-Date }                 
             $ProcessUsers.Add($ReportLine)
             $i++
            }
       } #End ForEach
      Write-Host "All done." $i "mailboxes processed and cleaned up. Details stored in c:\temp\SearchMailboxRemovals.csv"
      $ProcessUsers | Export-CSV -NoTypeInformation c:\temp\SearchMailboxRemovals.csv
      }
  "1" {
      Write-Host "OK. Maybe later? Messages not removed from mailboxes"
      }
  "2" {
      Write-Host "Cancelled. Messages not removed from mailboxes"
      }
} #End Switch
