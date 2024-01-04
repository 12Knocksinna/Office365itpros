# PurgeMessagesWithSearchMailbox.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/PurgeMessagesWithSearchMailbox.PS1
#
# A script to purge messages from Exchange Online using the famous Search-Mailbox cmdlet. The script can either delete items or report an estimate
# V1.0 September 2021
# V1.1 4-Jan-2024
# ---------------------------------------
# Runs like .\PurgeMessagesWithSearchMailbox.PS1 -DeleteItems [$True/$False]
# $True will force Search-Mailbox to remove items from Exchange Online mailboxes, $False will run an estimate search against the mailboxes

Param ([parameter (mandatory)] [ValidateSet($True,$False)] $DeleteItems )

If ($DeleteItems -ne $True -and $DeleteItems -ne $False) { Write-Host "No mode specified - please rerun" ; break }

$Status = Get-ConnectionInformation -ErrorAction SilentlyContinue
If (!($Status)) {
  Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

Clear-Host
Switch ($DeleteItems) {
  $True {
     $Mode = "delete"
     $SearchMessage = "Searching mailbox to delete items" }
  $False {
     $Mode = "estimate"
     $SearchMessage = "Searching mailbox to estimate items" }
} #EndSwitch

# Some information to identify the messages we want to purge
$BadSender = "no-reply@microsoft.com"
# Date range for the search - make this as precise as possible
$StartDate = "1-Jan-2019"
$EndDate = "13-Sep-2024"
$BodyText = "Your month in review"
$Subject = "MyAnalytics"

$SearchQuery = "From:" + $BadSender + " Sent:" + $StartDate + ".." + $EndDate
If ([string]::IsNullOrWhiteSpace($BodyText) -eq $False) { $SearchQuery = $SearchQuery + ' Body: "*' + $BodyText + '*"' }
If ([string]::IsNullOrWhiteSpace($Subject) -eq $False) { $SearchQuery = $SearchQuery + ' Subject:"*' + $Subject + '*"' }

Write-Host ("Finding user and shared mailboxes to run query {0} against in (1) mode" -f $SearchQuery, $Mode)
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox, SharedMailbox -ResultSize Unlimited
Clear-Host
$CSVOutput = "C:\temp\ExoSearchRemovals.CSV"
$Report = [System.Collections.Generic.List[Object]]::new();
$Successes=0
$Failures=0
$TotalItems=0
$ProgDelta = 100/($Mbx.Count); $CheckCount = 0; $MbxNumber = 0
ForEach ($M in $Mbx) {
   $MbxNumber++
   $MbxStatus = $M.DisplayName + " ["+ $MbxNumber +"/" + $Mbx.Count + "]"
   Write-Progress -Activity $SearchMessage -Status $MbxStatus -PercentComplete $CheckCount
   $CheckCount += $ProgDelta

   If ($DeleteItems -eq $True) { #Go ahead and delete items
      $SearchOutput = Search-Mailbox -Identity $M.ExternalDirectoryObjectId -SearchQuery $SearchQuery -DeleteContent -Force -WarningAction SilentlyContinue -SearchDumpster
    }
    ElseIf ($DeleteItems -eq $False) { # Estimate Search
      $SearchOutput = Search-Mailbox -Identity $M.ExternalDirectoryObjectId -SearchQuery $SearchQuery -EstimateResultOnly -Force -WarningAction SilentlyContinue -SearchDumpster
   }
   Switch ($SearchOutput.Success) { 
     "True" { # Success
        $Successes++
        $TotalItems = $TotalItems + $SearchOutput.ResultItemsCount 
        $ReportLine = [PSCustomObject][Ordered]@{ 
          Name     = $M.DisplayName
          UPN      = $M.UserPrincipalName
          ObjectId = $M.ExternalDirectoryObjectId
          Items    = $SearchOutput.ResultItemsCount 
          ItemSize = $SearchOutput.ResultItemsSize.Substring(0,$SearchOutput.ResultItemsSize.IndexOf("(")) 
          Status   = "Success" 
          Mode     = $Mode 
          Date     = (Get-Date -format g)}
     $Report.Add($ReportLine) }
     "False" { # Whoops!
        $Failures++
        $ReportLine = [PSCustomObject][Ordered]@{ 
          Name     = $M.DisplayName
          UPN      = $M.UserPrincipalName
          ObjectId = $M.ExternalDirectoryObjectId
          Items    = 0 
          ItemSize = 0 
          Status   = "Failure"
          Mode     = $Mode 
          Date     = (Get-Date -format g)}
       $Report.Add($ReportLine) }
     } # End Switch
    
} # End ForEach
Switch ($DeleteItems) {
   $True   { Write-Host ("{0} mailboxes processed and {1} items removed." -f $Mbx.Count, $TotalItems) }
   $False  { Write-Host ("{0} mailboxes processed and {1} items found. " -f $Mbx.Count, $TotalItems) }
}
$Report | Where-Object {$_.Items -gt 0} | Out-GridView
Export-CSV -Notypeinformation $CSVOutput 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
