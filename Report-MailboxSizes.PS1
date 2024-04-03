# Report-MailboxSizes.PS1
# Script to demonstrate how to report mailbox statistics using Graph usage data
# V1.0 20-Nov-2023
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-MailboxSizes.PS1

Connect-ExchangeOnline
Connect-MgGraph -NoWelcome -Scopes Reports.Read.All

# Before doing anything, we need to make sure that the Graph returns non-obfuscated user data
$ObfuscatedReset = $False
If ((Get-MgBetaAdminReportSetting).DisplayConcealedNames -eq $True) {
    $Parameters = @{ displayConcealedNames = $False }
    Update-MgBetaAdminReportSetting -BodyParameter $Parameters
    $ObfuscatedReset = $True
}

Write-Host "Fetching Microsoft 365 usage data"
Get-MgReportEmailActivityUserDetail -Period 'D180' -Outfile EmailActivity.CSV 
[array]$EmailActivityData = Import-CSV EmailActivity.CSV
Get-MgReportMailboxUsageDetail -Period 'D180' -Outfile MailboxUsage.CSV
[array]$MailboxUsage = Import-CSV MailboxUsage.CSV

[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited
Write-Host ("Processing {0} mailboxes..." -f $Mbx.count)
$OutputReport = [System.Collections.Generic.List[Object]]::new()

ForEach ($M in $Mbx) {
  $MbxSize = $Null
  $UserStats = $EmailActivityData | Where-Object {$_.'User Principal Name' -eq $M.UserPrincipalName}
  $MailboxStats = $MailboxUsage | Where-Object {$_.'User Principal Name' -eq $M.UserPrincipalName}
  If ([string]::IsNullOrWhiteSpace($MailboxStats.'Last Activity Date')) {
     $LastActivityDate = $Null
     $DaysSinceActivity = $Null
  } Else {              
     $DaysSinceActivity = (New-TimeSpan $MailboxStats.'Last Activity Date').Days
     $LastActivityDate = Get-Date ($MailboxStats.'Last Activity Date') -format 'dd-MMM-yyyy'
  }    
  [long]$MbxStorage =  $MailboxStats.'Storage Used (Byte)'  
  If ($MbxStorage -le 1GB) {
       $MailboxSize = [Math]::Round(($MbxStorage/1MB), 2)
       [string]$MbxSize = $MailboxSize.toString() + " MB"             
  } Else {
       $MailboxSize = [Math]::Round(($MbxStorage/1GB), 2)
       [string]$MbxSize = $MailboxSize.toString() + " GB"      
  }
   
  $ReportLine = [PSCustomObject]@{
    UPN               = $M.UserPrincipalName
    Name              = $M.DisplayName
    Items             = $MailboxStats.'ItemCount'
    Size              = $MbxSize
    LastActivity      = $LastActivityDate
    DaysSinceActivity = $DaysSinceActivity 
    'Send Count'      = $UserStats.'Send Count'
    'Receive Count'   = $UserStats.'Receive Count'
   } 
   $OutputReport.Add($ReportLine)
 }
$OutputReport | Format-Table Name, UPN, Items, Size, LastActivity 

# Now reset the obfuscated user data setting if necessary

If ($ObfuscatedReset -eq $True) {
    $Parameters = @{ displayConcealedNames = $True }
    Update-MgBetaAdminReportSetting -BodyParameter $Parameters
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
