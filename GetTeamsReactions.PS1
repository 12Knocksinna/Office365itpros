# GetTeamsReactions.PS1
# Report teams reactions found in messages
# https://github.com/12Knocksinna/Office365itpros/blob/master/GetTeamsReactions.PS1

$Modules = @( "ExchangeOnlineManagement" )
# Requires -Modules $Modules

$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online management module and then restart the script"; break}

Connect-MgGraph -Scope Channel.ReadBasic.All

Write-Host "Finding audit records for teams reactions..."
$StartDate = (Get-Date).AddDays(-90)
$EndDate = (Get-Date).AddDays(1)

[array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Formatted -ResultSize 5000 -Operations ReactedToMessage
If (!($Records)) {Write-Host "Error: Can't find any audit records to process - exiting" ; break}
Write-Host ("Processing {0} audit records..." -f $Records.count)
$Report = [System.Collections.Generic.List[Object]]::new()
$i = 0
ForEach ($Rec in $Records) {
  $i++
  Write-Host ("Processing record {0} of {1}" -f $i, $Records.count)
  $AuditData = $Rec.AuditData | ConvertFrom-Json
  If ($AuditData.ChannelName) { # Reaction to channel message
     $Type = "Channel"
     $Team = Get-MgTeam -TeamId $AuditData.AADGroupId | Select-Object -ExpandProperty DisplayName
     $Channel = Get-MgTeamChannel -TeamId $AuditData.AADGroupId | Where-Object {$_.id -eq $AuditData.ChannelGuid} | Select-Object -ExpandProperty DisplayName
     $MessageThread = $AuditData.TeamGuid
     $MessageId = $AuditData.MessageId
  }
  Else {
     $Team = $Null; $Channel = $Null
     $Type = "Chat"  
     $Channel = $AuditData.CommunicationType  
     $MessageThread = $AuditData.ChatThreadId
     $MessageId = $AuditData.MessageId
  } # End If
  
  # Check for reactions removed
 If (!($AuditData.MessageReactionType)) {
    $AuditText = $Auditdata.Extraproperties | Where-Object {$_.Key -eq "Source"} | Select-Object -ExpandProperty Value
    $Reaction = ("{0} ({1})" -f $AuditText, $AuditData.OldValue) 
 } Else {
    $Reaction =  $AuditData.MessageReactionType }

 $ReportLine = [PSCustomObject][Ordered]@{  
     Date             = $Rec.CreationDate
     User             = $Rec.UserIds
     Reaction         = $Reaction
     Type             = $Type
     Team             = $Team
     Channel          = $Channel
     Thread           = $MessageThread
     Message          = $MessageId
      }
   $Report.Add($ReportLine)

} # End ForEach Records

$Report | Group-Object Reaction | Sort-Object Count -Descending | Format-Table Name, Count
$Report | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
