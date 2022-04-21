# ReportArchivedTeams.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportArchivedTeams.PS1
# Script to show how to report Teams archiving actions extracted from the Office 365 audit log
# Needs a connection to the Exchange Online management and Teams modules

$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
If (!($ModulesLoaded -match "MicrosoftTeams")) {Write-Host "Please connect to the Microsoft Teams module and then restart the script"; break}
# Some lines to create the c:\temp\ directory. Comment them out if you don't want this to happen
$Path = "C:\Temp"
If(!(test-path $path)) {
    New-Item -ItemType Directory -Force -Path $path | Out-Null }
# Define output CSV file
$OutputCsvFile = "c:\temp\TeamsArchiveAuditRecords.csv"

CLS; Write-Host "Searching Audit Records to find Teams archiving operations"
$StartDate = (Get-Date).AddDays(-90); $EndDate = (Get-Date) 
[array]$Records = (Search-UnifiedAuditLog -Operations TeamSettingChanged -StartDate $StartDate -EndDate $EndDate -ResultSize 5000)
If ($Records.Count -eq 0) {
    Write-Host "No audit records for team archiving found." }
Else {
    Write-Host "Processing" $Records.Count "team archiving audit records..."
    $Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
    ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      If ($AuditData.Name -eq "Team is archived") {  # It's an archival team setting record
        Switch ($AuditData.NewValue) {
         "False"  { $Action = "Restored team" }
         "True" { $Action = "Archived team" }
        } #end switch

         Write-Host "Checking channels for" $AuditData.TeamName
         $TeamId = (Get-Team -DisplayName $AuditData.TeamName).GroupId
         If ($TeamId) {
            [array]$TeamChannels = Get-TeamChannel -GroupId $TeamId
            [array]$StandardChannels = $TeamChannels | ? {$_.Membershiptype -eq "Standard"}
            [array]$SharedChannels = $TeamChannels | ? {$_.Membershiptype -eq "Shared"}
            [array]$PrivateChannels = $TeamChannels | ? {$_.Membershiptype -eq "Private"}
            $ChannelStatus = "OK" }
         Else {
            $ChannelStatus = "Team might be deleted" }

         $ReportLine = [PSCustomObject] @{
           TimeStamp = Get-Date($AuditData.CreationTime) -format g
           User      = $AuditData.UserId
           Team      = $AuditData.TeamName
           Action    = $Action 
           Channels  = $TeamChannels.Count 
           Regular   = $StandardChannels.Count
           Private   = $PrivateChannels.Count 
           Shared    = $SharedChannels.Count 
           Status    = $ChannelStatus
           }       
         $Report.Add($ReportLine) 
      } #end if
   } #end foreach
} #end else

$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation $OutputCsvFile
Write-Host ("All done. {0} audit events output to {1}" -f $Report.Count, $OutputCsvFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
