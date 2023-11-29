# SetAutoReplyTeamsGroups.PS1
# https://github.com/12Knocksinna/Office365itpros/edit/master/SetAutoReplyTeamsGroups.PS1
# An example of how to to set an auto-reply on team-enabled Microsoft 365 groups
# Uses the Microsoft Graph PowerShell SDK (uses the Directory.Read.All and ChannelSettings.Read.All permissions) and the Exchange Online Management module.

Clear-Host
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {
   Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break
}
# Check that we're connected to the Microsoft Graph and if not, connect...
$Status = (Get-MgContext).Account
If (!($Status)) {
   Connect-MgGraph -Scopes Directory.Read.All, ChannelSettings.Read.All, Team.ReadBasic.All -NoWelcome
}

Write-Host "Finding team-enabled Microsoft 365 Groups..."
[array]$Teams = Get-MgTeam -All
If (!($Teams)) {
      Write-Host "No teams found - exiting"; break
   } Else { 
      $Teams = $Teams.Value 
}

Write-Host ("{0} team-enabled groups found... now to check for email-enabled channels." -f $Teams.count)
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Team in $Teams) { # Need to do it this way because Get-MgTeamChannel returns the group email address for the General channel
    Write-Host "Processing" $Team.DisplayName
    $Uri = "https://graph.microsoft.com/v1.0/teams/$($team.id)/channels"
    [array]$Channels = Invoke-MgGraphRequest -Uri $Uri -Method Get
    ForEach ($Channel in $Channels.Value) {
       If (!([string]::IsNullOrWhiteSpace($Channel.Email))) {
        $ReportLine = [PSCustomObject]@{ 
          Team    = $Team.DisplayName
          Channel = $Channel.DisplayName
          Email   = $Channel.Email
          TeamId  = $Team.Id } 
        $Report.Add($ReportLine)  } 
    } # End for each channel
} # End for each team

[array]$GeneralChannels = $Report | Where-Object {$_.Channel -eq "General"}

ForEach ($Team in $Teams) {
   $EmailAddress = $Null
   $EmailAddress = $GeneralChannels | Where-Object {$_.TeamId -eq $Team.Id}| Select-Object -ExpandProperty Email
   If ($EmailAddress) { 
     Write-Host "Setting auto reply with mailto: for" $Team.DisplayName 
     $InternalMessage = 'Please! We use Teams for communication, so send your message to <a href = "mailto:' + $EmailAddress + '">Teams</a> and it will be dealt with there.'
     Set-MailboxAutoReplyConfiguration -Identity $Team.Id -ExternalMessage "Sorry, this mailbox doesn't accept email" `
      -AutoReplyState Enabled -InternalMessage $InternalMessage 
   } Else {
      Write-Host ("Using default auto-reply for {0}" -f $Team.DisplayName)
      Set-MailboxAutoReplyConfiguration -Identity $Team.Id `
         -ExternalMessage "Sorry, this mailbox doesn't accept email" -ExternalAudience All `
         -AutoReplyState Enabled -InternalMessage "Please use Teams to communicate with us" 
   }
} 
# End For each team

# Comment these lines out if you don't want the script to create a temp directory to store its output files
$path = "C:\Temp"
If(!(test-path $path)) {
   New-Item -ItemType Directory -Force -Path $path | Out-Null }
$Report | Export-CSV -NoTypeInformation c:\temp\ListofEmailEnabledTeamChannels.csv
Write-Host "A listing of email-enabled teams channels is in c:\temp\ListofEmailEnabledTeamChannels.csv"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
