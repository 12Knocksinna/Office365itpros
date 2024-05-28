# AuditRecordsTeamsMeetings.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/AuditRecordsTeamsMeetings.PS1
# A script to show how to find and interpret audit records captured for Teams meetings (advanced auditing)
# Updated 24-May-2024 to use the Microsoft Graph PowerShell SDK

$ModulesLoaded = Get-Module | Select-Object -ExpandProperty Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {
    Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break
}
Connect-MgGraph -NoWelcome -Scopes User.Read.All

# Start and end date for the audit scan. By default, we look for 90 days, but you can choose any value you like up to 365 (if you have Office 365 E5)
$StartDate = (Get-Date).AddDays(-90); $EndDate = (Get-Date).AddDays(1) # Set your own date span here!
$OutputCSVFile = "C:\temp\AuditEventsTeamsMeetings.csv"

# Find the audit records
Write-Host "Looking for Teams meeting audit records..."
[array]$Records = Search-UnifiedAuditLog -Operations MeetingDetail, MeetingParticipantDetail -StartDate $StartDate -EndDate $EndDate -Formatted `
   -ResultSize 5000 -SessionCommand ReturnLargeSet
If (!($Records)) {
   Write-Host "No audit records found - exiting!"; break
} Else {
   $Records = $Records | Sort-Object Identity -Unique | Sort-Object { $_.CreationDate -as [datetime]} -Descending
}
Write-Host "Processing" $Records.Count "Teams meeting audit records..."
# Process the records
$MeetingRecords = [System.Collections.Generic.List[Object]]::new() 	
ForEach ($Rec in $Records) {
   $AuditData = $Rec.AuditData | ConvertFrom-Json
   $User = $Null; $Organizer = $Null
   Switch ($Rec.Operations) {
       "MeetingDetail" { # A meeting record     
      [datetime]$StartTime = Get-Date($AuditData.StartTime)
      [datetime]$EndTime = Get-Date($AuditData.EndTime)   
      $TimeSpent = $EndTime - $StartTime 
      $MeetingDuration = ("{0:hh\:mm\:ss}" -f $TimeSpent)
      $Organizer = (Get-MgUser -UserId $AuditData.Organizer.UserObjectId).DisplayName
      $DataLine = [PSCustomObject] @{
         'Audit timestamp'   = Get-Date($Rec.CreationDate).ToLocalTime()
         User                = $Rec.UserIds
         MeetingId           = $AuditData.Id
         Start               = Get-Date($AuditData.StartTime).ToLocalTime()
         End                 = Get-Date($AuditData.EndTime).ToLocalTime()
         'Meeting duration'  = $MeetingDuration.ToString()
         Organizer           = $Organizer
         Modalities          = $AuditData.Modalities
         'Meeting type'      = $AuditData.CommunicationSubType
         Type                = "Meeting"
         MeetingURL          = $AuditData.MeetingURL
         Operation           = $Rec.Operations 
      }
     }
     "MeetingParticipantDetail" { # A meeting participant record   
         # Resolve user name from the object identifier logged for participant
         [datetime]$StartTime = Get-Date($AuditData.JoinTime)
         [datetime]$EndTime = Get-Date($AuditData.LeaveTime)   
         $TimeSpent = $EndTime - $StartTime 
         $MeetingDuration = ("{0:hh\:mm\:ss}" -f $TimeSpent)
      If ($AuditData.Attendees.RecipientType -eq "User") {
            $User = (Get-MgUser -UserId $AuditData.Attendees.UserObjectid).UserPrincipalName }
      Else {
            $User = $AuditData.Attendees.DisplayName }
      If ($User -eq "b1902c3e-b9f7-4650-9b23-5772bd429747") { 
            $User = "Teams Meeting Recording Bot" 
      }
       $DataLine = [PSCustomObject] @{
         'Audit timestamp'   = $Rec.CreationDate
         User                = $User
         MeetingId           = $AuditData.MeetingDetailId
         Start               = Get-Date($AuditData.JoinTime).ToLocalTime()
         End                 = Get-Date($AuditData.LeaveTime).ToLocalTime()
         Duration            = $MeetingDuration.ToString()
         Role                = $AuditData.Attendees.Role
         DetailId            = $AuditData.MeetingDetailId
         Artifacts           = $AuditData.ArtifactsShared.ArtifactSharedName -join ", "
         UserInfo            = $AuditData.ExtraProperties.Value
         Type                = "Participant"
         Operation           = $Rec.Operations }
     }
    } # End Switch
    $MeetingRecords.Add($DataLine) 

} #End For

$MeetingRecords |  Sort-Object {$_.Date -as [datetime]}, MeetingId, Operation | `
   Select-Object Start, End, User, MeetingType, Organizer, Type, MeetingId | Out-GridView
$MeetingRecords |  Sort-Object {$_.Date -as [datetime]}, MeetingId, Operation | Export-CSV -NoTypeInformation $OutputCSVFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
