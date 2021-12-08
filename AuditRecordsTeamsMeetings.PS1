# AuditRecordsTeamsMeetings.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/AuditRecordsTeamsMeetings.PS1
# A script to show how to find and interpret audit records captured for Teams meetings (advanced auditing)
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
If (!($ModulesLoaded -match "AzureAD")) { Write-Host "Please connect to the Azure AD or Azure AD Preview module and then restart the script"; break }

# Start and end date for the audit scan. By default, we look for 39 days, but you can choose any value you like up to 365 (if you have Office 365 E5)
$StartDate = (Get-Date).AddDays(-30); $EndDate = (Get-Date) # Set your own date span here!
$OutputCSVFile = "C:\temp\AuditEventsTeamsMeetings.csv"

# Find the audit records
[array]$Records = (Search-UnifiedAuditLog -Operations MeetingDetail, MeetingParticipantDetail -StartDate $StartDate -EndDate $EndDate -Formatted -ResultSize 5000)

If (!($Records)) {Write-Host "No audit records found - exiting!"; break}

$MeetingRecords = [System.Collections.Generic.List[Object]]::new() 	
ForEach ($Rec in $Records) {
   $AuditData = $Rec.AuditData | ConvertFrom-Json
   $User = $Null; $Organizer = $Null
   Switch ($Rec.Operations) {
    "MeetingDetail" { # A meeting record     
      [datetime]$StartTime = Get-Date($AuditData.StartTime)
      [datetime]$EndTime = Get-Date($AuditData.EndTime)   
      $TimeSpent = $EndTime - $StartTime 
      $Organizer = (Get-AzureADUser -ObjectId $AuditData.Organizer.UserObjectId).DisplayName
      $DataLine = [PSCustomObject] @{
         Date                = $Rec.CreationDate
         User                = $Rec.UserIds
         MeetingId           = $AuditData.Id
         MeetingType         = $AuditData.ItemName
         Start               = $AuditData.StartTime
         End                 = $AuditData.EndTime
         MeetingTime         = $TimeSpent.ToString()
         Organizer           = $Organizer
         Modalities          = $AuditData.Modalities
         MeetingURL          = $AuditData.MeetingURL
         Operation           = $Rec.Operations }
     }
     "MeetingParticipantDetail" { # A meeting participant record   
       [datetime]$StartTime = Get-Date($AuditData.JoinTime)
       [datetime]$EndTime = Get-Date($AuditData.LeaveTime)   
       $TimeSpent = $EndTime - $StartTime
       # Resolve user name from the object identifier logged for participant
       If ($AuditData.Attendees.RecipientType -eq "User") {
          $User = (Get-AzureADUser -ObjectId $AuditData.Attendees.UserObjectid).UserPrincipalName }
       Else {
          $User = $AuditData.Attendees.DisplayName }
       If ($User -eq "bdd75849-e0a6-4cce-8fc1-d7c0d4da43e5") { $User = "Teams Meeting Recording Bot" }
       $DataLine = [PSCustomObject] @{
         Date                = $Rec.CreationDate
         User                = $User
         MeetingId           = $AuditData.MeetingDetailId
         MeetingType         = $AuditData.ItemName
         Start               = $AuditData.JoinTime
         End                 = $AuditData.LeaveTime
         "User Time"         = $TimeSpent.ToString()
         Role                = $AuditData.Attendees.Role
         DetailId            = $AuditData.MeetingDetailId
         Artifacts           = $AuditData.ArtifactsShared.ArtifactSharedName -join ", "
         UserInfo            = $AuditData.ExtraProperties.Value
         Operation           = $Rec.Operations }
     }
    } # End Switch
    $MeetingRecords.Add($DataLine) 

} #End For

$MeetingRecords | Sort MeetingId, Operation | Out-GridView
$MeetingRecords | Sort MeetingId, Operation | Export-CSV -NoTypeInformation $OutputCSVFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
