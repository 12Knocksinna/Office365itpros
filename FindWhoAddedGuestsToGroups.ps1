# A script to find and report who added new guest members to Microsoft 365 Groups (and Teams) over the last 90 days
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindWhoAddedGuestsToGroups.ps1
$EndDate = (Get-Date).AddDays(1); $StartDate = (Get-Date).AddDays(-90); $NewGuests = 0
$Records = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "Add Member to Group" -ResultSize 2000 -Formatted)
If ($Records.Count -eq 0) {
   Write-Host "No Group Add Member records found." }
 Else {
   Write-Host "Processing" $Records.Count "audit records..."
   $Report = @()
   ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      # Only process the additions of guest users to groups
      If ($AuditData.ObjectId -Like "*#EXT#*") {
         $TimeStamp = Get-Date $Rec.CreationDate -format g
         # Try and find the timestamp when the Guest account was created in AAD
         Try {$AADCheck = (Get-Date(Get-AzureADUser -ObjectId $AuditData.ObjectId).RefreshTokensValidFromDateTime -format g) }
           Catch {Write-Host "Azure Active Directory record for" $AuditData.ObjectId "no longer exists" }
         If ($TimeStamp -eq $AADCheck) { # It's a new record, so let's write it out
            $NewGuests++
            $ReportLine = [PSCustomObject][Ordered]@{
              TimeStamp   = $TimeStamp
              User        = $AuditData.UserId
              Action      = $AuditData.Operation
              Group       = $AuditData.modifiedproperties.newvalue[1]
              Guest       = $AuditData.ObjectId }      
           $Report += $ReportLine }}
      }}
Write-Host $NewGuests "new guest records found..."
$Report | Sort GroupName, Timestamp | Get-Unique -AsString | Format-Table Timestamp, Groupname, Guest

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
