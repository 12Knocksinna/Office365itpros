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
