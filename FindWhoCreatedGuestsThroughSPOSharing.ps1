# A script to find who created new guest accounts in an Office 365 tenant through SharePoint Online sharing invitations
$EndDate = (Get-Date).AddDays(1); $StartDate = (Get-Date).AddDays(-90); $NewGuests = 0
$Records = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "SharingInvitationCreated" -ResultSize 2000 -Formatted)
If ($Records.Count -eq 0) {
   Write-Host "No Sharing Invitations records found." }
 Else {
   Write-Host "Processing" $Records.Count "audit records..."
   $Report = @()
   ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      # Only process the additions of guest users to groups
      If ($AuditData.TargetUserOrGroupName -Like "*#EXT#*") {
         $TimeStamp = Get-Date $Rec.CreationDate -format g
         # Try and find the timestamp when the invitation for the Guest user account was accepted from AAD object
         Try {$AADCheck = (Get-Date(Get-AzureADUser -ObjectId $AuditData.TargetUserOrGroupName).RefreshTokensValidFromDateTime -format g) }
           Catch {Write-Host "Azure Active Directory record for" $AuditData.UserId "no longer exists" }
          If ($TimeStamp -eq $AADCheck) { # It's a new record, so let's write it out 
            $NewGuests++
            $ReportLine = [PSCustomObject][Ordered]@{
              TimeStamp   = $TimeStamp
              Action      = $AuditData.Operation
              URL         = $AuditData.ObjectId
              Site        = $AuditData.SiteUrl
              Document    = $AuditData.SourceFileName
              Guest       = $AuditData.TargetUserOrGroupName }      
           $Report += $ReportLine }}
      }}
$Report | Format-Table TimeStamp, Guest, Document -AutoSize
