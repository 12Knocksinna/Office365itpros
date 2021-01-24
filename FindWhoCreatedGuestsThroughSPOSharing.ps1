# A script to find who created new guest accounts in an Office 365 tenant through SharePoint Online sharing invitations
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindWhoCreatedGuestsThroughSPOSharing.ps1
$EndDate = (Get-Date).AddDays(1); $StartDate = (Get-Date).AddDays(-90); $NewGuests = 0
$Records = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "SharingInvitationCreated" -ResultSize 2000 -Formatted)
If ($Records.Count -eq 0) {
   Write-Host "No Sharing Invitations records found." }
 Else {
   Write-Host "Processing" $Records.Count "audit records..."
   $Report = [System.Collections.Generic.List[Object]]::new()
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
              TimeStamp    = $TimeStamp
              InvitingUser = $AuditData.UserId
              Action       = $AuditData.Operation
              URL          = $AuditData.ObjectId
              Site         = $AuditData.SiteUrl
              Document     = $AuditData.SourceFileName
              Guest        = $AuditData.TargetUserOrGroupName }      
           $Report.Add($ReportLine)}}
      }}
$Report | Format-Table TimeStamp, Guest, Document -AutoSize

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
