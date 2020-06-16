# GetHoldsOnMailbox.PS1
# Example used in Chapter 19 of Office 365 for IT Pros.
# Takes holds in place for a mailbox and interprets them
# Needs a connection to both Exchange Online and the Compliance Center endpoint.
$User = Read-Host "Enter User to check"
$Mbx = Get-Mailbox -Identity $User -ErrorAction SilentlyContinue
If ($Mbx -eq $Null) 
   { Write-Host $User "is not a valid user"
     Break
      }

$OrgHolds = (Get-OrganizationConfig).InPlaceHolds

If ($OrgHolds.Count -gt 0) {
   Write-Host "The following organization-wide mailbox holds are in force..."
     ForEach ($Hold in $OrgHolds) { 
     $RetentionPolicy = (($Hold -Split ":")[0].Substring(3))
     $HoldType = $Hold.Substring(0,3)
     Switch ($HoldType)
     {
       "grp" {$Type = "Groups"}
       "mbx" {$Type = "Mailbox"
              Get-RetentionCompliancePolicy -Identity $RetentionPolicy | Select Name, Workload }
   }}}

$MbxHolds = $Mbx.InPlaceHolds
If ($MbxHolds.Count -gt 0) {
   Write-Host ""
   Write-Host "The following specific holds are in place on the mailbox..."
   ForEach ($MHold in $MbxHolds) { 
   $RetentionPolicy = (($MHold -Split ":")[0].Substring(3))
   $HoldType = $MHold.Substring(0,3)
   Switch ($HoldType)
     {
       "grp" {$Type = "Groups"}
       "skp" {$Type = "Skype IM Conversations"}
       "uni" {$Type = "Unified Hold" 
              $CaseHold = $RetentionPolicy.SubString(1)
              $Text = (Get-CaseHoldPolicy -Identity $CaseHold -ErrorAction SilentlyContinue).Name
              Write-Host "eDiscovery Case:" $Text  }
       "mbx" {$Type = "Mailbox"
              Get-RetentionCompliancePolicy -Identity $RetentionPolicy | Select Name, Workload }
      }
    # There might be an old Exchange in place hold...
    $int = $Mhold.substring(0,1)
    If ([bool]($int -as [int]) -eq $True)
    {  $Text = (Get-MailboxSearch -InPlaceHoldIdentity $MHold -ErrorAction SilentlyContinue).Name
       Write-Host "Exchange In-Place Hold:" $Text }
  }}

If ($Mbx.LitigationHoldEnabled -eq $True) {
   Write-Host "Litigation hold is enabled on the mailbox" $Mbx.DisplayName }
