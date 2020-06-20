# ReportRestoreRecoverableItemsAudit.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportRestoreRecoverableItemsAudit.PS1
# Needs connection to Exchange Online and Azure AD
# Find audit records for Restore-RecoverableItems cmdlet (GUI or cmdlet)
$StartDate = (Get-Date).AddDays(-90) ; $EndDate = Get-Date
$Records = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations Restore-RecoverableItems -ResultSize 3000) 
If ($Records.Count -eq 0) {
    Write-Host "No audit records for restore deleted items found." }
Else {
    CLS
    $Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
    ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      $TimeStamp   = Get-Date($AuditData.CreationTime) -format g
      $TargetMailbox = ($Auditdata.Parameters | ?{$_ -Match "Identity"}).Value 
      # Audit record holds Azure AD account identifier (GUID) for target mailbox, so translate it.
      If ($TargetMailbox -ne $Null) { 
          $TargetMailbox = Get-AzureADUser -ObjectId $TargetMailbox | Select -ExpandProperty UserPrincipalName }
      Else { 
          $TargetMailbox = "Unknown user" }
      $SourceFolder = ($Auditdata.Parameters | ?{$_ -Match "SourceFolder"}).Value
      $EntryID = ($Auditdata.Parameters | ?{$_ -Match "EntryID"}).Value
      $SearchStart = ($Auditdata.Parameters | ?{$_ -Match "FilterStartTime"}).Value
      $SearchEnd  = ($Auditdata.Parameters | ?{$_ -Match "FilterEndTime"}).Value
      $ReportLine = [PSCustomObject] @{
           TimeStamp     = $TimeStamp
           User          = $AuditData.UserId
           TargetMailbox = $TargetMailbox
           EntryID       = $EntryID
           SourceFolder  = $SourceFolder
           SearchStart   = $SearchStart
           SearchEnd     = $SearchEnd
            }        
      $Report.Add($ReportLine) }
}

$Report | Format-Table TimeStamp, User, TargetMailbox, SourceFolder
