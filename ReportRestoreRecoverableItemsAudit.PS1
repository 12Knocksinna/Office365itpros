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
      # Audit record holds Azure AD account identifier (GUID) for target mailbox, so translate it - but sometimes the record holds a mailbox alias.
      If (-not($TargetMailbox -Like "*.*"))   { 
          $TargetMailbox = Get-AzureADUser -ObjectId $TargetMailbox | Select -ExpandProperty UserPrincipalName }
      $SourceFolder = ($Auditdata.Parameters | ?{$_ -Match "SourceFolder"}).Value
      If ($SourceFolder -eq $Null) { $SourceFolder = "Recoverable Items" }
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

$SortedDate = @{e={$_.TimeStamp -as [DateTime]}; descending = $True}
$Report = $Report |  Sort EntryId -Unique # Get rid of duplicate records
$Report | Sort $SortedDate | Format-Table TimeStamp, User, TargetMailbox, SourceFolder

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
