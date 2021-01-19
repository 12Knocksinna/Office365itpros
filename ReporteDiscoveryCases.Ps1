# Report eDiscovery Cases
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReporteDiscoveryCases.Ps1
# Used in Chapter 20 of Office 365 for IT Pros. Need to be connected to the Security and Compliance Center endpoint (use Connect-IPPSSession in EXO module)
$ClosedCases = 0 ;$ActiveHolds = 0
$Report = [System.Collections.Generic.List[Object]]::new()
Write-Host "Finding eDiscovery Cases"
$Cases = Get-ComplianceCase -ErrorAction SilentlyContinue
Write-Host $Cases.Count " cases found. Now extracting information."

ForEach ($Case in $Cases) {
    Write-Host "Processing eDiscovery Case:" $Case.Name
    $CaseMembers = (Get-ComplianceCaseMember -Case $Case.Name | Select Name, WindowsLiveId)
    $Names = $Null
    $First = $True
    # Figure out display name for case managers - nicer than their email address
    ForEach ($M in $CaseMembers) {
       If ($First) { 
          $Names = $M.Name
          $First = $False }
       Else { 
          $Names = $Names, $M.Name -Join ", " }    }
    If ($Case.Status -eq "Closed") {
        $ReportLine = [PSCustomObject]@{
           Case        = $Case.Name
           Status      = $Case.Status
           Created     = $Case.CreatedDateTime
           ClosedBy    = $Case.ClosedBy
           Closed      = $Case.ClosedDateTime
           Members     = $Names    }
        $Report.Add($ReportLine)
        $ClosedCases++   }  
   ElseIf ($Case.Status = "Open") {
      $HoldPolicies = (Get-CaseHoldPolicy -Case $Case.Name | % {Get-CaseHoldPolicy $_.Name -Case $_.CaseId -DistributionDetail})   
      ForEach ($Hold in $HoldPolicies) {
        $HoldRule = Get-CaseHoldRule -Policy $Hold.Name 
        $ActiveHolds++
      $i = 0 # Section of code to highlight inactive mailboxes that are under hold
      $Mbxes = $Null
      $CountMbx = 0
      ForEach ($H in $Hold.ExchangeLocation) {
      $Len = $Hold.ExchangeLocation[$i].DisplayName | Measure-Object -Character | Select -Expandproperty Characters
      If ($Hold.ExchangeLocation[$i].DisplayName.Substring(0,1) -eq ".") {
          $Mbx = ($Hold.ExchangeLocation[$i].DisplayName.Substring(1,$Len - 1)) + " (Inactive); "
          $Mbxes = $Mbxes + $Mbx }
      Else  {
          $Mbxes = $Mbxes + ($Hold.ExchangeLocation[$i].DisplayName) + "; " }
      $CountMbx++     
      $i++ }
 #    Write out the report line
      $ReportLine = [PSCustomObject]@{
           Case           = $Case.Name
           Status         = $Case.Status
           Created        = Get-Date ($Case.CreatedDateTime) -format g
           ClosedBy       = $Case.ClosedBy
           Closed         = $Case.ClosedDateTime
           Members        = $Names
           Hold           = $Hold.Name
           HoldEnabled    = $Hold.Enabled
           HoldCreatedby  = $Hold.CreatedBy
           HoldModifiedby = $Hold.LastModifiedBy
           Mailboxes      = $Mbxes
           MailboxCount   = $CountMbx
           SPOSites       = ($Hold.SharePointLocation.Name) -Join ","
           Query          = $HoldRule.ContentMatchQuery
           HoldCreated    = Get-Date($HoldRule.WhenCreatedUTC) -format g
           HoldModifued   = Get-Date($HoldRule.WhenChangedUTC) -format g }
        $Report.Add($ReportLine)  
   }
}}
CLS
Write-Host "EDiscovery Cases found: " $Cases.Count
Write-Host "Active Cases:           " ($Cases.Count - $ClosedCases)
Write-Host "Closed Cases:           " $ClosedCases
Write-Host "Active Holds:           " $ActiveHolds 
$Report | Sort Status, Case | Format-Table Case, Status, Created, HoldCreated

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
