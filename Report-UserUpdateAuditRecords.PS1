# Report-UserUpdateAuditRecords.PS1
# An example of how to find unified audit records for user account update events and analyze what actions/changes
# are performed in each audit event
#
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-UserUpdateAuditRecords.PS1

# Connect to Exchange Online and the Microsoft Graph
Connect-MgGraph -Scopes Directory.Read.All
Connect-ExchangeOnline

# Find some audit records to play with from the last 180 days
Write-Host "Searching for audit records for update user events..."
[array]$Records = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-180) -EndDate (Get-Date).AddDays(1) -Formatted -ResultSize 5000  -Operations "update user"
If (!($Records)) { 
   Write-Host "No audit records found for update user events - exiting" ; break
} Else {
   Write-Host ("{0} audit records found - processing..." -f $Records.count) }

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
 $AuditData = $Rec.AuditData | ConvertFrom-Json
 $OldValue = $Null; $NewValue = $Null
 [array]$ModifiedProperties = $AuditData.ModifiedProperties
 [array]$Properties = $ModifiedProperties | Where-Object { $_.Name -eq 'Included Updated Properties' } | Select-Object -ExpandProperty NewValue
 If (!([string]::IsNullOrEmpty($Properties))) {
  # Create a list of updated properties
  [array]$ListOfProperties = $Properties.Split(",")
  ForEach ($P in $ListOfProperties) {
    $Property = $P.Trim()         
    $Command = '$NewValue = (($modifiedproperties' + " | where {`$_.Name -eq '" + $Property + "'}).NewValue -replace '[\[\]]', '').Trim()"
    Invoke-Expression $Command
    $Command = '$OldValue = (($modifiedproperties' + " | where {`$_.Name -eq '" + $Property + "'}).OldValue -replace '[\[\]]', '').Trim()"
    Invoke-Expression $Command
    $NewValue = $NewValue -Replace  '["]'
    $OldValue = $OldValue -Replace  '["]'
    If ($Property -eq "StrongAuthenticationPhoneAppDetail") {
      # Handle an update of the MFA authentication method
         $MFADetails = $OldValue.replace("}", '')
         $MFADetails = $MFADetails.replace("{", '').trim()
         [array]$MFAData = $MFADetails.Split(",")
         $OldValue = ("{0} {1}" -f ($MFAData[0].toString()),  ($MFAData[3].toString().trim()))
         $MFADetails = $NewValue.replace("}", '')
         $MFADetails = $MFADetails.replace("{", '').trim()
         [array]$MFAData = $MFADetails.Split(",")
         $NewValue = ("{0} {1}" -f ($MFAData[0].toString()),  ($MFAData[3].toString().trim())) 
   }
    If ($Property -eq "AssignedLicense") {
    # License updates - this is an example of where we could do more to extract data from the audit records
       $NewValue = "License Update " + $NewValue }
    If ($Property -eq "AssignedPlan") {
       $NewValue = "See audit record for assigned plan information" }
    If ($Property -eq "LicenseAssignmentDetail") {
       $NewValue = "See audit record for license assignment detail" }
    If ([string]::IsNullOrEmpty($OldValue)) {
       $OldValue = "Not filled" }
    # Some audit events are caused by background processing - so let's find out what background process is involved for these ecents
    If ($Rec.UserIds.SubString(0,16) -eq "ServicePrincipal") {
       $UserId = Get-MgServicePrincipal -ServicePrincipalId $Rec.Userids.Substring(17,($Rec.Userids.Length-17)) | Select-Object -ExpandProperty DisplayName
    } Else {
    # It's a normal user who performed the action
       $UserId = $Rec.UserIds }
    # Report what happened
    $ReportLine = [PSCustomObject][Ordered]@{
       Timestamp   = $Rec.CreationDate
       User        = $UserId
       Account     = $AuditData.ObjectId
       Property    = $Property
       'Old value' = $OldValue
       'New value' = $NewValue }
    $Report.Add($ReportLine)
  } # End For to extract updated properties
 } # End If Properties
} # End For audit events

$Report | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
