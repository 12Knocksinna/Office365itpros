# CheckRetentionPolicyUpdates.PS1
# Script to check for changes made to a retention policy
# https://github.com/12Knocksinna/Office365itpros/blob/master/CheckRetentionPolicyUpdates.PS1
CLS
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}

# Look for audit records over the last 30 days
$StartDate = (Get-Date).AddDays(-30); $EndDate = (Get-Date) 
$OutputCSVFile = "C:\Temp\RetentionPolicyUpdates.csv"
$AuditRulesReport = "C:\Temp\RetentionPolicyRulesUpdates.csv"

Write-Host "Checking for Retention Policies"
# Build a hash table of retention policies we can use to resolve Guids into policy names
$RetentionPolicies = @{}
Try {
    [array]$RP = Get-RetentionCompliancePolicy }
Catch {
    Write-Host "Error fetching details of the tenant's retention policies - please make sure your session is connected to the compliance endpoint" ; break}
 
$RP.ForEach( {
       $RetentionPolicies.Add([String]$_.Guid, $_.Name) } )

# Now do the same for the app-specific retention policies (Teams private channels and Yammer)

Try {
    [array]$RP = Get-AppRetentionCompliancePolicy }
Catch {
    Write-Host "Error fetching details of the tenant's app retention policies - please make sure your session is connected to the compliance endpoint" ; break}

$RP.ForEach( {
       $RetentionPolicies.Add([String]$_.Guid, $_.Name) } )

# Find the audit records
Write-Host "Finding audit records"
[array]$Records = (Search-UnifiedAuditLog -Operations SetRetentionCompliancePolicy, SetRetentionComplianceRule	-StartDate $StartDate -EndDate $EndDate -Formatted -ResultSize 2000)
If (!($Records)) {Write-Host "No audit records found - exiting!"; break}

# Strip out records for policy updates (operation = SetRetentionCompliancePolicy)
[array]$AuditRecords = $Records | ? {$_.Operations -eq "SetRetentionCompliancePolicy"}

# Build array of retention rule changes
[array]$RuleRecords = $Records | ? {$_.Operations -eq "SetRetentionComplianceRule"}
$AuditRules = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Rule in $RuleRecords) {
    $AuditData = $Rule.AuditData | ConvertFrom-Json
    $DataLine = [PSCustomObject] @{
         Date                = $Rule.CreationDate
         User                = $AuditData.UserId
         Policy              = $AuditData.ExtendedProperties | ?{$_.Name -eq "PolicyName"} | Select -ExpandProperty Value
         RetentionAction     = $AuditData.ExtendedProperties | ?{$_.Name -eq "RetentionAction"} | Select -ExpandProperty Value
         RetentionDuration   = $AuditData.ExtendedProperties | ?{$_.Name -eq "RetentionDuration"} | Select -ExpandProperty Value
         RetentionType       = $AuditData.ExtendedProperties | ?{$_.Name -eq "RetentionType"} | Select -ExpandProperty Value
         Actions             = $AuditData.Parameters | ?{$_.Name -eq "CmdletOptions"} | Select -ExpandProperty Value   }
    $AuditRules.Add($DataLine) 
}

# Report audit records for policy updates
$AuditReport = [System.Collections.Generic.List[Object]]::new() 
ForEach ($AuditRecord in $AuditRecords) {
    $AuditData = $AuditRecord.AuditData | ConvertFrom-Json
    $PolicyDetails = $AuditData.Parameters | ?{$_.Name -eq "CmdletOptions"} | Select -ExpandProperty Value

    $PolicyName = $Null; $PolicyGuid = $Null; $Encodedtext = $Null
    If ($PolicyDetails -Like "*RetryDistribution*") { # The change is to restart distributions to target locations
       $Start = $PolicyDetails.IndexOf('"')+1
       $End = $PolicyDetails.IndexOf("-Retry")-13
       $PolicyName = $PolicyName = $PolicyDetails.SubString($Start,$End) }
   Else { # Update made to the policy
      $Start = $PolicyDetails.IndexOf('"')+1
      $EncodedText = $PolicyDetails.SubString($Start,48)
      $PolicyGuid = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($EncodedText))
      $PolicyName =  $RetentionPolicies.Item($PolicyGuid) # See if we can find the retention policy name
   }

    $DataLine = [PSCustomObject] @{
         Date                = $AuditRecord.CreationDate
         User                = $AuditData.UserId
         Policy              = $PolicyName  
         PolicyGuid          = $PolicyGuid
         DetailsLogged       = $PolicyDetails
         EC                  = $EncodedText }
    $AuditReport.Add($DataLine) 
}

$AuditReport | Export-CSV -NoTypeInformation $OutputCSVFile
$AuditRules | Export-CSV -NoTypeInformation $AuditRulesReport
$AuditReport | Out-GridView
Write-Host "All done. Policy update report available in" $OutputCSVFile "and details of audit rule updates in" $AuditRulesReport

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
