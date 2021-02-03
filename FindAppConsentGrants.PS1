# FindAppConsentGrants.PS1
# Find consent grants for app permissions from records captured in the Office 365 audit log
# Requires the Exchange Online management module
# 
$Records = Search-UnifiedAuditLog -StartDate ((Get-Date).AddDays(-90)) -EndDate ((Get-Date).AddDays(1)) -ResultSize 5000 -Operations "Consent to application."

If ($Records) {
   $Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report
   ForEach ($Rec in $Records) {
     $Auditdata = $Rec.Auditdata | ConvertFrom-Json
     $ReportLine = [PSCustomObject]@{ 
         User            = $Auditdata.UserId
         Date            = Get-Date ($Auditdata.CreationTime) -format g
         ObjectId        = $Auditdata.ObjectId
         AppId           = $Auditdata.ObjectId.Split(";")[0]
         AdminConsent    = $Auditdata.ModifiedProperties | ?{$_.Name -eq "ConsentContext.IsAdminConsent"} | Select -ExpandProperty NewValue
         ForAllUsers     = $Auditdata.ModifiedProperties | ?{$_.Name -eq "ConsentContext.OnBehalfOfAll"} | Select -ExpandProperty NewValue
         Tags            = $Auditdata.ModifiedProperties | ?{$_.Name -eq "ConsentContext.Tags"} | Select -ExpandProperty NewValue
         Details        =  $Auditdata.ExtendedProperties | ?{$_.Name -eq "additionalDetails"} | Select -ExpandProperty Value } 
     $Report.Add($ReportLine) 
}}

$Report | Out-GridView
