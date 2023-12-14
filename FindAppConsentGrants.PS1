# FindAppConsentGrants.PS1
# Find consent grants for app permissions from records captured in the Office 365 audit log
# Requires the Exchange Online management module
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindAppConsentGrants.PS1
# V1.1 14-Dec-2023 - Updated to use SessionCommand

$Status = Get-ConnectionInformation -ErrorAction SilentlyContinue
If (!($Status)) {
  Connect-ExchangeOnline -SkipLoadingCmdletHelp
}
# Connect to Graph SDK to get application information
Connect-MgGraph -Scopes Application.Read.All -NoWelcome

[array]$Apps = Get-MgApplication -All
[array]$Sps = Get-MgServicePrincipal -All
$SPHash = @{}
ForEach ($SP in $SPs) {
  $SPHash.Add($SP.AppId,$SP.DisplayName)
}
$AppHash = @{}
ForEach ($App in $Apps) {
  $AppHash.Add($App.AppId,$App.DisplayName)
}

Write-Host "Searching for audit records..."
[array]$Records = Search-UnifiedAuditLog -StartDate ((Get-Date).AddDays(-90)) -EndDate ((Get-Date).AddDays(1)) `
  -ResultSize 5000 -Operations "Consent to application." -SessionCommand ReturnLargeSet

If ($Records) {
   $Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report
   ForEach ($Rec in $Records) {
      $AppName = $Null
      $Auditdata = $Rec.Auditdata | ConvertFrom-Json
      $AppId = $Auditdata.ObjectId.Split(";")[0]
      $AppName = $SpHash[$AppId]
      If (!($AppName)) {
          $AppName = $AppHash[$AppId]
      }
      $Tag =  $Auditdata.ModifiedProperties | Where-Object {$_.Name -eq "ConsentContext.Tags"} | Select-Object -ExpandProperty NewValue
      If ($Tag -eq "WindowsAzureActiveDirectoryIntegratedApp") {
          $AppType = "Enterprise app"
      } Else {
          $AppType = "Registered app" 
      }
      $ReportLine = [PSCustomObject]@{ 
         User            = $Auditdata.UserId
         Date            = Get-Date ($Auditdata.CreationTime) -format g
         ObjectId        = $Auditdata.ObjectId
         AppId           = $AppId
         AppName         = $AppName
         AdminConsent    = $Auditdata.ModifiedProperties | Where-Object {$_.Name -eq "ConsentContext.IsAdminConsent"} | Select-Object -ExpandProperty NewValue
         ForAllUsers     = $Auditdata.ModifiedProperties | Where-Object {$_.Name -eq "ConsentContext.OnBehalfOfAll"} | Select-Object -ExpandProperty NewValue
         AppType            = $AppType
         Details        =  $Auditdata.ExtendedProperties | Where-Object {$_.Name -eq "additionalDetails"} | Select-Object -ExpandProperty Value } 
      $Report.Add($ReportLine) 
    }
}

$Report = $Report | Sort-Object {$_.Date -as [datetime]}
$Report | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
