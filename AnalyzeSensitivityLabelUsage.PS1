# AnalyzeSensitivityLabelUsage.PS1
# A script to analyze the usage of sensitivity labels based on Office 365 audit log data
# https://github.com/12Knocksinna/Office365itpros/blob/master/AnalyzeSensitivityLabelUsage.PS1

Connect-ExchangeOnline
Connect-IPPSSession

Write-Host "Retrieving sensitivity labels used in the tenant"
$Labels = @{}
[array]$LabelSet = Get-Label | Select-Object ImmutableId, DisplayName
If (!($LabelSet)) { Write-Host "Can't find any sensitivity labels - exiting"; break }
ForEach ($L in $LabelSet) { $Labels.Add([string]$L.ImmutableId, [string]$L.DisplayName) }

$Operations = ("SensitivityLabelUpdated", "SensitivityLabelApplied", "FileSensitivityLabelApplied", "MIPLabel")
$StartDate = (Get-Date).AddDays(-90)
$EndDate = (Get-Date).AddDays(1)
[Array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Formatted -ResultSize 5000 -Operations $Operations
If (!($Records)) { Write-Host "No audit records for sensitivity label application found - exiting" ; break }

$Records = $Records | Where-Object {$_.RecordType -ne "ComplianceDLPExchange"}
$Report = [System.Collections.Generic.List[Object]]::new() 

ForEach ($Rec in $Records) {
   $AuditData = $Rec.AuditData | ConvertFrom-Json
   $LabelRemoved = $Null; $LabelAdded = $Null; $Type = $Null; $LabelRemoved = $Null; $Item = $Null; $Site = $Null

   If ($AuditData.Application -ne "Outlook") {
   Switch ($Rec.Operations) {
    "FileSensitivityLabelApplied" {
      $Type = "Default label applied by policy"
      $LabelAdded = $Labels[$AuditData.DestinationLabel]
      $Application = $AuditData.Workload
      $ObjectId = $AuditData.ObjectId
      $Item = $AuditData.DestinationFileName 
      $Site = $AuditData.SiteUrl
    }   
    "SensitivityLabelApplied" {
     $Type = "Label assigned by user"
     $LabelAdded = $Labels[$AuditData.SensitivityLabelEventData.SensitivityLabelId]
     $Application = $AuditData.Application
     $ObjectId = [System.Web.HttpUtility]::UrlDecode($AuditData.ObjectId)
     $Item = $ObjectId.Split('/')[-1]	
     $Site = "https://" + $ObjectId.Split("/")[2] + "/sites/" + $ObjectId.Split("/")[4] + "/"
    }
     "SensitivityLabelUpdated" {
     $Type = "Label updated by user"
     $LabelAdded = $Labels[$AuditData.SensitivityLabelEventData.SensitivityLabelId]
     $LabelRemoved = $Labels[$AuditData.SensitivityLabelEventData.OldSensitivityLabelId]
     $Application = $AuditData.Application
     $ObjectId =  [System.Web.HttpUtility]::UrlDecode($AuditData.ObjectId)
     $Item = $ObjectId.Split('/')[-1]
     $Site = "https://" + $ObjectId.Split("/")[2] + "/sites/" + $ObjectId.Split("/")[4] + "/"
    }
    "MIPLabel" {
     $Type = "Email labeled"
     $LabelAdded = $Labels[$AuditData.LabelId]
     $Application = "Exchange Online"
     $ObjectId = "Email"
     $Item = "Email"
     $Site = "N/A"
     }
   } #End Switch

 If ($UserId -eq "app@sharepoint") {
     $Type = "Default label applied by document library" 
 } ElseIf ($UserId -eq "SHAREPOINT\system") { 
   $Type = "Label applied by auto-label policy" }
 If ($ObjectId -like "*/personal/*") { #Fix-up for OneDrive accounts
    $Site = "https://" + $ObjectId.Split("/")[2] + "/personal/" + $ObjectId.Split("/")[4] + "/" }

  $DataLine = [PSCustomObject] @{
       Timestamp    = Get-Date($Rec.CreationDate) -format g
       User         = $AuditData.UserId
       Operation    = $Rec.Operations
       LabelAdded   = $LabelAdded
       LabelRemoved = $LabelRemoved
       Application  = $Application
       Type         = $Type
       Site         = $Site
       Object       = $ObjectId
       Item         = $Item } 

 $Report.Add($DataLine) 
 } #End if
} # End ForEach

# Analysis
Write-Host ""
Write-Host "Most commonly used sensitivity labels"
Write-Host "-------------------------------------"
$Report | Group-Object LabelAdded | Sort-Object Count -Descending | Format-Table Name, Count
Write-Host ""
Write-Host "Most prolific applier of sensitivity labels"
Write-Host "-------------------------------------------"
$Report | Group-Object User | Sort-Object Count -Descending | Format-Table Name, Count

$Report | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
