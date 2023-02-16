# ReportSensitivityLabelSettings.PS1
# A script to report the settings of sensitivity labels
# V1.0 15-Feb-2023
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportSensitivityLabelSettings.PS1

Function CheckRightsIdentity {
 [cmdletbinding()]
    Param([string]$RightsIdentity)

  $RightsCheck = $False
  If ($RightsIdentity -eq "AuthenticatedUsers") {
     $RightsCheck = $True
     Return $RightsCheck }

 If ($RightsIdentity -like "*@*")  {
   $Domain = $RightsIdentity.Split("@")[1].toLower()
   If ($Domain -in $Domains) { # It's an email address belonging to our tenant
      $CheckUser = Get-ExoRecipient -Identity $RightsIdentity -ErrorAction SilentlyContinue
      If (!($CheckUser)) { $RightsCheck = $False; Return $RightsCheck 
      } Else { $RightsCheck = $True; Return $RightsCheck }
   } 
 }

 # Check a domain name https://learn.microsoft.com/en-gb/archive/blogs/tip_of_the_day/cloud-tip-of-the-day-use-powershell-to-check-domain-availability-for-office-365-and-azure
 # Don't bother for Outlook.com or microsoft.com - always Unmanaged for some reason
 Switch ($RightsIdentity) {
  "outlook.com"  {
     $RightsCheck = $True 
     $Global:CheckMessage = "consumer domain"
     break }
  "microsoft.com" {
     $RightsCheck = $True; break }
  "gmail.com" { 
     $RightsCheck = $True 
     $Global:CheckMessage = "consumer domain" 
     break }
  "google.com"  {
     $RightsCheck = $True 
     $Global:CheckMessage = "consumer domain" 
     break }
  "yandex.com"  {
     $RightsCheck = $True 
     $Global:CheckMessage = "consumer domain" 
     break }
  default {
     $Uri =  ('https://login.microsoftonline.com/getuserrealm.srf?login=user@{0}&xml=1' -f $RightsIdentity)
     $CheckDomain = Invoke-WebRequest -Uri $Uri
     If ($CheckDomain.StatusCode -eq 200) {
       $NamespaceType = ([xml]($CheckDomain.Content)).RealmInfo.NameSpaceType
       If ($NameSpaceType -eq "Managed") { # Azure AD domain
         $RightsCheck = $True; Return $RightsCheck 
       } Else {
         $RightsCheck = $False; Return $RightsCheck }
      }}
   } #End Switch
       
Return $RightsCheck
}

# End functions and start of processing

$HtmlHead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.pass{background: #B7EB83;}
	   td.warn{background: #FFF275;}
	   td.fail{background: #FF2626; color: #ffffff;}
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Microsoft 365 Sensitivity Labels Settings Report</h1></p>
           <p><h3>Generated: " + (Get-Date -format 'dd-MMM-yyyy hh:mm tt') + "</h3></p></div>"

$Version = "1.0"
$HtmlReportFile = "c:\temp\SenstivityLabelSettings.html"
$CSVReportFile = "c:\temp\LabelInfo.CSV"

$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) { Write-Host "Please connect to Exchange Online management before running the script" ; break }
# Connect to compliance endpoint
Connect-IPPSSession

[array]$Global:Domains = (Get-AcceptedDomain).DomainName

# Get details of sensitivity labels
[array]$Labels = Get-Label
If (!($Labels)) { Write-Host "Can't get details of sensitivity labels - exiting" ; break }
Write-Host ("Processing details for {0} sensitivity labels" -f $Labels.count)

$LabelInfo = [System.Collections.Generic.List[Object]]::new() 
$NewLine = "`r`n"
[int]$i = 0
ForEach ($Label in $Labels) {
 $i++
 Write-Host ("Processing label: {0} ({1}/{2})" -f $Label.DisplayName, $i, $Labels.count)
 $LabelActions = (Get-Label -Identity $Label.ImmutableId).LabelActions | Convertfrom-JSON
 ForEach ($Section in $LabelActions) {
    $SectionSettings = $Section.Settings
      ForEach ($Setting in $SectionSettings) { 
        $Value = $Null; $DisplayName = $Null
        Switch ($Setting.Key) {
         "rightsdefinitions" {
          [array]$Rights = $Setting.Value | ConvertFrom-Json
          [int]$r = 0
          ForEach ($Right in $Rights) {
           $r++
           $RightsCheck = $False; $CheckMessage = $Null
           $RightsCheck = CheckRightsIdentity($Right.Identity)
           $RightsFormat = ("{0} rights: {1}" -f $Right.Identity, $Right.Rights)
           If ($RightsCheck -eq $False) { $RightsFormat = $RightsFormat + " (Warning: Rights assigned to unverified holder)" }
           If ($CheckMessage) { $RightsFormat = ("{0} (Warning: {1})" -f $RightsFormat, $CheckMessage) }
           if ($r -lt $Rights.count) { $RightsFormat = $RightsFormat + $Newline }
           $Value += $RightsFormat } 
         }
         "protectionlevel" {
          $Protection = $Setting.Value | ConvertFrom-Json
          $Value = ("Id {0} {1} ({2})" -f $Protection.Id, $Protection.DisplayName, $Protection.Description)
         }
          Default {
           $Value = $Setting.Value }
         }
         If ($Label.ParentLabelDisplayName) { $DisplayName = ("{0}/{1}" -f $Label.ParentLabelDisplayName, $Label.DisplayName) 
          } Else { $DisplayName = $Label.DisplayName }          
         $DataLine  = [PSCustomObject] @{
            Label    = $DisplayName
            Applies  = $Label.ContentType
            Priority = $Label.Priority
            Section  = $Section.Type
            Subtype  = $Section.SubType
            Key      = $Setting.Key
            Value    = $Value }
         $LabelInfo.Add($DataLine)
      } #End ForEach setting
   } #End ForEach section
} #End ForEach label

$LabelInfo | Out-GridView
$LabelInfo | Export-CSV -NoTypeInformation $CSVReportFile
# Alternative $Labelinfo | Export-Excel LabelInfo.xlsx

# Generate array of unique label names from the report
[array]$LabelNames = $LabelInfo.Label | Sort-Object -Unique
[string]$HtmlReport = $Null

ForEach ($L in $LabelNames) {
   $DataToReport = $LabelInfo | Where-Object {$_.Label -eq $L} 
   $HtmlHeading = ("<p><h1>Settings for Sensitivity label <b>{0}</b></h1></p><p><h2>Label Scope: <i>{1}</i>  Label Priority: <i>{2}</i></p></h2>" -f $L, $DataToReport[0].Applies, $DataToReport[0].Priority)
   $DataToReport = $DataToReport | Select-Object Section, Subtype, Key, Value
   $HtmlData = $DataToReport | ConvertTo-Html -Fragment
   $HtmlReport = $HtmlReport + $HtmlHeading + $HtmlData 
}

# Create the HTML report
$Htmltail = "<p>Report created for: " + ((Get-OrganizationConfig).DisplayName) + "</p><p>" +
             "<p>Number of sensitivity labels: " + $Labels.count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------" +
             "<p>Microsoft 365 Sensitivity Labels Settings Report <b>" + $Version + "</b>"	
$HtmlReport = $HtmlHead + $HtmlReport + $HtmlTail
$HtmlReport | Out-File $HtmlReportFile  -Encoding UTF8
Write-Host ("The sensitivity Labels settings report files are available in {0} (HTML) and {1} (CSV)" -f $HtmlReportFile, $CSVReportFile)
# Check if any warnings are found
[array]$Warnings = $LabelInfo | Where-Object {$_.Value -like "*warning: rights*"} | Select-Object Label, Applies, Priority, Value
If ($Warnings) {
Write-Host " "
Write-Host "You should check these warnings flagged for rights assignments:"
$Warnings | Format-List }

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
