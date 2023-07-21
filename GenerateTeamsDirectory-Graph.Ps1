# GenerateTeamsDirectory-Graph.Ps1
# A script to generate a list of Teams with deep links that can be turned into a teams directory with clickable hyperlinks
# V2.0 6 November 2023
# https://github.com/12Knocksinna/Office365itpros/blob/master/GenerateTeamsDirectory-Graph.ps1
# This version uses the Microsoft Graph PowerShell SDK instead of cmdlets from the Teams and Exchange Online management module
# 
# Tony Redmond

Connect-MgGraph -Scopes Group.Read.All, Directory.Read.All, GroupMember.Read.All, Team.ReadBasic.All, TeamSettings.Read.All
Select-MgProfile Beta

$Tenant = Get-MgOrganization
$Today = (Get-Date)
$Date = Get-Date($Today) -format f

$ReportHeading = "Teams Organizational Directory for " + $Tenant.DisplayName

$DeepLinkPrefix = "https://teams.microsoft.com/l/team/"
$ReportFile = "c:\temp\TeamsDirectory.html"
$ExcelFile = "c:\temp\Teams Directory.xlsx"
$htmlhead="<!DOCTYPE html>
           <html>
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
           <p><h1>$ReportHeading</h1></p>
           <p><h3>Generated: " + $date + "</h3></p></div>"
		
Write-Host "Fetching List of Teams"
[array]$Teams = Get-MgTeam -All | Sort DisplayName
If (!($Teams)) { Write-Host "No teams found - exiting" ; break} 
   Else {Write-Host ("Processing {0} teams" -f $Teams.count) }

[int]$i = 0; [int]$Public = 0 ; [int]$Private = 0; [int]$HiddenMembership = 0

$Report = [System.Collections.Generic.List[Object]]::new() # Create output file
$ReportHTML = [System.Collections.Generic.List[Object]]::new() # Create output file
Write-Host ("Processing {0} teams" -f $Teams.count)

ForEach ($T in $Teams) {
   $i++
   Write-Host ("Processing team {0} {1}/{2}" -f $t.displayname, $i, $teams.count)
   $InternalId = Get-MgTeam -TeamId $T.Id | Select-Object -ExpandProperty InternalId
   $DeepLinkHTML = '<p><a href="' + $DeepLinkPrefix + $InternalId + "/conversations?groupId=" + $T.Id + "&tenantId=" + $Tenant.Id + '">Link to team</a></p>'
   $DeepLink = $DeepLinkPrefix + $InternalId + "/conversations?groupId=" + $T.Id + "&tenantId=" + $Tenant.Id

   # Find team owners
   $Data = [System.Collections.Generic.List[Object]]::new() 
   [array]$Owners =  Get-MgGroupOwner -GroupId $T.Id | Select-Object -ExpandProperty Id
   ForEach ($Owner in $Owners) {
        $ON = Get-MgUser -UserId $Owner
        $DataLine = [PSCustomObject] @{
          Owner            = $ON.DisplayName
          Email            = $ON.Mail}
        $Data.Add($Dataline)
   }
   $OwnerDisplayNames = $Data.Owner -join ", "
   $OwnerSMTPAddress = $Data.Email -join ", "
      
   [array]$GroupMembers = Get-MgGroupMember -GroupId $T.Id | Select-Object -ExpandProperty Id
   $GroupMemberCount = $GroupMembers.Count
   $MemberCount = 0; $ExternalMemberCount = 0
   ForEach ($User in $GroupMembers) {
      $GM = Get-MgUser -UserId $User 
      If ($GM.UserType -eq "member") { $MemberCount++ } Else { $ExternalMemberCount++ }
   }

   Switch ($T.visibility) {
     "Public" { $Public++ }
     "Hiddenmembership" { $HiddenMembership++ }
     "Private"  { $Private++ }
   }

   # Generate a line for this group for our Excel worksheet
   $ReportLine = [PSCustomObject][Ordered]@{
          Team                   = $T.DisplayName
          Description            = $T.Description
          'Link to access team'  = $DeepLink
          Owners                 = $OwnerDisplayNames
          'Owner Email'          = $OwnerSMTPAddress
          Members                = $GroupMemberCount
          'Tenant Members'       = $MemberCount
          'External Members'     = $ExternalMemberCount
          Access                 = $T.visibility }
   # And store the line in the report object
   $Report.Add($ReportLine)     

 # Generate a line for this group for our HTML report
   $ReportHTMLLine = [PSCustomObject][Ordered]@{
          Team                   = $T.DisplayName
          Description            = $T.Description
          'Link to access team'  = $DeepLinkHTML
          Owners                 = $OwnerDisplayNames
          'Owner Email'          = $OwnerSMTPAddress
          Members                = $GroupMemberCount
          'Tenant Members'       = $MemberCount
          'External Members'     = $ExternalMemberCount
          Access                 = $T.visibility }
   # And store the line in the report object
   $ReportHTML.Add($ReportHTMLLine)     
} # End Foreach Teams

#End of processing teams - now create the HTML report and CSV file
$htmlbody = $ReportHTML | ConvertTo-Html -Fragment
$htmltail = "<p>Report created for: " + $Tenant.DisplayName + "
             </p>
             <p>Number of teams scanned    :            " + $Teams.Count + "</p>" + 	
            "<p>Number of private teams    :            " + $Private + "</p>" +
            "<p>Number of teams with hidden membership: " + $HiddenMembership + "</p>" +
            "<p>Number of public teams:                 " + $Public + "</p></html>"
$htmlreport = $htmlhead + $htmlbody + $htmltail

# Make sure that we output a working hyperlink for the deeplink
Add-Type -AssemblyName System.Web
[System.Web.HttpUtility]::HtmlDecode($htmlreport) | Out-File $ReportFile  -Encoding UTF8
$Report | Export-Excel -Path $ExcelFile -WorksheetName "Teams Directory" -Title "Teams Directory" -TitleBold -TableName "TeamsDirectory" -TableStyle Dark1 

Write-Host ("Output files are available in {0} (CSV) and {1} (HTML)" -f $ExcelFile, $ReportFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
