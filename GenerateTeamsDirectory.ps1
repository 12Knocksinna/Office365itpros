# GenerateTeamsDirectory.Ps1
# A script to generate a list of Teams with deep links that can be turned into a teams directory with clickable hyperlinks
# V1.0 16 Apr 2019
# https://github.com/12Knocksinna/Office365itpros/blob/master/GenerateTeamsDirectory.ps1
# Tony Redmond
Connect-ExchangeOnline
Connect-MgGraph -Scopes Group.Read.All, Directory.Read.All -NoWelcome
# Example Link https://teams.microsoft.com/l/team/19%3aiVtGhQV1iyIVXHgaR6wGxW6Y3QbaziQSC2y8ke0qnxQ1%40thread.tacv2/conversations?groupId=96054cd2-8c97-4975-98ae-64a2a2ef05d2&tenantId=22e90715-3da6-4a78-9ec6-b3282389492b

$Organization = Get-MgOrganization
$OrgName = $Organization.DisplayName
$Date = Get-Date -Format 'dd-MMM-yyyy hh:mm'
$TenantId = $Organization.Id

$ReportFile = "c:\temp\ListOfTeams.html"
$CSVFile = "c:\temp\ListofTeams.csv"
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
           <p><h1>Teams Organization Directory for " + $OrgName + "</h1></p>
           <p><h3>Generated: " + $date + "</h3></p></div>"
		
Write-Host "Fetching details of Teams"
[array]$Teams = Get-MgGroup -Filter "resourceProvisioningOptions/any(x:x eq 'Team')" -All | Sort-Object DisplayName
$Public = 0; $Private = 0; [int]$i=0
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file
Write-Host ("Processing directory information for {0} teams" -f $Teams.Count)
ForEach ($T in $Teams) {
    $i++
    $InternalId = Get-MgTeam -TeamId $T.Id | Select-Object -ExpandProperty InternalId
    $DeepLink = ("https://teams.microsoft.com/l/team/{0}/conversations?groupId={1}&tenantId={2}" -f $InternalId, $T.Id, $TenantId)
    $JoinLink = ("<a href={0}>Join Link</a>" -f $DeepLink)
    $G = (Get-UnifiedGroup -Identity $T.Id | Select-Object ManagedBy, ManagerSmtp, GroupMemberCount, GroupExternalMemberCount, AccessType)
    $Access = $G.AccessType
    Write-Host ("Team ({0}/{1}): {2} access type: {3}" -f $i, $Teams.count, $T.DisplayName, $G.AccessType)
    If ($Access -eq "Public") { 
        $Public++ 
    }
# Figure out who owns the underlying Office 365 group
   $ManagerDetails = $G.ManagedBy
    If ([string]::IsNullOrWhiteSpace($ManagerDetails) -and [string]::IsNullOrEmpty($ManagerDetails)) {
      $ManagerDetails = $Null
      Write-Host $T.DisplayName "has no owners!" -ForegroundColor Red
    } Else {
      [array]$ManagerDetails = (Get-ExoMailbox -Identity $G.ManagedBy[0]) | Select-Object DisplayName, PrimarySmtpAddress
    }
    If ($Null -eq $ManagerDetails) {
      $ManagedBy = "No Group Owners" 
      $ManagerSmtp = $Null 
    } Else {
      $ManagedBy = $ManagerDetails.DisplayName
      $ManagerSmtp = $ManagerDetails.PrimarySmtpAddress 
    }

   # Generate a line for this group for our report
   $ReportLine = [PSCustomObject][Ordered]@{
          Team                = $T.DisplayName
          Description         = $T.Description
          JoinLink            = $JoinLink
          Owner               = $ManagedBy
          OwnerSMTP           = $ManagerSmtp 
          Members             = $G.GroupMemberCount
          ExternalGuests      = $G.GroupExternalMemberCount
          Access              = $Access 
          Deeplink            = $Deeplink 
    }
   # And store the line in the report object
   $Report.Add($ReportLine)     
}
#End of processing teams - now create the HTML report and CSV file

$Private = $Teams.Count - $Public
$htmlbody = $Report | ConvertTo-Html -Fragment
$htmltail = "<h3>Report created for: " + $OrgName + "
             </h3>
             <p>Number of teams scanned    : " + $Teams.Count + "</p>" + 	
            "<p>Number of private teams    : " + $Private + "</p>" +
            "<p>Number of public teams     : " + $Public + "</p></html>"
$htmlreport = $htmlhead + $htmlbody + $htmltail
Add-Type -AssemblyName System.Web
[System.Web.HttpUtility]::HtmlDecode($htmlreport) | Out-File $ReportFile -Encoding UTF8
$Report | Export-CSV -NoTypeInformation $CSVFile

Write-Host ""
Write-Host ("All done. directory information for {0} teams processed. The output files are in {1} and {2}" -f $Teams.count, $ReportFile, $CsvFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
