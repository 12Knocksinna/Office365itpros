# GenerateTeamsDirectory.Ps1
# A script to generate a list of Teams with deep links that can be turned into a teams directory with clickable hyperlinks
# V1.0 16 Apr 2019
# https://github.com/12Knocksinna/Office365itpros/blob/master/GenerateTeamsDirectory.ps1
# Tony Redmond

# Example Link https://teams.microsoft.com/l/team/19%3a099f14ec32cc4793a7ef99238dbaac86%40thread.skype/conversations?groupId=5b617155-a124-4e32-a230-022dfe0b80ac&tenantId=b662313f-14fc-43a2-9a7a-d2e27f4f3478
$OrgName = (Get-OrganizationConfig).Name
$Today = (Get-Date)
$Date = (Get-Date).ToShortDateString()
$TenantId = "&tenantId=b662313f-14fc-43a2-9a7a-d2e27f4f3478"
$DeepLinkPrefix = "https://teams.microsoft.com/l/team/19%3aa68e793c288743329333fb32d5d010ad%40thread.skype/conversations?groupId="
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
           <p><h1>List of Teams</h1></p>
           <p><h3>Generated: " + $date + "</h3></p></div>"
		
Write-Host "Fetching List of Teams"
$Teams = Get-Team | Sort DisplayName
$Public = 0; $Private = 0
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file
Write-Host "Processing" $Teams.Count "teams"
ForEach ($T in $Teams) {
   $DeepLink = $DeepLinkPrefix + $T.GroupId + $TenantId
   $G = (Get-UnifiedGroup -Identity $T.GroupId | Select ManagedBy, ManagerSmtp, GroupMemberCount, GroupExternalMemberCount, AccessType)
   $Access = $G.AccessType
   Write-Host "Team" $T.DisplayName "access type" $G.AccessType
   If ($Access -eq "Public") { $Public++ }
# Figure out who owns the underlying Office 365 group
   $ManagerDetails = $G.ManagedBy
    If ([string]::IsNullOrWhiteSpace($ManagerDetails) -and [string]::IsNullOrEmpty($ManagerDetails)) {
      $ManagerDetails = $Null
      Write-Host $T.DisplayName "has no owners!" -ForegroundColor Red}
    Else {
      $ManagerDetails = (Get-Mailbox -Identity $G.ManagedBy[0]) | Select DisplayName, PrimarySmtpAddress}
    If ($ManagerDetails -eq $Null) {
      $ManagedBy = "No Group Owners" 
      $ManagerSmtp = $Null }
    Else {
      $ManagedBy = $ManagerDetails.DisplayName
      $ManagerSmtp = $ManagerDetails.PrimarySmtpAddress }

   # Generate a line for this group for our report
   $ReportLine = [PSCustomObject][Ordered]@{
          Team                = $T.DisplayName
          Description         = $T.Description
          JoinLink            = $DeepLink
          Owner               = $ManagedBy
          OwnerSMTP           = $ManagerSmtp 
          Members             = $G.GroupMemberCount
          ExternalGuests      = $G.GroupExternalMemberCount
          Access              = $Access }
   # And store the line in the report object
   $Report.Add($ReportLine)     }
#End of processing teams - now create the HTML report and CSV file
$Private = $Teams.Count - $Public
$htmlbody = $Report | ConvertTo-Html -Fragment
$htmltail = "<p>Report created for: " + $OrgName + "
             </p>
             <p>Number of teams scanned    : " + $Teams.Count + "</p>" + 	
            "<p>Number of private teams    : " + $Private + "</p>" +
            "<p>Number of public teams     : " + $Public + "</p></html>"
$htmlreport = $htmlhead + $htmlbody + $htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8
$Report | Export-CSV -NoTypeInformation $CSVFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
