# ReportTeamsPolicyAssignments.PS1
# Generate a report about major Teams policies assigned to user accounts
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportTeamsPolicyAssignments.PS1
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "MicrosoftTeams")) {Write-Host "Please connect to the Microsoft Teams module and then restart the script"; break}
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online management module and then restart the script"; break}
[array]$Users = Get-CsOnlineUser -ResultSize 5000
# Filter the set to get Teams users - this will filter out all but cloud-only Teams users. If you don't want to use the filter, comment it out.
$Users = $Users | Where-Object {$_.InterpretedUserType -eq "PureOnlineTeamsOnlyUser" -or $_.InterpretedUserType -eq "PureOnlineTeamsOnlyUserFailedPublishingToAAD"} | Sort-Object DisplayName
If (!($Users)) {Write-Host "No users found - exiting"; break }
$Report = [System.Collections.Generic.List[Object]]::new()
# Process each user to fetch their policy assignments
ForEach ($User in $Users) {
    $TenantDefaultString = "Tenant Default" 
    $TeamsMeetingPolicy = $TenantDefaultString
    $TeamsMessagingPolicy = $TenantDefaultString
    $TeamsAppSetupPolicy = $TenantDefaultString
    $TeamsAppPermissionsPolicy = $TenantDefaultString
    $TeamsEncryptionPolicy = $TenantDefaultString
    $TeamsUpdatePolicy = $TenantDefaultString
    $TeamsChannelsPolicy = $TenantDefaultString
    $TeamsFeedbackPolicy = $TenantDefaultString
    $TeamsLiveEventsPolicy = $TenantDefaultString
    If ($User.TeamsMeetingPolicy) {$TeamsMeetingPolicy = $User.TeamsMeetingPolicy}
    If ($User.TeamsMessagingPolicy) {$TeamsMessagingPolicy = $User.TeamsMessagingPolicy}
    If ($User.TeamsAppSetupPolicy) {$TeamsAppSetupPolicy = $User.TeamsAppSetupPolicy}
    If ($User.TeamsAppPermissionPolicy) {$TeamsAppPermissionsPolicy = $User.TeamsAppPermissionPolicy}
    If ($User.TeamsEnhancedEncryptionPolicy) {$TeamsEncryptionPolicy = $User.TeamsEnhancedEncryptionPolicy}
    If ($User.TeamsUpdateManagementPolicy) {$TeamsUpdatePolicy = $User.TeamsUpdateManagementPolicy}
    If ($User.TeamsChannelsPolicy) {$TeamsChannelsPolicy = $User.TeamsChannelsPolicy}
    If ($User.TeamsFeedbackPolicy) {$TeamsFeedbackPolicy = $User.TeamsFeedbackPolicy}
    If ($User.TeamsMeetingBroadcastPolicy) {$TeamsLiveEventsPolicy = $User.TeamsMeetingBroadcastPolicy}
    # Output a report line
    $ReportLine = [PSCustomObject][Ordered]@{  
        User                         = $User.DisplayName
        UPN                          = $User.UserPrincipalName
        "Messaging Policy"           = $TeamsMessagingPolicy
        "Meeting Policy"             = $TeamsMeetingPolicy
        "App Setup Policy"           = $TeamsAppSetupPolicy
        "App Permissions Policy"     = $TeamsAppPermissionsPolicy
        "Enhanced Encryption Policy" = $TeamsEncryptionPolicy
        "Update Policy"              = $TeamsUpdatePolicy
        "Channels Policy"            = $TeamsChannelsPolicy
        "Feedback Policy"            = $TeamsFeedbackPolicy
        "Live Events"                = $TeamsLiveEventsPolicy
	"InterpretedUserType"        = $User.InterpretedUserType
    }
 #  Add it to the report
    $Report.Add($ReportLine) 
 }
$CSVOutput = "c:\temp\TeamsPolicyAssignments.CSV"
$ReportFile = "c:\temp\TeamsPolicyAssignments.html"
# Create the HTML report
$OrgDisplayName = (Get-OrganizationConfig).DisplayName
$CreationDate = Get-Date -format g
$Version = "1.0"
$htmlhead="<html>
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
           <p><h1>Teams Policy Assignment Report</h1></p>
           <p><h2><b>For the " + $OrgDisplayName + " organization</b></h2></p>
           <p><h3>Generated: " + (Get-Date -format g) + "</h3></p></div>"

$htmlbody1 = $Report | ConvertTo-Html -Fragment
$htmltail = "<p>Report created for: " + $OrgDisplayName + "</p>" +
             "<p>Created: " + $CreationDate + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of Teams users found:    " + $Users.Count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Teams Policy Assignment Report<b> " + $Version + "</b>"	
# Generate the HTML file
$htmlreport = $htmlhead + $htmlbody1 + $htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8
Write-Host ("All done. Teams policies for {0} users analyzed. CSV file is available at {1} and a HTML report at {2}" -f $Users.Count, $CSVOutput, $ReportFile)
$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation $CSVOutput

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
