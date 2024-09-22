# TeamsGroupsActivityReport.PS1
# A script to check the activity of Microsoft 365 Groups and Teams and report the groups and teams that might be deleted because they're not used.
# We check the group mailbox to see what the last time a conversation item was added to the Inbox folder. 
# Another check sees whether a low number of items exist in the mailbox, which would show that it's not being used.
# We also check the group document library in SharePoint Online to see whether it exists or has been used in the last 90 days.
# And we check Teams compliance items to figure out if any chatting is happening.

# Created 29-July-2016  Tony Redmond 
# V2.0 5-Jan-2018
# V3.0 17-Dec-2018
# V4.0 11-Jan-2020
# V4.1 15-Jan-2020 Better handling of the Team Chat folder
# V4.2 30-Apr-2020 Replaced $G.Alias with $G.ExternalDirectoryObjectId. Fixed problem with getting last conversation from Groups where no conversations are present.
# V4.3 13-May-2020 Fixed bug and removed the need to load the Teams PowerShell module
# V4.4 14-May-2020 Added check to exit script if no Microsoft 365 Groups are found
# V4.5 15-May-2020 Some people reported that Get-Recipient is unreliable when fetching Groups, so added code to revert to Get-UnifiedGroup if nothing is returned by Get-Recipient
# V4.6 8-Sept-2020 Better handling of groups where the SharePoint team site hasn't been created
# V4.7 13-Oct-2020 Teams compliance records are now in a different location in group mailboxes
# V4.8 16-Dec-2020 Some updates after review of code to create 5.0 (Graph based version)
# 
# https://github.com/12Knocksinna/Office365itpros/blob/master/TeamsGroupsActivityReport.ps1
# The Graph-based version of this script (much faster) is available in https://github.com/12Knocksinna/Office365itpros/blob/master/TeamsGroupsActivityReportV5.PS1
#
Clear-Host
# Check that we are connected to Exchange Online, SharePoint Online, and Teams
Write-Host "Checking that prerequisite PowerShell modules are loaded..."
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
If (!($ModulesLoaded -match "Microsoft.Online.SharePoint.PowerShell")) {Write-Host "Please connect to the SharePoint Online module and then restart the script"; break}

# Comment these lines out if you don't want the script to create a temp directory to store its output files
$path = "C:\Temp"
If(!(test-path $path)) {
   New-Item -ItemType Directory -Force -Path $path | Out-Null }

$OrgName = (Get-OrganizationConfig).Name  
       
# OK, we seem to be fully connected to both Exchange Online and SharePoint Online...
Write-Host "Checking Microsoft 365 Groups and Teams in the tenant:" $OrgName
# Setup some stuff we use
$WarningDate = (Get-Date).AddDays(-90); $WarningEmailDate = (Get-Date).AddDays(-365); $Today = (Get-Date); $Date = $Today.ToShortDateString()
$TeamsEnabled = $False; $ObsoleteSPOGroups = 0; $ObsoleteEmailGroups = 0
$Report = [System.Collections.Generic.List[Object]]::new(); $ReportFile = "c:\temp\GroupsActivityReport.html"
$CSVFile = "c:\temp\GroupsActivityReport.csv"
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
           <p><h1>Microsoft 365 Groups and Teams Activity Report</h1></p>
           <p><h3>Generated: " + $date + "</h3></p></div>"
		
# Get a list of Groups in the tenant
Write-Host "Extracting list of Microsoft 365 Groups for checking..."
[Int]$GroupsCount = 0; [int]$TeamsCount = 0; $TeamsList = @{}; $UsedGroups = $False
$Groups = Get-Recipient -RecipientTypeDetails GroupMailbox -ResultSize Unlimited | Sort-Object DisplayName
$GroupsCount = $Groups.Count
# If we don't find any groups (possible with Get-Recipient on a bad day), try to find them with Get-UnifiedGroup before giving up.
If ($GroupsCount -eq 0) { # 
   Write-Host "Fetching Groups using Get-UnifiedGroup"
   $Groups = Get-UnifiedGroup -ResultSize Unlimited | Sort-Object DisplayName 
   $GroupsCount = $Groups.Count; $UsedGroups = $True
   If ($GroupsCount -eq 0) {
     Write-Host "No Microsoft 365 Groups found; script exiting" ; break} 
} # End If

Write-Host "Populating list of Teams..."
If ($UsedGroups -eq $False) { # Populate the Teams hash table with a call to Get-UnifiedGroup
   Get-UnifiedGroup -Filter {ResourceProvisioningOptions -eq "Team"} -ResultSize Unlimited | `
      ForEach-Object { $TeamsList.Add($_.ExternalDirectoryObjectId, $_.DisplayName) } 
} Else { # We already have the $Groups variable populated with data, so extract the Teams from that data
   $Groups | Where-Object {$_.ResourceProvisioningOptions -eq "Team"} | `
      ForEach-Object { $TeamsList.Add($_.ExternalDirectoryObjectId, $_.DisplayName) } 
}
$TeamsCount = $TeamsList.Count

Clear-Host
# Set up progress bar
$ProgDelta = 100/($GroupsCount); $CheckCount = 0; $GroupNumber = 0

# Main loop
ForEach ($Group in $Groups) { #Because we fetched the list of groups with Get-Recipient, the first thing is to get the group properties
   $G = Get-UnifiedGroup -Identity $Group.DistinguishedName
   $GroupNumber++
   $GroupStatus = $G.DisplayName + " ["+ $GroupNumber +"/" + $GroupsCount + "]"
   Write-Progress -Activity "Checking group" -Status $GroupStatus -PercentComplete $CheckCount
   $CheckCount += $ProgDelta;  $ObsoleteReportLine = $G.DisplayName;    $SPOStatus = "Normal"
   $SPOActivity = "Document library in use"; $SPOStorage = 0
   $NumberWarnings = 0;   $NumberofChats = 0;  $TeamsEnabled = $False;  $LastItemAddedtoTeams = "N/A";  $MailboxStatus = $Null; $ObsoleteReportLine = $Null
# Check who manages the group
  $ManagedBy = $G.ManagedBy
  If ([string]::IsNullOrWhiteSpace($ManagedBy) -and [string]::IsNullOrEmpty($ManagedBy)) {
     $ManagedBy = "No owners"
     Write-Host $G.DisplayName "has no group owners!" -ForegroundColor Red}
  Else {
     $ManagedBy = (Get-ExoMailbox -Identity $G.ManagedBy[0]).DisplayName}
# Group Age
  $GroupAge = (New-TimeSpan -Start $G.WhenCreated -End $Today).Days
# Fetch information about activity in the Inbox folder of the group mailbox  
   $Data = (Get-ExoMailboxFolderStatistics -Identity $G.ExternalDirectoryObjectId -IncludeOldestAndNewestITems -FolderScope Inbox)
   If ([string]::IsNullOrEmpty($Data.NewestItemReceivedDate)) {$LastConversation = "No items found"}           
   Else {$LastConversation = Get-Date ($Data.NewestItemReceivedDate) -Format g }
   $NumberConversations = $Data.ItemsInFolder
   $MailboxStatus = "Normal"
  
   If ($Data.NewestItemReceivedDate -le $WarningEmailDate) {
      Write-Host "Last conversation item created in" $G.DisplayName "was" $Data.NewestItemReceivedDate "-> Obsolete?"
      $ObsoleteReportLine = $ObsoleteReportLine + " Last Outlook conversation dated: " + $LastConversation + "."
      $MailboxStatus = "Group Inbox Not Recently Used"
      $ObsoleteEmailGroups++
      $NumberWarnings++ }
   Else
      {# Some conversations exist - but if there are fewer than 20, we should flag this...
      If ($Data.ItemsInFolder -lt 20) {
           $ObsoleteReportLine = $ObsoleteReportLine + " Only " + $Data.ItemsInFolder + " Outlook conversation item(s) found."
           $MailboxStatus = "Low number of conversations"
           $NumberWarnings++}
      }

# Loop to check audit records for activity in the group's SharePoint document library
   If ($null -ne $G.SharePointSiteURL) {
      $SPOStorage = (Get-SPOSite -Identity $G.SharePointSiteUrl).StorageUsageCurrent
      $SPOStorage = [Math]::Round($SpoStorage/1024,2) # SharePoint site storage in GB
      $AuditCheck = $G.SharePointDocumentsUrl + "/*"
      $AuditRecs = $Null
      $AuditRecs = (Search-UnifiedAuditLog -RecordType SharePointFileOperation -StartDate $WarningDate -EndDate $Today -ObjectId $AuditCheck -ResultSize 1)
      If ($null -eq $AuditRecs) {
         #Write-Host "No audit records found for" $SPOSite.Title "-> Potentially obsolete!"
         $ObsoleteSPOGroups++   
         $ObsoleteReportLine = $ObsoleteReportLine + " No SPO activity detected in the last 90 days." }          
       }
   Else
       {
# The SharePoint document library URL is blank, so the document library was never created for this group
        #Write-Host "SharePoint team site never created for the group" $G.DisplayName 
        $ObsoleteSPOGroups++  
        $AuditRecs = $Null
        $ObsoleteReportLine = $ObsoleteReportLine + " SPO document library never created." 
       }
# Report to the screen what we found - but only if something was found...   
  If ($ObsoleteReportLine -ne $G.DisplayName)
     {
     Write-Host $ObsoleteReportLine 
     }
# Generate the number of warnings to decide how obsolete the group might be...   
  If ($null -eq $AuditRecs) {
       $SPOActivity = "No SPO activity detected in the last 90 days"
       $NumberWarnings++ 
}
   If ($null -eq $G.SharePointDocumentsUrl) {
       $SPOStatus = "Document library never created"
       $NumberWarnings++
   }
  
    $Status = "Pass"
    If ($NumberWarnings -eq 1)
       {
       $Status = "Warning"
    }
    If ($NumberWarnings -gt 1)
       {
       $Status = "Fail"
    } 

# If the group is team-enabled, find the date of the last Teams conversation compliance record
If ($TeamsList.ContainsKey($G.ExternalDirectoryObjectId) -eq $True) {
    $TeamsEnabled = $True
    [datetime]$DateOldTeams = "1-Jun-2021" # After this date, Microsoft should have moved the old Teams data to the new location
    $CountOldTeamsData = $False

# Start by looking in the new location (TeamsMessagesData in Non-IPMRoot)
    $TeamsChatData = (Get-ExoMailboxFolderStatistics -Identity $G.ExternalDirectoryObjectId -IncludeOldestAndNewestItems -FolderScope NonIPMRoot | `
      Where-Object {$_.FolderType -eq "TeamsMessagesData" })
    If ($TeamsChatData.ItemsInFolder -gt 0) {$LastItemAddedtoTeams = Get-Date ($TeamsChatData.NewestItemReceivedDate) -Format g}
    $NumberOfChats = $TeamsChatData.ItemsInFolder
    
# If the script is running before 1-Jun-2021, we need to check the old location of the Teams compliance records
If ($Today -lt $DateOldTeams) {
     $CountOldTeamsData = $True
     $OldTeamsChatData = (Get-ExoMailboxFolderStatistics -Identity $G.ExternalDirectoryObjectId -IncludeOldestAndNewestItems -FolderScope ConversationHistory)
     ForEach ($T in $OldTeamsChatData) { # We might have one or two subfolders in Conversation History; find the one for Teams
     If ($T.FolderType -eq "TeamChat") {
        If ($T.ItemsInFolder -gt 0) {$OldLastItemAddedtoTeams = Get-Date ($T.NewestItemReceivedDate) -Format g}
        $OldNumberofChats = $T.ItemsInFolder
}}}

If ($CountOldTeamsData -eq $True) { # We have counted the old date, so let's put the two sets together
   $NumberOfChats = $NumberOfChats + $OldNumberOfChats
   If (!$LastItemAddedToTeams) { $LastItemAddedToTeams = $OldLastItemAddedToTeams }
} # End if

If (($TeamsEnabled -eq $True) -and ($NumberOfChats -le 100)) { Write-Host "Team-enabled group" $G.DisplayName "has only" $NumberOfChats "compliance record(s)" }      
} # End if Processing Teams data

# Generate a line for this group and store it in the report
    $ReportLine = [PSCustomObject][Ordered]@{
          GroupName           = $G.DisplayName
          ManagedBy           = $ManagedBy
          Members             = $G.GroupMemberCount
          ExternalGuests      = $G.GroupExternalMemberCount
          Description         = $G.Notes
          MailboxStatus       = $MailboxStatus
          TeamEnabled         = $TeamsEnabled
          LastChat            = $LastItemAddedtoTeams
          NumberChats         = $NumberofChats
          LastConversation    = $LastConversation
          NumberConversations = $NumberConversations
          SPOActivity         = $SPOActivity
          SPOStorageGB        = $SPOStorage
          SPOStatus           = $SPOStatus
          WhenCreated         = Get-Date ($G.WhenCreated) -Format g
          DaysOld             = $GroupAge
          NumberWarnings      = $NumberWarnings
          Status              = $Status}
   $Report.Add($ReportLine)   
#End of main loop
}

If ($TeamsCount -gt 0) { # We have some teams, so we can calculate a percentage of Team-enabled groups
    $PercentTeams = ($TeamsCount/$GroupsCount)
    $PercentTeams = ($PercentTeams).tostring("P") }
Else {
    $PercentTeams = "No teams found" }

# Create the HTML report
$htmlbody = $Report | ConvertTo-Html -Fragment
$htmltail = "<p>Report created for: " + $OrgName + "
             </p>
             <p>Number of groups scanned: " + $GroupsCount + "</p>" +
             "<p>Number of potentially obsolete groups (based on document library activity): " + $ObsoleteSPOGroups + "</p>" +
             "<p>Number of potentially obsolete groups (based on conversation activity): " + $ObsoleteEmailGroups + "<p>"+
             "<p>Number of Teams-enabled groups    : " + $TeamsCount + "</p>" +
             "<p>Percentage of Teams-enabled groups: " + $PercentTeams + "</body></html>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------"+
             "<p>Microsoft 365 Groups and Teams Activity Report <b>V4.8</b>"	
$htmlreport = $htmlhead + $htmlbody + $htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8
$Report | Export-CSV -NoTypeInformation $CSVFile
$Report | Out-GridView
# Summary note
Clear-Host
Write-Host " "
Write-Host "Results"
Write-Host "-------"
Write-Host "Number of Microsoft 365 Groups scanned                          :" $GroupsCount
Write-Host "Potentially obsolete groups (based on document library activity):" $ObsoleteSPOGroups
Write-Host "Potentially obsolete groups (based on conversation activity)    :" $ObsoleteEmailGroups
Write-Host "Number of Teams-enabled groups                                  :" $TeamsList.Count
Write-Host "Percentage of Teams-enabled groups                              :" $PercentTeams
Write-Host " "
Write-Host "Summary report in" $ReportFile "and CSV in" $CSVFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
