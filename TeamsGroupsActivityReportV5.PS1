# TeamsGroupsActivityReportV5.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/TeamsGroupsActivityReportV5.PS1
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
# V5.0 21-Dec-2020 Use Graph API to get Groups and Teams data
# V5.1 21-Jan-2021 Add check for archived teams
#

CLS
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
       
$OrgName = (Get-OrganizationConfig).Name
# OK, we seem to be fully connected to both Exchange Online and SharePoint Online...
Write-Host "Checking Microsoft 365 Groups and Teams in the tenant:" $OrgName

# Setup some stuff we use
$WarningDate = (Get-Date).AddDays(-90); $WarningEmailDate = (Get-Date).AddDays(-365); $Today = (Get-Date); $Date = $Today.ToShortDateString()
$TeamsGroups = 0;  $TeamsEnabled = $False; $ObsoleteSPOGroups = 0; $ObsoleteEmailGroups = 0
[datetime]$DateOldTeams = "30-Jun-2021" # After this date, Microsoft should have moved the old Teams data to the new location
$Version = "V5.1"

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
[Int]$GroupsCount = 0; [int]$TeamsCount = 0; $UsedGroups = $False

Cls
# Define the values applicable for the application used to connect to the Graph - these will be different for your tenant
$AppId = "0ade5c24-d775-4017-a824-0a993c60787d"
$TenantId = "b762313f-14fc-43a2-9a7a-d2e27f4f3478"
$AppSecret = '7RplGSLWoSs~y5uHYy2041-jbm.4~_r.~q'

# Construct URI and body needed for authentication
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials"
}

# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
# Unpack Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token
$Headers = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $Token" 
            'ConsistencyLevel' = "eventual" }

$S1 = Get-Date #Start of processing

   # Get SharePoint site usage data
   $SPOUsageReportsURI = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D7')"
   $SPOUsage = (Invoke-RestMethod -Uri $SPOUsageReportsURI -Headers $Headers -Method Get -ContentType "application/json") -Replace "...Report Refresh Date", "Report Refresh Date" | ConvertFrom-Csv 
   $DataTable = @{} # Build hashtable with SharePoint usage information per site
   ForEach ($Site in $SPOUsage) {
    If ($Site."Root Web Template" -eq "Group") {
     If ([string]::IsNullOrEmpty($Site."Last Activity Date")) { # No activity for this site 
        $LastActivityDate = $Null }
     Else {
        $LastActivityDate = Get-Date($Site."Last Activity Date") -format g
        $LastActivityDate = $LastActivityDate.Split(" ")[0] }
     $SiteDisplayName = $Site."Owner Display Name".IndexOf("Owners") # Extract site name
     If ($SiteDisplayName -ge 0) {
         $SiteDisplayName = $Site."Owner Display Name".SubString(0,$SiteDisplayName) }
     Else { 
         $SiteDisplayName = $Site."Owner Display Name" }
     $StorageUsed = [string]([math]::round($Site."Storage Used (Byte)"/1GB,2)) + " GB"
     $SingleSiteData = @{
       'DisplayName'      = $SiteDisplayName
       'LastActivityDate' = $LastActivityDate
       'FileCount'        = $Site."File Count" 
       'StorageUsed'      = $StorageUsed }
     $DataTable.Add([String]$Site."Site URL",$SingleSiteData)
     }
}

# Create list of Microsoft 365 Groups in the tenant. We also get a list of Teams. In both cases, we build a hashtable
# to store the object identifier and display name for the group. If more than 100 groups exist, fetch and continue processing using the NextLink
Write-Host "Fetching data about Microsoft 365 Groups..."
$uri = "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(a:a eq 'unified')"
$Groups = Get-GraphData -AccessToken $Token -Uri $uri
$GroupsList = [System.Collections.Generic.List[Object]]::new()
ForEach ($Group in $Groups) {
   # Get Group Owners
   $Uri = "https://graph.microsoft.com/v1.0/groups/" + $Group.Id + "/owners?`$count=true"
   $GroupData = Invoke-RestMethod -Headers $Headers -Uri $Uri -UseBasicParsing -Method "GET"
   If ($GroupData."@odata.count" -eq 0) {
       $OwnerNames = "No owners found" }
   Else  { # Extract owner names
       $x = 0; $OwnerNames = $Null
       ForEach ($Owner in $GroupData.Value) {
          $uri = "https://graph.microsoft.com/v1.0/users/" + $Owner.id
          $OwnerName = Get-GraphData -AccessToken $Token -Uri $uri
            $OwnerName = $OwnerName.DisplayName ; $x++ 
          If ($x -eq 1) { $OwnerNames = $OwnerName } Else { $OwnerNames = $OwnerNames + "; " + $OwnerName }
   }}
   # Get extended group properties 
   $Uri = "https://graph.microsoft.com/v1.0/groups/" + $Group.Id + "?`$select=visibility,description,assignedlabels"
   $GroupData = Get-GraphData -AccessToken $Token -Uri $uri
   $Visibility  = $GroupData.Visibility
   $Description = $GroupData.Description
   $GroupLabel  = $GroupData.AssignedLabels.DisplayName
   # Get SharePoint site URL
   $SPOUrl = $Null; $SPODocLib = $Null; $SPOQuotaUsed = 0; $SPOLastDateActivity = $Null
   $Uri = "https://graph.microsoft.com/v1.0/groups/" + $Group.Id + "/drive"
   $SPOData = Get-GraphData -AccessToken $Token -Uri $uri
   If ($SPOData) {
       $SPOUrl = $SPOData.WebUrl.SubString(0,$SpoData.weburl.IndexOf("/Shared%20Documents"))
       $SPODocLib = $SPOUrl + "/Shared Documents" 
       $SPOQuotaUsed = [Math]::Round($SPOData.quota.used/1Gb,2)
       $SPOLastDateActivity = Get-Date ($SPOData.lastModifiedDateTime) -format g }
   Else { CLS; Write-Host "Continuing to fetch information about Microsoft 365 Groups..." } # Get rid of mucky SharePoint error  308 
     # Get Member and Guest member counts
   $Uri = "https://graph.microsoft.com/beta/groups/" + $Group.Id + "/Members/Microsoft.Graph.User/`$count?`$filter=UserType eq 'Guest'"
   $GuestMemberCount = Get-GraphData -AccessToken $Token -Uri $uri
   $Uri = "https://graph.microsoft.com/beta/groups/" + $Group.Id + "/Members/Microsoft.Graph.User/`$count?`$filter=UserType eq 'Member'"
   $GroupMemberCount = Get-GraphData -AccessToken $Token -Uri $uri   
   # Update list with group information
   $ReportLine = [PSCustomObject][Ordered]@{
       DisplayName      = $Group.DisplayName
       ObjectId         = $Group.Id
       ManagedBy        = $OwnerNames
       GroupMembers     = $GroupMemberCount
       GuestMembers     = $GuestMemberCount
       SharePointURL    = $SPOUrl
       SharePointDocLib = $SPODocLib
       LastSPOActivity  = $SPOLastDateActivity
       WhenCreated      = Get-Date ($Group.createdDateTime) -format g 
       Visibility       = $Visibility
       Description      = $Description
       Label            = $GroupLabel }
  $GroupsList.Add($ReportLine) 
} 
$GroupsList = $GroupsList | Sort DisplayName

# Get Teams
$uri = "https://graph.microsoft.com/beta/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"
$Teams = Get-GraphData -AccessToken $Token -Uri $uri
$TeamsHash = @{}
$Teams.ForEach( {
   $TeamsHash.Add($_.Id, $_.DisplayName) } )

# All groups and teams found...
$TeamsCount = $TeamsHash.Count
$GroupsCount = $GroupsList.Count
If (!$GroupsCount) {Write-Host "No Microsoft 365 Groups can be found - exiting"; break}
If (!$TeamsCount) {Write-Host "No Microsoft Teams found in the tenant - continuing..." }

$S2 = Get-Date # End of fetching
CLS
Write-Host "Fetching data for" $GroupsCount "Microsoft 365 Groups took" ($S2 - $S1).TotalSeconds "seconds"

# Set up progress bar and create output list
$ProgDelta = 100/($GroupsCount); $CheckCount = 0; $GroupNumber = 0
$Report = [System.Collections.Generic.List[Object]]::new(); $ReportFile = "c:\temp\GroupsActivityReport.html"
# Main loop
ForEach ($G in $GroupsList ) { #Because we fetched the list of groups with a Graph call, the first thing is to get the group properties
   $GroupNumber++
   $GroupStatus = $G.DisplayName + " ["+ $GroupNumber +"/" + $GroupsCount + "]"
   Write-Progress -Activity "Checking group" -Status $GroupStatus -PercentComplete $CheckCount
   $CheckCount += $ProgDelta;  $ObsoleteReportLine = $G.DisplayName;    $SPOStatus = "Normal"
   $SPOActivity = "Document library in use"
   $NumberWarnings = 0;   $NumberofChats = 0;  $TeamsChatData = $Null;  $TeamsEnabled = $False;  $LastItemAddedtoTeams = "N/A";  $MailboxStatus = $Null; $ObsoleteReportLine = $Null
# Group Age
  $GroupAge = (New-TimeSpan -Start $G.WhenCreated -End $Today).Days
# Fetch information about activity in the Inbox folder of the group mailbox  
   $Data = (Get-ExoMailboxFolderStatistics -Identity $G.ObjectId -IncludeOldestAndNewestITems -FolderScope Inbox)
   If ([string]::IsNullOrEmpty($Data.NewestItemReceivedDate)) {$LastConversation = "No items found"}           
   Else {$LastConversation = Get-Date ($Data.NewestItemReceivedDate) -Format g }
   $NumberConversations = $Data.ItemsInFolder
   $MailboxStatus = "Normal"

   If ($Data.NewestItemReceivedDate -le $WarningEmailDate) {
      Write-Host "Last conversation item created in" $G.DisplayName "was" $Data.NewestItemReceivedDate "-> Obsolete?"
      $ObsoleteReportLine = "Last Outlook conversation dated: " + $LastConversation + "."
      $MailboxStatus = "Group Inbox Not Recently Used"
      $ObsoleteEmailGroups++
      $NumberWarnings++ }
   Else
      {# Some conversations exist - but if there are fewer than 20, we should flag this...
      If ($Data.ItemsInFolder -lt 20) {
           $ObsoleteReportLine = $ObsoleteReportLine + "Only " + $Data.ItemsInFolder + " Outlook conversation item(s) found."
           $MailboxStatus = "Low number of conversations"
           $NumberWarnings++}
      }

# Check for activity in the group's SharePoint site
   $SPOFileCount = 0; $SPOStorageUsed = "N/A"; $SPOLastActivityDate = $Null; $DaysSinceLastSPOActivity = "N/A"
   If ($G.SharePointURL -ne $Null) {    
      If ($Datatable[$G.SharePointURL]) { # Look up hash table to find usage information for the site
        $ThisSiteData = $Datatable[$G.SharePointURL]
        $SPOFileCount = $ThisSiteData.FileCount
        $SPOStorageUsed = $ThisSiteData.StorageUsed
        $SPOLastActivityDate = $ThisSiteData.LastActivityDate 
        If ($SPOLastActivityDate -ne $Null) {
           $DaysSinceLastSPOActivity = (New-TimeSpan -Start $SPOLastActivityDate -End $Today).Days }
   }
   Else { # The SharePoint document library URL is blank, so the document library was never created for this group
        $ObsoleteSPOGroups++;  
        $ObsoleteReportLine = $ObsoleteReportLine + " SharePoint document library never created." 
       }}
   If ($DaysSinceLastSPOActivity -gt 90) { # No activity in more than 90 days
       $ObsoleteSPOGroups++; $ObsoleteReportLine = $ObsoleteReportLine + " No SPO activity detected in the last 90 days." }  

# Generate the number of warnings to decide how obsolete the group might be...   
   If ($G.SharePointDocLib -eq $Null) {
       $SPOStatus = "Document library never created"
       $NumberWarnings++ }
  
    $Status = "Pass"
    If ($NumberWarnings -eq 1)
       {
       $Status = "Warning"
    }
    If ($NumberWarnings -gt 1)
       {
       $Status = "Fail"
    } 
# Write-Host "Processing" $G.DisplayName
# If the group is team-enabled, find the date of the last Teams conversation compliance record
If ($TeamsHash.ContainsKey($G.ObjectId) -eq $True) {
    $TeamsEnabled = $True
    $CountOldTeamsData = $False

# Start by looking in the new location (TeamsMessagesData in Non-IPMRoot)
    $TeamsChatData = (Get-ExoMailboxFolderStatistics -Identity $G.ObjectId -IncludeOldestAndNewestItems -FolderScope NonIPMRoot | ? {$_.FolderType -eq "TeamsMessagesData" })
    If ($TeamsChatData.ItemsInFolder -gt 0) {$LastItemAddedtoTeams = Get-Date ($TeamsChatData.NewestItemReceivedDate) -Format g}
    $NumberOfChats = $TeamsChatData.ItemsInFolder
    
# If the script is running before 30-Jun-2021, we need to check the old location of the Teams compliance records
If ($Today -lt $DateOldTeams) {
     $CountOldTeamsData = $True
     $OldTeamsChatData = (Get-ExoMailboxFolderStatistics -Identity $G.ObjectId -IncludeOldestAndNewestItems -FolderScope ConversationHistory)
     ForEach ($T in $OldTeamsChatData) { # We might have one or two subfolders in Conversation History; find the one for Teams
     If ($T.FolderType -eq "TeamChat") {
        If ($T.ItemsInFolder -gt 0) {$OldLastItemAddedtoTeams = Get-Date ($T.NewestItemReceivedDate) -Format g }
     $OldNumberofChats = $T.ItemsInFolder
}}}

If ($CountOldTeamsData -eq $True) { # We have counted the old date, so let's put the two sets together
   $NumberOfChats = $NumberOfChats + $OldNumberOfChats
   If (!$LastItemAddedToTeams) { $LastItemAddedToTeams = $OldLastItemAddedToTeams }}
} # End if
    
If (($TeamsEnabled -eq $True) -and ($NumberOfChats -le 100)) { Write-Host "Team-enabled group" $G.DisplayName "has only" $NumberOfChats "compliance record(s)" } 

   # Discover if team is archived
   $Uri = "https://graph.microsoft.com/v1.0/teams/" + $Group.Id
   $TeamDetails = Get-GraphData -AccessToken $Token -Uri $Uri
   If ($TeamDetails.IsArchived -eq $False) { $DisplayName = $G.DisplayName }
   Else { $DisplayName = $G.DisplayName + " (Archived team)" }
# End if Processing Teams data

# Generate a line for this group and store it in the report
    $ReportLine = [PSCustomObject][Ordered]@{
          GroupName           = $DisplayName
          ManagedBy           = $G.ManagedBy
          Visibility          = $G.Visibility
          Members             = $G.GroupMembers
          ExternalGuests      = $G.GuestMembers
          Description         = $G.Description
          MailboxStatus       = $MailboxStatus
          TeamEnabled         = $TeamsEnabled
          LastChat            = $LastItemAddedtoTeams
          NumberChats         = $NumberOfChats
          LastConversation    = $LastConversation
          NumberConversations = $NumberConversations
          LastSPOActivity     = $SPOActivity
          SPOStorageGB        = $SPOStorage
          SPOStatus           = $SPOStatus
          WhenCreated         = $G.WhenCreated
          DaysOld             = $GroupAge
          NumberWarnings      = $NumberWarnings
          Label               = $G.Label
          Status              = $Status}
   $Report.Add($ReportLine)  
$S3 = Get-Date
$TotalSeconds = [math]::round(($S3-$S2).TotalSeconds,2)
$SecondsPerGroup = [math]::round(($TotalSeconds/$GroupNumber),2)
Write-Host "Processed" $GroupNumber "groups in" $TotalSeconds "- Currently processing at" $SecondsPerGroup "seconds per group"
#End of main loop
}

$OverallElapsed = [math]::round(($S3-$S1).TotalSeconds,2)

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
             "<p>Microsoft 365 Groups and Teams Activity Report <b>" + $Version + "</b>"	
$htmlreport = $htmlhead + $htmlbody + $htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8
$Report | Export-CSV -NoTypeInformation $CSVFile
$Report | Out-GridView
# Summary note
CLS
Write-Host " "
Write-Host "Results - Teams and Microsoft 365 Groups Activity Report" $Version
Write-Host "-------------------------------------------------------------"
Write-Host "Number of Microsoft 365 Groups scanned                          :" $GroupsCount
Write-Host "Potentially obsolete groups (based on document library activity):" $ObsoleteSPOGroups
Write-Host "Potentially obsolete groups (based on conversation activity)    :" $ObsoleteEmailGroups
Write-Host "Number of Teams-enabled groups                                  :" $TeamsCount
Write-Host "Percentage of Teams-enabled groups                              :" $PercentTeams
Write-Host " "
Write-Host "Total Elapsed time: " $OverAllElapsed " seconds"
Write-Host "Summary report in" $ReportFile "and CSV in" $CSVFile

#+-------------------------- Functions etc. -------------------------

function Get-GraphData {
# Based on https://danielchronlund.com/2018/11/19/fetch-data-from-microsoft-graph-with-powershell-paging-support/
# GET data from Microsoft Graph.
    param (
        [parameter(Mandatory = $true)]
        $AccessToken,

        [parameter(Mandatory = $true)]
        $Uri
    )

    # Check if authentication was successful.
    if ($AccessToken) {
        # Format headers.
       $Headers = @{
         'Content-Type'  = "application\json"
         'Authorization' = "Bearer $AccessToken" 
         'ConsistencyLevel' = "eventual"  }

        # Create an empty array to store the result.
        $QueryResults = @()

        # Invoke REST method and fetch data until there are no pages left.
        do {
            $Results = ""
            $StatusCode = ""

            do {
                try {
                    $Results = Invoke-RestMethod -Headers $Headers -Uri $Uri -UseBasicParsing -Method "GET" -ContentType "application/json"

                    $StatusCode = $Results.StatusCode
                } catch {
                    $StatusCode = $_.Exception.Response.StatusCode.value__

                    if ($StatusCode -eq 429) {
                        Write-Warning "Got throttled by Microsoft. Sleeping for 45 seconds..."
                        Start-Sleep -Seconds 45
                    }
                    else {
                        Write-Error $_.Exception
                    }
                }
            } while ($StatusCode -eq 429)

            if ($Results.value) {
                $QueryResults += $Results.value
            }
            else {
                $QueryResults += $Results
            }

            $uri = $Results.'@odata.nextlink'
        } until (!($uri))

        # Return the result.
        $QueryResults
    }
    else {
        Write-Error "No Access Token"
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
