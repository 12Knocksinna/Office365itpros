# TeamsGroupsActivityReport.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/TeamsGroupsActivityReportV5.PS1
# A script to check the activity of Microsoft 365 Groups and Teams and report the groups and teams that might be deleted because they're not used.
# We check the group mailbox to see what the last time a conversation item was added to the Inbox folder. 
# Another check sees whether a low number of items exist in the mailbox, which would show that it's not being used.
# We also check the group document library in SharePoint Online to see whether it exists or has been used in the last 180 days.
# And we check Teams usage data to figure out if any chatting is happening.

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
# V5.2 02-Feb-2021 Add option to import Teams usage data from CSV exported from Teams admin center
# V5.3 10-Nov-2021 Removed processing for old Teams compliance records 
# V5.4 12-Mar-2022 Changed check for Teams usage hash table to avoid errors and added explicit check for Teams data file downloaded from the TAC
# V5.5 15-Jun-2022 Recoded way that check to renew access token worked and incorporated automatic fetch of latest Teams usage data
# V5.6 19-Jul-2022 Add email address of team owner to output and updated activity range for usage reports to 180 days
# V5.7 17-Aug-2022 Amended how archived groups are processed.
# V5.8 15-Sep-2022 Improved check for groups with no owners
# V5.9 05-Jan-2023 Increased access token lifetime from 50 to 57 minutes and created new CheckAccessToken function.
# V5.10 23-Jan-2023 Include code to handle bad data returned by Exchange Online for mailbox inbox statistics
# V5.11 01-Nov-2023 
#
# https://github.com/12Knocksinna/Office365itpros/blob/master/TeamsGroupsActivityReport.ps1
#
=======
# Needs the following Graph application permissions:
# Group.Read.All, Reports.Read.All, User.Read.All, GroupMember.Read.All, Sites.Read.All, Organization.Read.All (or Directory.Read.All)
#
#+-------------------------- Functions etc. -------------------------

<# 
Extra line

Here's some more stuff

Is it online and safe?

#>


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

function GetAccessToken {
# function to return an Oauth access token

# Define the values applicable for the application used to connect to the Graph - These values will not work
$AppId = "328e1143-88e3-492b-bf82-24c4a47ada63"
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3478"
$AppSecret = 'ei_7Q~mY8SLKxKJHkY.x-WTWT0ncfaqu8ETtS'

# Construct URI and body needed for authentication
$Uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$Body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials"
}

# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
# Unpack Access Token
$Token = ($tokenRequest.Content | ConvertFrom-Json).access_token

Write-Host ("Retrieved new access token at {0}" -f (Get-Date)) -foregroundcolor red

$Global:Headers = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $Token" 
            'ConsistencyLevel' = "eventual" }
Return $Token
}

Function CheckAccessToken {
# Function to check if the access token needs to be refreshed. If it does, request a new token
# This often needs to happen when the script processes more than a few thousands groups

$TimeNow = (Get-Date)
if($TimeNow -ge $TokenExpiredDate) {
  $Global:Token = GetAccessToken
  $Global:TokenExpiredDate = (Get-Date).AddMinutes($TimeToRefreshToken) 
  #Write-Host "Requested new access token - expiration at" $TokenExpiredDate 
 }

Return $Token
}

Function GetTeamsStats {
# Function to retrieve per-team usage stats so that there's no need for the admin to download the report. The output is a hash table that
# we check for Teams data

$Uri = "https://graph.microsoft.com/beta/reports/getTeamsTeamActivityDetail(period='D180')?`$format=application/json"
[array]$TeamsData = Get-GraphData -Uri $Uri  -AccessToken $Token
$PerTeamStats = [System.Collections.Generic.List[Object]]::new() 
$TeamsDataHash = @{}
ForEach ($Team in $TeamsData) {
   If (!([string]::IsNullOrWhiteSpace($Team.LastActivityDate))) {
      $DaysSinceActive = (New-Timespan -Start ($Team.LastActivityDate -as [datetime]) -End ($Team.Reportrefreshdate -as [datetime])).Days
      $LastActiveDate = Get-Date ($Team.lastActivityDate) -format dd-MMM-yyyy 
    } Else { $DaysSinceActive = "> 90"
       $LastActiveDate = "More than 90 days ago" 
    }
   $ReportLine  = [PSCustomObject] @{   
      Team            = $Team.teamName
      Privacy         = $Team.teamType
      TeamId          = $Team.teamId
      LastActivity    = $LastActiveDate
      ReportPeriod    = $Team.Details.reportPeriod
      DaysSinceActive = $DaysSinceActive
      ActiveUsers     = $Team.Details.activeUsers
      Posts           = $Team.Details.postMessages
      ChannelMessages = $Team.Details.channelmessages
      Replies         = $Team.Details.replyMessages
      Urgent          = $Team.Details.urgentMessages
      Mentions        = $Team.Details.mentions
      Guests          = $Team.Details.guests
      ActiveChannels  = $Team.Details.activeChannels
      Reactions       = $Team.Details.reactions }
 $PerTeamStats.Add($ReportLine)
 # Update hash file
 $DataLine  = [PSCustomObject] @{  
      Id              = $Team.TeamId
      DisplayName     = $Team.TeamName
      Privacy         = $Team.TeamType
      Posts           = $Team.Details.postMessages
      Replies         = $Team.Details.replyMessages
      Messages        = $Team.Details.channelmessages
      LastActivity    = $LastActiveDate
      DaysSinceActive = $DaysSinceActive }    
   $TeamsDataHash.Add([string]$Team.TeamId, $DataLine)
} #end ForEach
   $TeamsDataHash
}
# ------

Clear-Host
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
# OK, we seem to be fully connected to Exchange Online.     

# Setup some stuff we use
$WarningDate = (Get-Date).AddDays(-90); $WarningEmailDate = (Get-Date).AddDays(-365); $Today = (Get-Date); $Date = $Today.ToShortDateString()
$TeamsEnabled = $False; $ObsoleteSPOGroups = 0; $ObsoleteEmailGroups = 0; $ArchivedTeams = 0
$SharedDocFolder = "/Shared%20Documents" # These values are to allow internationalization of the SPO document library URL. For French, this would be "/Documents%20partages" 
$SharedDocFolder2 = "/Shared Documents"  # Add both values
$Version = "V5.11"
$TimeToRefreshToken = "57" # Token lifetime
Clear-Host
Write-Host ("Teams and Groups Activity Report {0} starting up..." -f $Version)
# Comment these lines out if you don't want the script to create a temp directory to store its output files
$path = "C:\Temp"
If(!(test-path $path)) {
   New-Item -ItemType Directory -Force -Path $path | Out-Null }
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
		
$Global:Token = GetAccessToken
If (!($Token)) {Write-Host "Can't get a valid Azure AD access token - exiting" ; break }
# Set the creation date of the initial access token (59min and 59sec life long) and add 57 minutes to renew the token later in the main loop
$Global:TokenCreationDate = (Get-Date)
$Global:TokenExpiredDate = (Get-date).AddMinutes($TimeToRefreshToken) 
    
# Get tenant name and other details
$Uri = "https://graph.microsoft.com/v1.0/organization"
$OrgData = Get-GraphData -Uri $Uri -AccessToken $Token
$OrgName = $OrgData.DisplayName

$S1 = Get-Date #Start of processing

# Get a list of Groups in the tenant
[Int]$GroupsCount = 0; [int]$TeamsCount = 0; $UsedGroups = $False

# Retrieve Teams usage data
$TeamsUsageHash = GetTeamsStats

# Get SharePoint site usage data
   Write-Host "Retrieving SharePoint Online site usage data..."
   $SPOUsageReportsURI = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D180')"
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
     # Update hash table with details of the site
     Try {
        $DataTable.Add([String]$Site."Site URL",$SingleSiteData) 
         }
     Catch { # Error for some reason - it doesn't make much difference
        Write-Host ("Couldn't add details for site {0} to the DataTable - continuing" -f $SiteDisplayName)
         }
     }
}

# Create list of Microsoft 365 Groups in the tenant. We also get a list of Teams. In both cases, we build a hashtable
# to store the object identifier and display name for the group. 

Write-Host ("Checking Microsoft 365 Groups and Teams for: {0}" -f $OrgName)
Write-Host "This phase can take some time because we need to fetch every group in the organization and"
Write-Host "then analyze its settings and activity. Please wait."
$uri = "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(a:a eq 'unified')"
[array]$Groups = Get-GraphData -AccessToken $Token -Uri $uri
If (!($Groups)) {Write-Host "Can't find any Microsoft 365 Groups - Check your access token"; break}
Write-Host ("Found {0} groups to process. Now analyzing each group." -f $Groups.Count)
$i = 0
$GroupsList = [System.Collections.Generic.List[Object]]::new()
ForEach ($Group in $Groups) { 
   $i++
   Write-Host ("Processing group {0} {1}/{2}" -f $Group.DisplayName, $i, $Groups.Count)
   $GroupOwnerEmail = $Null; $OwnerNames = $Null 
   # Get Group Owners
   $Uri = "https://graph.microsoft.com/v1.0/groups/" + $Group.Id + "/owners?"
   [array]$GroupData = Get-GraphData -Uri $Uri -AccessToken $Token
  If ($GroupData[0].'@odata.type' -eq '#microsoft.graph.user') { # Found a group owner, so extract names
      $OwnerNames = $GroupData.DisplayName -join ", " 
      $GroupOwnerEmail = $GroupData[0].Mail }
   Else { # ownerless group
       $OwnerNames = "No owners found"
       $GroupOwnerEMail = $Null }
   # Get extended group properties 
   $Uri = "https://graph.microsoft.com/v1.0/groups/" + $Group.Id + "?`$select=visibility,description,assignedlabels"
   $GroupData = Get-GraphData -AccessToken $Token -Uri $uri
   $Visibility  = $GroupData.Visibility
   $Description = $GroupData.Description
   $GroupLabel  = $GroupData.AssignedLabels.DisplayName
   # Get SharePoint site URL
   $SPOUrl = $Null; $SPODocLib = $Null; $SPOQuotaUsed = 0; $SPOLastDateActivity = $Null
   $Uri = "https://graph.microsoft.com/v1.0/groups/" + $Group.Id + "/drive"
   [array]$SPOData = Get-GraphData -AccessToken $Token -Uri $uri
   [int]$LLVValue = $SPOData.WebUrl.IndexOf($SharedDocFolder) # Can we find a local-language value for the document library in the data returned?
   If (($SPOData.id) -and ($SPOData.DriveType -eq "documentLibrary")) { # Using language-specific values to identify the document library defined
       If ($LLVValue -gt 0) {  # If we have a local language value, parse it to extract the document library URL
          $SPOUrl = $SPOData.WebUrl.SubString(0,$SPOData.WebUrl.IndexOf($SharedDocFolder))
          $SPODocLib = $SPOUrl + $SharedDocFolder2 
          $SPOQuotaUsed = [Math]::Round($SPOData.quota.used/1Gb,2)
          $SPOLastDateActivity = Get-Date ($SPOData.lastModifiedDateTime) -format g 
       } # End if
       Else  { # Just report what we read from the Graph
         $SPOUrl = $SPOData.WebUrl
         $SPODocLib = $SPOUrl + $SharedDocFolder2 
         $SPOQuotaUsed = [Math]::Round($SPOData.quota.used/1Gb,2)
         $SPOLastDateActivity = Get-Date ($SPOData.lastModifiedDateTime) -format g 
       } # End Else
   } # End if against $SPOData
   Else { 
       Clear-Host; Write-Host "Continuing to fetch information about Microsoft 365 Groups..." 
   } # Get rid of mucky SharePoint error  308 

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
       GroupContact     = $GroupOwnerEmail
       GroupMembers     = $GroupMemberCount
       GuestMembers     = $GuestMemberCount
       SharePointURL    = $SPOUrl
       SharePointDocLib = $SPODocLib
       LastSPOActivity  = $SPOLastDateActivity
       WhenCreated      = Get-Date ($Group.createdDateTime) -format g 
       WhenRenewed      = Get-Date ($Group.renewedDateTime) -format g
       Visibility       = $Visibility
       Description      = $Description
       Label            = $GroupLabel }
  $GroupsList.Add($ReportLine) 
  
  # Check for expired access token and renew if necessary
  $Global:Token = CheckAccessToken
} 
$GroupsList = $GroupsList | Sort-Object DisplayName

Write-Host "Getting information about team-enabled groups..."
# Get Teams
$uri = "https://graph.microsoft.com/V1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"
[array]$Teams = Get-GraphData -AccessToken $Token -Uri $uri
$TeamsHash = @{}
$Teams.ForEach( {
   $TeamsHash.Add($_.Id, $_.DisplayName) } )

# All groups and teams found...
$TeamsCount = $Teams.Count
$GroupsCount = $GroupsList.Count
If (!$GroupsCount) {Write-Host "No Microsoft 365 Groups can be found - exiting"; break}
If (!$TeamsCount) {Write-Host "No Microsoft Teams found in the tenant - continuing..." }

$S2 = Get-Date # End of fetching
Clear-Host
Write-Host "Fetching data for" $GroupsCount "Microsoft 365 Groups took" ($S2 - $S1).TotalSeconds "seconds"

# Set up progress bar and create output list
$ProgDelta = 100/($GroupsCount); $CheckCount = 0; $GroupNumber = 0
$Report = [System.Collections.Generic.List[Object]]::new(); $ReportFile = "c:\temp\GroupsActivityReport.html"
# Main loop
ForEach ($G in $GroupsList ) { #Because we fetched the list of groups with a Graph call, the first thing is to get the group properties
   $GroupNumber++
   $DisplayName = $G.DisplayName
   $SPOStatus = "SPO: OK"; $MailboxStatus = "Mbx: OK"; $TeamsStatus = "Teams: OK"
   $GroupStatus = $DisplayName + " ["+ $GroupNumber +"/" + $GroupsCount + "]"
   Write-Progress -Activity "Analyzing and reporting group" -Status $GroupStatus -PercentComplete $CheckCount
   $CheckCount += $ProgDelta;  $ObsoleteReportLine = $DisplayName
   $SPOActivity = "Document library in use"
   $NumberWarnings = 0;   $NumberofChats = 0;  $TeamsChatData = $Null;  $TeamsEnabled = $False;  $LastItemAddedtoTeams = "N/A";  $ObsoleteReportLine = $Null

   # Check for expired access token and renew if necessary
   $Global:Token = CheckAccessToken
  
# Group Age
  $GroupAge = (New-TimeSpan -Start $G.WhenCreated -End $Today).Days
# Team-enabled or not?
$GroupIsTeamEnabled = $False
If ($TeamsHash.ContainsKey($G.ObjectId) -eq $True) {$GroupIsTeamEnabled = $True}

 If ($GroupIsTeamEnabled -eq $False) { # Not a Teams-enabled group, so look at the Inbox etc.
 # Fetch information about activity in the Inbox folder of the group mailbox  
 Try {  
   [array]$Data = (Get-ExoMailboxFolderStatistics -Identity $G.ObjectId -IncludeOldestAndNewestITems -FolderScope Inbox -ErrorAction SilentlyContinue) }
   Catch {
      Write-Host ("Can't read information from the group mailbox for {0} - continuing." -f $G.DisplayName)
      $Data = $Null
   }
   If ([string]::IsNullOrEmpty($Data.NewestItemReceivedDate)) {
       $LastConversation = "No items found"
       $NumberConversations = 0
   } Else {
     Try {
       $LastConversation = (Get-Date ($Data.NewestItemReceivedDate) -Format g )
       $NumberConversations = $Data.ItemsInFolder }
     Catch [System.Management.Automation.ParameterBindingException] {
       # Caused by the Exo cmdlets returning bad dates, so we need to do this...
       [string]$LastDate = $Data.NewestItemReceivedDate[0]
       [datetime]$LastDateConversation = $LastDate
       $LastConversation = Get-Date($LastDateConversation) -format g
       [string]$NumberConversations = $Data.ItemsInfolder[0].toString()
       $Error.Clear()
        }
     Catch {
      Write-Host "Error converting date values"
     }
   }

    If ($G.resourceBehaviorOptions -eq "CalendarMemberReadOnly") {
     $LastConversation = "Yammer group" }

   If ($LastConversation -le $WarningEmailDate) {
      # Write-Host "Last conversation item created in" $G.DisplayName "was" $LastConversation "-> Obsolete?"
      $ObsoleteReportLine = "Last Outlook conversation dated: " + $LastConversation + "."
      $MailboxStatus = "Group Inbox Not Recently Used"
      $ObsoleteEmailGroups++
      $NumberWarnings++ }
   Else
      {# Some conversations exist - but if there are fewer than 20, we should flag this...
      If ($NumberConversations -lt 20) {
           $ObsoleteReportLine = $ObsoleteReportLine + "Only " + $NumberConversations + " Outlook conversation item(s) found."
           $MailboxStatus = "Low number of conversations"
           $NumberWarnings++}
      }
  } # End If
  Else { # It's a team-enabled group, so we don't need to check the mailbox and so populate the values appropriately
     $LastConversation = "Teams-enabled group"
     $NumberConversations = "N/A" 
  } #End Else

# Check for activity in the group's SharePoint site
   $SPOFileCount = 0; $SPOStorageUsed = "N/A"; $SPOLastActivityDate = $Null; $DaysSinceLastSPOActivity = "N/A"
   If ($Null -ne $G.SharePointURL) {    
      If ($Datatable[$G.SharePointURL]) { # Look up hash table to find usage information for the site
        $ThisSiteData = $Datatable[$G.SharePointURL]
        $SPOFileCount = $ThisSiteData.FileCount
        $SPOStorageUsed = $ThisSiteData.StorageUsed
        $SPOLastActivityDate = $ThisSiteData.LastActivityDate 
        If ($Null -ne $SPOLastActivityDate) {
           $DaysSinceLastSPOActivity = (New-TimeSpan -Start $SPOLastActivityDate -End $Today).Days }
   }
   Else { # The SharePoint document library URL is blank, so the document library was never created for this group
        $ObsoleteSPOGroups++;  
        $ObsoleteReportLine = $ObsoleteReportLine + " SharePoint document library never created." 
       }}
   If ($DaysSinceLastSPOActivity -gt 90) { # No activity in more than 90 days
       $ObsoleteSPOGroups++; $ObsoleteReportLine = $ObsoleteReportLine + " No SPO activity detected in the last 90 days." }   

# Generate warnings for SPO 
   If ($Null -ne $G.SharePointDocLib) {
       $SPOStatus = "Document library never created"
       $NumberWarnings++ }

# Write-Host "Processing" $G.DisplayName
# If the group is team-enabled, find the date of the last Teams conversation compliance record
If ($GroupIsTeamEnabled -eq $True) { # We have a team-enabled group
    $TeamsEnabled = $True; $CountOldTeamsData = $False; $NumberOfChats = 0; $LastItemAddedToTeams = $Null
    If (-not $TeamsUsageHash.ContainsKey($G.ObjectId)) { # Check do we have Teams usage data stored in a hash table 
    # Nope, so we have to get the data from Exchange Online by looking in the TeamsMessagesData file in the non-IPM root
       Write-Host "Checking Exchange Online for Teams activity data..."
       Try {
        [array]$TeamsChatData = (Get-ExoMailboxFolderStatistics -Identity $G.ObjectId -IncludeOldestAndNewestItems -FolderScope NonIPMRoot -ErrorAction SilentlyContinue | `
          Where-Object {$_.FolderType -eq "TeamsMessagesData" })
           }
       Catch { # Report the error
         Write-Host ("Error fetching Team message data for {0} - continuing" -f $G.DisplayName) 
             }    
       If ($TeamsChatData.ItemsInFolder -gt 0) {$LastItemAddedtoTeams = Get-Date ($TeamsChatData.NewestItemReceivedDate) -Format g}
       $NumberOfChats = $TeamsChatData.ItemsInFolder }
   Else { # Read the data from the Teams usage data
    # Write-Host "Reading from Teams Hash Table for" $ThisTeamData.DisplayName
    $ThisTeamData = $TeamsUsageHash[$G.ObjectId]
    $NumberOfChats = [int]$ThisTeamData.Posts + [int]$ThisTeamData.Replies
    $LastItemAddedToTeams = Get-Date ($ThisTeamData.LastActivity) -Format g
   } #End Else
} # End if
    
#  Increase warnings if Teams activity is low
   If ($NumberOfChats -lt 20) { 
      $NumberWarnings++
      $TeamsStatus = "Low number of Teams conversations" }

  # If (($TeamsEnabled -eq $True) -and ($NumberOfChats -le 100)) { Write-Host "Team-enabled group" $G.DisplayName "has only" $NumberOfChats "compliance record(s)" }    
   # Discover if team is archived
   If ($TeamsEnabled -eq $True) {
     $Uri = "https://graph.microsoft.com/v1.0/teams/" + $G.ObjectId
     $TeamDetails = Get-GraphData -AccessToken $Token -Uri $Uri 
     Switch ($TeamDetails.IsArchived) {
      $False { $DisplayName = $G.DisplayName }
      $True  { $DisplayName = $G.DisplayName + " (Archived team)" 
               $ArchivedTeams++}
   }
 }
# End of Processing Teams data

# Calculate status
$Status = $MailboxStatus,$SpoStatus,$TeamsStatus -join ", "
$OverallStatus = "Pass"
If ($NumberWarnings -gt 1) { $OverallStatus = "Issues"}
If ($NumberWarnings -gt 2) { $OverallStatus = "Fail" }
    
# Generate a line for this group and store it in the report
    $ReportLine = [PSCustomObject][Ordered]@{
          GroupName               = $DisplayName
          ManagedBy               = $G.ManagedBy
	  ContactEmail            = $G.GroupContact
          Visibility              = $G.Visibility
          Members                 = $G.GroupMembers
          "External Guests"       = $G.GuestMembers
          Description             = $G.Description
          "Sensitivity Label"     = $G.Label
          "Team Enabled"          = $TeamsEnabled
          "Last Teams message"    = $LastItemAddedtoTeams
          "Number Teams messages" = $NumberOfChats
          "Last Email Inbox"      = $LastConversation
          "Number Email Inbox"    = $NumberConversations
          "Last SPO Activity"     = $SPOLastActivityDate
          "SPO Storage Used (GB)" = $SPOStorageUsed
          "Number SPO Files"      = $SPOFileCount
	  "SPO Site URL"          = $G.SharePointURL
          "Date Created"          = $G.WhenCreated
          "Days Old"              = $GroupAge       
           NumberWarnings         = $NumberWarnings
           Status                 = $Status
           "Overall Result"       = $OverallStatus }
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
Clear-Host
Write-Host " "
Write-Host "Results - Teams and Microsoft 365 Groups Activity Report" $Version
Write-Host "-------------------------------------------------------------"
Write-Host ("Number of Microsoft 365 Groups scanned:                          {0}" -f $GroupsCount)
Write-Host ("Potentially obsolete groups (based on document library activity: {0}" -f $ObsoleteSPOGroups)
Write-Host ("Potentially obsolete groups (based on conversation activity):    {0}" -f $ObsoleteEmailGroups)
Write-Host ("Number of Teams-enabled groups:                                  {0}" -f $TeamsCount)
Write-Host ("Number of archived teams:                                        {0}" -f $ArchivedTeams)
Write-Host ("Percentage of Teams-enabled groups:                              {0}" -f $PercentTeams)
Write-Host " "
Write-Host "Total Elapsed time: " $OverAllElapsed "seconds"
Write-Host "Summary report in" $ReportFile "and CSV in" $CSVFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
