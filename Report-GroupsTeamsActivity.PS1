# TeamsGroupsActivityReport.PS1
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
# V5.12 15-Feb-2024 Add check for Teams that haven't had a channel message for more than 90 days
# V5.13 18-Mar-2024 Include workaround because of problems with the SharePoint usage API
# V6.0  05-Jul-2024 Conversion to use the Microsoft Graph PowerShell SDK
# 
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-GroupsTeamsActivity.PS1
# See https://office365itpros.com/2022/03/14/microsoft-365-groups-teams-activity-report/ for more information
#
=======
# Uses an Entra ID registered app with consent for the following Graph application permissions:
# Group.Read.All:           Read groups
# Reports.Read.All:         Read usage reports
# User.Read.All:            Read users
# GroupMember.Read.All:     Read group membership 
# Sites.Read.All:           Read SharePoint Online sites for groups
# Organization.Read.All:    Read details of tenant
# Teams.ReadBasic.All:      Read basic Teams information
#
#+-------------------------- Functions etc. -------------------------

Function Get-TeamsStats {
# Function to retrieve per-team usage stats so that there's no need for the admin to download the report. The output is a hash table that
# we check for Teams data
[array]$TeamsData = $null
Remove-Item $TempDataFile -ErrorAction SilentlyContinue
[array]$TeamsData = Get-MgReportTeamActivityDetail -Period D180 -OutFile $TempDataFile
[array]$TeamsData = Import-Csv -Path $TempDataFile | Sort-Object 'Team Name'
$PerTeamStats = [System.Collections.Generic.List[Object]]::new() 
$TeamsDataHash = @{}
ForEach ($Team in $TeamsData) {
    If (!([string]::IsNullOrWhiteSpace($Team.'Last Activity Date'))) {
        $DaysSinceActive = (New-Timespan -Start ($Team.'Last Activity Date' -as [datetime]) -End ($Team.'Report Refresh date' -as [datetime])).Days
        $LastActiveDate = Get-Date ($Team.'Last Activity Date') -format dd-MMM-yyyy 
    } Else { 
        $DaysSinceActive = "> 90"
        $LastActiveDate = "More than 90 days ago" 
    }
    $ReportLine  = [PSCustomObject] @{   
        Team            = $Team.'Team Name'
        Privacy         = $Team.'Team Type'
        TeamId          = $Team.'Team Id'
        LastActivity    = $LastActiveDate
        ReportPeriod    = $Team.'Report Period'
        DaysSinceActive = $DaysSinceActive
        ActiveUsers     = $Team.'Active Users'
        Posts           = $Team.'Post Messages'
        ChannelMessages = $Team.'Channel Messages'
        Replies         = $Team.'Reply Messages'
        Urgent          = $Team.'Urgent Messages'
        Mentions        = $Team.Mentions
        Guests          = $Team.Guests
        ActiveChannels  = $Team.'Active Channels'
        ActiveGuests    = $Team.'Active external users'
        Reactions       = $Team.Reactions 
    }
    $PerTeamStats.Add($ReportLine)
    # Update hash file
    $DataLine  = [PSCustomObject] @{  
        Id              = $Team.'Team Id'
        DisplayName     = $Team.'Team Name'
        Privacy         = $Team.'Team Type'
        Posts           = $Team.'Post Messages'
        Replies         = $Team.'Reply Messages'
        Messages        = $Team.'Channel messages'
        LastActivity    = $LastActiveDate
        DaysSinceActive = $DaysSinceActive 
    }    
    $TeamsDataHash.Add([string]$Team.'Team Id', $DataLine)
} #end ForEach
   $TeamsDataHash
}
# ------

Clear-Host
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $ModulesLoaded) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}    

$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f3478"
$AppId = "828e1143-88e3-492b-bf82-24c4a47ada63"
$CertificateThumbprint = "F79286DB88C21491110109A0222348FACF694CBD"
# Connect to the Microsoft Graph
Connect-MgGraph -NoWelcome -AppId $AppId -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId

$Global:TempDataFile = 'c:\temp\TempData.csv'
$DataObfuscationOn = $False
$Headers = @{}
$Headers.Add("consistencyLevel", "eventual")
$Headers.Add("Content-Type", "application/json")

# We need full access to report data, so check the setting and update if necessary
If ((Get-MgBetaAdminReportSetting).DisplayConcealedNames -eq $True) {
    $DataObfuscationOn = $True
    $Parameters = @{ displayConcealedNames = $False }
    Update-MgBetaAdminReportSetting -BodyParameter $Parameters
}

# Setup some stuff we use
$WarningEmailDate = (Get-Date).AddDays(-365); $Today = (Get-Date)
$TeamsEnabled = $False; $ObsoleteSPOGroups = 0; $ObsoleteEmailGroups = 0; $ArchivedTeams = 0
$SharedDocFolder = "/Shared%20Documents" # These values are to allow internationalization of the SPO document library URL. For French, this would be "/Documents%20partages" 
$SharedDocFolder2 = "/Shared Documents"  # Add both values
$Version = "V6.00"

$RunDate = Get-Date -format 'dd-MMM-yyyy HH:mm'
# Get tenant name 
$OrgName = Get-MgOrganization | Select-Object -ExpandProperty DisplayName
$HeaderInfo = ("Report generated for the {0} tenant on {1}" -f $OrgName, $RunDate)

$htmlhead="<html>
	<style>
	BODY{font-family: Arial; font-size: 8pt;}
	H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	TD{border: 1px solid #969595; padding: 5px; }
	</style>
	<body>
        <div align=center>
        <p><h1>Microsoft 365 Groups and Teams Activity Report</h1></p>
        <p><h3>" + $HeaderInfo + "</h3></p></div>"
		
Clear-Host
Write-Host ("Teams and Groups Activity Report {0} starting up..." -f $Version)

$S1 = Get-Date #Start of processing

# Get a list of Groups in the tenant
[Int]$GroupsCount = 0; [int]$TeamsCount = 0

# Retrieve Teams usage data
$TeamsUsageHash = Get-TeamsStats

# Get SharePoint site usage data
Write-Host "Retrieving SharePoint Online site usage data..."

Remove-Item $TempDataFile -ErrorAction SilentlyContinue
[array]$SPOUsage = Get-MgReportSharePointSiteUsageDetail -Period D180 -Outfile $TempDataFile
[array]$SPOUsage = Import-Csv -Path $TempDataFile | Sort-Object 'Owner display name'

# This code inserted to fetch site information and build a hash table of SharePoint site URLs
# to fix problem where Microsoft's API doesn't turn the URL in the usage data
$DataTable = @{}
[array]$SPOSiteData = Get-MgSite -All -PageSize 999 
Write-Host ("Processing details for {0} sites..." -f $SPOSiteData.count)
$SPOSiteLookup = @{}
ForEach ($S in $SPOSiteData) {
    $SPOSiteLookup.Add([string]$S.id.split(",")[1], [string]$S.weburl)
}

ForEach ($Site in $SPOUsage) {
    If ($Site."Root Web Template" -eq "Group") {
        If ([string]::IsNullOrEmpty($Site."Last Activity Date")) { # No activity for this site 
            $LastActivityDate = $Null 
        } Else {
            $LastActivityDate = Get-Date($Site."Last Activity Date") -format g
            $LastActivityDate = $LastActivityDate.Split(" ")[0] 
        }
        $SiteDisplayName = $Site."Owner Display Name".IndexOf("Owners") # Extract site name
        If ($SiteDisplayName -ge 0) {
            $SiteDisplayName = $Site."Owner Display Name".SubString(0,$SiteDisplayName) 
        } Else { 
            $SiteDisplayName = $Site."Owner Display Name" 
        }
        $StorageUsed = [string]([math]::round($Site."Storage Used (Byte)"/1GB,2)) + " GB"
        # Fix inserted here because Microsoft has screwed up the SharePoint site usage API
        $SiteURL = $Site.'Site URL'
        If ([string]::IsNullOrEmpty($SiteURL)) {
            $SiteURL = $SPOSiteLookup[$Site.'Site Id']
        }
        $SingleSiteData = @{
            'DisplayName'      = $SiteDisplayName
            'LastActivityDate' = $LastActivityDate
            'FileCount'        = $Site."File Count" 
            'StorageUsed'      = $StorageUsed }
        # Update hash table with details of the site
        Try {
            $DataTable.Add([String]$SiteURL,$SingleSiteData) 
        }
        Catch { # Error for some reason - it doesn't make much difference
            Write-Host ("Couldn't add details for site {0} to the DataTable - continuing" -f $SiteDisplayName)
        }
    }
}

# Create list of Microsoft 365 Groups in the tenant. We also get a list of Teams. In both cases, we build a hashtable
# to store the object identifier and display name for the group. 

Write-Host ("Checking Microsoft 365 Groups and Teams for: {0}" -f $OrgName)
Write-Host "This phase can take some time because we need to fetch every group in the organization to be able to"
Write-Host "analyze its settings and activity. Please wait..."
[array]$Groups = Get-MgGroup -Filter "groupTypes/any(a:a eq 'unified')" -PageSize 999 -All `
    -Property id, displayname, visibility, assignedlabels, description, createdDateTime, renewedDateTime, drive `
    -Sort "displayName DESC" -ConsistencyLevel "eventual" -CountVariable NumberGroups
If (!($Groups)) {
    Write-Host "Can't find any Microsoft 365 Groups"; break
}
Write-Host ("Found {0} groups to process. Now analyzing each group." -f $Groups.Count)
$i = 0
$GroupsList = [System.Collections.Generic.List[Object]]::new()
ForEach ($Group in $Groups) { 
    $i++
    Write-Host ("Processing group {0} {1}/{2}" -f $Group.DisplayName, $i, $Groups.Count)  
    # Check for expired access token and renew if necessary
    $GroupOwnerEmail = $Null; $OwnerNames = $Null 
    # Get Group Owners
    [array]$GroupData = Get-MgGroupOwner -GroupId $Group.Id
    [array]$GroupOwners = $GroupData.AdditionalProperties
    If ($GroupOwners) {
        If ($GroupOwners[0].'@odata.type' -eq '#microsoft.graph.user') { # Found a group owner, so extract names
            $OwnerNames = $GroupOwners.displayName -join ", " 
            $GroupOwnerEmail = $GroupOwners[0].mail 
        } Else { # ownerless group
            $OwnerNames = "No owners found"
            $GroupOwnerEMail = $null 
        }
    }  
    # Get SharePoint site URL
    $SPOUrl = $Null; $SPODocLib = $Null; $SPOLastDateActivity = $Null
   
    $Uri = ("https://graph.microsoft.com/v1.0/groups/{0}/drive" -f $Group.Id)
    Try {
        [array]$SPOData = Invoke-MgGraphRequest -Uri $Uri -Method Get
    } Catch {
        Write-Host ("Error fetching SharePoint data for {0}. Site migh be archived. Continuing..." -f $Group.DisplayName)
        $SPOData = $null
    }
    If ($SPOData) {
        [int]$LLVValue = $SPOData.WebUrl.IndexOf($SharedDocFolder) # Can we find a local-language value for the document library in the data returned?
        If (($SPOData.id) -and ($SPOData.DriveType -eq "documentLibrary")) { # Using language-specific values to identify the document library defined
            If ($LLVValue -gt 0) {  # If we have a local language value, parse it to extract the document library URL
                $SPOUrl = $SPOData.WebUrl.SubString(0,$SPOData.WebUrl.IndexOf($SharedDocFolder))
                $SPODocLib = $SPOUrl + $SharedDocFolder2 
                # $SPOQuotaUsed = [Math]::Round($SPOData.quota.used/1Gb,2)
                $SPOLastDateActivity = Get-Date ($SPOData.lastModifiedDateTime) -format 'dd-MMM-yyyy HH:mm'
            } Else  { # Just report what we read from the Graph
                $SPOUrl = $SPOData.WebUrl
                $SPODocLib = $SPOUrl + $SharedDocFolder2 
                # $SPOQuotaUsed = [Math]::Round($SPOData.quota.used/1Gb,2)
                $SPOLastDateActivity = Get-Date ($SPOData.lastModifiedDateTime) -format 'dd-MMM-yyyy HH:mm'
            } 
        }
    } Else {
        # Drive is probably locked because the site is archived, so get what we can...
        $Data = $SPOSiteData | Where-Object {$_.Name -eq $Group.displayname}
        $SPOUrl = $Data.WebUrl
        $SPODocLib = $null
        $SPOLastDateActivity = "Site archived"
    }

   # Get Member and Guest member counts
   [array]$Members = Get-MgGroupMember -GroupId $Group.Id -All -PageSize 500
   [array]$Members = $Members.AdditionalProperties
   [array]$GuestMembers = $Members | Where-Object {$_.userPrincipalName -like "*#EXT#*"}
 
   # Update list with group information
   $ReportLine = [PSCustomObject][Ordered]@{
       DisplayName      = $Group.DisplayName
       ObjectId         = $Group.Id
       ManagedBy        = $OwnerNames
       GroupContact     = $GroupOwnerEmail
       GroupMembers     = $Members.count
       GuestMembers     = $GuestMembers.count
       SharePointURL    = $SPOUrl
       SharePointDocLib = $SPODocLib
       LastSPOActivity  = $SPOLastDateActivity
       WhenCreated      = Get-Date ($Group.createdDateTime) -format 'dd-MMM-yyyy HH:mm'
       WhenRenewed      = Get-Date ($Group.renewedDateTime) -format 'dd-MMM-yyyy HH:mm'
       Visibility       = $Group.visibility
       Description      = $Group.Description
       Label            = ($Group.assignedLabels.displayName -join ", ")
    }
  $GroupsList.Add($ReportLine) 

} 

Write-Host "Getting information about team-enabled groups..."
# Get Teams
[array]$Teams = Get-MgTeam -All -PageSize 999 | Sort-Object DisplayName
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
Write-Host ("Fetching data for {0} Microsoft 365 Groups took {1} seconds" -f $GroupsCount, (($S2 - $S1).TotalSeconds))

# Set up progress bar and create output list
$ProgDelta = 100/($GroupsCount); $CheckCount = 0; $GroupNumber = 0
$Report = [System.Collections.Generic.List[Object]]::new()
# Main loop
ForEach ($G in $GroupsList ) { #Because we fetched the list of groups with a Graph call, the first thing is to get the group properties
    $GroupNumber++
    $DisplayName = $G.DisplayName
    $SPOStatus = "SPO: OK"; $MailboxStatus = "Mbx: OK"; $TeamsStatus = "Teams: OK"
    $GroupStatus = $DisplayName + " ["+ $GroupNumber +"/" + $GroupsCount + "]"
    Write-Progress -Activity "Analyzing and reporting group" -Status $GroupStatus -PercentComplete $CheckCount
    $CheckCount += $ProgDelta;  $ObsoleteReportLine = $DisplayName
    $NumberWarnings = 0;   $NumberofChats = 0;  $TeamsChatData = $Null;  $TeamsEnabled = $False;  [string]$LastItemAddedtoTeams = "N/A";  $ObsoleteReportLine = $Null
 
# Group Age
    $GroupAge = (New-TimeSpan -Start $G.WhenCreated -End $Today).Days
# Team-enabled or not?
    $GroupIsTeamEnabled = $False
    If ($TeamsHash[$G.ObjectId]) {
        $GroupIsTeamEnabled = $True
    }

    If ($GroupIsTeamEnabled -eq $False) { # Not a Teams-enabled group, so look at the Inbox etc.
    # Fetch information about activity in the Inbox folder of the group mailbox  
    Try {  
        [array]$Data = (Get-ExoMailboxFolderStatistics -Identity $G.ObjectId -IncludeOldestAndNewestITems -FolderScope Inbox -ErrorAction SilentlyContinue) 
    } Catch {
        Write-Host ("Can't read information from the group mailbox for {0} - continuing." -f $G.DisplayName)
        $Data = $Null
    }
    If ([string]::IsNullOrEmpty($Data.NewestItemReceivedDate)) {
       $LastConversation = "No items found"
       $NumberConversations = 0
    } Else {
        Try {
            $LastConversation = Get-Date ($Data.NewestItemReceivedDate) -Format 'dd-MMM-yyyy HH:mm'
            $NumberConversations = $Data.ItemsInFolder 
        } Catch [System.Management.Automation.ParameterBindingException] {
            # Caused by the Exo cmdlets returning bad dates, so we need to do this...
            [string]$LastDate = $Data.NewestItemReceivedDate[0]
            [datetime]$LastDateConversation = $LastDate
            $LastConversation = Get-Date($LastDateConversation) -format 'dd-MMM-yyyy HH:mm'
            [string]$NumberConversations = $Data.ItemsInfolder[0].toString()
            $Error.Clear()
        } Catch {
            Write-Host "Error converting date values"
        }
    }

    If ($G.resourceBehaviorOptions -eq "CalendarMemberReadOnly") {
        $LastConversation = "Yammer group"
    }

    If ($LastConversation -le $WarningEmailDate) {
      # Write-Host "Last conversation item created in" $G.DisplayName "was" $LastConversation "-> Obsolete?"
        $ObsoleteReportLine = ("Last Outlook conversation dated {0}." -f $LastConversation)
        $MailboxStatus = "Group Inbox Not Recently Used"
        $ObsoleteEmailGroups++
        $NumberWarnings++ 
    } Else { # Some conversations exist - but if there are fewer than 20, we should flag this...
        If ($NumberConversations -lt 20) {
           $ObsoleteReportLine = $ObsoleteReportLine + "Only " + $NumberConversations + " Outlook conversation item(s) found."
           $MailboxStatus = "Low number of conversations"
           $NumberWarnings++
        }
    }
  }  Else { # It's a team-enabled group, so we don't need to check the mailbox and so populate the values appropriately
        $LastConversation = "Teams-enabled group"
        $NumberConversations = "N/A" 
  } 

# Check for activity in the group's SharePoint site
   $SPOFileCount = 0; $SPOStorageUsed = "N/A"; $SPOLastActivityDate = $Null; $DaysSinceLastSPOActivity = "N/A"
   If ($Null -ne $G.SharePointURL) {    
      If ($Datatable[$G.SharePointURL]) { # Look up hash table to find usage information for the site
        $ThisSiteData = $Datatable[$G.SharePointURL]
        $SPOFileCount = $ThisSiteData.FileCount
        $SPOStorageUsed = $ThisSiteData.StorageUsed
        $SPOLastActivityDate = $ThisSiteData.LastActivityDate 
        If ($Null -ne $SPOLastActivityDate) {
           $DaysSinceLastSPOActivity = (New-TimeSpan -Start $SPOLastActivityDate -End $Today).Days 
        }
   } Else { # The SharePoint document library URL is blank, so the document library was never created for this group
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
    $TeamsEnabled = $True; $NumberOfChats = 0; [string]$LastItemAddedToTeams = $Null
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
       If ($TeamsChatData.ItemsInFolder -gt 0) {
            [datetime]$LastItemAddedtoTeams = $TeamsChatData.NewestItemReceivedDate
            $LastItemAddedtoTeams = Get-Date ($LastItemAddedtoTeams) -Format 'dd-MMM-yyyy HH:mm'
       }
       $NumberOfChats = $TeamsChatData.ItemsInFolder 
    } Else { # Read the data from the Teams usage data
        $ThisTeamData = $TeamsUsageHash[$G.ObjectId]
        $NumberOfChats = [int]$ThisTeamData.Posts + [int]$ThisTeamData.Replies
        [int]$DaysSinceTeamsActivity = 0; [string]$LastItemAddedToTeams = $null
        If (!([string]::IsNullOrWhitespace($ThisTeamData.LastActivity)) -and ($ThisTeamData.LastActivity -ne 'More than 90 days ago')) {
            [datetime]$LastItemAddedToTeams = [datetime]::ParseExact($ThisTeamData.LastActivity, "dd-MMM-yyyy", $null)
            # [datetime]$LastItemAddedToTeams = $ThisTeamData.LastActivity
            $LastItemAddedToTeams = Get-Date ($LastItemAddedtoTeams) -Format 'dd-MMM-yyyy'
            [int]$DaysSinceTeamsActivity = (New-TimeSpan $LastItemAddedtoTeams).Days
        }
    } #End Else
} # End if
    
#  Increase warnings if Teams activity is low
   If ($NumberOfChats -lt 20) { 
        $NumberWarnings++
        $TeamsStatus = "Low number of Teams conversations" 
    }
    If ($DaysSinceTeamsActivity -gt 90) {
        $NumberWarnings++
        $TeamsStatus = ("{0} last item added {1} days ago" -f $TeamsStatus, $DaysSinceTeamsActivity)
    }

   # Discover if team is archived
   If ($TeamsEnabled -eq $True) {
        Try {
            $TeamDetails = Get-MgTeam -TeamId $G.ObjectId -Property IsArchived, Id, displayName
        } Catch {
            Out-Null
        }
        If ([string]::IsNullOrWhitespace($TeamDetails)) {
            Write-Host ("Error reading team details for {0} ({1})" -f $G.displayName, $G.ObjectId)
            Write-Host "Check the Teams admin center for this team - it might be deleted but still shows up in the teams list"
        } Else {
            Switch ($TeamDetails.IsArchived) {
                $False { 
                    $DisplayName = $G.DisplayName 
                }
                $True { 
                    $DisplayName = $G.DisplayName + " (Archived team)" 
                    $ArchivedTeams++
                }
            }
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
    "Overall Result"       = $OverallStatus 
}
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

# Generate the HTML report
$HTMLReport = $HTMLHead + $HTMLBody + $HTMLTail
$HTMLReportFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Teams and Groups Activity Report.html"
$HTMLReport | Out-File $HTMLReportFile -Encoding utf8

# Generate the report in either Excel worksheet or CSV format, depending on if the ImportExcel module is available
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Teams and Groups Activity Report.xlsx"
    $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Teams and Groups activity" -Title ("Teams and Groups Activity Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "Microsoft365LicensingReport" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Teams and Groups Activity Report.CSV"
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}

$Report | Out-GridView
# Summary details
Clear-Host
Write-Host " "
Write-Host ("Results - Teams and Microsoft 365 Groups Activity Report {0}" -f $Version)
Write-Host "-------------------------------------------------------------"
Write-Host ("Number of Microsoft 365 Groups scanned:                          {0}" -f $GroupsCount)
Write-Host ("Potentially obsolete groups (based on document library activity: {0}" -f $ObsoleteSPOGroups)
Write-Host ("Potentially obsolete groups (based on conversation activity):    {0}" -f $ObsoleteEmailGroups)
Write-Host ("Number of Teams-enabled groups:                                  {0}" -f $TeamsCount)
Write-Host ("Number of archived teams:                                        {0}" -f $ArchivedTeams)
Write-Host ("Percentage of Teams-enabled groups:                              {0}" -f $PercentTeams)
Write-Host " "
Write-Host "Total Elapsed time: " $OverAllElapsed "seconds"

If ($ExcelGenerated -eq $true) {
    Write-Host ("Teams and Groups activity report available in HTML {0} and Excel workbook {1}" -f $HTMLReportFile, $ExcelOutputFile)
} Else {
    Write-Host ("Teams and Groups activity report available in HTML {0} and CSV file {1}" -f $HTMLReportFile, $CSVOutputFile)
}

# Reverse report data obfuscation if it was set
If ($DataObfuscationOn -eq $True) {
    $Parameters = @{ displayConcealedNames = $True }
    Update-MgBetaAdminReportSetting -BodyParameter $Parameters
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
