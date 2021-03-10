# ReportM365GroupMemberships.PS1
# A script to report the membership of all Microsoft 365 Groups in a tenant
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportM365GroupMemberships.PS1

CLS
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
$S1 = Get-Date       
$OrgName = (Get-OrganizationConfig).Name
$Version = "1.0"
$ReportFile = "c:\temp\M365MembersReport.html"
$CSVFileSummary = "c:\temp\M365MembersSummaryReport.csv"
$CSVFileMembers = "c:\temp\M365MembersReport.csv"

$MemberList = [System.Collections.Generic.List[Object]]::new()
$SummaryData = [System.Collections.Generic.List[Object]]::new()
$UserNoGroups = 0; $UserWithGroups = 0
$CreationDate = Get-Date -format g
CLS
Write-Host "Fetching user information from Azure Active Directory..."
$Users = Get-AzureADUser -All:$true
# Now get rid of all the accounts created for room and resource mailboxes, service accounts etc.
$Users = $Users |?{($_.UserType -eq "Member" -and $_.AssignedLicenses -ne $Null) -or ($_.UserType -eq "Guest")}
# Get a list of Teams and put them into a hash table so that we can mark the groups we process as being team-enabled
$Teams = Get-UnifiedGroup -Filter {ResourceProvisioningOptions -eq "Team"} -ResultSize Unlimited | Select ExternalDirectoryObjectId, DisplayName
$TeamsHash = @{}
$Teams.ForEach( {
   $TeamsHash.Add($_.ExternalDirectoryObjectID, $_.DisplayName) } )
CLS
# Set up progress bar
$ProgDelta = 100/($Users.Count); $CheckCount = 0; $UserNumber = 0
ForEach ($User in $Users) {
  $UserNumber++
  $UserStatus = $User.DisplayName + " ["+ $UserNumber +"/" + $Users.Count + "]"
  Write-Progress -Activity "Checking groups for user" -Status $UserStatus -PercentComplete $CheckCount
  $CheckCount += $ProgDelta
  $UserType = "Tenant user"
  $DN = (Get-Recipient -Identity $User.UserPrincipalName).DistinguishedName
  $Groups = (Get-Recipient -ResultSize Unlimited -RecipientTypeDetails GroupMailbox -Filter "Members -eq '$DN'" | Select DisplayName, Notes, ExternalDirectoryObjectId, ManagedBy, PrimarySmtpAddress)
  If ($Groups) { # We found some groups for this recipient - process them
   $Groups = $Groups | Sort DisplayName
   $AllGroups = $Groups.DisplayName -Join ", "
   $g = 0
   ForEach ($Group in $Groups)  { 
     $g++
     $Owners = $Group.Managedby -join ", "
     If ($TeamsHash[$Group.ExternalDirectoryObjectId]) { $GroupName = $Group.DisplayName + " (** Team **)" } Else { $GroupName = $Group.DisplayName }
     $MemberLine = [PSCustomObject][Ordered]@{  # Write out details of the group
       "User"              = $User.DisplayName
       UPN                 = $User.UserPrincipalName
       "User type"         = $User.UserType
       "Group Name"        = $GroupName
       "Group Description" = $Group.Notes
       "Group Email"       = $Group.PrimarySmtpAddress
       "Group Owners"      = $Owners         }
      $MemberList.Add($MemberLine) 
    } #End For
    $SummaryLine = [PSCustomObject][Ordered]@{  # Write out summary record for the user
       "User"              = $User.DisplayName
       UPN                 = $User.UserPrincipalName
       "User type"         = $User.UserType
       "Groups count"      = $g
       "Member Of"         = $AllGroups     }
    $SummaryData.Add($SummaryLine) 
  } # End if
  Else { #No groups found for this user, so just write a summary record
   $SummaryLine = [PSCustomObject][Ordered]@{  
       "User"              = $User.DisplayName
       UPN                 = $User.UserPrincipalName
       "User type"         = $UserType
       "Groups count"      = 0
       "Member Of"           = "No groups found for user"     }
    $SummaryData.Add($SummaryLine) 
  } #End Else
} #End For 
$SummaryData = $SummaryData | Sort "Groups Count" -Descending
$GCount = $MemberList | Sort "Group Email" -unique
$UsersNoGroups = ($SummaryData | ? {$_."Groups Count" -eq 0}).Count
$UsersWithGroups = ($SummaryData.Count - $UsersNoGroups)
$S2 = Get-Date
$TotalSeconds = [math]::round(($S2-$S1).TotalSeconds,2)
$SecondsPerUser = [math]::round(($TotalSeconds/$Users.count),2)

# Create the HTML report
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
           <p><h1>Microsoft 365 Groups and Teams Membership Listing</h1></p>
           <p><h2><b>All groups in the " + $Orgname + " organization</b></h2></p>
           <p><h3>Generated: " + (Get-Date -format g) + "</h3></p></div>"

$htmlbody1 = $MemberList | ConvertTo-Html -Fragment
$htmlbody1 = $htmlbody1 + '<div class="page-break"></div>'
$htmlbody2 = $SummaryData | ConvertTo-Html -Fragment

$htmltail = "<p>Report created for: " + $OrgName + "</p>" +
             "<p>Created: " + $CreationDate + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of users in groups:      " + $UsersWithGroups + "</p>" +
             "<p>Number of users not in groups:  " + $UsersNoGroups + "<p>"+
             "<p>Number of Microsoft 365 Groups: " + $GCount.Count + "</p>" +
              "<p>Number of Microsoft Teams:     " + $Teams.Count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Microsoft 365 Group Membership Report <b>" + $Version + "</b>"	

$htmlreport = $htmlhead + $htmlbody1 + "<p><p>" + $htmlbody2 + $htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8

$MemberList | Export-CSV -NoTypeInformation $CSVFileMembers
$SummaryData | Export-CSV -NoTypeInformation $CSVFileSummary
CLS
Write-Host "Microsoft 365 Group Membership Report - Job Complete"
Write-Host "----------------------------------------------------"
Write-Host " "
Write-Host "Outputs:"
Write-Host "--------"
Write-Host "HTML report available in" $ReportFile
Write-Host " "
Write-Host "Contains all the data generated by the script."
Write-Host " "
Write-Host "CSV file for members in groups available in" $CSVCileMembers
Write-Host " "
Write-Host "Lists details of group membership for individual user accounts."
Write-Host " "
Write-Host "CSV summary report available in"  $CSVFileSummary
Write-Host " "
Write-Host "Summarizes the groups that users belong to."
Write-Host " "
Write-Host ("Total processing time {0} seconds ({1} seconds per user) for {2} user accounts and {3} Microsoft 365 Groups" -f $TotalSeconds, $SecondsPerUser, $Users.Count, $Gcount.count)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
