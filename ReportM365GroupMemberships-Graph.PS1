# ReportM365GroupMemberships-Graph.PS1
# A script to report the membership of all Microsoft 365 Groups in a tenant using Graph calls instead of PowerShell
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportM365GroupMemberships-Graph.PS1

# Needs the following Graph permissions:
# Group.Read.All, GroupMember.Read.All, User.Read.All
#

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
# End Functions


Clear-Host
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
# The need for the Exchange Online module only exists to get the organization name. If you want to hard code this variable, you can eliminate the need to load the EXO module
$OrgName = (Get-OrganizationConfig).Name
$S1 = Get-Date       
$Version = "1.0"
$ReportFile = "c:\temp\M365MembersReport.html"
$CSVFileSummary = "c:\temp\M365MembersSummaryReport.csv"
$CSVFileMembers = "c:\temp\M365MembersReport.csv"

$MemberList = [System.Collections.Generic.List[Object]]::new()
$SummaryData = [System.Collections.Generic.List[Object]]::new()
$CreationDate = Get-Date -format g

# Define the values applicable for the application used to connect to the Graph - these variables differ from tenant to tenant and from app to app
$AppId = "xxe1143-88e3-492b-bf82-24c4a47ada63"
$TenantId = "xxxxxd"
#$AppSecret = '7RplGSLWoSs~y4uHYy2041-jbm.4~_r.~q'
$AppSecret = "xxxxxx-a9_~dtG_k81UJEK-3Xv.J2X9F"

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

Clear-Host
Write-Host "Fetching user information from Azure Active Directory..."
$Uri = "https://graph.microsoft.com/v1.0/users?&`$select=displayName,usertype,assignedlicenses,id,mail,userprincipalname"
$Users = Get-GraphData -AccessToken $Token -Uri $uri
# Now get rid of all the accounts created for room and resource mailboxes, service accounts etc.
$Users = $Users | Where-Object {($_.UserType -eq "Member" -and $_.AssignedLicenses -ne $Null) -or ($_.UserType -eq "Guest")}

# Get a list of Teams and put them into a hash table so that we can mark the groups we process as being team-enabled
$uri = "https://graph.microsoft.com/V1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"
$Teams = Get-GraphData -AccessToken $Token -Uri $uri
$TeamsHash = @{}
$Teams.ForEach( {
   $TeamsHash.Add($_.Id, $_.DisplayName) } )

# Set up progress bar
$ProgDelta = 100/($Users.Count); $CheckCount = 0; $UserNumber = 0
ForEach ($User in $Users) {
  $UserNumber++
  $UserStatus = $User.DisplayName + " ["+ $UserNumber +"/" + $Users.Count + "]"
  Write-Progress -Activity "Checking groups for user" -Status $UserStatus -PercentComplete $CheckCount
  $CheckCount += $ProgDelta
  $UserType = "Tenant user"
  # Get Microsoft 365 groups this user is a member of and sort by displayname
  $Uri = "https://graph.microsoft.com/v1.0/users/" + $user.id +"/transitiveMemberOf"
  $Groups = Get-GraphData -AccessToken $Token -Uri $Uri
  $Groups = $Groups | Where-Object {$_.Grouptypes -eq "unified"} | Sort-Object DisplayName
  If ($Groups) { # We found some groups for this recipient - process them
   $AllGroups = $Groups.DisplayName -Join ", "
   $g = 0
   ForEach ($Group in $Groups)  { 
     $g++
     $Uri = "https://graph.microsoft.com/v1.0/groups/" + $Group.Id + "/owners"
     $GroupData = Get-GraphData -Uri $Uri -AccessToken $Token
     If ($GroupData.count -eq 0) {
       $Owners = "No owners found" }
     Else  { # Extract owner names
        $Owners = $GroupData.DisplayName -join ", " }
     If ($TeamsHash[$Group.Id]) { $GroupName = $Group.DisplayName + " (** Team **)" } Else { $GroupName = $Group.DisplayName }
     $MemberLine = [PSCustomObject][Ordered]@{  # Write out details of the group
       "User"              = $User.DisplayName
       UPN                 = $User.UserPrincipalName
       "User type"         = $User.UserType
       "Group Name"        = $GroupName
       "Group Description" = $Group.Description
       "Group Email"       = $Group.Mail
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

$SummaryData = $SummaryData | Sort-Object "Groups Count" -Descending
$GCount = $MemberList | Sort-Object "Group Email" -unique
$UsersNoGroups = ($SummaryData | Where-Object {$_."Groups Count" -eq 0}).Count
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
             "<p>Microsoft 365 Group Membership Report (Graph)<b>" + $Version + "</b>"	

$htmlreport = $htmlhead + $htmlbody1 + "<p><p>" + $htmlbody2 + $htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8

$MemberList | Export-CSV -NoTypeInformation $CSVFileMembers
$SummaryData | Export-CSV -NoTypeInformation $CSVFileSummary
Clear-Host
Write-Host "Microsoft 365 Group Membership Report (Graph Version) - Job Complete"
Write-Host "--------------------------------------------------------------------"
Write-Host " "
Write-Host "Outputs:"
Write-Host "--------"
Write-Host "HTML report available in" $ReportFile
Write-Host " "
Write-Host "Contains all the data generated by the script."
Write-Host " "
Write-Host "CSV file for members in groups available in" $CSVFileMembers
Write-Host " "
Write-Host "Lists details of group membership for individual user accounts."
Write-Host " "
Write-Host "CSV summary report available in"  $CSVFileSummary
Write-Host " "
Write-Host "Summarizes the groups that users belong to."
Write-Host " "
Write-Host ("Total processing time {0} seconds ({1} seconds per user) for {2} user accounts and {3} Microsoft 365 Groups" -f $TotalSeconds, $SecondsPerUser, $Users.Count, $Gcount.count)


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
