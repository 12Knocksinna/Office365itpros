# ReportDLMembershipsCountsGraph.PS1
# Report the membership and counts for distribution lists in Exchange Online - the Graph version. The pure PowerShell version is in
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportDLMembershipsCountsGraph.PS1
# 
# The registered app needs Directory.Read.All, Group.Read.All, User.Read.All application permissions.

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

$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
# OK, we seem to be fully connected to Exchange Online. 

# Define all the stuff necessary to use a registered app to interact with the Graph APIs. You need to update these values to work with your tenant.
$AppId = "87c31534-ca1f-4d46-959a-6159fcb2f77a"
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3478"
$AppSecret = "7xP4Nj~kiU.yBXY9~yQB3sMrvpLv5Rx_._"

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

Write-Host "Finding Exchange Online Distribution Lists..."
# Find distribution lists - excluding room lists
[array]$DLs = Get-DistributionGroup -RecipientTypeDetails "MailUniversalDistributionGroup", "MailUniversalSecurityGroup" -ResultSize Unlimited
# This is the Graph method -includes security groups, but it's faster - use it if you like
# $Uri = "https://graph.microsoft.com/V1.0/groups?`$filter=Mailenabled eq true AND NOT groupTypes/any(c:c+eq+'Unified')&`$count=true"
# [array]$DLs = Get-GraphData -AccessToken $Token -Uri $uri

If (!($DLs)) { Write-Host "No distribution lists found... sorry! "; break }
 Else { Write-Host ("{0} distribution lists found" -f $DLs.count) }
CLS; $Report = [System.Collections.Generic.List[Object]]::new()

# Loop down through each DL and fetch the membership using the Graph transitivemembers call to return a complete set.
$DLNumber = 0
ForEach ($DL in $DLs) {
$DLNumber++
  $ProgressBar = "Processing distribution list " + $DL.DisplayName + " (" + $DLNumber + " of " + $DLs.Count + ")" 
  Write-Progress -Activity "Analzying membership of distribution list " -Status $ProgressBar -PercentComplete ($DLNumber/$DLs.Count*100)
# Retrieve transitive membership for the distribution list
  $Uri = "https://graph.microsoft.com/v1.0/groups/" + $DL.ExternalDirectoryObjectId + "/transitiveMembers"
  [array]$Members = Get-GraphData -AccessToken $Token -Uri $uri
  $CountContacts = 0; $CountTenantMembers = 0; $CountGuests = 0; $CountGroups = 0
  $MembersNames = [System.Collections.Generic.List[Object]]::new()
  $CountOfMembers = $Members.Count
  # Loop through each member and figure out what type of member they are and their display name
  ForEach ($Member in $Members) {
   Switch ($Member."@odata.type") {
    "#microsoft.graph.orgContact" { # Mail contact
      $MemberDisplayName = $Member.DisplayName
      $CountContacts++ }
    "#microsoft.graph.user" { # Tenant user (including guests
      $MemberDisplayName = $Member.DisplayName
      If ($Member.UserPrincipalName -Like "*#EXT#*") { $CountGuests++ }
       Else { $CountTenantMembers++ }
    }
    "#microsoft.graph.group" { #Another group
      $MemberDisplayName = $Member.DisplayName
      $CountGroups++ }    
   } #End Switch
    # Update member table
    $MemberData = [PSCustomObject][Ordered]@{  
       MemberName = $MemberDisplayName
       MemberId   = $Member.Id }
    $MembersNames.Add($MemberData)   
} #End Foreach

   # Remove any duplicates and sort the member list alphabetically
   $MembersNames = $MembersNames | Sort MemberId -Unique | Sort MemberName
   $OutputNames =  $MembersNames.MemberName -join ", " 
   
   #  Figure out DL manager names - can't be done using the Graph at present
   $ManagerList = [System.Collections.Generic.List[Object]]::new()
    ForEach ($Manager in $DL.ManagedBy) {
       $Recipient = Get-Recipient -Identity $Manager -ErrorAction SilentlyContinue
       If (!($Recipient)) { # Can't resolve manager
           $Recipient = "Unknown user" }
       $ManagerLine = [PSCustomObject][Ordered]@{  
         DisplayName = $Recipient.DisplayName
         UPN         = $Recipient.WIndowsLiveID }
       $ManagerList.Add($ManagerLine) 
    } # End processing managers
    $Managers = $ManagerList.DisplayName -join ", "

   $ReportLine = [PSCustomObject][Ordered]@{  
       DLName             = $DL.DisplayName
       ManagedBy          = $Managers
       "Members"          = $CountOfMembers
       "Tenant Users"     = $CountTenantMembers
       "Groups"           = $CountGroups
       "Guest members"    = $CountGuests
       "Mail contacts"    = $CountContacts
       "Member names"     = $OutputNames  }

   $Report.Add($ReportLine) 

} # End processing DLs

# Create output files
$DLCSVOutput = "c:\temp\DLData.CSV"
$ReportFile = "c:\temp\DLData.html"
# Create the HTML report
$OrgName = (Get-OrganizationConfig).Name
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
           <p><h1>Distribution List Manager Report</h1></p>
           <p><h2><b>For the " + $Orgname + " organization</b></h2></p>
           <p><h3>Generated: " + (Get-Date -format g) + "</h3></p></div>"

$htmlbody1 = $Report | ConvertTo-Html -Fragment

$htmltail = "<p>Report created for: " + $OrgName + "</p>" +
             "<p>Created: " + $CreationDate + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of distribution lists found:    " + $DLs.Count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Distribution List Manager Report<b> " + $Version + "</b>"	

$htmlreport = $htmlhead + $htmlbody1 + $htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8

Write-Host ("All done. {0} distribution lists analyzed. CSV file is available at {1} and a HTML report at {2}" -f $DLs.Count, $DLCSVOutput, $ReportFile)
$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation $DLCSVOutput

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
