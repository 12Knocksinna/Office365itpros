# ReportDLMembershipsCountsGraph.PS1
# Report the membership and counts for distribution lists in Exchange Online - the Graph version. The pure PowerShell version is in
# 
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

# Define all the stuff necessary to use a registered app to interact with the Graph APIs
$AppId = "77c31534-ca1f-4d46-959a-6159fcb2f77a"
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f3478"
$AppSecret = "7FP4Nj~kiU.yBXY9~yQB3sMrvpLv5Rx_._"

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

# Find all distribution lists
Write-Host "Finding Exchange Online Distribution Lists..."
# Find distribution lists, excluding room lists
$DLs = Get-DistributionGroup -ResultSize Unlimited -Filter {RecipientTypeDetails -ne "Roomlist"} | Select DisplayName, ExternalDirectoryObjectId, ManagedBy
If (!($DLs)) { Write-Host "No distribution lists found... sorry! "; break }
 Else { Write-Host ("{0} distribution lists found" -f $DLs.count) }
CLS; $DLNumber = 0
$Report = [System.Collections.Generic.List[Object]]::new()
$DLCSVOutput = "c:\temp\DLMemberCounts.CSV"

# Loop down through each DL and fetch the membership using the Graph transitivemembers call to return a complete set.
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

   # Remove any duplicates
   $MembersNames = $MembersNames | Sort MemberId -Unique
   $OutputNames =  $MembersNames.MemberName -join ", " 

   $ReportLine = [PSCustomObject][Ordered]@{  
       DLName             = $DL.DisplayName
       ManagedBy          = $DL.ManagedBy -join ", "
       "Members"          = $CountOfMembers
       "Tenant Users"     = $CountTenantMembers
       "Groups"           = $CountGroups
       "Guest members"    = $CountGuests
       "Mail contacts"    = $CountContacts
       "Member names"     = $OutputNames  }

   $Report.Add($ReportLine) 

} # End processing DLs

Write-Host ("All done. {0} distribution lists analyzed. Output is in {1}" -f $DLs.Count, $DLCSVOutput)
$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation $DLCSVOutput

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
