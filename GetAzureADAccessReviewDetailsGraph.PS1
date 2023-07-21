# GetAzureADAccessReviewDetailsGraph.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/GetAzureADAccessReviewDetailsGraph.PS1
# A script to show how to navigate the Graph interface for Azure AD Access Reviews, especially a review set up for all Groups and Teams
# V1.0 15-Jan-2021
# App must be assigned the AccessReview.ReadWrite.All, AccessReview.ReadWrite.Members, Group.Read.All, and User.Read.All (application) permissions to work
#
function Get-GraphData {
# GET data from Microsoft Graph.
# Based on https://danielchronlund.com/2018/11/19/fetch-data-from-microsoft-graph-with-powershell-paging-support/
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
            'ConsistencyLevel' = "eventual"   }

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
# Define the values applicable for the application used to connect to the Graph (change these details for your tenant and registered app)
$AppId = "0ade5c24-d775-4017-a824-0a993c60787d"
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3428"
$AppSecret = '7RplGSLWoSs~y4uHYy2041-jbm.4~_s.~q'

$OutputCSV = "c:\temp\AzureADGuestReviews.CSV"

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

Write-Host "Fetching Azure AD Access Review Data..."
# Get Access Reviews currently running 
$Uri = "https://graph.microsoft.com/beta/identityGovernance/accessReviews/definitions"
$AccessData = Get-GraphData -AccessToken $Token -Uri $uri 
# Find the access review for Teams and Groups    
# This check depends on the display name you assign to the access review in the Azure AD admin center
$Id = $Accessdata | Where-Object {$_.displayname -eq "Review guest access across Microsoft 365 Groups"} | Select-Object -ExpandProperty Id
If ($Id -eq $Null) { Write-Host "Can't find access review - please check the display name"; break }
# Find the instances (groups being reviewed)
$Uri = "https://graph.microsoft.com/beta/identityGovernance/accessReviews/definitions/" + $Id +"/instances"
$AccessData = Get-GraphData -AccessToken $Token -Uri $uri   
$CountOfGroups = $AccessData.Count

$Report = [System.Collections.Generic.List[Object]]::new()
$ApproveCount = 0; $DenyCount = 0; $NoDecision = 0
ForEach ($Instance in $AccessData) {
    $InstanceId = $Instance.Id
    # Get group id
    $Start = $Instance.scope.query.IndexOf("s/")
    $End = $Instance.scope.query.IndexOf("/members")
    $GroupId = $Instance.scope.query.substring($Start + 2,$End - 13)
    $Uri = "https://graph.microsoft.com/v1.0/groups/" + $GroupId
    $GroupDetails = Get-GraphData -AccessToken $Token -Uri $uri 
    $GroupName = $GroupDetails.DisplayName
    # Now get the instances (people being reviewed) and what's happened to each
    $GroupUnderReview = 0 # Flag to track if a group has started review. Set if a Deny or Approve decision is made
    $Uri = "https://graph.microsoft.com/beta/identityGovernance/accessReviews/definitions/" + $Id +"/instances/" + $instanceId + "/decisions"
    $InstanceData = Get-GraphData -AccessToken $Token -Uri $uri 
    Write-Host "Number of Guests to review in" $GroupName ":" $InstanceData.id.Count
    If ($InstanceData.id.Count -gt 0) {
     ForEach ($Decision in $InstanceData) {
       Switch ($Decision.decision) {  # Only generate report if guests are present to review
         "Deny"  {
             $DenyCount++; $GroupUnderReview = 1
             $Verdict = $Decision.decision
             $User = $decision.target.userprincipalname
             $Name = $decision.target.userdisplayname
             $justification = $decision.justification
             $Recommendation = $Decision.recommendation
             $reviewer = $decision.reviewedby.displayname
             $ReviewDate = (get-date $decision.revieweddatetime -format g) }
         "Approve" {
             $ApproveCount++
             $Verdict = $Decision.decision; $GroupUnderReview = 1
             $User = $decision.target.userprincipalname
             $Name = $decision.target.userdisplayname
             $justification = $decision.justification
             $Recommendation = $Decision.recommendation
             $reviewer = $decision.reviewedby.displayname
             $ReviewDate = (get-date $decision.revieweddatetime -format g) }
         "NotReviewed" {
             $NoDecision++
             $Verdict = $Decision.decision
             $User = $decision.target.userprincipalname
             $Name = $decision.target.userdisplayname
             $justification = $decision.justification
             $Recommendation = $Decision.recommendation
             $reviewer = $decision.reviewedby.displayname
             $ReviewDate = "No decision made" }
         } #End Switch
     # Report decision
       $ReportLine = [PSCustomObject] @{
           User           = $User
           Name           = $Name
           Verdict        = $Verdict
           Recommendation = $Recommendation.Trim()
           Reviewer       = $Reviewer
           ReviewDate     = $ReviewDate
           Justification  = $Justification.Trim()
           Group          = $GroupName }             
      $Report.Add($ReportLine) 
     } # End report if any guests are found to review
    } #End ForEach InstanceData
    If ($GroupUnderReview -eq 1) { $GroupsWithReview++ }
} #End ForEach AccessData

# Quick way of reporting counts for the various verdicts is to group the report data, but we want some nice figures 
# $Report |  Group-Object Verdict | format-Table Name, Count
$CountOfApprovals = ($Report | Where-Object {$_.Verdict -eq "Approve"} | Measure)
$CountOfDeny = ($Report | Where-Object {$_.Verdict -eq "Deny"} | Measure ) 
$CountOfNoDecision = ($Report | Where-Object {$_.Verdict -eq "NotReviewed"} | Measure)    
Write-Host "Number of Groups with reviews for Guest Members: " $CountOfGroups

CLS
Write-Host ""
Write-Host "Decision Profile"
Write-Host "----------------"
Write-Host ""
Write-Host "Total Groups with Guests:                 " $CountOfGroups
Write-Host "Groups started reviews:                   " $GroupsWithReview
Write-Host "Groups not started reviews                " ($CountOfGroups - $GroupsWithReview)
Write-Host "Total decisions to be made:               " $Report.Count
Write-Host ("Review decisions to approve guest access:  {0} ({1})" -f $CountOfApprovals.Count, ($CountOfApprovals.Count/$Report.Count).ToString("P") )
Write-Host ("Review decisions to deny guest access:     {0} ({1})" -f $CountOfDeny.Count, ($CountOfDeny.Count/$Report.Count).ToString("P") )
Write-Host ("No decisions made so far:                  {0} ({1})" -f $CountOfNoDecision.Count, ($CountOfNoDecision.Count/$Report.Count).ToString("P") )
Write-Host " "
Write-Host "A CSV file for current Access Review decision status is available in" $OutputCSV

# Output files
$Report | Sort User | Export-CSV -NoTypeInformation $OutputCSV
$Report | Sort User | Out-GridView


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
