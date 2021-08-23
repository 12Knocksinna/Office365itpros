# ReportDelveInsightsDisabled.PS1
# Report Microsoft 365 accounts that are disabled for the Delve Document Insights feature
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportDelveInsightsDisabled.PS1
# The app registered in Azure AD needs the User.Read.All application permission.

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

# Define all the stuff necessary to use a registered app to interact with the Graph APIs. Amend these values for your tenant and app!
$AppId = "76c31534-ca1f-4d46-959a-6159fcb2f77a"
$TenantId = "aa62313f-14fc-43a2-9a7a-d2e27f4f3478"
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

CLS;$Report = [System.Collections.Generic.List[Object]]::new();$CSVOutput = "C:\temp\DelveDisabledAccounts.CSV"

Write-Host "Finding mailboxes to check"
[array]$Mbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
ForEach ($M in $Mbx) {
   $uri = "https://graph.microsoft.com/v1.0/users/" + $M.ExternalDirectoryObjectId + "/settings"
   $Settings = Get-GraphData -AccessToken $Token -Uri $uri
   If ($Settings.contributionToContentDiscoveryDisabled -eq $True) { # Delve document insights feature turned off so report this mailbox
      Write-Host "Delve turned off for" $M.DisplayName
      $ReportLine = [PSCustomObject][Ordered]@{ 
         Name     = $M.DisplayName
         UPN      = $M.UserPrincipalName
         ObjectId = $M.ExternalDirectoryObjectId
         DelveOff = $Settings.contributionToContentDiscoveryDisabled }
     $Report.Add($ReportLine) 
     } #End if
} #End ForEach

If ($Report) {
   Write-Host ("All done. {0} accounts discovered with Delve document insights disabled. Details in {1}." -f $Report.Count, $CSVOutput) 
   $Report | Out-GridView
   $Report | Export-CSV -NoTypeInformation $CSVOutput
   }
Else {
   Write-Host "No accounts found with Delve document insights disabled" }

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
