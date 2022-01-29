# GetServicePrincipalSignIns-Graph.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/GetServicePrincipalSignIns-Graph.PS1
# Example of how to extract and analyze service principal sign-in data from Azure AD using the Graph API
# Requires the Directory.Read.All and AuditLog.Read.All application permissions

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

# Define the values applicable for the application used to connect to the Graph - these variables vary from tenant to tenant and app to app
$AppId = ""
$TenantId = ""
$AppSecret = ''

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

# Define variables
CLS;$Report = [System.Collections.Generic.List[Object]]::new();$CSVOutput = "C:\temp\SPSignInData.CSV"
# Define start and end date for query. Add Z to each sortable date to make the Graph query happy
$StartDate = (Get-Date).AddDays(-7); $EndDate = (Get-Date -format s) + "Z"
$StartDate = (Get-Date $StartDate -format s) + "Z"
# Build Uri for the Graph query
$Uri = "https://graph.microsoft.com/beta/auditLogs/signIns?&`$filter=createdDateTime ge " + $StartDate + " and createdDateTime le " + $EndDate + " and signInEventTypes/any(z:z eq 'servicePrincipal')"
# Execute the query
Write-Host "Querying Azure AD for service principal sign-in records from" $StartDate "to" $EndDate
[Array]$SpSignInData = Get-GraphData -Uri $Uri -AccessToken $Token
If (!($SpSignInData)) { Write-Host "No service principal sign in data found - exiting" ; break }
Write-Host "Processing" $SpSignInData.Count "sign-in records for service principals"
# Process the information which came back
ForEach ($Sp in $SpSignInData) { # Process the records
    $StatusCode = "Success"; $StatusReason = $Null
   If ($Sp.Status.ErrorCode -ne 0) { 
       $StatusCode = $Sp.Status.ErrorCode 
       $StatusReason = $Sp.Status.FailureReason }
      $ReportLine  = [PSCustomObject][Ordered]@{ 
         Date           = Get-Date($Sp.createdDateTime) -format g
         SPName         = $Sp.servicePrincipalName
         App            = $Sp.AppDisplayName
         AppId          = $Sp.AppId
         Location       = $Sp.Location.City
         State          = $Sp.Location.State
         ipAddress      = $Sp.IpAddress
         SpId           = $Sp.ServicePrincipalId
         Resource       = $Sp.ResourceDisplayName
         Status         = $StatusCode
         Reason         = $StatusReason
         }
      $Report.Add($ReportLine) 
} #End ForEach

# Report what we've found
Write-Host " "
Write-Host "Summary of Service Principal sign-in activity"
Write-Host "From" $StartDate "to" $EndDate
Write-Host "Output CSV file: " $CSVOutput
Write-Host ""

$Report | Group SpName | Sort Count -Descending | Select Name, Count
$Report | Export-CSV -NoTypeInformation $CSVOutput

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
