# GetTenantFeatureUpdatesGraph.PS1
# Example of using the Office 365 Service Communications Graph API to retrieve information about feature updates coming to a tenant
# The registered app needs the ServiceMessage.Read.All permission to read service feature updates 
# https://github.com/12Knocksinna/Office365itpros/blob/master/GetTenantFeatureUpdatesGraph.PS1

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

Function ConvertFrom-Html
{
    [CmdletBinding(SupportsShouldProcess = $True)]
    Param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Html
    )

    $HtmlObject = New-Object -Com "HTMLFile"
    $HtmlObject.IHTMLDocument2_write($Html)
    return $HtmlObject.documentElement.innerText
}

$CSVOutputFile = "C:\temp\MessagesRequiringAction.csv"
$HTMLOutputFile = "C:\temp\MessagesRequiringAction.Html"
$Now = Get-Date

# Change these values to match your tenant and app details
$AppId = "b6bd07f5-63be-4b96-b569-470d93401d50"
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f3478"
$AppSecret = '12EJ.O2~1.HUFJXRcJ-8o4S2e~q_16-YJw'

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
            'Authorization' = "Bearer $Token" }

# Get feature updates with action required within the next 180 days - the filter clause sets the number of days to look back
$DaysRange = (Get-Date).AddDays(180)
$DaysRangeZ = Get-Date($DaysRange) -format s

# Fetch messages and filter to get just those which have an action required date set
$Uri = "https://graph.microsoft.com/beta/admin/serviceAnnouncement/messages?`$filter=actionRequiredByDateTime le $DaysRangeZ" + "Z"
[array]$Messages = Get-GraphData -AccessToken $Token -Uri $uri
If (!($Messages)) {Write-Host "No update messages found with action required in the next 180 days - exiting"; break}
# Make sure that the messages are sorted
$Messages = $Messages | Sort {$_.actionRequiredByDateTime -as [datetime]} -Descending
$MessageData = [System.Collections.Generic.List[Object]]::new()

# Go through the messages and extract information of interest
ForEach ($Message in $Messages) {
 $Status = "Action due"
 $Tags = $Message.Tags -join ", "
 $Services = $Message.Services -join ", "
 If ($Message.actionRequiredByDateTime) {
    $TimeToGo = New-TimeSpan ($Message.actionRequiredByDateTime)
    $FormattedTime = "{0:dd}d:{0:hh}h:{0:mm}m" -f $TimeToGo
    [datetime]$MessageDate = $Message.actionRequiredByDateTime
    If ($Now -ge $MessageDate) {
          $Status = "Action overdue"
          $FormattedTime = $FormattedTime + " (o/d)" }
    }
    Else { $FormattedTime = "N/A" }
 
 $RoadmapId = $Null; $BlogLink = $Null; $WebLink = $Null
 If ($Message.Details -ne $Null) {
    $RoadmapId = $Message.Details |?{$_.Name -eq "RoadmapIds"} | Select -ExpandProperty Value
    $BlogLink  = $Message.Details |?{$_.Name -eq "BlogLink"} | Select -ExpandProperty Value
    $WebLink   = $Message.Details |?{$_.Name -eq "ExternalLink"} | Select -ExpandProperty Value }

 $Description = ConvertFrom-html -html $Message.Body.Content

 $ReportLine = [PSCustomObject][Ordered]@{  
     Title       = $Message.Title
     Id          = $Message.Id
     Services    = $Services
     Category    = $Message.Category
     Severity    = $Message.Severity
     ActionBy    = Get-Date($Message.actionRequiredByDateTime) -format g
     TimeToGo    = $FormattedTime
     Status      = $Status
     StartDate   = Get-Date($Message.startDateTime) -format g
     EndDate     = Get-Date($Message.endDateTime) -format g
     LastUpdate  = Get-Date($Message.lastModifiedDateTime) -format g  
     Description = $Description
     RoadmapId   = $RoadmapId
     BlogLink    = $BlogLink
     WebLink    = $WebLink
     Tags        = $Tags }
  $MessageData.Add($ReportLine)
}   # End ForEach
  
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
# OK, we seem to be fully connected to Exchange Online, so we can fetch the organization name (to make the report prettier)     
$OrgName = (Get-OrganizationConfig).Name

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
           <p><h1>Microsoft 365 Features Action Required Report</h1></p>
           <p><h2><b>for the " + $Orgname + " organization</b></h2></p>
           <p><h3>Generated: " + (Get-Date -format g) + "</h3></p></div>"

$htmlbody1 = $MessageData | ConvertTo-Html -Fragment
$htmltail = "<p>Report created for: " + $OrgName + "</p>" +
            "<p>Created: " + $Now + "<p>"	

$htmlreport = $htmlhead + $htmlbody1 + $htmltail
$htmlreport | Out-File $HTMLOutputFile -Encoding UTF8

$MessageData | Export-CSV -NoTypeInformation $CSVOutputFile
CLS
# And report out
Write-Host "All done. Output files are" $CSVOutputFile "and" $HTMLOutputFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
