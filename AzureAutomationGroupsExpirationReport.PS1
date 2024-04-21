# https://github.com/12Knocksinna/Office365itpros/blob/master/AzureAutomationGroupsExpirationReport.PS1
# See https://office365itpros.com/2022/04/08/post-teams-channels/ for more information
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
# Set up our connections
$Connection = Get-AutomationConnection -Name AzureRunAsConnection
$Certificate = Get-AutomationCertificate -Name AzureRunAsCertificate
$GraphConnection = Get-MsalToken -ClientCertificate $Certificate -ClientId $Connection.ApplicationID -TenantId  $Connection.TenantID 
$Token = $GraphConnection.AccessToken
$Headers = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $Token" 
            'ConsistencyLevel' = "eventual" }

$AzConnection = Connect-AzAccount -Tenant $Connection.TenantId -ApplicationId `
    $Connection.ApplicationId -CertificateThumbPrint $Connection.CertificateThumbprint `
    -ServicePrincipal
Write-Host "AZ Connection is" $AzConnection
# Get username and password from Key Vault - You need to set up your own Key Vault and populate 
# it with secrets to make this work...
$UserName = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -Name "ExoAccountName" -AsPlainText
$UserPassword = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -name "ExoAccountPassword" -AsPlainText
# Create credentials object from the username and password
[securestring]$SecurePassword = ConvertTo-SecureString $UserPassword -AsPlainText -Force
[pscredential]$UserCredentials = New-Object System.Management.Automation.PSCredential ($UserName, $SecurePassword)
# Get Site URL to use with PnP connection
$SiteURL = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -name "SPOSiteURL" -AsPlainText
# Target channel identifier for incoming webhook connector
$TargetChannel = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -name "IncomingWebhookId" -AsPlainText
# Target team and channel in that team to which we post a message containing the report
$TargetTeamId = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -name "TargetTeamId" -AsPlainText
$TargetTeamChannel = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -name "TargetChannelID" -AsPlainText
Write-Output "Credentials" $UserCredentials
Write-Output "Target Site" $SiteURL
Write-Output "Target team and channel" $TargetTeamId " " $TargetTeamChannel

# Get set of groups with an expiration date set. Can't check for null as the ExpirationDateTime property doesn't support this
$uri = "https://graph.microsoft.com/beta/groups?`$filter=ExpirationDateTime ge 2014-01-01T00:00:00Z AND groupTypes/any(a:a eq 'unified')&`$count=true"
[array]$Groups = Get-GraphData -AccessToken $Token -Uri $uri
If (!($Groups)) { 
    Write-Output "No groups found subject to the expiration policy - exiting" ; break 
}
$Report = [System.Collections.Generic.List[Object]]::new(); $Today = (Get-Date)
ForEach ($G in $Groups) {
        $Days = (New-TimeSpan -Start $G.CreatedDateTime -End $Today).Days  # Age of group
        #$LastRenewed = $G.RenewedDateTime
        #$NextRenewalDue = $G.ExpirationDateTime
        $DaysLeft = (New-TimeSpan -Start $Today -End $G.ExpirationDateTime).Days
        $GroupsInPolicy++
        $ReportLine = [PSCustomObject]@{
           Group                   = $G.DisplayName
           Created                 = Get-Date($G.CreatedDateTime) -format g
          "Age in days"            = $Days
          "Last renewed"           = Get-Date($G.RenewedDateTime) -format g
          "Next renewal"           = Get-Date($G.ExpirationDateTime) -format g
          "Days before expiration" = $DaysLeft}
          $Report.Add($ReportLine)
} # End Foreach
Write-Output "Total Groups covered by expiration policy:" $Groups.Count
Write-Output ""
# Write out details to show that the job works!
$Report | Sort-Object "Days before expiration" | Format-Table Group, "Last renewed", "Next renewal", "Days before expiration" -AutoSize
# Create data to store in SharePoint Online
# First, the CSV file
$SDate = Get-Date -format yyyyMMddHHmmss
[string]$SourceDocument = "Microsoft 365 Groups Expiration Report " + $SDate + ".csv"
[string]$HTMLDocument = "Microsoft 365 Groups Expiration Report " + $SDate + ".html"
$Report | Sort-Object "Days before expiration" | Export-CSV -NoTypeInformation $SourceDocument

# Connect to SharePoint Online using PnP 
$PnpConnection = Connect-PnPOnline $SiteURL -Credentials $UserCredentials -ReturnConnection
# Add a document title
$Values = @{"Title" = 'Microsoft 365 Groups Expiration Report (CSV)'}
# Add the file to the General folder
$FileAddStatus = (Add-PnPFile -Folder "Shared Documents/General" -Path $SourceDocument -Connection $PnpConnection -Values $Values | Out-Null)
$FileAddStatus
$NewFileUri = $SiteUrl + "/Shared Documents/General/" + $HTMLDocument
# Now the HTML file
Connect-MgGraph -ClientID $Connection.ApplicationId -TenantId $Connection.TenantId -CertificateThumbprint $Connection.CertificateThumbprint
$Organization = Get-MgOrganization
$TenantName = $Organization.DisplayName
$Title = "Microsoft 365 Groups Expiration Report"
$cssString = @'
<style type="text/css">
.tftable {table-layout:fixed;width: 40%;font-family:"Segoe UI";font-size:12px;color:#333333;border-width: 1px;border-color: #729ea5;border-collapse: collapse;}
.tftable th {width: 30%;font-size:12px;background-color:#acc8cc;border-width: 1px;padding: 8px;border-style: solid;border-color: #729ea5;text-align:left;}
.tftable tr {background-color:#d4e3e5;}
.tftable td {width: 10%font-size:12px;border-width: 1px;padding: 8px;border-style: solid;border-color: #729ea5;}
.tftable tr:hover {background-color:#ffffff;}

table.center {
  margin-left: auto; 
  margin-right: auto;
}
</style>
'@
$Body = "<html><head><title>$($Title)</title>"
$Body += '<meta http-equiv="Content-Type content="text/html; charset=ISO-8859-1 />'
$Body += $cssString
$Body += '</head><body><p><font face="Segoe UI"><h1>Microsoft 365 Groups Expiration Report</h1></font></p><p><font face="Segoe UI"><h2>Tenant: ' + ($TenantName) + '</h2></p><p><font face="Segoe UI"><h3>Generated: ' + $Today + '</h3></font></p>'
$Body += '<table class="tftable">'
$Body += "<colgroup><col/><col/><col/><col/><col/><col/></colgroup> <tr><th>Group</th><th>Created</th><th>Age in days</th><th>Last renewed</th><th>Next renewal</th><th>Days before expiration</th></tr>"
$Report = $Report | Sort-Object Group
ForEach ($R in $Report) {
   $Body += "<tr><td>$($R.Group)</td><td>$($R.Created)</td><td>$($R.'Age in days')</td><td>$($R.'Last Renewed')</td><td>$($R.'Next Renewal')</td><td>$($R.'Days before expiration')</td></tr>"
}
$Body += "</table>" 
$Body += '<p><font face="Segoe UI"><h3>End of Report<h3></font></p>'
$Body += '<p><font size="2" face="Segoe UI">'
$Body += '</body></html>'
$Body | Out-File $HTMLDocument
# And write to SharePoint Online
$Values = @{"Title" = 'New Microsoft 365 Groups Expiration Report (HTML)'}
$FileAddStatus = (Add-PnPFile -Folder "Shared Documents/General" -Path $HTMLDocument -Connection $PnpConnection -Values $Values | Out-Null)
# Post to a Teams channel using an incoming webhook connector
$GroupWebHookData = 'The new report is available in <a href="' + $NewFileUri + '">' + 'Microsoft 365 Groups Expiration Report</a>'

Write-Host $GroupWebHookData

$DateNow = Get-Date -format g
$Notification = @"
    {
        "@type": "MessageCard",
        "@context": "https://schema.org/extensions",
        "summary": "Microsoft 365 Groups",
        "themeColor": "0072C6",
        "title": "Notification: New Microsoft 365 Groups Expiration Report is available",
         "sections": [
            {
                "facts": [
                    {
                        "name": "Tenant:",
                        "value": "TENANT"
                    },
                    {
                        "name": "Date:",
                        "value": "DATETIME"
                    }],
                    "markdown" : "true"                   
                  }],
                "potentialAction": [{
                       "@type": "OpenUri",
                       "name": "Download the report",
                       "targets": [{
                           "os": "default",
                           "uri": "URI"
                 }],
    }  ]
   } 
"@

$NotificationBody = $Notification.Replace("TENANT","$TenantName").Replace("DATETIME","$DateNow").Replace("URI","$NewFileUri")
$Command = (Invoke-RestMethod -uri $TargetChannel -Method Post -body $NotificationBody -ContentType 'application/json')           

Write-Host "Command result" $Command

# Post to a Teams channel using PnP
Submit-PnPTeamsChannelMessage -Team $TargetTeamId -Channel $TargetTeamChannel -Message $Body -ContentType Html -Important
Disconnect-PnpOnline

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
