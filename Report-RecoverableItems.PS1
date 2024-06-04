# Report-RecoverableItems.PS1

function FormatFileSize {
    # Format File Size nicely
    param (
            [parameter(Mandatory = $true)]
            $InFileSize
        ) 
    
     If ($InFileSize -lt 1KB) { # Format the size of a document
            $FileSize = $InFileSize.ToString() + " B" } 
          ElseIf ($InFileSize -lt 1MB) {
            $FileSize = $InFileSize / 1KB
            $FileSize = ("{0:n2}" -f $FileSize) + " KB"} 
          Elseif ($InFileSize -lt 1GB) {
            $FileSize = $InFileSize / 1MB
            $FileSize = ("{0:n2}" -f $FileSize) + " MB" }
         Elseif ($InFileSize -ge 1GB) {
            $FileSize = $InFileSize / 1GB
            $FileSize = ("{0:n2}" -f $FileSize) + " GB" }
      Return $FileSize    
} 

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

# Set these values to those appropriate in your tenant (these values won't work)
$AppId = "ae3c3720-a554-4b76-8078-2c26bcc38bef"
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3487"
$AppSecret = "gg98Q~OCXoDFBHlXmCeDA8hVrZNvRexpFbTxAbix"

# Make sure that the app has the Mail.Read.All application permission

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
$global:token = ($tokenRequest.Content | ConvertFrom-Json).access_token

$Headers = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $Token" 
            'ConsistencyLevel' = "eventual" }

If (!($Token)) { Write-Host "Can't get Entra ID access token - exiting" ; break }

# Connect to Exchnage Online, if we're not already connected
$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}
# Connect to the Graph with an Entra ID app that has the Mail.Read permission
Connect-MgGraph -AppId $AppId -Scopes $Scopes -TenantId $TenantId -CertificateThumbprint $Thumbprint

[array]$NoReportFolders = @("Audits", "Calendar Logging", "SubstrateHolds")
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited
If (!($Mbx)) {
    Write-Host "No mailboxes found - exiting!"; break
}  

Write-Host ("Processing {0} mailboxes..." -f $Mbx.Count)
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {
    Write-Host ("Processing mailbox {0}" -f $M.UserPrincipalName)
   # Graph request 
   # $Uri = ("https://graph.microsoft.com/v1.0/users/{0}/MailFolders/RecoverableItemsRoot/childfolders" -f $M.ExternalDirectoryObjectId)

    Try {
       $RecoverableItemsRoot = (Get-MgUserMailFolder -Userid $M.ExternalDirectoryObjectId -MailFolderId 'RecoverableItemsRoot').id
       [array]$Folders = Get-MgUserMailFolderChildFolder -MailFolderId $RecoverableItemsRoot -UserId $M.ExternalDirectoryObjectId
    }
    Catch {
        Write-Host ("Error {0} when processing mailbox {1}" -f $_.Exception.Message, $M.UserPrincipalName)
        Continue
    }
  
    Write-Host ("Found {0} folders in  Recoverable Items" -f $Folders.Count)
    
    ForEach ($Folder in $Folders){
        # Ignore the folders used for mailbox auditing and calendar logging
        If ($Folder.displayName -in $NoReportFolders) {
        # Write-Host "Ignored folder" $Folder.displayName
            Continue
        }
        If ($Folder.TotalItemCount -gt 0) {
            Write-Host ("Processing folder {0} in mailbox" -f $Folder.displayName, $M.displayName)
        # Graph API request 
        # $Uri = ("https://graph.microsoft.com/v1.0/users/{0}/mailfolders/{1}/Messages/?`$select=sender,createdDateTime,subject&`$expand=singleValueExtendedProperties(`$filter=Id%20eq%20'LONG%200x0E08')" -f $M.ExternalDirectoryObjectId, $Folder.id)
       
            [array]$Items = Get-MgUserMailFolderMessage -UserId $M.ExternalDirectoryObjectId -MailFolderId $Folder.id -All `
                -Property sender,createdDateTime,subject`
                -ExpandProperty "singleValueExtendedProperties(`$filter=Id eq 'LONG 0x0E08')"
            # If some items are returned, report them
            Write-Host ("Found {0} items in folder {1}" -f $Items.Count, $Folder.displayName)
            ForEach ($Item in $Items) {
                [long]$ItemFileSize = $Item.singleValueExtendedProperties.value
                $ReportLine = [PSCustomObject][Ordered]@{
                    Mailbox     = $M.UserPrincipalName
                    Folder      = $Folder.displayName
                    Subject     = $Item.Subject
                    Sender      = $Item.sender.emailAddress.address
                    Created     = Get-Date($Item.createdDateTime).ToLocalTime()
                    Size        = FormatFileSize -InFileSize $ItemFileSize
                }
                $Report.Add($ReportLine)   
            }
        }
    }
}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.