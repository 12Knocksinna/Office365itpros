# CleanUpMailbox-Graph.PS1
# A simple PowerShell script to show how to use Graph queries to clean up mailbox items
# 
# https://github.com/12Knocksinna/Office365itpros/blob/master/CleanUpMailbox-Graph.PS1
# See https://practical365.com/mailbox-contents-report/ for an article describing how to use the script
#
# Parameters to tell the script what to do. 
# Folder is the name of a specific folder to check. Set this to All to search all mailbox folders. 
# DeleteItems parameter controls if we delete the found messages or just report them. 
# SearchQuery is the KQL keyword to search message body and attachments (optional). 
# Mailbox is All to search all mailboxes or a mailbox alias, DN, or SMTP address to search an individual mailbox
# SenderAddress is the sender address to look for (optional)
# MessageSubject is the message subject to look for (mandatory)
# StartDate is the start date for the search (mandatory)
# EndDate is the end date for the search (optional). By default, this will be today's date if not specified and a start date is present.

# Requires an Azure AD registered app with Mail.ReadWrite permission
# 
# Example usage
# .\CleanUpMailbox-Graph.PS1 -Folder All -SearchQuery "Special Offer" -MessageSubject Offer -DeleteItems Y -Mailbox All -StartDate 1-Apr-2022 -EndDate 30-Jul-2022

param (
        [parameter(Mandatory = $true)]$Folder,
		
        [parameter(Mandatory = $true)]$SearchQuery,

        [parameter(Mandatory = $true)]$Mailbox,

        [parameter()]$SenderAddress,

        [parameter(Mandatory = $true)]$MessageSubject,

        [parameter(Mandatory = $true)]$StartDate,

        [parameter()]$EndDate,

        [parameter(Mandatory = $true)]
        [ValidateSet('Y','N')]
        $DeleteItems,
	
	[parameter(Mandatory = $true)]
	$AppId,
	
	[parameter(Mandatory = $true)]
	$TenantId,
	
	[parameter(Mandatory = $true)]
	$AppSecret
    )

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

Function UnpackFolders {
# Unpack a set of folders to return their ids and displaynames - we go down 4 levels, which is quite enough
# Input parameter is the identifier of a top-level mailbox folder

param (
  [parameter(mandatory = $True)]
   $FolderId, 
  [parameter(mandatory = $true) ]
  $UserId
)

$Level3 = $Null; $Level4 = $Null; $Level2 = $Null; $NFF2 = $Null; $NFF3 = $Null
# Get folders in the child folder
   [array]$Output = $Null
   $Uri = $("https://graph.microsoft.com/v1.0/users/{0}/MailFolders/{1}/childfolders" -f $UserId, $FolderId)
   [array]$Level1 = Get-GraphData -Uri $Uri -AccessToken $Token
   $Output = $Level1
   $Level2 = $Level1 | Where-Object {$_.ChildFolderCount -gt 0}
   If ($Level2) {
      ForEach ($NF2 in $Level2) {
       $Uri = $Uri = $("https://graph.microsoft.com/v1.0/users/{0}/MailFolders/{1}/childfolders" -f $UserId, $NF2.Id)
       [array]$NFF2 = Get-GraphData -Uri $Uri -AccessToken $Token
       $Output = $Output + $NFF2 }}
   $Level3 = $NFF2 | Where-Object {$_.ChildFolderCount -gt 0}
   If ($Level3) {
     ForEach ($NF3 in $Level3) {
      $Uri = $Uri = $("https://graph.microsoft.com/v1.0/users/{0}/MailFolders/{1}/childfolders" -f $UserId, $NF3.Id)
       [array]$NFF3 = Get-GraphData -Uri $Uri -AccessToken $Token
     $Output = $Output + $NFF3 }}
   $Level4 = $NFF3 | Where-Object {$_.ChildFolderCount -gt 0}
   If ($Level4) {
     ForEach ($NF4 in $Level4) {
      $Uri = $Uri = $("https://graph.microsoft.com/v1.0/users/{0}/MailFolders/{1}/childfolders" -f $UserId, $NF4.Id)
       [array]$NFF4 = Get-GraphData -Uri $Uri -AccessToken $Token
     $Output = $Output + $NFF4 }
 }
  Return $Output
}

# End Functions

# Check that we have the necessary Exchange Online module loaded
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}

# Set these values to those appropriate in your tenant
# Removing AppId, TenantID, and AppSecret variables, and pass them as parameters

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

If (!($Token)) { Write-Host "Can't get access token - exiting" ; break }

# Prepare search filter
$SearchFilter = "subject:$MessageSubject" 

If ($SenderAddress) {
  $SearchFilter = $SearchFilter + " AND from:$SenderAddress" }

If ($StartDate) {
   $StartDateFilter = (Get-Date $StartDate).toString('yyyy-MM-dd') }
 If (($StartDateFilter) -and (!($EndDate))) { # if we have a start date but no end date, set to today's date
   $EndDate = Get-Date }
 If ($EndDate) {
   $EndDateFilter = (Get-Date $EndDate).toString('yyyy-MM-dd') }

If (($StartDateFilter) -and ($EndDateFilter)) {
  $SearchFilter = $SearchFilter + " AND received>=$StartDateFilter AND received<=$EndDateFilter" }

If ($SearchQuery) {
  $SearchFilter = $SearchFilter + " AND '" + $SearchQuery + "'" }

Write-Host "Search criteria:"
Write-Host "----------------"
Write-Host "Search filter:      $SearchFilter"
Write-Host "Target mailboxes:   $Mailbox"
Write-Host "Target folder:      $Folder"
Write-Host "Sender address:     $SenderAddress"
Write-Host "Date from:         " (Get-Date($StartDateFilter) -format dd-MMM-yyyy)
Write-Host "End date:          " (Get-Date($EndDateFilter) -format dd-MMM-yyyy)
Write-Host "Delete found items " $DeleteItems
Write-Host ""

Write-Host "Finding target mailboxes..."
If ($Mailbox -eq "All") {   
   [array]$Mbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox }
Else {
   [array]$Mbx = Get-ExoMailbox -Identity $Mailbox }
If (!($Mbx)) { Write-Host "No mailboxes found - exiting"; break }
Write-Host ("{0} mailboxes found." -f $Mbx.count)

$DeletionsList = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {
  $DataTable = @{}
  If ($Folder -eq "All") { # Process all folders
   # Get list of folders in the mailbox
   $Uri = $("https://graph.microsoft.com/v1.0/users/{0}/MailFolders?includeHiddenFolders=true" -f $M.ExternalDirectoryObjectId)
   [array]$AllFolders = Get-GraphData -Uri $Uri -AccessToken $Token
   $AllFolders = $AllFolders | Sort-Object Id -Unique
   # Build a hash table of folder ids and display names
   ForEach ($F in $AllFolders) {
     $DataTable.Add([String]$F.Id,[String]$F.DisplayName) }
   # Find folders with child folders
   [array]$FoldersWithChildFolders = $AllFolders | Where-Object {$_.ChildFolderCount -gt 0}

   ForEach ($ChildFolder in $FoldersWithChildFolders) {
   [array]$ChildFolders = UnpackFolders -FolderId $ChildFolder.Id -UserId $M.ExternalDirectoryObjectId
   ForEach ($F in $ChildFolders) {
     Try {
        $DataTable.Add([String]$F.Id,[String]$F.DisplayName) }
     Catch {}
    }
   }

   # Build Uri to look for matching messages across all folders
   $Uri = 'https://graph.microsoft.com/v1.0/users/' + $M.ExternalDirectoryObjectId + "/messages?`$search=" + '"' + $SearchFilter + '"' + "&`$select=id,parentfolderid,receivedDateTime,subject,from" 
} # End if
  Else { # Process an individual folder
   # Find the target folder
   $Uri = $("https://graph.microsoft.com/v1.0/users/{0}/mailFolders?`$filter=displayName eq '{1}'" -f $M.ExternalDirectoryObjectId, $Folder) 
   [Array]$TargetFolder = Get-GraphData -AccessToken $Token -Uri $Uri
   If (!($TargetFolder)) { Write-Host ("Can't find the {0} folder - exiting" -f $Folder); break }
   Write-Host ""
   Write-Host ( $("Mailbox        {0}" -f $M.DisplayName))
   Write-Host ( $("Target folder  {0}" -f $TargetFolder.displayName))
   Write-Host ( $("Unread items   {0}" -f $TargetFolder.unreadItemCount))
   Write-Host ( $("Total items    {0}" -f $TargetFolder.totalItemCount))
   Write-Host ""
   If ($TargetFolder.totalItemCount -eq 0) { Write-Host ("No items are in the {0} folder..." -f $Folder) }
   # Build Uri to find matching messages in the target folder
   $Uri = 'https://graph.microsoft.com/v1.0/users/' + $M.ExternalDirectoryObjectId + "/mailfolders/" + $TargetFolder.Id + "/messages?`$search=" + '"' + $SearchFilter + '"' + "&`$select=id,parentfolderid,receivedDateTime,subject,from" 
  
 } #End Else
 
[int]$i = 0; $Action = "Delete"
If ($DeleteItems -eq "N") { $Action = "Report only" }

Write-Host ("Searching for matching messages in mailbox {0}... ({1}/{2})" -f $M.DisplayName, $counter, $Mbx.Count)
# Get messages that aren't in user folders that aren't Deleted Items
[Array]$Messages = Get-GraphData -Uri $Uri -AccessToken $Token

# If processing all folders, search Deleted Items too
# Using Well-known folder names https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.wellknownfoldername?view=exchange-ews-api
If ($Folder -eq "All") {
   $Uri = 'https://graph.microsoft.com/v1.0/users/' + $M.ExternalDirectoryObjectId + "/mailfolders('DeletedItems')/messages?`$search=" + '"' + $SearchFilter + '"' + "&`$select=id,parentfolderid,receivedDateTime,subject,from"
   [array]$DeletedItemsMessages = Get-GraphData -Uri $Uri -AccessToken $Token
   $Messages = $Messages + $DeletedItemsMessages }

# Fetch messages in the Deletions folder in Recoverable Items
# No point in including them if we're deleting items because the Graph won't let you delete them
# See https://www.michev.info/Blog/Post/3849/can-you-delete-mailbox-items-on-hold-via-the-graph-api
# Only fetch these messages if we're in report only mode and processing all folders
If (($Action -eq "Report only") -and ($Folder -eq "All")) {
  $Uri = 'https://graph.microsoft.com/v1.0/users/' + $M.ExternalDirectoryObjectId + "/mailfolders('RecoverableItemsDeletions')/messages?`$search=" + '"' + $SearchFilter + '"' + "&`$select=id,parentfolderid,receivedDateTime,subject,from" 
  [array]$Deletions = Get-GraphData -Uri $Uri -AccessToken $Token
  $Messages = $Messages + $Deletions
  # This code is to retrieve the display name of the Deletions folder and insert it into the hash table used for folder lookups
  $Uri = "https://graph.microsoft.com/v1.0/users/" + $M.ExternalDirectoryObjectId + "/mailfolders('RecoverableItemsDeletions')"
  [array]$DeletionFolderData = Get-GraphData -Uri $Uri -AccessToken $Token
  $DeletionFolderId = $DeletionFolderData[0].id
  $DeletionFolderName = $DeletionFolderData[0].DisplayName
  $DataTable.Add([String]$DeletionFolderId,[String]$DeletionFolderName)
 }

# Sometimes the set of messages returns includes an item with @odata.context rather than @odata.etag. This filter removes those records
 $Messages = $Messages | Where-Object {$_.id -ne $Null}

If (($Messages.Count -gt 0) -and ($Null -eq $Messages."@odata.context")) { #We have some messages to delete or report

  Write-Host ("Found {0} matching message(s) in mailbox {1} " -f $Messages.count, $M.DisplayName)
  ForEach ($Message in $Messages) {
    $i++
    Write-Host ("Processing Message {0} ({1})" -f $i, $Action)
    # Log details of what happened to a message
   $FolderName = $Folder 
   If ($Folder -eq "All") { #Resolve parent folder name
     Try {
        $FolderName = $DataTable[$Message.ParentFolderId] }
     Catch {
        $FolderName = "Unresolved folder name" }
    }
    If ($Message.From.EmailAddress.Address -like "*ExchangeLabs*")  {
          $MessageSender = $M.PrimarySmtpAddress }
    Else {
          $MessageSender = $Message.From.EmailAddress.Address }
       
    If ([string]::IsNullOrEmpty($Message.ReceivedDateTime)) { 
          $ReceivedDate = "Not noted" }
    Else {
          $ReceivedDate =  Get-Date ($Message.ReceivedDateTime) -format g }

       $DeletionLine = [PSCustomObject][Ordered]@{  # Write out details of the group
          Mailbox             = $M.DisplayName
          UPN                 = $M.UserPrincipalName
          "User type"         = $M.RecipientTypeDetails
          Subject             = $Message.Subject
          Folder              = $FolderName
          From                = $MessageSender
          ReceivedDate        = $ReceivedDate
          Id                  = $Message.Id  
          ProcessDate         = Get-Date -format u
          Action              = $Action }
       $DeletionsList.Add($DeletionLine) 
       If ($Action -eq "Delete") {
          # This puts the deleted item into the Deletions sub-folder of Recoverable Items
          $Uri = $("https://graph.microsoft.com/v1.0/users/{0}/messages/{1}" -f $M.ExternalDirectoryObjectId, $Message.Id)
          $Status = Invoke-RestMethod $Uri -Method 'Delete' -Headers $Headers  }
       }
    } #End If check that some items exist
} #End loop through Mbx

Write-Host ""
Write-Host ("{0} messages were found and processed" -f $DeletionsList.count)
Write-Host ""
Write-Host "Information about the messages is available in c:\temp\DeletionsList.csv"
Write-Host ""

$DeletionsList | Select-Object Mailbox, UPN, Subject, Folder, From, ReceivedDate, Action | Out-GridView 
$DeletionsList | Export-CSV -NoTypeInformation c:\temp\DeletionsList.csv

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
