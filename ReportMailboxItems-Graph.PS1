# ReportMailboxItems-Graph.PS1
# A script to demonstrate how to use Graph API queries to read the folders and folder items from an Exchange Online mailbox.
# By default, the script looks for items more than a year old (easily changed) in a single mailbox.
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportMailboxItems-Graph.PS1
# See https://practical365.com/mailbox-contents-report/ for an article describing the use of the script.
#+-------------------------- Functions etc. -------------------------

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

Clear-Host
# Check that we have the necessary Exchange Online module loaded
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {
    Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break
}

# Set these values to those appropriate in your tenant (these values won't work)
$AppId = "d48578ac-7cb4-4b5a-a296-f19218a03f11"
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3487"
$AppSecret = "bzS8Q~9EDXMUrUOJUbZXTTiJp7lTFdkWskETObRU"

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

If (!($Token)) { Write-Host "Can't get Azure AD access token - exiting" ; break }

# Define the set of folders that we're not interested in processing
[array]$FoldersToIgnore = "Deleted Items", "Sync Issues", "Conversation History", "Server Failures", "Local Failures", "Drafts", "Conflicts", "Junk Email", "Social Activity Notifications"

$Mailbox = Read-Host "Enter the name of the mailbox to report"      
$YearAgo = (Get-Date).AddDays(-365)
$YearAgo = (Get-Date -Date $YearAgo -Format s) + "Z"

[array]$Mbx = Get-ExoMailbox -Identity $Mailbox -ErrorAction SilentlyContinue
If (!($Mbx)) { 
   Write-Host ("A mailbox called {0} cannot be found - exiting" -f $Mailbox); break 
}
If ($Mbx.RecipientTypeDetails -ne 'UserMailbox') { Write-Host "This script can only process user mailboxes - exiting"; break }

[datetime]$StartTime = Get-Date
$FolderList = [System.Collections.Generic.List[Object]]::new()
Write-Host ("Checking mailbox information for mailbox {0} ({1})" -f $Mbx.DisplayName, $Mbx.UserPrincipalName)   
# Get list of folders in the mailbox
$Uri = $("https://graph.microsoft.com/v1.0/users/{0}/MailFolders?includeHiddenFolders=true" -f $Mbx.ExternalDirectoryObjectId)
[array]$AllFolders = Get-GraphData -Uri $Uri -AccessToken $Token
# Get list of folders that we want to process
$AllFolders = $AllFolders | Where-Object {$_.displayName -Notin $FoldersToIgnore} | Sort-Object DisplayName
# Build a table of folder ids and display names
ForEach ($F in $AllFolders) {
   $FolderLine = [PSCustomObject][Ordered]@{
      Id          = $F.Id
      DisplayName = $F.DisplayName } 
   $FolderList.Add($FolderLine)
} 
# Find folders with child folders
[array]$FoldersWithChildFolders = $AllFolders | Where-Object {$_.ChildFolderCount -gt 0}

ForEach ($ChildFolder in $FoldersWithChildFolders) {
  [array]$ChildFolders = UnpackFolders -FolderId $ChildFolder.Id -UserId $Mbx.ExternalDirectoryObjectId 
  ForEach ($ChildFolderProcess in $ChildFolders) {
     $FolderLine = [PSCustomObject][Ordered]@{
      Id          = $ChildFolderProcess.Id
      DisplayName = $ChildFolderProcess.DisplayName } 
     $FolderList.Add($FolderLine)
  } #End Foreach ChildFolder
}
   
# Remove any blank folders that might have been returned by the Graph
$FolderList = $FolderList | Where-Object {$_.DisplayName -ne $Null }| Sort-Object DisplayName
Clear-Host

If (!($FolderList)) { Write-Host ("Can't find any folders in the {0} mailbox - script exiting..." -f $Mbx.displayname); break }
$ReportList = [System.Collections.Generic.List[Object]]::new()
$ProgDelta = 100/($FolderList.Count); $i = 0; $Progress = 0
ForEach ($Folder in $FolderList) {
  $i++
  $Status = $Folder.DisplayName + " [ Folder "+ $i +"/" + $FolderList.Count + " Items found so far " + $ReportList.count + "]"
  Write-Progress -Activity "Reporting items in folder more than a year old" -Status $Status -PercentComplete $Progress
  $Progress += $ProgDelta

  # Find messages from a year ago
  $Uri = 'https://graph.microsoft.com/v1.0/users/' + $Mbx.ExternalDirectoryObjectId + "/mailfolders/" + $Folder.Id + `
     "/Messages/?select=receivedDateTime,subject,sender,id,importance,torecipients,ccrecipients,bccrecipients,bodypreview&`$filter=receivedDateTime le $($YearAgo)&`$expand=singleValueExtendedProperties(`$filter=Id%20eq%20'LONG%200x0E08')"
  [array]$Messages = Get-GraphData -Uri $Uri -AccessToken $Token
  If (!($Messages[0].'@odata.context')) { # Check that the Graph has returned some messages or just a note that it hasn't found anything
    Write-Host ("Found {0} messages in the {1} folder" -f $Messages.count, $Folder.DisplayName) 
    ForEach ($Message in $Messages) {
     $CcRecipients = $Null; $ToRecipients = $Null
     $BodyText = ($Message.BodyPreview -replace '</p>',"`r`n" -replace "<[^>]+>",'' -replace "&nbsp;",' ').trim()
     If ($Message.torecipients.emailaddress.name.count -gt 0) {
        $ToRecipients = $Message.torecipients.emailaddress.name -join ", " }
     Else {
        $ToRecipients = $Message.torecipients.emailaddress.name }
     If ($Message.ccrecipients.emailaddress.name.count -gt 0) {
        $CcRecipients = $Message.ccrecipients.emailaddress.name -join ", " }
     Else {
        $CcRecipients = $Message.ccrecipients.emailaddress.name }
     [long]$Size = $Message.singleValueExtendedProperties.value
     $MsgSize = FormatFileSize $Size
     # Add each message to the list
       $ReportLine = [PSCustomObject][Ordered]@{  
         Folder     = $Folder.DisplayName
         Received   = $Message.receivedDateTime
         Subject    = $Message.subject
         To         = $ToRecipients
         CC         = $CcRecipients
         Sender     = $message.sender.emailaddress.address
         SenderName = $message.sender.emailaddress.name
         Importance = $Message.importance
         Size       = $MsgSize
         Body       = $BodyText
         id         = $Message.Id 
         FolderId   = $Folder.Id
         UserId     = $Mbx.ExternalDirectoryObjectId }
       $ReportList.Add($ReportLine)
    } # End Foreach Message 
  } # End if Messages
} # End Foreach Folder

[datetime]$EndTime = Get-Date
$Elapsed = $EndTime - $StartTime
Write-Host ("Elapsed time {0}" -f $Elapsed)
Write-Host ""
Write-Host ("{0} messages found" -f $ReportList.Count)
$ExcelFileName = "c:\temp\MailboxData_" + $Mbx.ExternalDirectoryObjectId + ".xlsx"
$ExcelTitle = "Mailbox Data for " + $Mbx.DisplayName
$WorkSheetTitle = "Created " + (Get-Date -format "dd-MMM-yyyy hhmm")
$ReportList | Export-Excel -Path $ExcelFileName -WorksheetName $WorkSheetTitle -Title $ExcelTitle -TitleBold -TableName "MailboxData" -TableStyle Medium28
Write-Host "Output Excel spreadsheet is available in" $ExcelFileName

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
