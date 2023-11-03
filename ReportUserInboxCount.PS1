# ReportUserInboxCount.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportUserInboxCount.PS1
# A script to show how to fetch mailbox folder statistics using a combination of PowerShell and the Microsoft Graph ListFolders API
# You need an Azure AD registered app with consent for the Mail.Read application permission.

Clear-Host
# Check that we are connected to Exchange Online
Import-Module ImportExcel -ErrorAction SilentlyContinue
$ExcelAvailable = $True
[array]$ModulesLoaded = Get-Module | Select-Object -ExpandProperty Name 
If  ("ExchangeOnlineManagement" -notin $ModulesLoaded) {
    Write-Host "Please connect to the Exchange Online Management module and then restart the script"
    Break
}
# See if we can use ImportExcel
If ("ImportExcel" -notin $ModulesLoaded) {
    $ExcelAvailable = $False
}

$ExcelOutput = "c:\temp\UserInboxCounts.xlsx"
$CSVOutput = "c:\temp\UserInboxCounts.csv"

# Define the values applicable for the application used to connect to the Graph - change these for your tenant
# to use the app id and app secret for the Azure AD registered app you want to use for this script
$AppId = "829e1143-88e3-492b-bf82-24c4a37ada63"
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3477"
$AppSecret = 'yVh8Q~oA9.QJchbNqSDd89Iz1-xNTZ9m8PPVYdgP'

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

Write-Host "Finding user mailboxes..."
[array]$Mbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Sort-Object DisplayName
If (!($Mbx)) { Write-Host "No mailboxes found... exiting!" ; break }
$Report = [System.Collections.Generic.List[Object]]::new()
$i = 1
ForEach ($M in $Mbx) {
    Write-Host ("Processing mailbox {0} of {1}: {2}" -f $i, $Mbx.Count, $M.DisplayName); $i++
    $Uri = "https://graph.microsoft.com/v1.0/users/" + $M.ExternalDirectoryObjectId + "/mailFolders?`$top=250"
    $FolderData = Invoke-RestMethod -Headers $Headers -Uri $Uri -UseBasicParsing -Method "GET" -ContentType "application/json"
    $InboxData = $FolderData.Value | Where-Object {$_.displayname -eq "Inbox"}
    $TotalMbxItems = ($FolderData.Value.totalitemcount | Measure-Object -Sum | Select-Object -ExpandProperty Sum)
    $TotalMbxSize = ($FolderData.Value.SizeInBytes | Measure-Object -Sum | Select-Object -ExpandProperty Sum)
    $ReportLine = [PSCustomObject][Ordered]@{  # Write out details of the mailbox
       "User"              = $M.DisplayName
       UPN                 = $M.UserPrincipalName
       InboxCount          = $InboxData.totalItemCount
       UnreadCount         = $InboxData.unreadItemCount
       TotalMbxFolders     = $FolderData.Value.Count
       TotalMbxItems       = $TotalMbxItems
       TotalMbxFolderSize  = [math]::Round($TotalMbxsize/1Mb,2)  }
      $Report.Add($ReportLine) 
}

# If you don't have the ImportExcel module available, you can change this to export to a CSV file using the Export-CSV cmdlet
$Report | Sort-Object UnreadCount -Descending
If ($ExcelAvailable -eq $True) {
    Export-Excel -Path $ExcelOutput -WorkSheetName "Inbox Counts" -TableStyle Medium16 `
      -Title "Inbox Data for Mailboxes"
    Write-Host ("Excel worksheet available in {0}" -f $ExcelOutput)
} Else {
    Export-CSV -NoTypeInformation $CSVOutput
    Write-Host ("CSV file available in {0}" -f $CSVOutput)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
