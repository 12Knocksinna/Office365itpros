# Report-CopilotInteractions.PS1
# A script to show how to fetch details about Copilot interactions from the compliance records 
# captured by the Microsoft 365 substrate in the TeamsMessagesData folder. The script uses the 
# Get-MailboxFolderStatistics cmdlet to find the TeamsMessagesData folder and extract its identifier.
# The identifier is then converted to an EntryId and used to fetch messages from the folder using
# Microsoft Graph PowerShell SDK cmdlets.

# V1.0 5-Sep-2024

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-CopilotInteractions.PS1

# Note that if you want to run the code in application mode to collect data from multiple mailboxes, 
# the signed-in account must hold the Exchange administrator role.
[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ('ExchangeOnlineManagement' -notin $Modules) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

# Scopes needed 
# User.Read - read user profiles
# Mail.Read - read 
# These are application permissions
Connect-MgGraph -NoWelcome User.Read, Mail.Read

# The example code runs with delegated permissions. To run in application mode to extract data from
# multiple mailboxes, use an Entra ID app with a certificate to authenticate or run the code in Azure
# Automation with a managed identity.

$CP0 = "Microsoft 365 Chat"
$CP1 = "Copilot in Word"
$CP2 = "Copilot in Outlook"
$CP3 = "Copilot in PowerPoint"
$CP4 = "Copilot in Excel"
$CP5 = "Copilot in Teams"
$CP6 = "Copilot in Stream"
$CP7 = "Copilot in OneNote"
$CP8 = "Copilot in Loop"

[datetime]$StartDate = (Get-Date).AddDays(-365)
[string]$StartDate = Get-Date $StartDate -Format "yyyy-MM-ddTHH:mm:ssZ"
[string]$EndDate = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"

$User = (Get-MgContext).Account
$User = Get-MgUser -UserId $User
[array]$Folders = Get-ExoMailboxFolderStatistics -Identity $User.Id -FolderScope NonIPMRoot | Select-Object Name, FolderId 
$TeamsMessagesData = $Folders | Where-Object {$_.Name -eq "TeamsMessagesData"}
If ($TeamsMessagesData) {
    $FolderId = $TeamsMessagesData.FolderId
} Else {
    Write-Host "TeamsMessagesData folder not found"
    Break
}

# The folder id is in StoreId format. This code converts it to an EntryId
# convert from base64 to bytes
$folderIdBytes = [Convert]::FromBase64String($folderId)

# convert byte array to string, remove '-' and ignore first byte
$folderIdHexString = [System.BitConverter]::ToString($folderIdBytes).Replace('-','')
$folderIdHexStringLength = $folderIdHexString.Length

# get hex entry id string by removing first and last byte
$entryIdHexString = $folderIdHexString.SubString(2,($folderIdHexStringLength-4))

# convert to byte array - two chars represents one byte
$entryIdBytes = [byte[]]::new($entryIdHexString.Length / 2)

For($i=0; $i -lt $entryIdHexString.Length; $i+=2){
    $entryIdTwoChars = $entryIdHexString.Substring($i, 2)
    $entryIdBytes[$i/2] = [convert]::ToByte($entryIdTwoChars, 16)
}

# convert bytes to base64 string
$entryIdBase64 = [Convert]::ToBase64String($entryIdBytes)

# count how many '=' contained in base64 entry id
$equalCharCount = $entryIdBase64.Length - $entryIdBase64.Replace('=','').Length

# trim '=', replace '/' with '-', replace '+' with '_' and add number of '=' at the end
$EntryId = $entryIdBase64.TrimEnd('=').Replace('/','_').Replace('+','-')+$equalCharCount

# Now convert the entryId to be a RestId using the translateExchangeIds API
$Body = @{}
[array]$InputId = $EntryId
$Body.Add("inputIds", $InputId)
$Body.Add("sourceIdType", "entryId")
$Body.Add("targetIdType", "restid")

$Data = Invoke-MgTranslateUserExchangeId -UserId $User.Id -BodyParameter $Body
$RestId = $Data.targetId

Write-Host ("The RestId for the TeamsMessagesData folder is {0}" -f $RestId)

Set-MgRequestContext -MaxRetry 10 -RetryDelay 15 | Out-Null
Write-Host "Fetching messages sent by Copilot from the TeamsMessagesData folder"
# Find messages sent by Copilot
[array]$Items = $null
[array]$Items = Get-MgUserMailFolderMessage -UserId $User.Id -MailFolderId 'TeamsMessagesData' -All -PageSize 500  `
    -Filter "(ReceivedDateTime ge $StartDate and ReceivedDateTime le $EndDate) `
    and (sender/emailAddress/name eq '$CP0' or sender/emailAddress/name eq '$CP1' or sender/emailAddress/name eq '$CP2' `
    or sender/emailAddress/name eq '$CP3' or sender/emailAddress/name eq '$CP4' or sender/emailAddress/name eq '$CP5' `
    or sender/emailAddress/name eq '$CP6' or sender/emailAddress/name eq '$CP7' or sender/emailAddress/name eq '$CP8')" `
    -Property Sender, SentDateTime, BodyPreview, ToRecipients
   
If (!($Items)) {
    Write-Host "Failed to find folder identifier for the TeamsMessagesData folder"
    Break
}

Write-Host "Finding messages received by Copilot..."
# Now try and find messages received by Copilot
[array]$ItemsReceived = Get-MgUserMailFolderMessage -UserId $User.Id -MailFolderId 'TeamsMessagesData' `
    -All -PageSize 500 -Property Sender, SentDateTime, BodyPreview, ToRecipients `
    -Filter "(ReceivedDateTime ge $StartDate and ReceivedDateTime le $EndDate) `
    AND (singleValueExtendedProperties/any(ep:ep/id eq 'String 0x0E04' and contains(ep/value,'Copilot in')))" 
   

# And because we have some prompts received by "Microsoft 365 Chat", we need to find them too

[array]$ItemsChat = Get-MgUserMailFolderMessage -UserId $User.Id -MailFolderId 'TeamsMessagesData' `
    -All -PageSize 500 -Property Sender, SentDateTime, BodyPreview, ToRecipients `
    -Filter "(ReceivedDateTime ge $StartDate and ReceivedDateTime le $EndDate) `
    AND (singleValueExtendedProperties/any(ep:ep/id eq 'String 0x0E04' and ep/value eq 'Microsoft 365 Chat'))" 
   

If ($ItemsReceived) {
    $Items = $Items + $ItemsReceived
}
If ($ItemsChat) {
    $Items = $Items + $ItemsChat
}

Write-Host ("Found {0} messages sent and received by Copilot in the TeamsMessagesData folder" -f $Items.Count)

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Item in $Items) {
   
    $ReportLine = [PSCustomObject][Ordered]@{ 
        Sender = $Item.Sender.emailaddress.Name
        To = $Item.Torecipients.emailaddress.name -join ","
        Sent = $Item.SentDateTime
        Body = $Item.BodyPreview
    }
    $Report.Add($ReportLine)
}

# Output what we found
$Report | Sort-Object {$_.Sent -as [datetime]} | Out-GridView -Title "Copilot Interactions" 
$ReportCopilot = $Report | Where-Object {$_.Sender -ne $User.displayName}
$ReportCopilot | Group-Object -Property Sender | Select-Object Name, Count | Sort-Object Count -Descending | Out-GridView -Title "Copilot Interactions by App"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.