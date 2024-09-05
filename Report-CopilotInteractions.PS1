# Report-CopilotInteractions.PS1

Connect-ExchangeOnline -SkipLoadingCmdletHelp
Connect-MgGraph -NoWelcome

$CP0 = "Microsoft 365 Chat"
$CP1 = "Copilot in Word"
$CP2 = "Copilot in Outlook"
$CP3 = "Copilot in PowerPoint"
$CP4 = "Copilot in Excel"

[datetime]$StartDate = (Get-Date).AddDays(-365)
[string]$StartDate = Get-Date $StartDate -Format "yyyy-MM-ddTHH:mm:ssZ"
[string]$EndDate = Get-Date $EndDate -Format "yyyy-MM-ddTHH:mm:ssZ"

$User = (Get-MgContext).Account
$User = Get-MgUser -UserId $User
[array]$Folders = Get-MailboxFolderStatistics -Identity $User.Id -FolderScope NonIPMRoot -ResultSize Unlimited | Select-Object Name, FolderId 
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

# count how many '=' contains base64 entry id
$equalCharCount = $entryIdBase64.Length - $entryIdBase64.Replace('=','').Length

# trim '=', replace '/' with '-', replace '+' with '_' and add number of '=' at the end
$EntryId = $entryIdBase64.TrimEnd('=').Replace('/','_').Replace('+','-')+$equalCharCount

$Body = @{}
[array]$InputId = $EntryId
$Body.Add("inputIds", $InputId)
$Body.Add("sourceIdType", "entryId")
$Body.Add("targetIdType", "restid")
[array]$Data = Invoke-MgGraphRequest -Method post -Uri $Uri -Body $Body
$FolderId = $Data.value.targetId

If ($FolderId) {
    Write-Host "Fetching messages from the TeamsMessagesData folder"
    [array]$Items = Get-MgUserMailFolderMessage -UserId $User.Id -MailFolderId $RestId -All -PageSize 999  `
        -Filter "(ReceivedDateTime ge $StartDate and ReceivedDateTime le $EndDate) `
        and (sender/emailAddress/name eq '$CP0' or sender/emailAddress/name eq '$CP1' or sender/emailAddress/name eq '$CP2' `
        or sender/emailAddress/name eq '$CP3' or sender/emailAddress/name eq '$CP4')" `
        -Property Sender, SentDateTime, BodyPreview, From
} Else {
    Write-Host "Failed to find folder identifier for the TeamsMessagesData folder"
    Break
}

Write-Host ("Found {0} messages in the TeamsMessagesData folder" -f $Items.Count)
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Item in $Items) {
   
    $ReportLine = [PSCustomObject][Ordered]@{ 
        Sender = $Item.Sender.emailaddress.Name
        From = $Item.from.emailaddress.name
        Sent = $Item.SentDateTime
        Body = $Item.BodyPreview
    }
    $Report.Add($ReportLine)
}

$Report | Out-GridView -Title "Copilot Interactions" 
$Report | Group-Object -Property Sender | Select-Object Name, Count | Sort-Object Count -Descending | Out-GridView -Title "Copilot Interactions by Sender"