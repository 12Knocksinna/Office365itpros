# Clean-MailboxMessages.PS1
# V1.0

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
Function FormatDuration ([TimeSpan]$Duration) {
    $h = $Duration.Days * 24 + $Duration.Hours
    $m = $Duration.Minutes

    switch ($h) {
        0 {$HourStr=$null}
        1 {$HourStr="$h hour"}
        {$_ -ge 1} {$HourStr="$h hours"}
    }
    switch ($m) {
        0 {$MinuteStr=$null}
        1 {$MinuteStr="$m minute"}
        {$_ -ge 1} {$MinuteStr="$m minutes"}
    }
    if ($HourStr -gt 0 -and $MinuteStr -gt 0) {
        [String]$DurationStr=$HourStr+" "+$MinuteStr
    } else {
        $DurationStr=$HourStr+$MinuteStr
    }

    return $DurationStr
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
           # Get-MgUserMailFolderChildFolder -UserId $userId -MailFolderId $mailFolderId
           $Uri = $("https://graph.microsoft.com/v1.0/users/{0}/MailFolders/{1}/childfolders" -f $UserId, $FolderId)
           [array]$Level1 = Invoke-MgGraphRequest -Uri $Uri -Method Get | Select-Object -ExpandProperty Value
           $Output = $Level1
           $Level2 = $Level1 | Where-Object {$_.ChildFolderCount -gt 0}
           If ($Level2) {
              ForEach ($NF2 in $Level2) {
               $Uri = $Uri = $("https://graph.microsoft.com/v1.0/users/{0}/MailFolders/{1}/childfolders" -f $UserId, $NF2.Id)
               [array]$NFF2 = Invoke-MgGraphRequest -Uri $Uri -Method Get | Select-Object -ExpandProperty Value
               $Output = $Output + $NFF2 }}
           $Level3 = $NFF2 | Where-Object {$_.ChildFolderCount -gt 0}
           If ($Level3) {
             ForEach ($NF3 in $Level3) {
              $Uri = $Uri = $("https://graph.microsoft.com/v1.0/users/{0}/MailFolders/{1}/childfolders" -f $UserId, $NF3.Id)
               [array]$NFF3 = Invoke-MgGraphRequest -Uri $Uri -Method Get | Select-Object -ExpandProperty Value
             $Output = $Output + $NFF3 }}
           $Level4 = $NFF3 | Where-Object {$_.ChildFolderCount -gt 0}
           If ($Level4) {
             ForEach ($NF4 in $Level4) {
              $Uri = $Uri = $("https://graph.microsoft.com/v1.0/users/{0}/MailFolders/{1}/childfolders" -f $UserId, $NF4.Id)
               [array]$NFF4 = Invoke-MgGraphRequest -Uri $Uri -Method Get | Select-Object -ExpandProperty Value
             $Output = $Output + $NFF4 }
         }
          Return $Output
 }
        
# End Functions

$AppId = "d58578ac-7cb4-4b5a-a296-f19218a03f11"
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f3478"
$Thumbprint = "32C9529B1FFD08BCD483A5D98807E47A472C5318"
Connect-MgGraph -AppId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint

# Define the set of folders that we're not interested in processing
[array]$FoldersToIgnore = "Deleted Items", "Sync Issues", "Conversation History", "Server Failures", "Local Failures", "Drafts", "Conflicts", "Junk Email", "Social Activity Notifications"
$YearAgo = (Get-Date).AddDays(-365)
$YearAgo = (Get-Date -Date $YearAgo -Format s) + "Z"

Clear-Host
# Check that we have the necessary Exchange Online module loaded
[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break
}

# Who are we processing?
$Mailbox = Read-Host "Enter the name of the mailbox to report"  
[array]$Mbx = Get-ExoMailbox -Identity $Mailbox -ErrorAction SilentlyContinue
If (!($Mbx)) { 
   Write-Host ("A mailbox called {0} cannot be found - exiting" -f $Mailbox); break 
}
If ($Mbx.RecipientTypeDetails -ne 'UserMailbox') { 
    Write-Host "This script can only process user mailboxes - exiting"; break 
} Else {
    Write-Host ("Processing mailbof for {0}" -f $Mbx.DisplayName) 
}
[datetime]$StartTime = Get-Date

$Uri = ("https://graph.microsoft.com/v1.0/users/{0}/MailFolders?includeHiddenFolders=true&`$Top=500" -f $Mbx.ExternalDirectoryObjectId)
[array]$AllFolders = Invoke-MgGraphRequest -Uri $Uri -Method Get | Select-Object -ExpandProperty Value
If ($AllFolders.Count -eq 0) { 
    Write-Host "No folders found in the mailbox - exiting"; break 
} Else {
    $AllFolders = $AllFolders | Where-Object {$_.displayName -Notin $FoldersToIgnore} | Sort-Object DisplayName
    Write-Host ("Found {0} folders to process in the mailbox" -f $AllFolders.Count)
    $FolderList = [System.Collections.Generic.List[Object]]::new()
}

# Build a table of folder ids and display names
ForEach ($F in $AllFolders) {
    $FolderLine = [PSCustomObject][Ordered]@{
       Id          = $F.Id
       DisplayName = $F.DisplayName } 
    $FolderList.Add($FolderLine)
} 

ForEach ($ChildFolder in $FoldersWithChildFolders) {
    [array]$ChildFolders = UnpackFolders -FolderId $ChildFolder.Id -UserId $Mbx.ExternalDirectoryObjectId 
    ForEach ($ChildFolderProcess in $ChildFolders) {
        $FolderLine = [PSCustomObject][Ordered]@{
            Id          = $ChildFolderProcess.Id
            DisplayName = $ChildFolderProcess.DisplayName } 
        $FolderList.Add($FolderLine)
    } 
}

# Remove any blank folders that might have been returned by the Graph
$FolderList = $FolderList | Where-Object {$_.DisplayName -ne $Null }| Sort-Object DisplayName

If (!($FolderList)) { 
    Write-Host ("Can't find any folders in the {0} mailbox - script exiting..." -f $Mbx.displayname); break 
}

$ReportList = [System.Collections.Generic.List[Object]]::new()
$ProgDelta = 100/($FolderList.Count); $i = 0; $Progress = 0
ForEach ($Folder in $FolderList) {
    $i++
    $Status = ("Processing folder: {0} ({1} of {2}. {3} items processed so far)" -f $Folder.DisplayName, $i, $FolderList.Count, $ReportList.count)
    Write-Progress -Activity "Reporting items in folder more than a year old" -Status $Status -PercentComplete $Progress
    $Progress += $ProgDelta
  
    # Find messages from a year ago
    
    [array]$Messages = Get-MgUserMailFolderMessage -UserId $Mbx.ExternalDirectoryObjectId -MailFolderId $Folder.id -All -PageSize 999 `
        -Property id, sender, receivedDateTime, createdDateTime, importance, subject, torecipients, ccrecipients,bccrecipients,bodypreview `
        -ExpandProperty "singleValueExtendedProperties(`$filter=Id eq 'LONG 0x0E08')" `
        -Filter "receivedDateTime le $($YearAgo)" `
  
    If ($Messages) { 
        Write-Host ("Found {0} messages in the {1} folder" -f $Messages.count, $Folder.DisplayName) 
        ForEach ($Message in $Messages) {
            $CcRecipients = $Null; $ToRecipients = $Null
            $BodyText = ($Message.BodyPreview -replace '</p>',"`r`n" -replace "<[^>]+>",'' -replace "&nbsp;",' ').trim()
            If ($Message.torecipients.emailaddress.name.count -gt 0) {
                $ToRecipients = $Message.torecipients.emailaddress.name -join ", " 
            } Else {
                $ToRecipients = $Message.torecipients.emailaddress.name 
            }
            If ($Message.ccrecipients.emailaddress.name.count -gt 0) {
                $CcRecipients = $Message.ccrecipients.emailaddress.name -join ", " 
            } Else {
            $CcRecipients = $Message.ccrecipients.emailaddress.name 
            }
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
$Elapsed = FormatDuration $Elapsed
Write-Host ("Elapsed time {0}" -f $Elapsed)
Write-Host ""
Write-Host ("{0} messages found" -f $ReportList.Count)


Write-Host "Generating report..."
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\MailboxData_" + $Mbx.ExternalDirectoryObjectId + ".xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    }
    $ReportList | Export-Excel -Path $ExcelOutputFile -WorksheetName "Mailbox Items" `
        -Title ("Mailbox Items {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "MailboxItems"
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\MailboxData_" + $Mbx.ExternalDirectoryObjectId + ".CSV"
    $ReportList | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}
 
If ($ExcelGenerated) {
    Write-Host ("An Excel report of calendar items is available in {0}" -f $ExcelOutputFile)
} Else {    
    Write-Host ("A CSV report of calendar items is available in {0}" -f $CSVOutputFile)
}  

Write-Host "All done..."

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.