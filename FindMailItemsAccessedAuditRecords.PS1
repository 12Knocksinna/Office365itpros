# FindMailItemsAccessedAuditRecords.PS1
# V1.0 8-Apr-2020
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindMailItemsAccessedAuditRecords.PS1
# A script to process Office 365 "MailItemsAccessed" audit log records and put them into a more digestible format.
# Tony Redmond 
CLS
$Now = Get-Date
Write-Host "Finding MailItemsAccessed records..."
$SearchStartDate = (Get-Date).AddDays(-10)  #For Message Trace
$SearchEndDate = (Get-Date).AddDays(+1)
# Edit this line to use whatever search terms you wnat to find audit records
$Records = (Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-4) -EndDate $SearchEndDate -Operations MailItemsAccessed -ResultSize 3000) 
If ($Records.Count -eq 0) {
    Write-Host "No audit records for mail access found." }
Else {
    CLS
    $ProgressDelta = 100/($Records.count); $PercentComplete = 0; $RecordNumber = 0; 
    Write-Host "Processing" $Records.Count "MailItemsAccessed audit records..."
    $Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
    # Scan each audit record to extract information
    ForEach ($Rec in $Records) {
      $RecordNumber++
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      $TimeStamp   = Get-Date($AuditData.CreationTime) -format g
      $Status = "Audit record for " + $AuditData.MailboxOwnerUPN + " " + $Timestamp + " ["+ $RecordNumber +"/" + $Records.Count + "]"
      Write-Progress -Activity "Checking audit record" -Status $Status -PercentComplete $PercentComplete
      $PercentComplete += $ProgressDelta
      Switch ($AuditData.LogonType) { #Easier to read than to remember logon type codes
          "0" {$LogonType = "User"}
          "1" {$LogonType = "Admin"}
          "2" {$LogonType = "Delegate"}
      } # End Switch
      $RecordType =  $Auditdata.OperationProperties |?{$_.Name -eq "MailAccessType"}
      $DaysSince = New-TimeSpan(Get-Date($AuditData.CreationTime))
      $i = 0
      Switch ($RecordType.Value) { #Figure out the folder names and extract message information for Bind events
        "Bind" { $FolderId = $AuditData.Folders.Id; $Folder = $AuditData.Folders.Path.Split("\")[1]
                 ForEach ($Msg in $Auditdata.Folders.FolderItems) {
                   $i++; $Subject = $Null
                   # Section to try and find the message subject using a message trace
                   $MsgId = $Msg.InternetMessageId.Substring(1, ($Msg.InternetMessageId.Length -2)) #Trim the message identifier
                   If ($DaysSince.Days -le 10) { # Within the last 10 days so try a message trace
                        $Subject = (Get-MessageTrace -MessageId $MsgId -StartDate $SearchStartDate -EndDate $SearchEndDate)
                      If ($Null -ne $Subject) {$Subject = $Subject[0].Subject } #If we get a subject, take the first because a trace might return multiple results
                      }   # Not worth tracing because it's too far back
                   Else
                         { $Subject = "Too far back to trace" } # Find message subject 
                   If ($Subject -eq $Null) { $Subject = "***** No trace data available ******" }  # catch all 
                   If ($i -eq 1) { $Messages = "(" + $i + ") " + $MsgId + " (" + $Subject + ")" } # Format output
                   Else { $Messages = $Messages + "; (" + $i + ") " + $MsgId + " (" + $Subject + ")" }
                 } #End Foreach           
               } #End of Bind-specific processing
        "Sync" { $FolderId = $AuditData.Item.Id; $Folder = $AuditData.Item.Parentfolder.Name 
                 $Messages = $Null }
      } #End Switch
      $Throttled = $Null
      $Throttled = $Auditdata.OperationProperties |?{$_.Name -eq "IsThrottled"}    
         
      $ReportLine = [PSCustomObject] @{
           TimeStamp   = $TimeStamp
           Mailbox     = $AuditData.MailboxOwnerUPN
           User        = $AuditData.UserId
           LogonType   = $LogonType
           FolderId    = $FolderId
           Folder      = $Folder
           Access      = $RecordType.Value
           Operation   = $RecordType.Name
           Throttled   = $Throttled.Value
           ClientIP    = $AuditData.ClientIPAddress
           ClientInfo  = $AuditData.ClientInfoString
           SessionId   = $AuditData.SessionId
           Operations  = $AuditData.OperationCount
           Messages    = $Messages }        
      $Report.Add($ReportLine) }
}
$Report | Sort {$_.TimeStamp -as [DateTime]} |Out-GridView
$Report | Sort {$_.TimeStamp -as [DateTime]} | Export-CSV -NoTypeInformation c:\temp\MailItemsAccessed.csv

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
