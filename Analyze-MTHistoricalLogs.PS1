# Analyze-MTHistoricalLogs.PS1
# Folder where the historical message tracking logs downloaded from Exchange Online are stored
$DataFolder = "c:\temp\MtData\"
$CSVFile = "c:\temp\HistoricalMessageTrace.CSV"

[array]$DataFiles = Get-ChildItem -Path $DataFolder | Select-Object -ExpandProperty Name

If (!($DataFiles)) {
    Write-Host "No historical message tracking logs to analyze - exiting"
    Break
}

Write-Host ("Preparing to process {0} historical message trace data files..." -f $DataFiles.count)
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report

ForEach ($File in $DataFiles) {
    $MtDataFile = $DataFolder + $File
    [array]$MtData = Import-CSV -Path $MtDataFile -Encoding unicode
        ForEach ($Line in $MtData) {
            If (!([string]::IsNullOrEmpty($Line.origin_timestamp_utc))) {
                [array]$RecipientStatus = $Line.Recipient_Status.split(";")
                # array of individual recipients for a message
                $RecipientInfo = [System.Collections.Generic.List[Object]]::new()
                ForEach ($RecipientDetail in $RecipientStatus) {
                    $Recipient = $RecipientDetail.Split("##")[0]
                    $RecipientOutcome = $RecipientDetail.Split("##")[1]
                    $RecipientLine = [PSCustomObject]@{ 
                        Recipient = $Recipient
                        Outcome   = $RecipientOutcome
                    }
                    $RecipientInfo.Add($RecipientLine)
                }    
                $SenderDomain  = $Line.Sender_address.Split("@")[1]
                $ReportLine = [PSCustomObject]@{ 
                      Timestamp        = $Line.origin_timestamp_utc
                      Sender           = $Line.sender_address
                      Subject          = $Line.message_subject
                      Recipient        = $Line.Recipient_Status
                      RecipientInfo    = $RecipientInfo
                      Outcome          = $RecipientOutcome 
                      Bytes            = $Line.total_bytes
                      Message_id       = $Line.message_id
                      Sender_Domain    = $SenderDomain
                      Client_IP        = $Line.original_client_ip
                      Direction        = $Line.directionality
                }
                $Report.Add($ReportLine) 
            }
        }
}

$Report | Sort-Object TimeStamp | Select-Object Timestamp, Sender, Subject, Recipient | Out-GridView
$Report | Export-CSV -NoTypeInformation $CSVFile

# Split into outbound and inbound files (when we have data files containing both types of data)
$OutboundEmail = $Report | Where-Object {$_.Direction -eq 'Originating'} | Sort-Object Timestamp 
$InboundEmail = $Report | Where-Object {$_.Direction -eq 'Incoming'} | Sort-Object Timestamp 

Write-Host ("{0} records found for inbound email" -f $InboundEmail.count)
Write-Host ("{0} records found for outbound email" -f $OutboundEmail.count)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
