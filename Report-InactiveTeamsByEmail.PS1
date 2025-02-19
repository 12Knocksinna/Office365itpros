# Report-InactiveTeamsByEmail.PS1
# an example of how to report inactive teams based on their activity as recorded in audit logs and 
# the usage report.
# 
# Needs Team.ReadBasic.All, AuditQuery.Read.All, AuditLog.Read.All, Reports.Read.All, Sites.Read.All, ReportSettings.ReadWrite.All
# Would like to use SharePoint Site Usage report but site URL is not included https://learn.microsoft.com/en-gb/graph/api/reportroot-getsharepointsiteusagedetail?view=graph-rest-1.0&WT.mc_id=M365-MVP-9501
# So we have to use the audit log to find file upload and file modified events
# See https://office365itpros.com/2024/02/19/sharepoint-usage-data-issue/ 

Connect-MgGraph 

[array]$Teams = Get-MgTeam -All -PageSize 500
If (!$Teams) {
    Write-Host "No Teams found"
    Break
}

Write-Output ("Found {0} Teams" -f $Teams.Count)

# Run an audit job to find SharePoint FileUpload and FileModified events
# SharePoint events at https://learn.microsoft.com/en-us/purview/audit-log-activities?WT.mc_id=M365-MVP-9501#file-and-page-activities

Set-MgRequestContext -MaxRetry 10 -RetryDelay 15 | Out-Null
$AuditQueryName = ("Audit Job SPO Operations created at {0}" -f (Get-Date))
$StartDate = (Get-Date).AddDays(-30)
$EndDate = (Get-Date).AddDays(1)
$AuditQueryStart = (Get-Date $StartDate -format s)
$AuditQueryEnd = (Get-Date $EndDate -format s)
[array]$AuditQueryOperations = "FileModified", "FileUploaded"
$AuditQueryParameters = @{}
#$AuditQueryParameters.Add("@odata.type","#microsoft.graph.security.auditLogQuery")
$AuditQueryParameters.Add("displayName", $AuditQueryName)
$AuditQueryParameters.Add("OperationFilters", $AuditQueryOperations)
$AuditQueryParameters.Add("filterStartDateTime", $AuditQueryStart)
$AuditQueryParameters.Add("filterEndDateTime", $AuditQueryEnd)

# Submit the audit query job
$AuditJob =  New-MgBetaSecurityAuditLogQuery -BodyParameter $AuditQueryParameters

# Check the audit query job status every 20 seconds until it completes
[int]$i = 1
[int]$SleepSeconds = 20
$SearchFinished = $false; [int]$SecondsElapsed = 20
Write-Host "Checking audit query status..."
Start-Sleep -Seconds 30
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}" -f $AuditJob.Id)
$AuditQueryStatus = Invoke-MgGraphRequest -Uri $Uri -Method GET
While ($SearchFinished -eq $false) {
    $i++
    Write-Host ("Waiting for audit search to complete. Check {0} after {1} seconds. Current state {2}" -f $i, $SecondsElapsed, $AuditQueryStatus.status)
    If ($AuditQueryStatus.status -eq 'succeeded') {
        $SearchFinished = $true
    } Else {
        Start-Sleep -Seconds $SleepSeconds
        $SecondsElapsed = $SecondsElapsed + $SleepSeconds
        $AuditQueryStatus = Invoke-MgGraphRequest -Uri $Uri -Method GET
    }
}

Write-Host "Fetching audit records found by the search..."
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}/records?`$Top=999" -f $AuditJob.Id)
[array]$SearchRecords = Invoke-MgGraphRequest -Uri $Uri -Method GET

[array]$AuditRecords = $SearchRecords.value
# Paginate to fetch all available audit records
$NextLink = $SearchRecords.'@odata.NextLink'
While ($null -ne $NextLink) {
    $SearchRecords = $null
    [array]$SearchRecords = Invoke-MgGraphRequest -Uri $NextLink -Method GET 
    $AuditRecords += $SearchRecords.value
    Write-Host ("{0} audit records fetched so far..." -f $AuditRecords.count)
    $NextLink = $SearchRecords.'@odata.NextLink' 
}

Write-Host ("Total of {0} audit records found" -f $AuditRecords.count) -ForegroundColor Red

# Parse the audit records and extract information about the sites where activities occurred.
$SPOAuditInfo = [System.Collections.Generic.List[Object]]::new()
ForEach ($Record in $AuditRecords) {
    $SiteUrl = $null

    Switch ($Record.Operation) {
        "FileModified" {
            If ($Record.AuditData.SourceRelativeURL -notlike "*PreservationHoldLibrary") {
                $SiteUrl = $Record.AuditData.SiteUrl
            }
        }
        "FileUploaded" {
            $SiteUrl = $Record.AuditData.SiteUrl
        }
    }
    $SPOAuditLine = [PSCustomObject]@{
        Id          = $Record.Id
        Creation    = Get-Date $Record.CreatedDateTime -format 'dd-MMM-yyyy HH:mm:ss'
        User        = $Record.UserId
        Operation   = $Record.Operation
        SiteURL     = $SiteUrl
    }
    $SPOAuditInfo.Add($SPOAuditLine)
}

# Now get the report data
$ObfuscationChanged = $false
If ((Get-MgAdminReportSetting).DisplayConcealedNames -eq $True) {
   $Parameters = @{ displayConcealedNames = $False }
   Update-MgAdminReportSetting -BodyParameter $Parameters
   $ObfuscationChanged = $true
}

$TempFile = "C:\Temp\TeamActivityDetail.csv"
Get-MgReportTeamActivityDetail -Period 'D30' -OutFile $TempFile
$TeamsData = Import-Csv -Path $TempFile

If ($ObfuscationChanged) {
    If ((Get-MgAdminReportSetting).DisplayConcealedNames -eq $False) {
       $Parameters = @{ displayConcealedNames = $True }
       Update-MgAdminReportSetting -BodyParameter $Parameters
    }
}

# For each team, check what activities we know about
$TeamReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Team in $Teams) {
    # Get the webURL for the Team site
    $WebURL = $null
    $Uri = ("https://graph.microsoft.com/v1.0/groups/{0}/sites/root?`$select=webUrl" -f $Team.id)
    Try {
        $Data = Invoke-MgGraphRequest -Uri $Uri -Method Get -ErrorAction SilentlyContinue
    } Catch {
        Write-Host ("Unable to find SharePoint site for {0}" -f $Team.DisplayName)
        Continue
    }
    $WebURL = $Data.WebUrl + "/"
    # Get team owners
    $OwnersDisplayNames = $null
    [array]$Owners = Get-MgGroupOwner -GroupId $Team.Id | Select-Object -ExpandProperty AdditionalProperties 
    $OwnersDisplayNames = $Owners.displayName -join ", "
    # Try and find any audit records for filesuploaded and updated for the site
    [array]$FilesUploaded = $SPOAuditInfo | Where-Object { $_.SiteURL -eq $WebURL -and $_.Operation -eq 'FileUploaded' }
    [array]$FilesModified = $SPOAuditInfo | Where-Object { $_.SiteURL -eq $WebURL -and $_.Operation -eq 'FileModified' }
    [array]$UsageData = $TeamsData | Where-Object { $_.'Team Id' -eq $Team.Id }
    If ($UsageData) {
        $ActiveUsers = $UsageData.'Active Users'
        $ActiveChannels = $UsageData.'Active Channels'	    
        $ChannelMessages = $UsageData.'Channel Messages'
        $Reactions = $UsageData.'Reactions'
        $MeetingsOrganized = $UsageData.'Meetings Organized'    
        $PostMessages = $UsageData.'Post Messages'
        $ReplyMessages = $UsageData.'Reply Messages'
        $UrgentMessages = $UsageData.'Urgent Messages'
        $Mentions = $UsageData.'Mentions'
        $Guests = $UsageData.'Guests'
        $ActiveSharedChannels = $UsageData.'Active Shared Channels'
    } Else {
        $ActiveUsers = 0
        $ActiveChannels = 0
        $ChannelMessages = 0
        $Reactions = 0
        $MeetingsOrganized = 0
        $PostMessages = 0
        $ReplyMessages = 0
        $UrgentMessages = 0
        $Mentions = 0       
        $Guests = 0
        $ActiveSharedChannels = 0
    }

    If ($UsageData.'Last Activity Date') {
        $LastActiveDate = Get-Date $UsageData.'Last Activity Date' -format dd-MMM-yyyy
    } Else {
        $LastActiveDate = "Never active"
    }   

    [int]$CountOfTeamActivities = $FilesUploaded.Count + $FilesModified.Count + $ActiveUsers + $Reactions + $ChannelMessages + $PostMessages + $ReplyMessages

    # Get created date - silly that it's not returned by default when you list teams.
    $CreatedDate = (Get-MgTeam -TeamId $Team.Id | Select-Object -ExpandProperty CreatedDateTime)
    $AgeInDays = (New-TimeSpan -Start $CreatedDate -End (Get-Date)).Days

    $TeamReportLine = [PSCustomObject]@{
        DisplayName         = $Team.DisplayName
        TeamId              = $Team.Id
        Created             = Get-Date $CreatedDate -format 'dd-MMM-yyyy HH:mm'
        'Age in Days'       = $AgeInDays
        Description         = $Team.Description
        Owners              = $OwnersDisplayNames
        SiteURL             = $WebURL
        FilesUploaded       = $FilesUploaded.Count
        FilesModified       = $FilesModified.Count
        'Last active date'  = $LastActiveDate
        'Active Users'      = $ActiveUsers
        'Active Channels'   = $ActiveChannels
        'Channel Messages'  = $ChannelMessages
        Reactions           = $Reactions
        'Meetings Organized' = $MeetingsOrganized
        'Post Messages'     = $PostMessages
        'Reply Messages'    = $ReplyMessages
        'Urgent Messages'   = $UrgentMessages
        Mentions            = $Mentions
        Guests              = $Guests
        'Active Shared Channels' = $ActiveSharedChannels
        'Count of Activities' = $CountOfTeamActivities
    }
    $TeamReport.Add($TeamReportLine)
}

$HtmlTeamReport = $TeamReport | Where-Object {$_.'Count of Activities' -le 100} | `
    Select-Object DisplayName, Created, 'Age in Days', 'Last active date', Owners, 'Count of Activities', TeamId, Description, SiteURL | Sort-Object DisplayName
$HtmlTeamReport = $HtmlTeamReport | ConvertTo-Html -Fragment

$MsgFrom = 'Customer.Services@office365itpros.com'
# Define some variables used to construct the HTML content in the message body
#HTML header with styles
$HtmlHead="<html>
    <style>
    BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    H4{font-size: 8px; padding-top: 4px;}
</style>"

$MsgSubject = "Inactive Teams for your review"

$ToRecipients = @{}
$ToRecipients.Add("emailAddress",@{'address'='tony.redmond@office365itpros.com'})
[array]$MsgTo = $ToRecipients
# Customize the message 
$HtmlHeaderUser = "<h2>Administrative alert: Inactive Teams based on 30-day lookback/h2>"   
# Add some content for the message - obviously, this is very customizable and should reflect what you want to say about the data being reported
$HtmlBody = "<body> <h1>Please Check these inactive teams.</h1><p></p>"
$HtmlBody = $HtmlBody + $HtmlTeamReport
$HtmlBody = $HtmlBody + "<p>These teams are highlighted because of their lack of activity in Teams messaging and SharePoint Online. Please check to ensure that they are still needed.</p>"
$HtmlBody = $HtmlBody + "<p><h4>Generated:</strong> $(Get-Date -Format g)</h4></p>"
$HtmlMsg = "</body></html>" + $HtmlHead + $Htmlheaderuser + $HtmlBody + "<p>"
# Construct the message body
$MsgBody = @{}
$MsgBody.Add('Content', "$($HtmlMsg)")
$MsgBody.Add('ContentType','html')

$Message = @{}
$Message.Add('subject', $MsgSubject)
$Message.Add('toRecipients', $MsgTo)    
$Message.Add('body', $MsgBody)
$Params = @{}
$Params.Add('message', $Message)
$Params.Add('saveToSentItems', $true)
$Params.Add('isDeliveryReceiptRequested', $true)    

Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.