# SketchPad.PS1
# Used for bits of PowerShell I am working on or want to note

# Get list to update metadata for the new item
$ListId = (Get-MgSiteList -SiteId $Site.Id -Filter "DisplayName eq 'Documents'").Id
[array]$ListItems = Get-MgSiteListItem -SiteId $Site.Id -ListId $ListId
$ListItem = $ListItems[-1]
$Body = @{}
$Body.Add("Title", "Hard Deleted Users Report Created by Azure Automation")
$Status = Update-MgSiteListItemField -SiteId $site.Id -ListId $listId -ListItemId $listItem.Id -BodyParameter $Body
If ($Status) {
    Write-Output ("Updated document metadata for item {0} with title {1}" -f $ListItem.Id, $Params.Title)
}




# Report all OneDrive accounts
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
    -ConsistencyLevel eventual -CountVariable UsersFound -All -PageSize 500
If (!$Users) {
    Write-Host "No user accounts found"
    Break
}
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    Try {
        $OneDrive = Get-MgUserDefaultDrive -UserId $User.Id -ErrorAction Stop
    } Catch {
        Write-Host ("Unable to find OneDrive for {0}" -f $User.UserPrincipalName)
        Continue
    }   
    $ReportLine = [PSCustomObject][Ordered]@{
        UserPrincipalName = $User.UserPrincipalName
        OneDriveUrl       = $OneDrive.WebUrl
        Created           = Get-Date $OneDrive.CreatedDateTime -format 'dd-MMM-yyyy HH:mm'
        Modified          = Get-Date $OneDrive.LastModifiedDateTime -format 'dd-MMM-yyyy HH:mm'
    }
    $Report.Add($ReportLine)
}

# --- Add multiple members from a Microsoft 365 Group to another group

$SourceGroup = Get-MgGroup -Filter "DisplayName eq 'Bala Group'"
$TargetGroup = Get-MgGroup -Filter "DisplayName eq 'Bedson Project'"
[array]$MembersSourceGroup = Get-MgGroupMember -GroupId $SourceGroup.Id -All | Select-Object -ExpandProperty Id
[array]$MembersTargetGroup = Get-MgGroupMember -GroupId $TargetGroup.Id -All | Select-Object -ExpandProperty Id
# Remove source members who are already members of the target group
$MembersSourceGroup = $MembersSourceGroup | Where-Object { $MembersTargetGroup -notcontains $_ }
$Data = [System.Collections.Generic.List[Object]]::new()
$MembersSourceGroup | ForEach-Object {$Data.Add("https://graph.microsoft.com/beta/directoryobjects/{0}" -f $_)}
While ($Data.count -ne 0) {
    $Parameters = @{"members@odata.bind" = $Data[0..19] }
    Update-MgGroup -GroupId $TargetGroup.Id -BodyParameter $Parameters
    If ($Data.count -gt 20) {
        $Data.RemoveRange(0.20)
    } Else {
        $Data.RemoveRange(0,$Data.count)
    }
}

$SelectedUsers = Get-MgUser -Filter "userType eq 'Member'"
$MsgFrom = 'Customer.Services@office365itpros.com'
# Define some variables used to construct the HTML content in the message body
# HTML header with styles
$HtmlHead="<html>
    <style>
    BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    H4{font-size: 8px; padding-top: 4px;}
</style>"

$HtmlBody = $null
$HtmlBody = $HtmlBody + "<body> <h1>Users</h1><p></p>"   

$HtmlBody = $HtmlBody + ($SelectedUsers| Sort-Object DisplayName | ConvertTo-HTML -Fragment -As Table -PreContent "<h2>Administrative alert: Inactive Teams based on 30-day lookback</h2>")
$HtmlBody = $HtmlBody + "<p>These users are member accounts</p>"
$HtmlBody = $HtmlBody + "<p><h4>Generated:</strong> $(Get-Date -Format 'dd-MMM-yyyy HH:mm')</h4></p>"

$HtmlMsg = $HtmlHead + $HtmlBody + "<p></body>"

$MsgSubject = "Member users"

$ToRecipients = @{}
$ToRecipients.Add("emailAddress", @{"address"="tony.redmond@office365itpros.com"} )
[array]$MsgTo = $ToRecipients

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


#-----------

$UPN = (Get-MgContext).Account
$StartTime = (Get-Date).AddDays(1).ToString("yyyy-MM-ddT00:00:00Z")
$EndTime = (Get-Date).AddDays(7).ToString("yyyy-MM-ddT00:00:00Z")

$ScheduledStartDateTime = @{}
$ScheduledStartDateTime.Add("dateTime", $StartTime)
$ScheduledStartDateTime.Add("timeZone", "UTC")
$ScheduledEndDateTime = @{}
$ScheduledEndDateTime.Add("dateTime", $EndTime)
$ScheduledEndDateTime.Add("timeZone", "UTC")    

$AutomaticRepliesSetting = @{}
$AutomaticRepliesSetting.Add("status", "alwaysEnabled")
$AutomaticRepliesSetting.Add("externalAudience", "all")
$AutomaticRepliesSetting.Add("scheduledEndDateTime", $ScheduledEndDateTime)
$AutomaticRepliesSetting.Add("scheduledStartDateTime", $ScheduledStartDateTime)
$AutomaticRepliesSetting.Add("internalReplyMessage", "I am out of the office until next week")
$AutomaticRepliesSetting.Add("externalReplyMessage", "I am out of the office until next week")

$AutoReply = @{}
$AutoReply.Add("@odata.context", "https://graph.microsoft.com/v1.0/$UPN/mailboxSettings")
$AutoReply.Add("automaticRepliesSetting", $AutomaticRepliesSetting)

Update-MgUserMailboxSetting -UserId $UPN -BodyParameter $AutoReply

$params = @{
	"@odata.context" = "https://graph.microsoft.com/v1.0/$metadata#Me/mailboxSettings"
	automaticRepliesSetting = @{
		status = "Scheduled"
		scheduledStartDateTime = @{
			dateTime = "2026-03-20T18:00:00.0000000"
			timeZone = "UTC"
		}
		scheduledEndDateTime = @{
			dateTime = "2026-03-28T18:00:00.0000000"
			timeZone = "UTC"
		}
        externalReplyMessage = "I am out of the office until next week"
        internalReplyMessage = "I am out of the office until next week"
        externalAudience = "all"
	}
}

#+------------- Copilot Interactions

Set-MgRequestContext -MaxRetry 10 -RetryDelay 15| Out-Null
$AuditQueryName = ("Copilot Interactions audit job created at {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm'))
$StartDate = (Get-Date).AddDays(-30)
$EndDate = (Get-Date).AddDays(1)
$AuditQueryStart = (Get-Date $StartDate -format s)
$AuditQueryEnd = (Get-Date $EndDate -format s)
[array]$AuditQueryOperations = "CopilotInteraction"
$AuditQueryParameters = @{}
#$AuditQueryParameters.Add("@odata.type","#microsoft.graph.security.auditLogQuery")
$AuditQueryParameters.Add("displayName", $AuditQueryName)
$AuditQueryParameters.Add("OperationFilters", $AuditQueryOperations)
$AuditQueryParameters.Add("filterStartDateTime", $AuditQueryStart)
$AuditQueryParameters.Add("filterEndDateTime", $AuditQueryEnd)

# Submit the audit query
$AuditJob = New-MgBetaSecurityAuditLogQuery -BodyParameter $AuditQueryParameters

# Check the audit query status every 20 seconds until it completes
[int]$i = 1
[int]$SleepSeconds = 20
$SearchFinished = $false; [int]$SecondsElapsed = 20
Write-Host "Checking audit query status..."
Start-Sleep -Seconds 30
# This cmdlet is not working...
#$AuditQueryStatus = Get-MgBetaSecurityAuditLogQuery -AuditLogQueryId $AuditJob.Id
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}" -f $AuditJob.id)
$AuditQueryStatus = Invoke-MgGraphRequest -Uri $Uri -Method Get

While ($SearchFinished -eq $false) {
    $i++
    Write-Host ("Waiting for audit search to complete. Check {0} after {1} seconds. Current state {2}" -f $i, $SecondsElapsed, $AuditQueryStatus.status)
    If ($AuditQueryStatus.status -eq 'succeeded') {
        $SearchFinished = $true
    } Else {
        Start-Sleep -Seconds $SleepSeconds
        $SecondsElapsed = $SecondsElapsed + $SleepSeconds
        # $AuditQueryStatus = Get-MgBetaSecurityAuditLogQuery -AuditLogQueryId $AuditJob.Id
        $AuditQueryStatus = Invoke-MgGraphRequest -Uri $Uri -Method Get
    }
}

# Fetch the audit records returned by the query
# This cmdlet isn't working either
# [array]$AuditRecords = Get-MgBetaSecurityAuditLogQueryRecord -AuditLogQueryId $AuditJob.Id -All -PageSize 999
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}/records" -f $AuditJob.Id)
[array]$SearchRecords = Invoke-MgGraphRequest -Uri $Uri -Method GET
[array]$AuditRecords = $SearchRecords.value

$NextLink = $SearchRecords.'@Odata.NextLink'
While ($null -ne $NextLink) {
    $SearchRecords = $null
    [array]$SearchRecords = Invoke-MgGraphRequest -Uri $NextLink -Method GET 
    $AuditRecords += $SearchRecords.value
    Write-Host ("{0} audit records fetched so far..." -f $AuditRecords.count)
    $NextLink = $SearchRecords.'@odata.NextLink' 
}

Write-Host ("Audit query {0} returned {1} records" -f $AuditQueryName, $AuditRecords.Count)`
$AuditRecords = $AuditRecords | Sort-Object CreatedDateTime -Descending

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $AuditRecords) {
    $AuditData = $Rec.AuditData
    $CopilotApp = 'Copilot for Microsoft 365'; $Context = $null; $CopilotLocation = $null

    $CopilotApp = $AuditData.copiloteventdata.apphost
    If ($null -eq $CopilotApp) {  
        Switch ($Auditdata.copiloteventdata.contexts.type) {
            "xlsx" {
                $CopilotApp = "Excel"
            }
            "docx" {
                $CopilotApp = "Word"
            }
            "pptx" {
                $CopilotApp = "PowerPoint"
            }
            "TeamsMeeting" {
                $CopilotApp = "Teams"
                $CopilotLocation = "Teams meeting"
            }
            "StreamVideo" {
                $CopilotApp = "Stream"
                $CopilotLocation = "Stream video player"
            }
            Default {
                $CopilotApp = "Copilot for Microsoft 365"
                $CopilotLocation = "Unknown"
            }
        }
    }

    If ($Auditdata.copiloteventdata.contexts.id -like "*https://teams.microsoft.com/*") {
        $CopilotApp = "Teams"
    } ElseIf ($AuditData.CopiloteventData.AppHost -eq "bizchat" -or $AuditData.CopiloteventData.AppHost -eq "Office") {
        $CopilotApp = "Copilot for Microsoft 365 Chat"
        $CopilotLocation = "Chat"
    }

    If ($Auditdata.CopilotEventData.contexts.id) {
        $Context = $Auditdata.CopilotEventData.contexts.id
    } ElseIf ($Auditdata.CopilotEventData.threadid) {
        $Context = $Auditdata.CopilotEventData.threadid
       # $CopilotApp = "Teams"
    }

    If ($Auditdata.copiloteventdata.contexts.id -like "*/sites/*") {
        $CopilotLocation = "SharePoint Online"
    } ElseIf ($Auditdata.copiloteventdata.contexts.id -like "*https://teams.microsoft.com/*") {
        $CopilotLocation = "Teams"
        If ($Auditdata.copiloteventdata.contexts.id -like "*ctx=channel*") {
            $CopilotLocation = "Teams Channel"
        } Else {
            $CopilotLocation = "Teams Chat"
        }
    } ElseIf ($Auditdata.copiloteventdata.contexts.id -like "*/personal/*") {
        $CopilotLocation = "OneDrive for Business"
    } 
    If ($CopilotApp -eq "Outlook") {
        $CopilotLocation = "Mailbox"
    }
    If ($CopilotApp -eq 'Loop') {
        $CopilotLocation = "SharePoint Embedded"
    }
       
    # Make sure that we report the resources used by Copilot and the action (like read) used to access the resource
    [array]$AccessedResources = $AuditData.copiloteventdata.accessedResources.name | Sort-Object -Unique
    [string]$AccessedResources = $AccessedResources -join ", "
    [array]$AccessedResourceLocations = $AuditData.copiloteventdata.accessedResources.id | Sort-Object -Unique
    [string]$AccessedResourceLocations = $AccessedResourceLocations -join ", "
    [array]$AccessedResourceActions = $AuditData.copiloteventdata.accessedResources.action | Sort-Object -Unique
    [string]$AccessedResourceActions = $AccessedResourceActions -join ", "

    $ReportLine = [PSCustomObject][Ordered]@{
        TimeStamp                       = (Get-Date $Rec.CreatedDateTime -format "dd-MMM-yyyy HH:mm:ss")
        User                            = $Rec.userPrincipalName
        UserId                          = $Rec.userId
        App                             = $CopilotApp
        'Resource Location'             = $CopilotLocation 
        'App context'                   = $Context   
        'Accessed Resources'            = $AccessedResources
        'Accessed Resource Locations'   = $AccessedResourceLocations
        Action                          = $AccessedResourceActions
    }
    $Report.Add($ReportLine)
}

[array]$UsersOfCopilot = $Report | Sort-Object UserId -Unique | Select-Object User, UserId
$CopilotUserData =  [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $UsersOfCopilot) {
    $CopilotChatInteractions = $Report | Where-Object {($_.UserId -eq $User.UserId) -and ($_.App -eq 'Copilot for Microsoft 365 Chat')} | Measure-Object | Select-Object -ExpandProperty Count
    $CopilotExcelInteractions = $Report | Where-Object {($_.UserId -eq $User.UserId) -and ($_.App -eq 'Excel')} | Measure-Object | Select-Object -ExpandProperty Count
    $CopilotLoopInteractions = $Report | Where-Object {($_.UserId -eq $User.UserId) -and ($_.App -eq 'Loop')} | Measure-Object | Select-Object -ExpandProperty Count
    $CopilotOneNoteInteractions = $Report | Where-Object {($_.UserId -eq $User.UserId) -and ($_.App -eq 'OneNote')} | Measure-Object | Select-Object -ExpandProperty Count
    $CopilotOutlookInteractions = $Report | Where-Object {($_.UserId -eq $User.UserId) -and ($_.App -eq 'Outlook')} | Measure-Object | Select-Object -ExpandProperty Count
    $CopilotPowerPointInteractions = $Report | Where-Object {($_.UserId -eq $User.UserId) -and ($_.App -eq 'PowerPoint')} | Measure-Object | Select-Object -ExpandProperty Count
    $CopilotWordInteractions = $Report | Where-Object {($_.UserId -eq $User.UserId) -and ($_.App -eq 'Word')} | Measure-Object | Select-Object -ExpandProperty Count
    $TotalInteractions = $Report | Where-Object {$_.UserId -eq $User.UserId} | Measure-Object | Select-Object -ExpandProperty Count
    $ReportLine = [PSCustomObject][Ordered]@{
        User                    = $User.User
        UserId                  = $User.UserId
        ChatInteractions        = $CopilotChatInteractions
        ExcelInteractions       = $CopilotExcelInteractions
        LoopInteractions        = $CopilotLoopInteractions
        OutlookInteractions     = $CopilotOutlookInteractions
        PowerPointInteractions  = $CopilotPowerPointInteractions
        WordInteractions        = $CopilotWordInteractions
        OtherInteractions       = $TotalInteractions - ($CopilotChatInteractions + $CopilotExcelInteractions + $CopilotLoopInteractions + $CopilotOneNoteInteractions + $CopilotOutlookInteractions + $CopilotPowerPointInteractions + $CopilotWordInteractions)
        TotalInteractions       = $TotalInteractions
    }
    $CopilotUserData.Add($ReportLine)
}

