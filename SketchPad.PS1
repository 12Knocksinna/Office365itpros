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

#+------------- Application Management Policy

$PasswordCredentials1 = @{}
$PasswordCredentials1.Add("restrictForAppsCreatedAfterDateTime", [System.DateTime]::Parse("2025-01-01T00:00:00Z"))
$PasswordCredentials1.Add("restrictionType", "passwordAddition")
$PasswordCredentials1.Add("maxLifetime", $null)

$PasswordCredentials2 = @{}
$PasswordCredentials2.Add("restrictionType", "customPasswordAddition")
$PasswordCredentials2.Add("maxLifetime", $null)
$PasswordCredentials2.Add("restrictForAppsCreatedAfterDateTime", [System.DateTime]::Parse("2025-01-01T00:00:00Z"))

[array]$PasswordCredentials = $PasswordCredentials1, $PasswordCredentials2

$ApplicationCredentials = @{}
$ApplicationCredentials.Add("passwordCredentials", $PasswordCredentials)

$ApplicationPolicyParameters = @{}
$ApplicationPolicyParameters.Add("isEnabled", $True)
$ApplicationPolicyParameters.Add("applicationRestrictions", $ApplicationCredentials)  
$ApplicationPolicyParameters.Add("ServicePrincipalRestrictions", $ApplicationCredentials)

Update-MgPolicyDefaultAppManagementPolicy -BodyParameter $ApplicationPolicyParameters

$Policy = Get-MgPolicyDefaultAppManagementPolicy

$Policy.applicationRestrictions.PasswordCredentials

#RestrictForAppsCreatedAfterDateTime RestrictionType        State
#----------------------------------- ---------------        -----
#01/01/2025 00:00:00                 passwordAddition       enabled
#01/01/2025 00:00:00                 customPasswordAddition enabled




$params = @{
	displayName = "Credential management policy"
	description = "Cred policy sample"
	isEnabled = $true
	restrictions = @{
		passwordCredentials = @(
			@{
				restrictionType = "passwordAddition"
				state = "enabled"
				maxLifetime = $null
				restrictForAppsCreatedAfterDateTime = [System.DateTime]::Parse("2025-04-01T10:37:00Z")
			}
			@{
				restrictionType = "passwordLifetime"
				state = "enabled"
				maxLifetime = "P90D"
				restrictForAppsCreatedAfterDateTime = [System.DateTime]::Parse("2025-03-01T00:00:00Z")
			}
			@{
				restrictionType = "symmetricKeyAddition"
				state = "enabled"
				maxLifetime = $null
				restrictForAppsCreatedAfterDateTime = [System.DateTime]::Parse("2019-10-19T10:37:00Z")
			}
			@{
				restrictionType = "symmetricKeyLifetime"
				state = "enabled"
				maxLifetime = "P90D"
				restrictForAppsCreatedAfterDateTime = [System.DateTime]::Parse("2014-10-19T10:37:00Z")
			}
		)
		keyCredentials = @(
		)
	}
}


$AppPolicyParameters = @{
    displayName = "Restrict App Secrets to 180 days"
	description = "This policy allows apps to have app secrets lasting for up to 180 days"
	isEnabled = $true
    restrictions = @{
		passwordCredentials = @(
			@{
				restrictionType = "passwordLifeTime"
				state = "enabled"
				maxLifetime = 'P180D'
				restrictForAppsCreatedAfterDateTime = [System.DateTime]::Parse("2025-01-01T00:00:00Z")
			}
            @{
                restrictionType = "passwordAddition"
                state = "disabled"
                maxLifetime = $null
                restrictForAppsCreatedAfterDateTime = [System.DateTime]::Parse("2025-01-01T00:00:00Z")
            }       
        )
    }
}


# Convert a PowerShell timespan to ISO8601 duration
Function Convert-TimeSpanToISO8601 {
    param (
        [Parameter(Mandatory=$true)]
        [TimeSpan]$TimeSpan
    )
    
    $duration = "P"
    if ($TimeSpan.Days -gt 0) {
        $duration += "$($TimeSpan.Days)D"
    }
    if ($TimeSpan.Hours -gt 0 -or $TimeSpan.Minutes -gt 0 -or $TimeSpan.Seconds -gt 0) {
        $duration += "T"
        if ($TimeSpan.Hours -gt 0) {
            $duration += "$($TimeSpan.Hours)H"
        }
        if ($TimeSpan.Minutes -gt 0) {
            $duration += "$($TimeSpan.Minutes)M"
        }
        if ($TimeSpan.Seconds -gt 0) {
            $duration += "$($TimeSpan.Seconds)S"
        }
    }
    return $duration
}

# Example usage
$timespan = New-TimeSpan -Days 1 -Hours 2 -Minutes 30 -Seconds 45
$iso8601Duration = Convert-TimeSpanToISO8601 -TimeSpan $timespan
Write-Output $iso8601Duration

# ------------------------ AuditLOgQuery Searches


$AuditJobName = ("SharePoint Audit job created at {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm'))
$EndDate = (Get-Date).AddHours(1)
$StartDate = (Get-Date $EndDate).AddDays(-180)
$AuditQueryStart = (Get-Date $StartDate -format s)
$AuditQueryEnd = (Get-Date $EndDate -format s)
[array]$AuditOperationFilters = "FileModified", "FileDeleted", "FileUploaded"
[array]$AuditobjectIdFilters = "https://redmondassociates.sharepoint.com/sites/blogsandprojects/*", "https://redmondassociates.sharepoint.com/sites/Office365Adoption/*"
[array]$AuditAdministrativeUnitIdFilters = "112f5e71-b430-4c83-945b-8b665c14ff25" -as [string]
[array]$AuditUserPrincipalNameFilters = "Ken.Bowers@office365itpros.com", "Lotte.Vetler@office365itpros.com", "tony.redmond@redmondassociates.org"

$AuditQueryParameters = @{}
$AuditQueryParameters.Add("@odata.type","#microsoft.graph.security.auditLogQuery")
$AuditQueryParameters.Add("displayName", $AuditJobName)
$AuditQueryParameters.Add("OperationFilters", $AuditOperationFilters)
$AuditQueryParameters.Add("filterStartDateTime", $AuditQueryStart)
$AuditQueryParameters.Add("filterEndDateTime", $AuditQueryEnd)
$AuditQueryParameters.Add("userPrincipalNameFilters", $AuditUserPrincipalNameFilters)
$AuditQueryParameters.Add("objectIdFilters", $AuditobjectIdFilters)
# $AuditQueryParameters.Add("administrativeUnitIdFilters", $AuditAdministrativeUnitIdFilters)

$Uri = "https://graph.microsoft.com/beta/security/auditLog/queries"
$AuditJob = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $AuditQueryParameters


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
$AuditRecords = [System.Collections.Generic.List[string]]::new()
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}/records?`$top=999" -f $AuditJob.Id)
[array]$AuditSearchRecords = Invoke-MgGraphRequest -Uri $Uri -Method GET
[array]$AuditRecords = $AuditSearchRecords.value

$NextLink = $AuditSearchRecords.'@Odata.NextLink'
While ($null -ne $NextLink) {
    $AuditSearchRecords = $null
    [array]$AuditSearchRecords = Invoke-MgGraphRequest -Uri $NextLink -Method GET 
    $AuditRecords += $AuditSearchRecords.value
    Write-Host ("{0} audit records fetched so far..." -f $AuditRecords.count)
    $NextLink = $AuditSearchRecords.'@odata.NextLink' 
}

Write-Host ("Audit query {0} returned {1} records" -f $AuditJobName, $AuditRecords.Count)
$AuditRecords = $AuditRecords | Sort-Object CreatedDateTime -Descending


$Uri = "https://graph.microsoft.com/beta/security/auditLog/queries"
$Data = Invoke-MgGraphRequest -Uri $Uri -Method GET
If ($Data) {
    Write-Output "Audit Jobs found"
    $Data.Value | ForEach-Object {
        Write-Host ("{0} {1}" -f $_.id, $_.displayName)
    }
} Else {
    Write-Output "No audit jobs found"
}   

# Full filter

$AuditJobName = ("Full audit job created at {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm'))
$EndDate = (Get-Date).AddHours(1)
$StartDate = (Get-Date $EndDate).AddDays(-180)
$AuditQueryStart = (Get-Date $StartDate -format s)
$AuditQueryEnd = (Get-Date $EndDate -format s)

$AuditQueryParameters = @{}
$AuditQueryParameters.Add("@odata.type","#microsoft.graph.security.auditLogQuery")
$AuditQueryParameters.Add("displayName", $AuditJobName)
$AuditQueryParameters.Add("filterStartDateTime", $AuditQueryStart)
$AuditQueryParameters.Add("filterEndDateTime", $AuditQueryEnd)


$Uri = "https://graph.microsoft.com/beta/security/auditLog/queries"
$AuditJob = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $AuditQueryParameters