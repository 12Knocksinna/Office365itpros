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

#----------- HTML header

$ReportTitle = "Audit Log Report"
$DateRun = Get-Date -Format "dd-MMM-yyyy HH:mm"

$HtmlHeader = @"
<html>
<head>
    <style>
        body { font-family: Arial; font-size: 10pt; }
        h1 { background-color: blue; color: white; padding: 10px; }
        h2 { font-size: 18px; padding-top: 10px; }
        h3 { font-size: 16px; padding-top: 8px; }
        h4 { font-size: 8px; padding-top: 4px; }
    </style>
</head>
<body>
    <h1>$ReportTitle</h1>
    <p>Date Run: $DateRun</p>
"@

# ---


[array]$Mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.PrimarySmtpAddress.Split('@')[1] -notin $Domains }

[array]$Domains = Get-AcceptedDomain 
$PrimaryDomain = $Domains | Where-Object { $_.Default -eq $true } | Select-Object -ExpandProperty DomainName
[array]$Domains = $Domains | Select-Object -ExpandProperty DomainName


Write-Host "Checking mailboxes..."
[array]$Mailboxes = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox, discoveryMailbox
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Mailbox in $Mailboxes) {
    $ExternalAddresses = $Mailbox.EmailAddresses | Where-Object { $_ -like "SMTP:*" -and ($_.Split(':')[1].Split('@')[1] -notin $Domains) }
    If ($ExternalAddresses) {
        $ReportLine = [PSCustomObject][Ordered]@{
            DisplayName             = $Mailbox.DisplayName
            PrimarySmtpAddress      = $Mailbox.PrimarySmtpAddress
            EmailAddresses          = $ExternalAddresses -join ", "
            Type                    = "mailbox"
            Identity                = $Mailbox.Alias
        }
        $Report.Add($ReportLine)
    }
}

Write-Host "Checking Microsoft 365 Groups..."
[array]$Groups = Get-UnifiedGroup -ResultSize Unlimited

ForEach ($Group in $Groups) {
    $ExternalAddresses = $Group.EmailAddresses | Where-Object { $_ -like "SMTP:*" -and ($_.Split(':')[1].Split('@')[1] -notin $Domains) }
    If ($ExternalAddresses) {
        $ReportLine = [PSCustomObject][Ordered]@{
            DisplayName             = $Group.DisplayName
            PrimarySmtpAddress      = $Group.PrimarySmtpAddress
            EmailAddresses  = $ExternalAddresses -join ", "
            Type                    = "group"
            Identity                = $Group.Alias
        }
        $Report.Add($ReportLine)
    }
}

Write-Host "Checking Distribution Lists..."
[array]$DLs = Get-DistributionGroup -ResultSize Unlimited

ForEach ($DL in $DLs) {
    $ExternalAddresses = $DL.EmailAddresses | Where-Object { $_ -like "SMTP:*" -and ($_.Split(':')[1].Split('@')[1] -notin $Domains) }
    If ($ExternalAddresses) {
        $ReportLine = [PSCustomObject][Ordered]@{
            DisplayName             = $DL.DisplayName
            PrimarySmtpAddress      = $DL.PrimarySmtpAddress
            EmailAddresses          = $ExternalAddresses -join ", "
            Type                    = "dl"
            Identity                = $DL.Alias
        }
        $Report.Add($ReportLine)
    }
}

Write-Host "Checking Dynamic distribution groups..."
[array]$DDLs = Get-DynamicDistributionGroup -ResultSize Unlimited

ForEach ($DDL in $DDLs) {
    $ExternalAddresses = $DDL.EmailAddresses | Where-Object { $_ -like "SMTP:*" -and ($_.Split(':')[1].Split('@')[1] -notin $Domains) }
    If ($ExternalAddresses) {
        $ReportLine = [PSCustomObject][Ordered]@{
            DisplayName             = $DDL.DisplayName
            PrimarySmtpAddress      = $DDL.PrimarySmtpAddress
            EmailAddresses          = $ExternalAddresses -join ", "
            Type                    = "ddl"
            Identity                = $DDL.Alias
        }
        $Report.Add($ReportLine)
    }
}

Write-Host ("{0} mailboxes, {1} groups, {2} distribution lists, and {3} dynamic distribution lists checked" -f $Mailboxes.Count, $Groups.Count, $DLs.Count, $DDLs.Count)
Write-Host ("Problems found in {0} objects" -f $Report.Count)

$Report | Format-Table -AutoSize

ForEach ($Object in $Report) {

    $UpdatePrimary = $false
    $NewPrimarySmtpAddress = $null

    # Check if primary SMTP address needs to be updated
    If ($Object.PrimarySmtpAddress.Split('@')[1] -notin $Domains) {
        Write-Host ("Primary SMTP address must be updated from {0}" -f $Object.PrimarySmtpAddress)
        $NewPrimarySmtpAddress = ("{0}@{1}" -f $Object.PrimarySmtpAddress.Split('@')[0], $PrimaryDomain)
        $UpdatePrimary = $true
    }

    If ($UpdatePrimary) {
        Write-Host ("Setting new primary SMTP address {0}" -f $NewPrimarySmtpAddress)
        Switch ($Object.Type) {
            "mailbox" {
                Set-Mailbox -Identity $Object.Identity -EmailAddresses @{Remove=$Object.PrimarySmtpAddress; Add=$NewPrimarySmtpAddress} -ErrorAction SilentlyContinue
                Set-Mailbox -Identity $Object.Identity -WindowsEmailAddress $NewPrimarySmtpAddress -ErrorAction SilentlyContinue
            }   
            "group" {
                Set-UnifiedGroup -Identity $Object.Identity -PrimarySmtpAddress $NewPrimarySmtpAddress -ErrorAction SilentlyContinue
            }
            "dl" {
                Set-DistributionGroup -Identity $Object.Identity -PrimarySmtpAddress $NewPrimarySmtpAddress -ErrorAction SilentlyContinue
            }
            "ddl" {
                Set-DynamicDistributionGroup -Identity $Object.Identity -PrimarySmtpAddress $NewPrimarySmtpAddress -ErrorAction SilentlyContinue
            }
        }
    }

    [array]$EmailAddresses = $Object.EmailAddresses -split ", "
    ForEach ($Address in $EmailAddresses) {
        If ($Address.Split('@')[1] -notin $Domains) {
            $AddressToRemove = $Address.Split(':')[1]
            Write-Host ("Removing address {0} from {1}" -f $Address, $Object.DisplayName)
            Switch ($Object.Type) {
            "mailbox" {
                Set-Mailbox -Identity $Object.Identity -EmailAddresses @{Remove=$AddressToRemove} -ErrorAction SilentlyContinue   
            }
            "group" {
                Set-UnifiedGroup -Identity $Object.Identity -EmailAddresses @{Remove=$AddressToRemove} -ErrorAction SilentlyContinue 
            }
             "dl" {
                Set-DistributionGroup -Identity $Object.Identity -EmailAddresses @{Remove=$AddressToRemove} -ErrorAction SilentlyContinue 
            }
             "ddl" {
                Set-DynamicDistributionGroup -Identity $Object.Identity -EmailAddresses @{Remove=$AddressToRemove} -ErrorAction SilentlyContinue 
            }
            }          
        }
    }

}


[array]$ExoTags = Get-RetentionPolicyTag 
[array]$M365Tags = Get-ComplianceTag

$RetentionTagsHash = @{}
ForEach ($Tag in $ExoTags) {
    $RetentionTagsHash.Add([string]$Tag.Guid, $Tag.Name)
}
ForEach ($Tag in $M365Tags) {
    $RetentionTagsHash.Add([string]$Tag.Guid, $Tag.Name)
}

Write-Host "Looking for audit records..."
[array]$Records = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) -Operations ApplyPriorityCleanup -ResultSize 5000 -Formatted

If ($Records.Count -eq 0) {
    Write-Host "No audit records found for ApplyPriorityCleanup operations"
    Break
} Else {
    $Records = $Records | Sort-Object Identity -Unique | Sort-Object { $_.CreationDate -as [datetime]} -Descending
    Write-Host ("Processing {0} audit records..." -f $Records.Count)
}

$PriorityCleanupReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
    $LabelApplied = $null; $LabelID = $null; $LabelRemoved = $null; [string]$TimeStamp = $null
    $AuditData = $Rec.AuditData | ConvertFrom-Json
    $LabelApplied = $AuditData.OperationProperties | Where-Object {$_.Name -eq 'TagName'} | Select-Object -ExpandProperty Value
    $LabelId = $AuditData.OperationProperties | Where-Object {$_.Name -eq 'TagId'} | Select-Object -ExpandProperty Value
    $LabelRemoved = $AuditData.OperationProperties | Where-Object {$_.Name -eq 'TagReplacedByPriorityCleanup'} | Select-Object -ExpandProperty Value
    $TimeStamp = Get-Date ($AuditData.CreationTime) -format 'dd-MMM-yyyy HH:mm'
    $ReportLine = [PSCustomObject][Ordered]@{
        TimeStamp       = $TimeStamp
        User            = $AuditData.UserId
        Action          = $AuditData.Operation
        Mailbox         = $AuditData.MailboxOwnerUPN
        Item            = $AuditData.Item.Subject
        'Label Applied' = $LabelApplied
        'Label Id'      = $LabelId
        'Label Removed' = $RetentionTagsHash[$LabelRemoved]  
    }
    $PriorityCleanupReport.Add($ReportLine)
}

$PriorityCleanupReport | Group-Object Mailbox -NoElement | Sort-Object Count -Descending | Format-Table Name, Count

$UserId = (Get-MgUser -UserId (Get-MgContext).Account).Id
# Create simple calendar appointment
$EventBody = @{}
$EventBody.Add("contentType", "HTML")
$EventBody.Add("content", "The TEC 2025 comference event starts with registration and breakfast at 8:30AM. The first session will commence at 9:30AM")

$EventStart = @{}
$EventStart.Add("dateTime", "2025-09-30T09:00:00")
$EventStart.Add("timeZone", "Central Standard Time")

$EventEnd = @{}
$EventEnd.Add("dateTime", "2025-10-01T17:00:00")
$EventEnd.Add("timeZone", "Central Standard Time")

$EventLocation = @{}
$EventLocation.Add("displayName", "Minneapolis")

$EventDetails = @{}
$EventDetails.Add("subject", "The Experts Conference 2025")
$EventDetails.Add("body", $EventBody)
$EventDetails.Add("start", $EventStart)
$EventDetails.Add("end", $EventEnd)
$EventDetails.Add("location", $EventLocation)
$EventDetails.Add("allowNewTimeProposals", $true)
$EventDetails.Add("transactionId", (New-Guid))

# hash table for attendees
$EventAttendees = @()

# Each attendde defined as email address and name
$Participant1 = @{}
$Participant1.add("address","lotte.vetler@office365itpros.com")
$Participant1.add("name", "Lotte Vetler")

$Participant2 = @{}
$Participant2.add("address","otto.flick@office365itpros.com")
$Participant2.add("name", "Otto.Flick")

$Participant3 = @{}
$Participant3.add("address","kim.akers@office365itpros.com")
$Participant3.add("name", "Kim Akers")

$EventAttendee1 = @{}
$EventAttendee1.add("emailaddress", $Participant1)
$EventAttendee1.Add("type", "required")

$EventAttendee2 = @{}
$EventAttendee2.add("emailaddress", $Participant2)
$EventAttendee2.Add("type", "optional")

$EventAttendee3 = @{}
$EventAttendee3.add("emailaddress", $Participant3)
$EventAttendee3.Add("type", "optional")

$EventAttendees = $EventAttendee1, $EventAttendee2, $EventAttendee3

$EventDetails.Add("attendees", $EventAttendees)

$Uri =("https://graph.microsoft.com/v1.0/users/{0}/calendar/events" -f $userId)
$NewEvent = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $EventDetails

$UpdateEventDetails = @{}
$UpdateEventDetails.Add("isOnlineMeeting", $true)
$UpdateEventDetails.Add("onlineMeetingProvider", "teamsForBusiness")
$UpdateEventDetails.Add("isReminderOn", $true)
$UpdateEventDetails.Add("reminderMinutesBeforeStart", 30)

$Uri = ("https://graph.microsoft.com/v1.0/{0}/events/{1}" -f $UserId, $NewEvent.Id)
$UpdatedEvent = Invoke-MgGraphRequest -Uri $Uri -Method PATCH -Body $UpdateEventDetails
    
$NewEvent = Update-MgUserEvent -UserId $userId -EventId $NewEvent.Id -BodyParameter $UpdateEventDetails

Update-MgUserEvent -Userid $Userid -Eventid $NewEvent.Id -IsOnlineMeeting:$true -Importance High -OnlineMeetingProvider 'TeamsforBusiness' -ReminderMinutesBeforeStart 30

# Update with attendees - rewrite attendee list

$Participant1 = @{}
$Participant1.Add("address","James.Ryan@office365itpros.com")
$Participant1.Add("name", "James Ryan")

$Attendee1 = @{}
$Attendee1.Add("type","required")
$Attendee1.Add("Emailaddress", $Participant1)

[array]$Participants = $Attendee1
$EventDetails = @{}
$EventDetails.Add("attendees", $Participants)


# New recurring event
[array]$DaysOfWeek = "Tuesday"
$RecurringPattern = @{}
$RecurringPattern.Add("type", "weekly")
$RecurringPattern.Add("interval", 1)
$RecurringPattern.Add("daysOfWeek",  $DaysOfWeek)
$RecurringPattern.Add("firstDayOfWeek", "monday")

$RecurringRange = @{}
$RecurringRange.Add("startdate", "2025-04-15T09:00:00")
$RecurringRange.Add("enddate", "2025-04-15T09:00:00")
$RecurringRange.Add("recurrenceRangeType", "endDate")

$RecurrenceRange = @{}
$RecurrenceRange.Add("pattern", $RecurringPattern)
$RecurrenceRange.Add("range", $RecurringRange)

$EventDetails = @{}
$EventDetails.Add("recurrence", $RecurrenceRange)

$EventBody = @{}
$EventBody.Add("contentType", "HTML")
$EventBody.Add("content", "Weekly update meeting")

$EventStart = @{}
$EventStart.Add("dateTime", "2025-04-15T09:00:00")
$EventStart.Add("timeZone", "UTC")

$EventEnd = @{}
$EventEnd.Add("dateTime", "2025-04-15T09:30:00")
$EventEnd.Add("timeZone", "UTC")

$EventLocation = @{}
$EventLocation.Add("displayName", "Royal Garden Hotel, London")

$EventDetails.Add("subject", "TEC Roadshow")
$EventDetails.Add("body", $EventBody)
$EventDetails.Add("start", $EventStart)
$EventDetails.Add("end", $EventEnd)
$EventDetails.Add("location", $EventLocation)
$EventDetails.Add("allowNewTimeProposals", $true)
$EventDetails.Add("transactionId", (New-Guid))

# hash table for attendees
$EventAttendees = @()

# Each attendde defined as email address and name
$Participant1 = @{}
$Participant1.add("address","lotte.vetler@office365itpros.com")
$Participant1.add("name", "Lotte Vetler")

$Participant2 = @{}
$Participant2.add("address","otto.flick@office365itpros.com")
$Participant2.add("name", "Otto.Flick")

$Participant3 = @{}
$Participant3.add("address","kim.akers@office365itpros.com")
$Participant3.add("name", "Kim Akers")

$EventAttendee1 = @{}
$EventAttendee1.add("emailaddress", $Participant1)
$EventAttendee1.Add("type", "required")

$EventAttendee2 = @{}
$EventAttendee2.add("emailaddress", $Participant2)
$EventAttendee2.Add("type", "optional")

$EventAttendee3 = @{}
$EventAttendee3.add("emailaddress", $Participant2)
$EventAttendee3.Add("type", "optional")

$EventAttendees = $EventAttendee1, $EventAttendee2, $EventAttendee3

$EventDetails.Add("attendees", $EventAttendees)

$Uri =("https://graph.microsoft.com/v1.0/users/{0}/calendar/events" -f $userId)
$NewEvent = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $EventDetails

# Doesn't work at present - The property 'attendees' does not exist on type 'microsoft.graph.attende
$UpdatedEvent = Update-MgUserEvent -UserId $Userid -Eventid $NewEvent.Id -Attendees $EventDetails



[array]$Records = Search-unifiedauditlog -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date) `
    -Formatted -ObjectIds "*.agent" -Operations FileUploaded -ResultSize 5000 -SessionCommand ReturnLargeset
If ($Records) {
    $Records = $records | Sort-Object Identity -Unique
    Write-Host ("{0} audit records found" -f $Records.Count)
} Else {
    Write-Host "No audit records found"
    Break
}

$AgentReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
    $AuditData = $Rec.AuditData | ConvertFrom-Json
 
    $ReportLine = [PSCustomObject][Ordered]@{
        TimeStamp       = Get-Date ($AuditData.CreationTime) -format 'dd-MMM-yyyy HH:mm'
        User            = $AuditData.UserId
        Action          = $AuditData.Operation
        SiteURL         = $AuditData.SiteURL
        Agent           = $AuditData.SourceFileName

    }
    $AgentReport.Add($ReportLine)
}
$AgentReport = $AgentReport | Sort-Object {$_.TimeStamp -as [datetime]} -Descending
$AgentReport | Out-GridView -Title "Custom SharePoint Agent Creation"

Write-Host ""
Write-Host "Custom agents created in these SharePoint Online sites"
$AgentReport | Group-Object SiteURL -NoElement | Sort-Object Count -Descending | Format-Table Name, Count
Write-Host ""
Write-Host "Custom agents created by these users"
$AgentReport | Group-Object User -NoElement | Sort-Object Count -Descending | Format-Table Name, Count


# More

# Example: Build attractive HTML report for app role assignment audit records

# Sample data (replace with your actual records)
$Records = @(
    [PSCustomObject]@{
        CreatedDateTime     = '13-Jun-2025 12:56:09'
        Action             = 'App role assignment added to service principal'
        Application        = 'SDKAutomation'
        User               = 'Tony.Redmond@redmondassociates.org'
        GrantSource        = 'Microsoft Graph permission'
        SourceId           = '5e1e9171-754d-478c-812c-f1755a9a4c2d'
        'New Permissions'  = 'Read audit logs data from all services'
        ServicePrincipalId = '553a9e20-35ee-4ed1-b53e-ed32133996ae'
        AuditRecordId      = '0298fe74-2160-4ec5-bf89-f0f50e7898e1'
        Operation          = 'Add app role assignment to service principal'
    }
    # Add more records as needed
)

# Define HTML style with improved row color for visibility
$HtmlStyle = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; background: #f4f6f8; color: #222; }
h1 { background: #0078d4; color: #fff; padding: 16px; border-radius: 6px 6px 0 0; margin-bottom: 0; }
table { border-collapse: collapse; width: 100%; background: #fff; border-radius: 0 0 6px 6px; overflow: hidden; }
th, td { padding: 10px 12px; text-align: left; }
th { background: #e5eaf1; color: #222; }
tr { background: #fff; color: #222; }
tr:nth-child(even) { background: #f0f4fa; color: #222; }
tr:hover { background: #d0e7fa; color: #222; }
.caption { font-size: 14px; color: #555; margin-bottom: 12px; }
</style>
"@

# Convert records to HTML table
$HtmlTable = $Report | Select-Object `
    CreatedDateTime, Action, Application, User, GrantSource, SourceId, 'New Permissions', ServicePrincipalId |
    ConvertTo-Html -Fragment -PreContent "<div class='caption'>App Role Assignment Audit Records</div>"

# Compose full HTML
$HtmlReport = @"
<html>
<head>
$HtmlStyle
<title>App Role Assignment Audit Report</title>
</head>
<body>
<h1>App Role Assignment Audit Report</h1>
<p>Report generated: $(Get-Date -Format 'dd-MMM-yyyy HH:mm')</p>
$HtmlTable
</body>
</html>
"@

# Output to file or display
$ReportPath = "$env:TEMP\AppRoleAssignmentAuditReport.html"
$HtmlReport | Out-File -FilePath $ReportPath -Encoding utf8
Write-Host "HTML report created: $ReportPath"
Start-Process $ReportPath


[array]$Users = Get-MgUser -All -filter "usertype eq 'Member' and accountEnabled eq true" `
     -Property "id,displayName"

Set-MgRequestContext -MaxRetry 3 -RetryDelay 3


[array]$Users = Get-MgUser -All -filter "usertype eq 'Member' and accountEnabled eq true" `
     -Property "id,displayName"     
[int]$Pause = 2500
[int]$i=0   
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    $i++
    Write-Host ("Checking user {0} {1}" -f $i, $User.DisplayName)
    $Uri = ("https://graph.microsoft.com/v1.0/users/{0}?`$select=id,displayName,userPrincipalName,lastSigninActivity" -f $User.Id)
    Try {
        $Data = Invoke-MgGraphRequest -Uri $Uri -Method GET -ResponseHeadersVariable $Response -ErrorAction Stop
        If ($Data) {
            $LastSignIn = $null
            $LastSignIn = $Data.signInActivity.lastSignInDateTime
            If ($null -ne $LastSignIn) {
                $LastSignIn = Get-Date $LastSignIn -Format 'dd-MMM-yyyy HH:mm'
            } Else {
                $LastSignIn = "Never"
            }
            $ReportLine = [PSCustomObject][Ordered]@{
                DisplayName       = $Data.displayName
                UserPrincipalName = $Data.userPrincipalName
                LastSignIn        = $LastSignIn
            }
            $Report.Add($ReportLine)
        } Else {
            Write-Host "No data found for user" $User.DisplayName
        }
    } Catch {
        Write-Host "Error getting user" $User.DisplayName
        Write-Host $_.Exception.Message
        Continue
    }   
    If ($i % 5 -eq 0 -and $i -ne $Users.count) {
        Write-Host "Processed $i users, pausing for $Pause milliseconds..."; Start-Sleep -Milliseconds $Pause
    }
}


Write-Host "Checkiung for OAuth2 Permission Grants..."
# Get oAuth2PermissionGrant of Principal consent type (to impersonate a specific user)
[array]$Grants = Get-MgOauth2PermissionGrant -Filter "consentType eq 'Principal'" -All

Write-Host "Finding service principals..."
# Find service principals and create a hash table for quick lookup
[array]$ServicePrincipals = Get-MgServicePrincipal -All
If ($ServicePrincipals) {
    $SPHash = @{}
    ForEach ($SP in $ServicePrincipals) {   
        $SPHash.Add($SP.Id, $SP.DisplayName)
        }
} Else {
    Write-Host "No service principals found"
    Break
}

Write-Host "Looking for licensed users..."
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All -PageSize 999 | Sort-Object displayName 
If ($Users) {
    $UserHash = @{}
    ForEach ($User in $Users) {     
        $UserHash.Add($User.Id, $User) }
} Else {
    Write-Host "No licensed users found"
    Break
}

Write-Host "Generating report for OAuth2 Permission Grants..."
# Generate a report
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Grant in $Grants) {
    $SP = $ServicePrincipals | Where-Object { $_.Id -eq $Grant.ClientId }
  
    If ($SP) {
        $UserDetails = $UserHash[$Grant.PrincipalId]
        $Resource = $SPHash[$Grant.ResourceId]
        $ReportLine = [PSCustomObject][ordered]@{
            Id              = $SP.Id
            DisplayName     = $SP.DisplayName
            AppId           = $SP.AppId
            ConsentType     = $Grant.ConsentType
            User            = $UserDetails.UserPrincipalName
            UserDisplayName = $UserDetails.DisplayName
            Scope           =  (($Grant.Scope.Trim().Split(" "))) -join ", "
            Resource        = $Resource
        }
        $Report.Add($ReportLine)
    }
}

$Report | Out-Gridview -Title "Specific principal delegated OAuth2 permission grants"

[array]$Book25 = import-csv book2025Buyers.csv
$Book25 = $Book25 | Sort-Object -Property Email -Unique
[array]$Book26 = import-csv book2026Buyers.csv
$Book26 = $Book26 | Sort-Object -Property Email -Unique
$Book26Hash = @{}
ForEach ($Buyer in $Book26) {
    $Book26Hash.Add($Buyer.EmailTrim().ToLower(), $Buyer)
}

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Buyer in $Book25) {
    $LookUpValue = $Buyer.Email.Trim().ToLower()
    If ($null -eq $Book26Hash[$Buyer.Email]) {
        $ReportLine = [PSCustomObject][Ordered]@{
            Name            = $Buyer.Buyer
            Email           = $Buyer.Email
            Country         = $Buyer.Country
            Date            = $Buyer.'Purchase Date'
            Price           = $Buyer.Price
            Tip             = $Buyer.'Tip ($)'
        }
        $Report.Add($ReportLine)
    }
}
 $report | Export-Csv -Path "ToBuy.csv" -NoTypeInformation -Encoding UTF8

# Convert UNIX epoch time (seconds since 1970-01-01) to PowerShell DateTime
$UnixEpochValue = 1752763429
$Date = [DateTimeOffset]::FromUnixTimeSeconds($UnixEpochValue).ToLocalTime().DateTime
Write-Host "UNIX epoch $UnixEpochValue is" $(Get-Date $Date -format 'dd-MMM-yyyy HH:mm')



[Array]$Users = Get-MgUser -All -Sort DisplayName
[Array]$ProcessedUsers = @()
[int]$i=0
Do {
    ForEach ($User in $Users) {
        $i++
        Try {
           #$Uri = "https://graph.microsoft.com/v1.0/users/{0}?`$select=id,displayName,userPrincipalName,signInActivity" -f $User.Id
           #$Data = Invoke-MgGraphRequest -Method GET -Uri $Uri
            $Data = Get-MgUser -UserId $User.Id -Property id,displayName,userPrincipalName,signInActivity -ErrorAction Stop
            $ProcessedUsers += $Data
            Write-Host ("Processed user {0} ({1})" -f $i, $User.DisplayName)
        } Catch {
            If ($_.Exception.Message -like "*Too many retries performed*") {
                Write-Host ("Detected: Too many retries performed error when processing user {0}." -f $User.DisplayName) -ForegroundColor Red
                # Wait 15 second then reattempt to process user
                Start-Sleep -Seconds 15
                Write-Host ("Retrying user {0}" -f $User.DisplayName)
                $Data = Get-MgUser -UserId $User.Id -Property id,displayName,userPrincipalName,signInActivity
                $ProcessedUsers += $Data
            } Else {
                Write-Host ("Other error: {0}" -f $_.Exception.Message)
            }
        }
    }
} While ($i -lt 500)


# Disable the People, Files, and Calendar Microsoft 365 Companion Apps from starting automatically
$RegistryKey = "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\Microsoft.M365Companions_8wekyb3d8bbwe"
[array]$AppStartUpIds = @("$RegistryKey\CalendarStartupId","$RegistryKey\FilesStartupId","$RegistryKey\PeopleStartupId")

ForEach ($AppStartupId in $AppStartUpIds) {
    Try {
        If (Test-Path $AppStartupId) {
            # Disable startup state for the app
            Write-Host ("Disabling startup state for the {0} companion app" -f $AppStartupId.Split("StartupId")[0].Split("\")[11]) -ForegroundColor Green
            Set-ItemProperty -Path $AppStartupId -Name "State" -Value 1 -Type DWORD -ErrorAction Stop
        } Else {
            Write-Host ("Couldn't find path to disable startup for the {0} companion app" -f $AppStartupId.Split("StartupId")[0].Split("\")[11]) -ForegroundColor Red
        }
    } Catch {
        Write-Error ("Failed to set State for {0} : {1}" -f $AppStartupId, $_)  
    }
}
Write-Host "Completed suppressing the startup of the Calendar, Files, and People companion apps"

#- Search for audit records for Researcher Copilot agent

[array]$Records = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date) `
    -Formatted -Operations "CopilotInteraction" -ResultSize 5000 -SessionCommand ReturnLargeSet
If ($Records) {
    $Records = $records | Sort-Object Identity -Unique  
    Write-Host ("{0} audit records found" -f $Records.Count)
    $Records = $Records | Sort-Object { $_.CreationDate -as [datetime]} -Descending
} Else {
    Write-Host "No audit records found - now scanning for Researcher Copilot agent records"
    Break
}   

$ResearcherReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {    
    $AuditData = $Rec.AuditData | ConvertFrom-Json
    If ($AuditData.AgentName -ne "Researcher") {
        Continue
    }
    [array]$CopilotResources = $AuditData.CopilotEventData.AccessedResources
    If ($CopilotResources) {
        $Resources = [System.Collections.Generic.List[Object]]::new()
        ForEach ($Resource in $CopilotResources) {
            If ($Resource.Action) {
                If ($Resource.SiteURL) {
                    $ResourceName = $Resource.SiteURL
                    $ResourceId  = $null
                } Else {
                    $ResourceName = $Resource.Name
                    $ResourceId   = $Resource.Id.Split("&")[0]
                }
                $ResourcesAccessed =  [PSCustomObject][Ordered]@{
                    Action = $Resource.Action
                    Name   = $ResourceName
                    Id     = $ResourceId
                }
                $Resources.Add($ResourcesAccessed)
            }   
        }
        $ReportLine = [PSCustomObject][Ordered]@{
            TimeStamp       = Get-Date ($AuditData.CreationTime) -format 'dd-MMM-yyyy HH:mm'
            User            = $AuditData.UserId
            Action          = $AuditData.Operation
            Resources       = $Resources.Name -join "; "
        }
        $ResearcherReport.Add($ReportLine)
        } Else {
            If ($null -ne $AuditData.CopilotEventData.Messages) {
                $ReportLine = [PSCustomObject][Ordered]@{
                    TimeStamp       = Get-Date ($AuditData.CreationTime) -format 'dd-MMM-yyyy HH:mm'
                    User            = $AuditData.UserId
                    Action          = $AuditData.Operation
                    Resources       = If ($AuditData.CopilotEventData.Messages.count -le 3) {$AuditData.CopilotEventData.Messages.Count}                                     
                                       Else {$AuditData.CopilotEventData.Messages.count.toString() + " (likely use of Claude LLM)"}
                }
                $ResearcherReport.Add($ReportLine)
            }
        }
}

[array]$Invitations = Get-MgAuditLogDirectoryAudit -All -Filter "activityDisplayName eq 'Invite external user to organization'" 
If ($Invitations) {
    Write-Host ("{0} invitation audit records found" -f $Invitations.Count)
} Else {
    Write-Host "No invitation audit records found"  
}
ForEach ($Invitation in $Invitations) {
    $Inviter = $Invitation.InitiatedBy.User.UserPrincipalName 
    $Invitee = $Invitation.TargetResources[0].DisplayName
    $InviteeUPN = $Invitation.TargetResources[0].UserPrincipalName 
    $InviteeID = $Invitation.TargetResources[0].Id
    $InvitationDate = Get-Date $Invitation.ActivityDateTime -Format 'dd-MMM-yyyy HH:mm'

    Write-Host ("{0} invited {1} ({2}) on {3}" -f $Inviter, $Invitee, $InviteeUPN, $InvitationDate)
}





<# 
.SYNOPSIS
  Export EnforcedTimestamps for all mailboxes with numeric EventType -> friendly name mapping.

.REQUIREMENTS
  - Exchange Online Management module (EXO v2+)
  - Permissions: Exchange Administrator (or equivalent)

.OUTPUT
  - CSV file in current directory: EnforcedTimestamps_<yyyyMMdd_HHmmss>.csv
#>

# --- 1) Connect to Exchange Online (comment out if you are already connected) ---
try {
    # If you use MFA, omit -UserPrincipalName and sign in interactively
    Connect-ExchangeOnline -ShowBanner:$false | Out-Null
} catch {
    Write-Warning "Could not connect to Exchange Online. If already connected, you can ignore this warning."
}


# --- 6) Optional: summary on screen ---
$export |
  Group-Object EventTypeName |
  Sort-Object Count -Descending |
  Format-Table Name, Count
``
$Uri = "https://graph.microsoft.com/v1.0/security/alerts_v2?`$filter=serviceSource eq 'DataLossPrevention'&`$orderby=createdDateTime desc&`$top=200"
[array]$Alerts = Invoke-MgGraphRequest -Method GET -Uri $Uri

[array]$DLPAlerts = Get-MgSecurityAlertV2 -Filter "serviceSource eq 'dataLossPrevention'" -PageSize 500 -All -Sort "CreatedDateTime Desc"

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Alert in $DLPAlerts) {   
    If ($Alert.LastUpdateDateTime) { $LastUpdated = Get-Date $Alert.LastUpdateDateTime -Format 'dd-MMM-yyyy HH:mm' 
    } Else { 
        $LastUpdated = "N/A" 
    }
    $ReportLine = [PSCustomObject][Ordered]@{
        Id                  = $Alert.id
        Title               = $Alert.title
        CreatedDateTime     = Get-Date $Alert.createdDateTime -Format 'dd-MMM-yyyy HH:mm'
        Severity            = $Alert.severity
        Status              = $Alert.status
        Category            = $Alert.category
        AssignedTo          = $Alert.AssignedTo
        LastUpdateDateTime  = $LastUpdated
        Classification      = $Alert.classification
    } 
    $Report.Add($ReportLine)
}

# Report SharePoint RCD Audit Events
[array]$Operations = "RestrictContentOrgWideSearchDisabled", "RestrictContentOrgWideSearchEnabled", "SharePointRCDUpdateJustification"
[array]$Records = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) -Formatted `
    -SessionCommand ReturnLargeSet -ResultSize 5000 -Operations $Operations
$Records = $Records | Sort-Object Identity -Unique

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
    $AuditData = $Rec.AuditData | ConvertFrom-Json
 
    Switch ($AuditData.Operation) {
        "RestrictContentOrgWideSearchEnabled" {
            $Action = ("Enabled RCD for {0}" -f $AuditData.ObjectId)
        }
        "SharePointRCDUpdateJustification" {
            $Action = ("RCD Justification: {0}" -f $AuditData.RCDJustification)
        }
        "RestrictContentOrgWideSearchDisabled" {
            $Action = ("Disabled RCD for {0}" -f $AuditData.ObjectId)
        }
        Default {
            $Action = $AuditData.Operation
        }
    }
    $ReportLine = [PSCustomObject][Ordered]@{
        TimeStamp       = Get-Date ($AuditData.CreationTime) -format 'dd-MMM-yyyy HH:mm'
        User            = $AuditData.UserId
        Action          = $AuditData.Operation
        SiteURL         = $AuditData.ObjectId
        EventInfo       = $Action

    }
    $Report.Add($ReportLine)
}
$Report = $Report | Sort-Object {$_.TimeStamp -as [datetime]} -Descending
$Report | Out-GridView -Title "Restricted Content Discovery Audit Records"

# Create a calendar event in a Teams channel calendar

# Get identifiers for the target team (group) and channel
$GroupId = Get-MgGroup -Filter "displayName eq 'Microsoft Graph Gurus'" | Select-Object -ExpandProperty Id
$ChannelId = Get-MgTeamChannel -TeamId $GroupId -Filter "displayName eq 'General'" | Select-Object -ExpandProperty Id

# Define event details
$EventDetails = @{}
$EventDetails.Add("subject", "Team Channel Meeting")
$EventBody = @{}
$EventBody.Add("contentType", "HTML")   
$EventBody.Add("content", "This is a meeting scheduled in the Teams channel calendar.")
$EventDetails.Add("body", $EventBody)       
$EventStart = @{}
$EventStart.Add("dateTime", "2026-01-15T10:00")
$EventStart.Add("timeZone", "UTC")
$EventDetails.Add("start", $EventStart)
$EventEnd = @{}         
$EventEnd.Add("dateTime", "2026-01-15T11:00")
$EventEnd.Add("timeZone", "UTC")
$EventDetails.Add("end", $EventEnd)
$EventLocation = @{}
$EventLocation.Add("displayName", "Teams Channel Meeting")
$EventDetails.Add("location", $EventLocation)
$EventDetails.Add("isOnlineMeeting", $true)
$EventDetails.Add("onlineMeetingProvider", "teamsForBusiness")  
$EventChannel = @{}
$EventChannel.Add("teamId", $GroupId)       
$EventChannel.Add("channelid", $ChannelId)
$EventDetails.Add("channelIdentity", $EventChannel)

# URI to create event in channel calendar
$Uri = ("https://graph.microsoft.com/v1.0/groups/{0}/calendar/events" -f $GroupId)
$NewEvent = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $EventDetails

# Client side filter for mailbox items by size


# PR_MESSAGE_SIZE
$SizePropId = "Integer 0x0e08"
# Size threshold
$Threshold = 0.5MB   
# Date range (example: last 30 days)
$since = (Get-Date).AddDays(-30).ToString('s') + "Z"

# Server-side filter - find messages with attachments received since $since
$filter = "hasAttachments eq true and receivedDateTime ge $since"

[array]$Messages = Get-MgUserMailFolderMessage -UserId $User.Id -MailFolderId 'Inbox' `
    -Filter $filter -All -Select id,subject,sender,receivedDateTime `
    -ExpandProperty "singleValueExtendedProperties(`$filter=id eq '$SizePropId')"

$Report = [System.Collections.Generic.List[Object]]::new()
# Client-side size check
$Messages | ForEach-Object {
    $Size = [int64]$_.SingleValueExtendedProperties[0].Value
    If ($Size -gt $Threshold) {
        $ReportLine = [PSCustomObject][Ordered]@{
            Subject          = $_.Subject
            From             = $_.Sender.EmailAddress.Address
            ReceivedDateTime = $_.ReceivedDateTime
            SizeMB           = [math]::Round($size / 1MB, 2)
            Id               = $_.Id
        }
        $Report.Add($ReportLine)
    }
}
    
$Report | Sort-Object SizeMB -Descending | Out-GridView -Title "Large Emails in Inbox over $($Threshold / 1MB) MB"

$ScopeName = "Permanent Employees”
$Membership = Get-AdaptiveScopeMembers -Identity $ScopeName -ErrorAction SilentlyContinue
If (!($Membership)) {
  Write-Output ("No members found for {0} scope" -f $ScopeName )
} Else { 
  Write-Output ("{0} members found in {1} scope" -f $Membership.count, $ScopeName)
}

$DateString = "04/22/2022 21:36:39 +00:00"
$ParsedDate = [DateTime]::ParseExact($DateString, "MM/dd/yyyy HH:mm:ss zzz", [System.Globalization.CultureInfo]::InvariantCulture)
Write-Host $ParsedDate


# UserConfiguration FAI Items in ueer mailboxes

$User = (Get-MgContext).Account
$UserId = (Get-MgUser -UserId $User).Id

# Calendar work hours
$Uri = ("https://graph.microsoft.com/beta/users/{0}/mailFolders/Calendar/userConfigurations/WorkHours" -f $UserId)
[array]$Data = Invoke-MgGraphRequest -URI $Uri -Method Get -OutputType PSObject
[string]$Base64Convert = [Text.Encoding]::Utf8.GetString([Convert]::FromBase64String($Data.xmldata))

# Parse calendar settings from the XML data
[xml]$doc = $Base64Convert

$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace('w','WorkingHours.xsd')

$root = $doc.SelectSingleNode('//w:WorkHoursVersion1',$ns)

$tz = $root.SelectSingleNode('w:TimeZone',$ns)
$tzName = $tz.SelectSingleNode('w:Name',$ns).InnerText
$tzBias = [int]$tz.SelectSingleNode('w:Bias',$ns).InnerText

$std = $tz.SelectSingleNode('w:Standard',$ns)
$stdBias = [int]$std.SelectSingleNode('w:Bias',$ns).InnerText
$stdChange = $std.SelectSingleNode('w:ChangeDate',$ns)
$stdChangeDate = $stdChange.SelectSingleNode('w:Date',$ns).InnerText
$stdChangeTime = $stdChange.SelectSingleNode('w:Time',$ns).InnerText
$stdChangeDayOfWeek = $stdChange.SelectSingleNode('w:DayOfWeek',$ns).InnerText

$dst = $tz.SelectSingleNode('w:DaylightSavings',$ns)
$dstBias = [int]$dst.SelectSingleNode('w:Bias',$ns).InnerText
$dstChange = $dst.SelectSingleNode('w:ChangeDate',$ns)
$dstChangeDate = $dstChange.SelectSingleNode('w:Date',$ns).InnerText
$dstChangeTime = $dstChange.SelectSingleNode('w:Time',$ns).InnerText
$dstChangeDayOfWeek = $dstChange.SelectSingleNode('w:DayOfWeek',$ns).InnerText

$timeslot = $root.SelectSingleNode('w:TimeSlot',$ns)
$start = $timeslot.SelectSingleNode('w:Start',$ns).InnerText
$end = $timeslot.SelectSingleNode('w:End',$ns).InnerText

$workdays = $root.SelectSingleNode('w:WorkDays',$ns).InnerText

$CalendarOptions = [PSCustomObject]@{
    TimeZoneName            = $tzName
    TimeZoneBiasMinutes     = $tzBias
    StandardBiasMinutes     = $stdBias
    StandardChangeDate      = $stdChangeDate
    StandardChangeTime      = $stdChangeTime
    StandardChangeDayOfWeek = $stdChangeDayOfWeek
    DaylightBiasMinutes     = $dstBias
    DaylightChangeDate      = $dstChangeDate
    DaylightChangeTime      = $dstChangeTime
    DaylightChangeDayOfWeek = $dstChangeDayOfWeek
    WorkStart               = $start
    WorkEnd                 = $end
    WorkDays                = $workdays
}

# Calendar options
$Uri = ("https://graph.microsoft.com/beta/users/{0}/mailFolders/Calendar/userConfigurations/Calendar" -f $UserId)
[array]$Data = Invoke-MgGraphRequest -URI $Uri -OutputType PsObject | Select-Object -ExpandProperty structuredData

ForEach ($Setting in $Data) {
    [string]$SettingName = $Setting.keyEntry.values
    [string]$SettingValue = $Setting.valueEntry.values
    Write-Host ("{0} : {1}" -f $SettingName, $SettingValue)
}