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

# Bring all the event details together into a hash table for the payload
$EventDetails = @{}
$EventDetails.Add("subject", "The Experts Conference (TEC) 2025" – Best Conference Ever)

$EventDetails.Add("body", $EventBody)
$EventDetails.Add("start", $EventStart)
$EventDetails.Add("end", $EventEnd)
$EventDetails.Add("location", $EventLocation)
$EventDetails.Add("allowNewTimeProposals", $true)
$EventDetails.Add("transactionId", (New-Guid))

# hash table for attendees
#$EventAttendees = @()

# Each attendde defined as email address and name
#$Participant1 = @{}
#$Participant1.add("address","lotte.vetler@office365itpros.com")
#$Participant1.add("name", "Lotte Vetler")

#$Participant2 = @{}
#$Participant2.add("address","otto.flick@office365itpros.com")
#$Participant2.add("name", "Otto.Flick")

#$Participant3 = @{}
#$Participant3.add("address","kim.akers@office365itpros.com")
#$Participant3.add("name", "Kim Akers")

#$EventAttendee1 = @{}
#$EventAttendee1.add("emailaddress", $Participant1)
#$EventAttendee1.Add("type", "required")

#$EventAttendee2 = @{}
#$EventAttendee2.add("emailaddress", $Participant2)
#$EventAttendee2.Add("type", "optional")

#$EventAttendee3 = @{}
#$EventAttendee3.add("emailaddress", $Participant2)
#$EventAttendee3.Add("type", "optional")

#$EventAttendees = $EventAttendee1, $EventAttendee2, $EventAttendee3

# Input array for event participants
$ParticipantsArray = @(
    @{ name = "Kim Akers"; emailaddress = "kim.akers@office365itpros.com"; state = "required" },
    @{ name = "Lotte Vetler"; emailaddress = "lotte.vetler@office365itpros.com"; state = "optional" },
    @{ name = "James Ryan"; emailaddress = "james.ryan@office365itpros.com"; state = "optional" },
    @{ name = "Otto Flick"; emailaddress = "otto.flick@office365itpros.com"; state = "required" }
)

# Create hash tables for each participant
$EventAttendees = @()
foreach ($Participant in $ParticipantsArray) {
    $ParticipantHash = @{
        address = $Participant.emailaddress
        name    = $Participant.name
    }

    $EventAttendee = @{
        emailaddress = $ParticipantHash
        type         = $Participant.state
    }

    $EventAttendees += $EventAttendee
}

# Add the attendees to the event
$EventDetails.Add("attendees", $EventAttendees)

$Uri =("https://graph.microsoft.com/v1.0/users/{0}/calendar/events/" -f $userId)
$NewEvent = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $EventDetails


 $UpdateEventDetails = @{}
 $UpdateEventDetails.Add("isOnlineMeeting", $true)
 $UpdateEventDetails.Add("onlineMeetingProvider", "teamsForBusiness")
 $UpdateEventDetails.Add("isReminderOn", $true)
 $UpdateEventDetails.Add("reminderMinutesBeforeStart", 90)
 $UpdateEventDetails.Add("categories", @("Conference", "Travel"))

    
 $UpdatedEvent = Invoke-MgGraphRequest -Uri $Uri -Method PATCH -Body $EventDetails
    
 $NewEvent = Update-MgUserEvent -UserId $userId -EventId $NewEvent.Id -BodyParameter $UpdateEventDetails

 Update-MgUserEvent -Userid $Userid -Eventid $NewEvent.Id -IsOnlineMeeting:$true -Importance High -OnlineMeetingProvider 'TeamsforBusiness' -ReminderMinutesBeforeStart 30

# Update with attendees - rewrite attendee list

$Participant1 = @{}
$Participant1.Add("address","Lotte.Vetler@office365itpros.com")
$Participant1.Add("name", "Lotte Vetler")

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

$Event = @{}
$Event.Add("recurrence", $RecurrenceRange)

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

$Event.Add("subject", "TEC Roadshow")
$Event.Add("body", $EventBody)
$Event.Add("start", $EventStart)
$Event.Add("end", $EventEnd)
$Event.Add("location", $EventLocation)
$Event.Add("allowNewTimeProposals", $true)
$Event.Add("transactionId", (New-Guid))

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

$Event.Add("attendees", $EventAttendees)

$Uri =("https://graph.microsoft.com/v1.0/users/{0}/calendar/events" -f $userId)
$NewEvent = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $Event

# Doesn't work at present - The property 'attendees' does not exist on type 'microsoft.graph.attende
$UpdatedEvent = Update-MgUserEvent -UserId $Userid -Eventid $NewEvent.Id -Attendees $Event



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


# Delete an event
$Uri = ("https://graph.microsoft.com/v1.0/users/{0}/events/{1}" -f $UserId, $NewEvent.Id)   
Invoke-MgGraphRequest -Uri $Uri -Method DELETE

# Permanently Delete an event   
$Uri = ("https://graph.microsoft.com/v1.0/users/{0}/events/{1}/permanentDelete" -f $UserId, $NewEvent.Id)
Invoke-MgGraphRequest -Uri $Uri -Method POST


# Cancel
$CancelMessage = @{}
$CancelMessage.Add("comment", "Cancelling because I can't go to TEC 2025! (sad)")


$Uri = ("https://graph.microsoft.com/v1.0/users/{0}/events/{1}/cancel" -f $UserId, $NewEvent.Id)
Invoke-MgGraphRequest -Uri $Uri -Method POST -Body $CancelMessage
