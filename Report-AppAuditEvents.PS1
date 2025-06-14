# Report-AppAuditEvents.PS1
# Version 1.0 14-June-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-AppAuditEvents.PS1

# This scripts uses the AuditLog Query Graph API to fetch audit events for app creation and updates to report the 
# information to administrators.
# Requires the AuditLogsQuery.Read.All permission in the Microsoft Graph API.
# Requires Mail.Send to send email
# Requires Application.Read.All to read applications and service principals
# Requires DelegatedPermissionGrant.Read.All to read delegated permission grants

If ([Environment]::UserInteractive) { 
    # We're running interactively...
    Write-Host "Running interactively..."
    Connect-MgGraph -NoWelcome -Scopes User.Read.All
    # Define message sender for email sent at the end of the script
    $MsgFrom = (Get-MgContext).Account
    Write-Host "Email will be sent from $MsgFrom"
} Else { 
    # We're not, so likely in Azure Automation
    Write-Output "Script executing as a Azure Automation runbook..."
    Connect-MgGraph -Identity -NoWelcome
    # Define the message sender - change this value to the email address of the account that will send the message
    $MsgFrom = 'Customer.Services@office365itpros.com'
    Write-Output "Email will be sent from $MsgFrom"
}

$AuditJobName = ("Audit job created at {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm'))
$AuditQueryStart = ((Get-Date).AddDays(-90).toString('yyyy-MM-ddTHH:mm:ss'))
$AuditQueryEnd = ((Get-Date).toString('yyyy-MM-ddTHH:mm:ss'))
[array]$AuditOperationFilters = "Add app role assignment to service principal.", "Add application.", "Add delegated permission grant.", "Update application - Certificates and secrets management", "Consent to application"

$AuditQueryParameters = @{}
$AuditQueryParameters.Add("@odata.type","#microsoft.graph.security.auditLogQuery")
$AuditQueryParameters.Add("displayName", $AuditJobName)
$AuditQueryParameters.Add("OperationFilters", $AuditOperationFilters)
$AuditQueryParameters.Add("filterStartDateTime", $AuditQueryStart)
$AuditQueryParameters.Add("filterEndDateTime", $AuditQueryEnd)

$Uri = "https://graph.microsoft.com/beta/security/auditLog/queries"
$AuditJob = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $AuditQueryParameters

If ($null -eq $AuditJob) {
    Write-Host "Failed to create audit query job. Please check the code and permissions and try again."
    Exit
}

# Check the audit query status every 20 seconds until it completes
[int]$i = 1
[int]$SleepSeconds = 20
$SearchFinished = $false; [int]$SecondsElapsed = 20
# Write-Host "Checking audit query status..."
# Initial wait to let the audit job spin up
Start-Sleep -Seconds 30
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
        $AuditQueryStatus = Invoke-MgGraphRequest -Uri $Uri -Method Get
    }
}

# Fetch audit records found by the search
$AuditRecords = [System.Collections.Generic.List[string]]::new()
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}/records?`$top=999" -f $AuditJob.Id)
[array]$AuditSearchRecords = Invoke-MgGraphRequest -Uri $Uri -Method GET
[array]$AuditRecords = $AuditSearchRecords.value

$NextLink = $AuditSearchRecords.'@Odata.NextLink'
While ($null -ne $NextLink) {
    $AuditSearchRecords = $null
    [array]$AuditSearchRecords = Invoke-MgGraphRequest -Uri $NextLink -Method GET 
    $AuditRecords += $AuditSearchRecords.value
    # Write-Host ("{0} audit records fetched so far..." -f $AuditRecords.count)
    $NextLink = $AuditSearchRecords.'@odata.NextLink' 
}

Write-Host ("Audit query {0} returned {1} records" -f $AuditJobName, $AuditRecords.Count)
If ($AuditRecords.Count -eq 0) {
    Write-Host "No audit records found for the specified time period and operations."
    Exit
} 
# Make sure that the audit records are sorted by date
$AuditRecords = $AuditRecords | Sort-Object CreatedDateTime -Descending

# Fetch all service principals to resolve names and populate a hash table for lookup    
$SP = Get-MgServicePrincipal -All | Select-Object Id, DisplayName | Sort-Object DisplayName
$SPHash = @{}
ForEach ($ServicePrincipal in $SP) {
    $SPHash.Add($ServicePrincipal.Id, $ServicePrincipal.DisplayName)
}
# Load the roles for the Microsoft Graph application to resolve app role assignments
$GraphApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
[array]$Roles = $GraphApp.AppRoles

$Report = [System.Collections.Generic.List[Object]]::new() 
[int]$i = 0
ForEach ($Record in $AuditRecords) {
    $i++
    Switch ($Record.Operation) {
        "Add app role assignment to service principal." { 
            $Record.ActivityDisplayName = "Add app role assignment to service principal" 
            $OperationDisplayName = "App role assignment added to service principal"
            $RoleCheck = $Roles | Where-Object {$_.Id -eq $record.auditdata.modifiedproperties[0].NewValue.Trim() }
            If ($RoleCheck) {
                $GrantSource = "Microsoft Graph permission"
                $SourceId = $RoleCheck.Id
                $Permissions = $RoleCheck.DisplayName.Trim()
            } Else {
                $GrantSource = $Record.auditdata.modifiedproperties[0].NewValue.Trim()
                $SourceId = $Record.AuditData.modifiedProperties[0].NewValue.Trim()
                $Permissions = $Record.auditdata.modifiedproperties[1].NewValue.Trim()
            }
            $GrantedTo = $Record.auditdata.modifiedproperties[6].NewValue.Trim()
            $SPIdGrantedTo = $Record.AuditData.modifiedProperties[5].NewValue.Trim()
        }
        "Add application." { 
            $OperationDisplayName = "New Application created"
            $App = Get-MgApplication -ApplicationId $Record.Auditdata.Target[1].ID
            $GrantedTo = $App.DisplayName
            $AppId = $App.AppId
            $SPIdGrantedTo = (Get-MgServicePrincipal -Filter "appId eq '$AppId'").Id
            $Permissions = $null
            $GrantSource = $null
            $SourceId = $App.Id
        }
        "Add delegated permission grant." { 
            $OperationDisplayName = "App received delegated permission grant"
            $GrantSource = $SPHash[$Record.AuditData.Target[1].id]
            $SourceId = $Record.AuditData.Target[1].id 
            [array]$OldPermissions = $Record.AuditData.modifiedProperties[0].OldValue.Trim().Split(' ')
            [array]$NewPermissions = $Record.AuditData.modifiedProperties[0].NewValue.Trim().Split(' ')
            $Permissions = $NewPermissions | Where-Object { $_ -notin $OldPermissions } 
            [string]$Permissions = $Permissions -join ', '
            $GrantedTo = $SPHash[$Record.AuditData.Target[0].id]
            $GrantedTo = $SPHash[$Record.AuditData.modifiedProperties[2].NewValue]
            $SPIdGrantedTo = $Record.AuditData.modifiedProperties[2].NewValue
        }
        # Yes, this operation does have a trailing space and a en dash (u+2013) in the name...
        "Update application – Certificates and secrets management " { 
            # Extract details of app secret or certificate added or removed from the application
            $Record.ActivityDisplayName = "Update application - Certificates and secrets management" 
            $OperationDisplayName = "Application updated with certificate or app secret"
            $AppId = $Record.AuditData.ObjectId.Split("_")[1]
            $App = Get-MgApplication -ApplicationId $AppId -ErrorAction SilentlyContinue
            $GrantedTo = $App.DisplayName
            $SPIdGrantedTo = $null
            $OldValue = $Record.auditdata.modifiedProperties[0].OldValue | ConvertFrom-Json
            $NewValue = $Record.auditdata.modifiedProperties[0].NewValue | ConvertFrom-Json
            $KeyString = $NewValue | Where-Object { $_ -notin $OldValue }
            If ($null -eq $KeyString) {
                $KeyString = "Can't determine new value for certificate or app secret"
            } Else  {
                # Remove the brackets
                $KeyString = $KeyString.Trim('[', ']')
                # Split into key-value pairs
                $Pairs = $KeyString -split ','
                # Create a hashtable for the properties
                $CertProps = @{}
                Foreach ($Pair in $pairs) {
                    $kv = $Pair -split '=', 2
                    If ($kv.Count -eq 2) {
                        $key = $kv[0].Trim()
                        $value = $kv[1].Trim()
                    $CertProps[$key] = $value
                    }
                }
                # Generate the output string
                $Permissions = ("Key Id: {0}, Key Type: {1}, Key Usage: {2}, Name: {3}" -f `
                $CertProps['KeyIdentifier'], $CertProps['KeyType'], $CertProps['KeyUsage'], $CertProps['DisplayName'])
            }
            $GrantSource = $null
            $SourceId = $null
        }

        "Consent to application." { 
            $OperationDisplayName = "Consent granted for permissions"
            If ($Record.AuditData.modifiedProperties[0].NewValue -eq $true) {
                $GrantSource = "Administrator consent"
            } Else {
                $SourceId = "User consent"
            }
            If ($Record.AuditData.modifiedProperties[4].NewValue -match 'Id:\s*([^\s,]+)') {
                $OAuth2Delegation = $matches[1]
                # Write-Host "Attempting to resolve OAuth2 delegation $OAuth2Delegation  $i"
                $Delegated = Get-MgOauth2PermissionGrant -OAuth2PermissionGrantId $OAuth2Delegation -ErrorAction SilentlyContinue
                If ($Delegated) {
                    [array]$AssignedPermissions = $Delegated.Scope.trim().split(" ") 
                    [string]$Permissions = $AssignedPermissions -join ', '
                } Else {
                    $Permissions = "Unable to resolve delegated permissions"
                }
            }
            $SourceId = $null
            $GrantedTo = $SPHash[$Record.AuditData.Target[1].id]
        }
        Default { # Just in case we get an unexpected audit record
            $OperationDisplayName = $Record.Operation
            $GrantSource = $null
            $SourceId = $null
            $Permissions = $null
            $GrantedTo = $null
            $SPIdGrantedTo = $null
        }
       
    }
    $ReportLine = [PSCustomObject] @{
        CreatedDateTime     = (Get-Date $Record.CreatedDateTime -format 'dd-MMM-yyyy HH:mm:ss')
        Action              = $OperationDisplayName
        Application         = $GrantedTo
        User                = $Record.UserPrincipalName
        GrantSource         = $GrantSource
        SourceId            = $SourceId
        'New Permissions'   = $Permissions
        ServicePrincipalId  = $SPIdGrantedTo
        AuditRecordId       = $Record.Id
        Operation           = $Record.Operation
    }
    $Report.Add($ReportLine)
}

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
    ConvertTo-Html -Fragment -PreContent "<div class='caption'>Critical App Management Audit Events for Tenant Administrators to Review</div>"

# Compose full HTML
$HtmlReport = @"
<html>
<head>
$HtmlStyle
<title>App Critical Audit Events Report</title>
</head>
<body>
<h1>Details of Critical Audit Events for App Management</h1>
<p>Report generated: $(Get-Date -Format 'dd-MMM-yyyy HH:mm')</p>
$HtmlTable
</body>
</html>
"@

$ReportFile = "$env:TEMP\CriticalAppManagementAuditEvents.html"
$CSVFile = "$env:TEMP\CriticalAppManagementAuditEvents.csv"
$HtmlReport | Out-File -FilePath $ReportFile -Encoding utf8
$Report | Export-Csv -Path $CSVFile -NoTypeInformation -Encoding UTF8
Write-Host "HTML report created: $ReportFile"
Dir $Env:TEMP\CriticalAppManagementAuditEvents.* | Format-Table Name, Length, LastWriteTime

$Attachment = (Get-Item -Path $CSVFile).Name
$EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($CSVFile))

$MsgAttachments = @(
    @{
	"@odata.type" = "#microsoft.graph.fileAttachment"
	Name = ($CSVFile -split '\\')[-1]
	ContentBytes = $EncodedAttachmentFile
	}
)

# Build the array of a single TO recipient detailed in a hash table - make sure that you add your preferred email address here.
# The address can be any emailable object, including a user, group, or distribution list.
$ToRecipients = @{}
$ToRecipients.Add("emailAddress",@{'address'="tony.redmond@office365itpros.com"})
[array]$MsgTo = $ToRecipients
# Define the message subject
$MsgSubject = "Important: Critical App Management Audit Events Report"

$MsgBody = @{}
$MsgBody.Add('Content', "$($HtmlReport)")
$MsgBody.Add('ContentType','html')

$Message = @{}
$Message.Add('subject', $MsgSubject)
$Message.Add('toRecipients', $MsgTo)
$Message.Add('body', $MsgBody)
$Message.Add('attachments', $MsgAttachments)

$EmailParameters = @{}
$EmailParameters.Add('message', $Message)
$EmailParameters.Add('saveToSentItems', $true)
$EmailParameters.Add('isDeliveryReceiptRequested', $true)

# Send the message
Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters



# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 