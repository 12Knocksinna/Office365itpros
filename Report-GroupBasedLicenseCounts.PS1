# Report-GroupBasedLicenseCounts.PS1
# Report on group-based licensing in a tenant by analyzing the licenses assigned through each group and sending a report by email.

# V1.0 24-Feb-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-GroupBasedLicenseCounts.PS1

# Requires Mail.Send to send email
# Requires Group.Read.All to read groups
# Requires GroupMember.Read.All to read group members
# Requires User.Read.All to read user data

# Download the Microsoft licensing data
Write-Host "Attempting to download Microsoft licensing data..."
[array]$ProductData = Invoke-RestMethod -Method Get -Uri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv" | ConvertFrom-CSV
If ($ProductData) {
    [array]$ProductInfo = $ProductData | Sort-Object GUID -Unique
    $ProductInfoHash = @{}
    ForEach ($P in $ProductInfo) {
        $ProductInfoHash.Add([string]$P.GUID, [string]$P.Product_Display_Name)
    }   
} Else {
    # if a local copy exists, you could plug it in here
    Write-Host "Unable to retrieve product data"
    Break
}

[array]$TenantSkus = Get-MgSubscribedSku -All | Select-Object SkuId, SkuPartNumber, ServicePlans, ConsumedUnits, PrepaidUnits

# Get all groups with assigned licenses
Write-Host "Finding groups used by group-based licensing..."
[array]$Groups = Get-MgGroup -All -Filter "assignedLicenses/`$count ne 0" `
    -ConsistencyLevel Eventual -CountVariable Count -Property Id, DisplayName, AssignedLicenses, LicenseProcessingState

If (!$Groups) {
    Write-Host "No groups found with assigned licenses"
    Break
} Else {
    Write-Host ("Found {0} groups with assigned licenses" -f $Groups.Count)
}

$Report = [System.Collections.Generic.List[Object]]::new()
$GroupBasedSkusReport = [System.Collections.Generic.List[Object]]::new()

Write-Host "Analyzing group-based licensing..."
# Two outputs - a list containing group information and another with all the license data
ForEach ($Group in $Groups) {
    # Resolve product identifiers to product names for each license
    $ProductNames = @()
    ForEach ($License in $Group.AssignedLicenses) {
        $ConsumedUnits = 0; $PrepaidUnits = 0
        $ProductNames += $ProductInfoHash[$License.SkuId]
        $ConsumedUnits = ($TenantSkus | Where-Object {$_.SkuId -eq $License.SkuId}).ConsumedUnits
        $PrepaidUnits = ($TenantSkus | Where-Object {$_.SkuId -eq $License.SkuId}).PrepaidUnits.Enabled
        $GroupBasedSkuLine = [PSCustomObject][Ordered]@{
            SkuId           = $License.SkuId
            'Product name'  = $ProductInfoHash[$License.SkuId]
            ConsumedUnits   = $ConsumedUnits
            PrepaidUnits    = $PrepaidUnits
            AvailableUnits  = $PrepaidUnits - $ConsumedUnits
            'GroupId'       = $Group.Id
        }
        $GroupBasedSkusReport.Add($GroupBasedSkuLine)
    }
    [array]$GroupMembers = Get-MgGroupMember -GroupId $Group.Id -All 
    $ReportLine = [PSCustomObject][Ordered]@{
        DisplayName         = $Group.DisplayName
        GroupId             = $Group.Id
        Licenses            = $ProductNames -join ', '
        Members             = $GroupMembers.Count
        'Licensing errors'  = (Get-MgGroupMemberWithLicenseError -GroupId $Group.Id).Count
        'Member names'      = ($GroupMembers.additionalProperties.displayName -join ', ')
        AssignedLicenses    = $Group.AssignedLicenses
        'Processing state'  = $Group.LicenseProcessingState.State
    }
    $Report.Add($ReportLine)
}

Write-Host "Checking for user accounts with license assignment errors..."
[array]$UsersWithErrors = Get-MgUser -All -PageSize 500 -Property AssignedLicenses, LicenseAssignmentStates, DisplayName | `
    Select-Object DisplayName, AssignedLicenses -ExpandProperty LicenseAssignmentStates | `
    Select-Object DisplayName, AssignedByGroup, State, Error, SkuId | Where-Object {$_.Error -ne 'None'}

# Remove errors that aren't associated with group-based licensing
$UsersWithErrors = $UsersWithErrors | Where-Object {$_.AssignedByGroup -ne $null}

# Build a hash table to lookup group names for user accounts with license assignment errors
$GroupsHash = @{}
ForEach ($Group in $Groups) {
    $GroupsHash.Add($Group.Id, $Group.DisplayName)
}   

# HTML style (basic)
$HtmlHead="<html>
    <style>
    BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    H4{font-size: 8px; padding-top: 4px;}
</style>"
$HtmlBody = $HtmlHead + "<h1>Group-based licensing report</h1><p>This report details licenses assigned to tenant accounts through group-based licensing."

# Build a HTML message body part by looping through the set of SKUs found for group-based licensing and 
# report what we find
ForEach ($Sku in $GroupBasedSkusReport) {
    $GroupData = $Report | Where-Object {$_.GroupId -eq $Sku.GroupId}
    $HtmlHeader = ("<h2>Product: <u>{0}</u>" -f $Sku.'Product name') + "</h2><p>"
    $HtmlHeader = $HtmlHeader + ("<p><h2>License assignment through the <u>{0}</u> group</h2>" -f $GroupData.DisplayName) + "</p>"
    $HtmlHeader = $HtmlHeader + ("<p><h3>Consumed units: {0} Prepaid Units: {1} Available Units: {2} Assigned through group: {3}</h3> " `
        -f $Sku.ConsumedUnits, $Sku.PrepaidUnits, $Sku.AvailableUnits, $GroupData.Members) + "</p>"
    If ($Sku.AvailableUnits -le 0) {
        $HtmlHeader = $HtmlHeader + "<p><strong>Warning: No more licenses availble</strong></p>"
    } ElseIf ($Sku.AvailableUnits -lt 10) {
        $HtmlHeader = $HtmlHeader + "<p><strong>Warning: Less than 10 licenses available</strong></p>"
    }
    If ($GroupData.Members -gt $Sku.PrepaidUnits) {
        $HtmlHeader = $HtmlHeader + "<p><strong>Warning: More group members than prepaid licenses</strong></p>"
    }
    $HtmlBody = $HtmlBody + $HtmlHeader + ("<p>Licenses assigned to the following members: {0}" -f $GroupData.'Member names') + "</p>"
}
If (!$UsersWithErrors) {
    Write-Host "No user accounts with license assignment errors found"
    $HtmlUsersError = "<p>No user accounts with license assignment errors found</p>"
} Else {
    $HtmlUsersError = "<h2>User accounts with license assignment errors</h2><p>"
    $HtmlUsersError = $UsersWithErrors | Select-Object DisplayName, @{name='Group';expression={$GroupsHash[$_.AssignedByGroup]}}, `
        @{name='Product';expression={$ProductInfoHash[$_.SkuId]}}, Error | ConvertTo-Html -Fragment
    $HtmlUsersError = "<h2>User Accounts with License Assignment Errors</h2><p>" + $HtmlUsersError
}

$HtmlMsg = $HtmlBody + $HtmlUsersError + "<p><h4>Generated:</strong> $(Get-Date -Format 'dd-MMM-yyyy HH:mm')</h4></p>""<p></body>"

$MsgSubject = "Group-based licensing report"
# Email sent using the account that signed into the Graph session. Change this if you want the sender to be a different account.
# Sending from a different account means that the script must run in app-only mode and have consent to use the Mail.Send application permission
$MsgFrom = (Get-MgContext).Account
$ToRecipients = @{}
# Add the email addresses of the recipients to the ToRecipients hash table - change this address to get the email to where you want it to go
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

Try {
    Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params -ErrorAction Stop
    Write-Host "Report sent by email"
} Catch {
    Write-Host "Error sending email: $($_.Exception.Message)"
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.