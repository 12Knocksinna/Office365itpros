# Report-DailySignIns.PS1
# An example of how to analyze and report daily sugn-ins recorded in the Entra ID sign-in audit log. The script generates a bunch of information
# from the sign-in data and sends email to the tenant administratords

# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-DailySignIns.PS1
# V1.0 1-Oct-2025

# Access to risky users information requires the IdentityRiskyUser.Read.All permission and an admin role like Global Reader or Security Reader

Connect-MgGraph -Scopes "AuditLog.Read.All", "User.Read.All", "Organization.Read.All", "IdentityRiskyUser.Read.All"

# Get tenant identifier
$TenantData = Get-MgOrganization
$TenantId = $TenantData.Id

# Set dates up to find what happened yesterday
$Today = Get-Date
$StartDate=$Today.AddDays(-30).ToString('yyyy-MM-ddT00:00:00Z')
$EndDate=$Today.ToString('yyyy-MM-ddT00:00:00Z')

Write-Output "Retrieving interactive sign-in records from $StartDate to $EndDate"
[array]$SignInRecords = Get-MgBetaAuditLogSignIn -All -Filter "createdDateTime ge $StartDate and createdDateTime lt $EndDate" -PageSize 999
If ($SignInRecords.Count -eq 0) {
    Write-Output "No sign-in records found"
    Break
}   Else {
    Write-Output ("Found {0} sign-in records to process..." -f $SignInRecords.Count)
    $Report = [System.Collections.Generic.List[Object]]::new()  
}

# Get Risky Users Information
[array]$RiskyUsers = Get-MgRiskyUser | Where-Object {$_.RiskDetail -eq "none" -and $_.RiskState -eq "atRisk"} | Sort-Object {$_.RiskLastUpdatedDateTime -as [datetime]} -Descending
ForEach ($RiskyUser in $RiskyUsers) {
    
    $ReportLine = [PSCustomObject]@{
        UserPrincipalName   = $RiskyUser.UserPrincipalName
        RiskLevel           = $RiskyUser.RiskLevel
        RiskState           = $RiskyUser.RiskState
        RiskDetail          = $RiskyUser.RiskDetail
        RiskLastUpdated     = $RiskyUser.RiskLastUpdatedDateTime
        DaysAtRisk          = (New-TimeSpan -Start $RiskyUser.RiskLastUpdatedDateTime -End (Get-Date)).Days
    }
    $Report.Add($ReportLine)
}   

[array]$AppSignIns = $SignInRecords | Where-Object { $_.AppDisplayName -ne $null -and $_.Status.ErrorCode -eq 0 } | Group-Object -Property AppDisplayName -NoElement 
[array]$UserSignIns = $SignInRecords | Where-Object { $_.UserDisplayName -ne $null -and $_.Status.ErrorCode -eq 0} | Group-Object -Property UserDisplayName -NoElement 
# Some failed sign-ins have a null userPrincipalName, notably when a tennat account fails with an attempt using B2B authentication with another tenant
[array]$FailedSignIns = $SignInRecords | Where-Object { $_.Status.ErrorCode -ne 0 -and -not [string]::IsNullOrWhiteSpace($_.userPrincipalName) } | Group-Object -Property UserDisplayName -NoElement 
[array]$SingleFactorSignIns = $SignInRecords | Where-Object { $_.AuthenticationRequirement -eq "singleFactorAuthentication" -and $_.Status.ErrorCode -eq 0 } 
[array]$MfaSignIns = $SignInRecords | Where-Object { $_.ConditionalAccessStatus -eq "Success" -and $_.Status.ErrorCode -eq 0 } 
[array]$IncomingGuestMemberSignIns = $SignInRecords | Where-Object { $_.UserType -eq "Guest" -and $_.Status.ErrorCode -eq 0 -and $_.ResourceTenantId -eq $TenantId }
[array]$OutboundGuestMemberSignIns = $SignInRecords | Where-Object { $_.UserType -eq "Guest" -and $_.Status.ErrorCode -eq 0 -and $_.ResourceTenantId -ne $TenantId }

[array]$MfaUsers = $MfaSignIns | Group-Object userPrincipalName -NoElement  | Select-Object -ExpandProperty Name
[array]$SingleFactorUsers = $SingleFactorSignIns | Group-Object userPrincipalName -NoElement  | Select-Object -ExpandProperty Name
[array]$SingleFactorApps = $SingleFactorSignIns | Group-Object AppDisplayName -NoElement | Select-Object -ExpandProperty Name
[int]$CASuccess = $SignInRecords | Where-Object { $_.ConditionalAccessStatus -eq "Success" -and $_.Status.ErrorCode -eq 0 } | Measure-Object | Select-Object -ExpandProperty Count
[array]$CAFailureEvents = $SignInRecords | Where-Object { $_.ConditionalAccessStatus -eq "Failure" }
[int]$CANotApplied = $SignInRecords | Where-Object { $_.ConditionalAccessStatus -eq "NotApplied" -and $_.Status.ErrorCode -eq 0 } | Measure-Object | Select-Object -ExpandProperty Count
[int]$CAFailures = $CAFailureEvents.count
[array]$CAFailureUsers = $CAFailureEvents | Group-Object UserDisplayName -NoElement  | Select-Object -ExpandProperty Name

# Build attachnent for the email using signin data

If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $true
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Daily Sign In Report.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    } 
    $SignInRecords | Export-Excel -Path $ExcelOutputFile -WorksheetName "Sign-in Records" -Title ("Daily Sign In Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "SignInRecords" -AutoSize -AutoFilter
    $AttachmentFile = $ExcelOutputFile
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Daily Sign In Report.CSV"
    $SignInRecords   | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
    $AttachmentFile = $CSVOutputFile
}

If ($ExcelGenerated) {
    Write-Output ("Excel worksheet output written to {0}" -f $ExcelOutputFile)
} Else {
    Write-Output ("CSV output file written to {0}" -f $CSVOutputFile)
} 

# Send the spreadsheet as an email attachment
$EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($AttachmentFile))
$MsgAttachments = @(
    @{
        '@odata.type' = '#microsoft.graph.fileAttachment'
        Name = (Split-Path $AttachmentFile -Leaf)
        ContentBytes = $EncodedAttachmentFile
        ContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
)

Write-Output ""
Write-Output ("Daily Sign-In Report for {0}" -f (Get-Date $EndDate -format 'dd MMMM yyyy'))
Write-Output ""
Write-Output "Top ten Applications signed into during the day"
Write-Output "-----------------------------------------------" 
$AppSignIns | Sort-Object Count -Descending | Select-Object -First 10 | Format-Table Name, Count -AutoSize
Write-Output ""
Write-Output "Top five users with successful sign-ins during the day"
Write-Output "------------------------------------------------------"
$UserSignIns | Sort-Object Count -Descending | Select-Object -First 5 | Format-Table Name, Count -AutoSize
Write-Output ""
Write-Output "Top five users with failed sign-ins during the day"
Write-Output "--------------------------------------------------"
$FailedSignIns | Sort-Object Count -Descending | Select-Object -First 5 | Format-Table Name, Count -AutoSize
Write-Output ""
Write-Output ("Number of users who signed in using single-factor authentication:  {0}" -f ($SingleFactorUsers.count))
Write-Output ("Number of users who signed in using multi-factor authentication:   {0}" -f ($MfaUsers.count))
Write-Output ("Accounts signed in using single-factor authentication:             {0}" -f ($SingleFactorUsers -join ", "))
Write-Output ("Apps accessed using single-factor authentication:                  {0}" -f ($SingleFactorApps -join ",  "))  
Write-Output ("Incoming guest user sign-ins:                                      {0}" -f ($IncomingGuestMemberSignIns.count))
Write-Output ("Outbound guest user sign-ins:                                      {0}" -f ($OutboundGuestMemberSignIns.count))
Write-Output ("Number of sign-ins where Conditional Access succeeded:             {0}" -f $CASuccess)
Write-Output ("Number of sign-ins where Conditional Access was not applied:       {0}" -f $CANotApplied)
Write-Output ("Number of sign-ins where Conditional Access failed:                {0}" -f $CAFailures) 
Write-Output ("Users with sign-ins where Conditional Access failed:               {0}" -f ($CAFailureUsers -join ", "))
Write-Output ""

# Prepare HTML fragments for inclusion in the email body
$TopApps = $AppSignIns | Sort-Object Count -Descending | Select-Object -First 10
$AppHTML = $TopApps |
    Select-Object @{Name='Application';Expression={$_.Name}},
                  @{Name='Sign In Count';Expression={$_.Count}} |
    ConvertTo-Html -As Table -Fragment

$TopUsers = $UserSignIns | Sort-Object Count -Descending | Select-Object -First 5
$UserHTML = $TopUsers |
    Select-Object @{Name='User';Expression={$_.Name}},
                  @{Name='Sign In Count';Expression={$_.Count}} |
    ConvertTo-Html -As Table -Fragment

$TopFailedSignInUsers = $FailedSignIns | Sort-Object Count -Descending | Select-Object -First 5 
$FailedSignInsHTML = $TopFailedSignInUsers |
    Select-Object @{Name='User';Expression={$_.Name}},
                  @{Name='Failed Sign In Count';Expression={$_.Count}} |
    ConvertTo-Html -As Table -Fragment

[string]$SingleFactorAppsList = $SingleFactorApps -join ", "
[string]$SingleFactorUsersList = $SingleFactorUsers -join ", "
[string]$CAFailureUsersList = $CAFailureUsers -join ", "

$MsgFrom = 'Customer.Services@office365itpros.com'
$DestinationEmailAddress = "Admin.Info@office365itpros.com"
# Build the array of a single TO recipient detailed in a hash table - change the sender and recipient to the appropriate recipient for your tenant
$ToRecipient = @{}
$ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
[array]$MsgTo = $ToRecipient
# Define the message subject
$MsgSubject = "Important: Daily Sign-in Report for the {0} tenant on {1}" -f $TenantData.DisplayName, (Get-Date $EndDate -format 'dd MMMM yyyy')
# Create the HTML content
$HtmlMsg = "</body></html><p>The output file for the <b>Daily Sign-in Report</b> is attached to this message. Please review the information at your convenience</p>"
$HtmlMsg += "<h2>Service Principals Analysis Summary</h2>"
$HtmlMsg += "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;'>"
$HtmlMsg += "<tr><th align='left'>Metric</th><th align='left'>Value</th></tr>"
$HtmlMsg += "<tr><td>Total Sign-ins</td><td>{0}</td></tr>" -f $SignInRecords.Count
$HtmlMsg += "<tr><td>Unique Applications Signed Into</td><td>{0}</td></tr>" -f $AppSignIns.Count
$HtmlMsg += "<tr><td>Unique Users Signed In</td><td>{0}</td></tr>" -f $UserSignIns.Count
$HtmlMsg += "<tr><td>Failed Sign-ins</td><td>{0}</  td></tr>" -f $FailedSignIns.Count
$HtmlMsg += "<tr><td>Single-Factor Authentication Sign-ins</td><td>{0}</td></tr>" -f $SingleFactorSignIns.Count
$HtmlMsg += "<tr><td>Unique Users using Single-Factor Authentication</td><td>{0}</td></tr>" -f $SingleFactorUsers.Count
$HtmlMsg += ("<tr><td>Unique Users using Single-Factor Authentication</td><td>{0}</td></tr>" -f $SingleFactorUsersList)
$HtmlMsg += "<tr><td>Unique Applications using Single-Factor Authentication</td><td>{0}</td></tr>" -f $SingleFactorApps.Count
$HtmlMsg += ("<tr><td>Unique Applications using Single-Factor Authentication</td><td>{0}</td></tr>" -f $SingleFactorAppsList)
$HtmlMsg += "<tr><td>Multi-Factor Authentication Sign-ins</td><td   >{0}</td></tr>" -f $MfaSignIns.Count
$HtmlMsg += "<tr><td>Multi-Factor Authentication Users</td><td>{0}</td></tr>" -f $MfaUsers.Count
$HtmlMsg += "<tr><td>Conditional Access Success Sign-ins</td><td>{0}</td></tr>" -f $CASuccess
$HtmlMsg += "<tr><td>Conditional Access Not Applied Sign-ins</td><td>{0}</td></tr>" -f $CANotApplied
$HtmlMsg += "<tr><td>Conditional Access Failure Sign-ins</td><td>{0}</td></tr>" -f $CAFailures
$htmlMsg += ("<tr><td>Users with Conditional Access Failure Sign-ins</td><td>{0}</td></tr>" -f $CAFailureUsersList)
$HtmlMsg += "</table>"      
$HtmlMsg += "<p><h3>Top ten Applications signed into during the day</h3></p>"
$HtmlMsg += $AppHTML 
$HtmlMsg += "<p><h3>Top five users with successful sign-ins during the day</h3></p>"
$HtmlMsg += $UserHTML
$HtmlMsg += "<p><h3>Top five users with failed sign-ins during the day</h3></p>"
$HtmlMsg += $FailedSignInsHTML
$htmlMsg += "<p><h3>Recent Risky Users Identified</h3></p>"
$htmlMsg += $Report | ConvertTo-Html -As Table -Fragment
$htmlMsg += "</body></html>"

# Construct the message body 	
$MsgBody = @{}
$MsgBody.Add('Content', "$($HtmlMsg)")
$MsgBody.Add('ContentType','html')
# Build the parameters to submit the message
$Message = @{}
$Message.Add('subject', $MsgSubject)
$Message.Add('toRecipients', $MsgTo)
$Message.Add('body', $MsgBody)
$Message.Add("attachments", $MsgAttachments)

$EmailParameters = @{}
$EmailParameters.Add('message', $Message)
$EmailParameters.Add('saveToSentItems', $true)
$EmailParameters.Add('isDeliveryReceiptRequested', $true)

# Send the message
Try {
    Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters -ErrorAction Stop
    Write-Output ("Daily sign-ins analysis report emailed to {0}" -f $ToRecipient.emailAddress.address)
    Write-Output "All done!"
} Catch {
    Write-Output "Unable to send email"
    Write-Output $_.Exception.Message
}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.