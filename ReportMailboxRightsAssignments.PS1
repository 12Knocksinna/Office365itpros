# ReportMailboxRightsAssignments.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportMailboxRightsAssignments.PS1
# Quick script to find audit records for rights assigned to mailboxes to allow us to notify the mailbox owner
# about the assignment by sending email using the Microsoft Graph SDK for PowerShell. Needs version 1.7.0 or later
# of the Microsoft.Graph.Mail and Microsoft.Graph.Authentication modules (both part of the SDK).
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online management module and then restart the script"; break}
# Make sure that we have valid credentials
If (-not $O365Cred) { #Make sure we have credentials
    $O365Cred = (Get-Credential)}
# Message is from the logged in account - we're going to compare email addresses later, so make sure that we know the primary SMTP address of the account 
# rather than just its user principal name (used to sign in).
[string]$MsgFrom = (Get-ExoMailbox -Identity $O365Cred.UserName).PrimarySmtpAddress
$MsgSubject = "Notification of permission change to your mailbox"
#HTML header with styles
$HtmlHead="<html>
     <style>
      BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    </style>"

# Find the audit records for the last 30 days
$StartDate = (Get-Date).AddDays(-30); $EndDate = (Get-Date) 
Write-Host "Searching for audit records for mailbox permission update events..."
[array]$Records = Search-UnifiedAuditLog -Operations Add-MailboxPermission, Add-RecipientPermission -StartDate $StartDate -EndDate $EndDate -Formatted -ResultSize 5000
# If we find some records, process them
If (!$Records) {  Write-Host "No audit records for mailbox permissions found."; break }
$Records = $Records | Where-Object {$_.UserIds -ne "NT AUTHORITY\SYSTEM (Microsoft.Exchange.Servicehost)"}
Write-Host $Records.Count "audit records found..."
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
    $AuditData          = $Rec.AuditData | ConvertFrom-Json
    $TargetMailbox      = $AuditData.ObjectId
    $Admin              = $AuditData.UserId
    $TrusteeDisplayName = $Null; $MailboxType = $Null; $MailboxType = $Null; $UserDetailsEmail = $Null; $Access = $Null
   
    $TrusteeDetails = Get-ExoMailbox -Identity $Trustee -ErrorAction SilentlyContinue -IncludeInactiveMailbox
    $UserDetails    = Get-ExoMailbox -Identity $TargetDN -ErrorAction SilentlyContinue -IncludeInactiveMailbox
   
    If ($TrusteeDetails -eq $Null) {$TrusteeDetailsName = $Trustee } Else {$TrusteeDetailsName = $TrusteeDetails.DisplayName}

    Switch ($Rec.Operations) { # Set up data for either a mailbox permission or recipient permission operation
        "Add-MailboxPermission" {
            $TargetDN           = $AuditData.Parameters | ? {$_.Name -eq "Identity"} | Select-Object -ExpandProperty Value
            $Trustee            = $AuditData.Parameters | ? {$_.Name -eq "User"} | Select-Object -ExpandProperty Value
            [string]$Access     = $AuditData.Parameters | ? {$_.Name -eq "AccessRights"} | Select-Object -ExpandProperty Value
            $FullAccessList     = Get-ExoMailboxPermission -Identity $UserDetails.PrimarySmtpAddress | ? {$_.User -ne "NT AUTHORITY\SELF"} | Select User, AccessRights
            }
        "Add-RecipientPermission" {
            $TargetDN           = $AuditData.Parameters | ? {$_.Name -eq "Identity"} | Select-Object -ExpandProperty Value
            $Trustee            = $AuditData.Parameters | ? {$_.Name -eq "Trustee"} | Select-Object -ExpandProperty Value
            [string]$Access     = $AuditData.Parameters | ? {$_.Name -eq "AccessRights"} | Select-Object -ExpandProperty Value
            $FullAccessList     = Get-ExoRecipientPermission -Identity $UserDetails.PrimarySmtpaddress | ? {$_.Trustee -ne "NT AUTHORITY\SELF"} | Select Trustee, AccessRights
            }
     } #End Switch

    If ($UserDetails -eq $Null) {$UserDetailsName = $TargetDN} Else {
        $UserDetailsName = $UserDetails.DisplayName
        $UserDetailsEmail = $UserDetails.PrimarySmtpAddress
          
        $FullAccessReport = [System.Collections.Generic.List[Object]]::new()
        ForEach ($AccessPermission in $FullAccessList) {
           [string]$AccessRightsGranted = $AccessPermission.AccessRights
           Switch ($Rec.Operations) {
              "Add-MailboxPermission" {
                 $ReportLine = [PSCustomObject][Ordered]@{ 
                  User   = $AccessPermission.User
                  Access = $AccessRightsGranted }
                  $FullAccessReport.Add($ReportLine)
               }
              "Add-RecipientPermission" {
                $ReportLine = [PSCustomObject][Ordered]@{ 
                  User   = $AccessPermission.Trustee
                  Access = $AccessRightsGranted }
                  $FullAccessReport.Add($ReportLine)
               }
           } #end Switch
         } #End Foreach access permission
     } #End if
    If ($UserDetails.RecipientTypeDetails -eq $Null) { $MailboxType = "Unknown"}
    
    $ReportLine = [PSCustomObject][Ordered]@{ 
           TargetMailbox  = $UserDetailsName
           MailboxType    = $UserDetails.RecipientTypeDetails
           TargetEmail    = $UserDetailsEmail
           AccessGranted  = $Access
           GrantedOn      = $Rec.CreationDate
           GrantedTo      = $TrusteeDetailsName
           GrantedBy      = $Admin
           Operation      = $Rec.Operations
           FullAccessList = $FullAccessReport
    } # End ReportLine
    $Report.Add($ReportLine)
} #End ForEach $Records

# We now have a set of records for adding mailbox permissions. Before we start to send email, discard amy 
# records for deleted mailboxes where we don't have an email address and shared mailboxes. We should be left with 
# a set of permission changes made to user mailboxes.
[array]$EmailUsers = $Report | Where-Object {$_.TargetEmail -ne $Null  -and $_.MailboxType -eq "UserMailbox"} | Sort-object TargetMailbox
# Drop messages to the account sending email as there's no point in telling them something they already know
[array]$EmailUsers = $EmailUsers | Where-Object {$_.TargetEmail -ne $MsgFrom}
If (!($EmailUsers)) { Write-Host "No notifications found to send after analyzing audit records - exiting"; break}

# Connect to the Microsoft Graph SDK for PowerShell
Write-Host "Connecting to the Microsoft Graph"
Connect-MgGraph -Scope "Mail.Send, Mail.ReadWrite"

ForEach ($User in $EmailUsers) {
        Write-Host "Sending notification email to" $User.TargetMailbox
        # Add the recipient using the mailbox's primary SMTP address
        $EmailAddress  = @{address = $User.TargetEmail} 
        $EmailRecipient = @{EmailAddress = $EmailAddress}  
        # Customize the message   
        $ChangeDate = Get-Date($User.GrantedOn) -format r
        Switch ($User.AccessGranted) {
           "FullAccess" { $PermissionText = "Full Access permission grants $($User.GrantedTo) access to all items in your mailbox" }
           "SendAs"     { $PermissionText = "Send As permission allows $($User.GrantedTo) to impersonate you and send email as if the messages come from you." }
        }     
        # Message body
       $HtmlContent = "<body><html>
       <h1>Notification of mailbox permission change made to your mailbox</h1>
       <h2><u>Please contact the Help Desk if this permission update is not approved by you</u></h2>
       <p><b><u>Details of Permission Change</b></u></p>
       <p><strong>Change made on:      </strong> $($ChangeDate)</p>  
       <p><strong>Permission added:    </strong> $($User.AccessGranted)</p>
       <p><strong>Granted to:          </strong> $($User.GrantedTo)</p>
       <p><i>$($PermissionText)</i></p>
       <p><strong>Permission added by: </strong> $($User.GrantedBy)</p></body></html></p>
       <p><p><strong>Current set of users with $($User.AccessGranted) permission for the mailbox</strong></p>
       <p>$($User.FullAccessList.User -join ", ")</p>"
       $HtmlMsg = $HtmlHead + $HtmlContent
       # Construct the message body
       $MessageBody = @{
             content = "$($HtmlMsg)"
             ContentType = 'html'
             }
       # Create a draft message in the signed-in user's mailbox
       $NewMessage = New-MgUserMessage -UserId $MsgFrom -Body $MessageBody -ToRecipients $EmailRecipient -Subject $MsgSubject 
       # Send the message
       Send-MgUserMessage -UserId $MsgFrom -MessageId $NewMessage.Id  
  } # End ForEach User
  
Write-Host "All done. Messages sent!"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
