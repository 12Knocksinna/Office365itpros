# SendWelcomeEmail-Runbook.PS1
# Example of how to send email using the Microsoft Graph SDK for PowerShell via an Azure Automation account runbook
# https://github.com/12Knocksinna/Office365itpros/blob/master/SendWelcomeEmail-Runbook.ps1
#
# Create an automation connection to our RunAs account
$Connection = Get-AutomationConnection -Name AzureRunAsConnection
# Get the certificate for the automation account
$Certificate = Get-AutomationCertificate -Name AzureRunAsCertificate
# Connect to the Graph using the automation account
Connect-MgGraph -ClientID $Connection.ApplicationId -TenantId $Connection.TenantId -CertificateThumbprint $Connection.CertificateThumbprint
$Organization = Get-MgOrganization
$TenantName = $Organization.DisplayName
$TenantDomain = $Organization.Verifieddomains | ? {$_.IsInitial -eq $True} | Select -ExpandProperty Name
# Connect to Exchange Online
Connect-ExchangeOnline –CertificateThumbprint $Connection.CertificateThumbprint –AppId $Connection.ApplicationID –ShowBanner:$false –Organization $TenantDomain
$Profile = (Get-MgProfile).Name
If ($Profile -ne "beta") { Select-MgProfile Beta }

# Who the message is from - this has to be an account in your tenant with a mailbox
$MsgFrom = "HR.Admin.Officer@offic365itpros.com"
# Fetch the attachment from a web file, create a locate copy in a throwaway folder on the sandbox machine and 
# load it into an encoded form to attach to our message
$WebAttachmentFile = "https://office365itpros.com/wp-content/uploads/2022/02/WelcomeToOffice365ITPros.docx"
New-Item -Path c:\TempForScriptxxx -ItemType directory -ErrorAction SilentlyContinue
$AttachmentFile = "c:\TempForScriptxxx\WelcomeNewEmployeeToOffice365itpros.docx"
Invoke-WebRequest -uri $WebAttachmentFile -OutFile $AttachmentFile
$Attachment = (Get-Item -Path $AttachmentFile).Name
$EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($AttachmentFile))

$MsgAttachment = @{
    "@odata.type"= "#microsoft.graph.fileAttachment"
    name = ($AttachmentFile -split '\\')[-1]
    contentBytes = $EncodedAttachmentFile
}
# Hardcoded tenant name for screen shot purposes, which is why it's commented out
# $TenantName = "an Azure Automation runbook on behalf of Practical365.com"
# Define some variables used to construct the HTML content in the message body
#HTML header with styles
$htmlhead="<html>
     <style>
      BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    </style>"
#Content for the message
$HtmlBody = "<body>
     <h1>Welcome to $($TenantName)</h1>
     <p><strong>Generated:</strong> $(Get-Date -Format g)</p>  
     <h2><u>We're Pleased to Have You Here</u></h2>
     <p><b>Welcome to your new Office 365 account</b></p>
     <p>You can open your account to access your email and documents by clicking <a href=http://www.portal.office.com>here</a> </p>
     <p>Have a great time and be sure to call the help desk if you need assistance. And be sure to read all the great articles about Office 365 published on Practical365.com.</p>"
$MsgSubject = "A warm welcome from $($TenantName)"

# Date to Check for new accounts - we use the last 90 days here, but that's easily changable.
[string]$CheckDate = (Get-Date).AddDays(-90)
# Find all mailboxes created in the target period
[array]$Users = (Get-ExoMailbox -Filter "WhenMailboxCreated -gt '$CheckDate'" -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties WhenMailboxCreated | Select WhenMailboxCreated, DisplayName, UserPrincipalName, PrimarySmtpAddress)

# Create and send welcome email message to each of the new mailboxes
ForEach ($User in $Users) {
      # Add the recipient using the mailbox's primary SMTP address
      $EmailAddress  = @{address = $User.PrimarySmtpAddress} 
      $EmailRecipient = @{EmailAddress = $EmailAddress}  
      # Customize the message 
      $htmlHeaderUser = "<h2>New User " + $User.DisplayName + "</h2>"    
      $HtmlMsg = "</body></html>" + $HtmlHead + $htmlheaderuser + $htmlbody + "<p>"
      # Construct the message body
      $MessageBody = @{
           content = "$($HtmlBody)"
           ContentType = 'html'
           }
     # Create a draft message in the signed-in user's mailbox
     $NewMessage = New-MgUserMessage -UserId $MsgFrom -Body $MessageBody -ToRecipients $EmailRecipient -Subject $MsgSubject -Attachments $MsgAttachment
     # Send the message
     Send-MgUserMessage -UserId $MsgFrom -MessageId $NewMessage.Id  
} # End ForEach User

Write-Output "All done. Messages sent!"
$Users | Format-Table DisplayName, PrimarySmtpAddress
# Cleanup
Remove-Item $AttachmentFile
Remove-Item "c:\TempForScriptxxx"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# checking the code
