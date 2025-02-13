# SendWelcomeMessage-UserMail.PS1
# Github link: https://github.com/12Knocksinna/Office365itpros/blob/master/SendWelcomeMessage-UserMail-Mg.PS1
# How to send a welcome message to new mailboxes using the Send-MgUserMail cmdlet from the Microsoft Graph SDK for PowerShell
# Needs version 1.7.0 or later of the Microsoft.Graph.Mail and Microsoft.Graph.Users.Actions modules

# V1.1 12-Feb-2025: Updated function to process attachments to make sure that content type is better defined

Function Update-MessageRecipients {
    [cmdletbinding()]
    Param(
        [array]$ListOfAddresses )
    ForEach ($SMTPAddress in $ListOfAddresses) {
        @{
            emailAddress = @{address = $SMTPAddress}
        }    
    }    
}

Function Update-MessageAttachments {
    [cmdletbinding()]
    Param(
        [array]$ListOfAttachments 
    )

    [array]$MsgAttachments = $null
    ForEach ($File in $ListOfAttachments) {
        $ConvertedContent = [Convert]::ToBase64String([IO.File]::ReadAllBytes($File))
        $FileExtension = [System.IO.Path]::GetExtension($File) 
        Switch ($FileExtension) {
            ".pdf" {
                $ContentType = "application/pdf"
            }
            ".docx" {
                $ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            }
            ".xlsx" {
                $ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }   
            ".pptx" {
                $ContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            }   
            ".jpg" {
                $ContentType = "image/jpeg"
            }   
            ".png" {
                $ContentType = "image/png"
            }   
            default {
                $ContentType = "application/octet-stream"
            }
        }
        $AttachmentDetails = @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            Name = $File
            ContentType =  $ContentType
            ContentBytes = $ConvertedContent
        }
        $MsgAttachments += $AttachmentDetails
    }
    Return $MsgAttachments  
}

# Main processing
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {
    Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break
}
# Now connect to the Microsoft Graph SDK for PowerShell
Connect-MgGraph -Scope Mail.Send

# Message is from the signed-in account - change this if necessary (and you're running in app-only mode)
$MsgFrom = (Get-MgContext).Account

# Define some variables used to construct the HTML content in the message body
#HTML header with styles
$htmlhead="<html><style>BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    </style>"

#Content for the message - obviously, this is very customizable and should reflect what you want to say to new users
$HtmlBody = "<body>
     <h1>Welcome to $($TenantName)</h1>
     <p><strong>Generated:</strong> $(Get-Date -Format g)</p>  
     <h2><u>We're Pleased to Have You Here</u></h2>
     <p><b>Welcome to your new Office 365 account</b></p>
     <p>You can open your account to access your email and documents by clicking <a href=http://www.portal.office.com>here</a> </p>
     <p>Have a great time and be sure to call the help desk if you need assistance. And be sure to read all the great articles about Office 365 published on Practical365.com.</p>"
$MsgSubject = "A warm welcome to $($TenantName)"

# Date to Check for new accounts - we use the last 7 days here, but that's easily changable.
[string]$CheckDate = (Get-Date).AddDays(-7)
# Find all mailboxes created in the target period
[array]$Users = Get-ExoMailbox -Filter "WhenMailboxCreated -gt '$CheckDate'" -RecipientTypeDetails UserMailbox -ResultSize Unlimited `
     -Properties WhenMailboxCreated | Select-Object WhenMailboxCreated, DisplayName, UserPrincipalName, PrimarySmtpAddress
If (!($Users)) { 
    Write-Host ("No mailboxes found that were created before {0}. Exiting!" -f $CheckDate) 
}

# Define attachments  we're only using one here, which we fetch from a web site
# if you want to add more files, add the file names to the $AttachmentsList array
$WebAttachmentFile = "https://office365itpros.com/wp-content/uploads/2022/02/WelcomeToOffice365ITPros.docx"
$AttachmentFile = "c:\temp\WelcomeNewEmployeeToOffice365itpros.docx"
Invoke-WebRequest -Uri $WebAttachmentFile -OutFile $AttachmentFile

# Add as many attachments as you want to this array and they'll be added to the message
[array]$AttachmentsList = "c:\temp\WelcomeNewEmployeeToOffice365itpros.docx"
[array]$MsgAttachments = Update-MessageAttachments -ListOfAttachments $AttachmentsList

# Populate CC Recipienmt List - any SMTP address will do
$CcRecipientList = @( 
        'Tenant.Admins@office365itpros.com'
        'Kim.Akers@Office365itpros.com' )
[array]$MsgCcRecipients = Update-MessageRecipients -ListOfAddresses $CcRecipientList

ForEach ($User in $Users) {
  $ToRecipientList   = @( $User.PrimarySmtpAddress )
  [array]$MsgToRecipients = Update-MessageRecipients -ListOfAddresses $ToRecipientList
  Write-Host "Sending welcome email to" $User.DisplayName
  # Customize the message 
  $htmlHeaderUser = "<h2>New User " + $User.DisplayName + "</h2>"    
  $HtmlMsg = "</body></html>" + $HtmlHead + $htmlheaderuser + $htmlbody + "<p>"
  # Construct the message body
  $MsgBody = @{
     Content = "$($HtmlMsg)"
     ContentType = 'html'   }

  $Message =  @{subject           = $MsgSubject}
  $Message += @{toRecipients      = $MsgToRecipients}  
  $Message += @{ccRecipients      = $MsgCcRecipients}  
  $Message += @{attachments       = $MsgAttachments}
  $Message += @{body              = $MsgBody}
  $Params   = @{'message'         = $Message}
  $Params  += @{'saveToSentItems' = $True}
  $Params  += @{'isDeliveryReceiptRequested' = $True}

  Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params
}
Write-Host "All done. Messages sent!"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
