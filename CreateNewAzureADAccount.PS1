# CreateNewAzureADAccount.PS1
# Example of how to create a new Azure AD Account and email creation details to tenant admins

Function Populate-MessageRecipient { # Build a list of recipients for a message
    [cmdletbinding()]
    Param(
        [array]$ListOfAddresses )
    ForEach ($SMTPAddress in $ListOfAddresses) {
        @{
            emailAddress = @{address = $SMTPAddress}
        }    
    }    
}

Connect-MgGraph -Scopes User.ReadWrite.All, Directory.ReadWrite.All
Select-MgProfile Beta

$Office365E3Sku = "6fd2c87f-b296-42f0-b197-1e91e994b900"

# Find admins to email
$AdminRoleHolders = [System.Collections.Generic.List[Object]]::new() 
[array]$AdminRoles = Get-MgDirectoryRole | Select-Object DisplayName, Id | Sort-Object DisplayName
$AdminRoles = $AdminRoles | Where-Object {$_.DisplayName -eq "Global Administrator" -or $_.DisplayName -eq "User Administrator"}
ForEach ($Role in $AdminRoles) {
    [array]$RoleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $Role.Id | ? {$_.AdditionalProperties."@odata.type" -eq "#microsoft.graph.user"}
    ForEach ($Member in $RoleMembers) {
      $UserDetails = Get-MgUser -UserId $Member.Id
      $ReportLine  = [PSCustomObject] @{   
         User   = $UserDetails.UserPrincipalName
         Id     = $UserDetails.Id
         Role   = $Role.DisplayName
         RoleId = $Role.Id
         Mail   = $UserDetails.Mail }
     $AdminRoleHolders.Add($ReportLine) }
}
$AdminRoleHolders = $AdminRoleHolders | Sort User
$Admins = $AdminRoleHolders | Sort User -Unique

Add-Type -AssemblyName 'System.Web'
$NewPassword = [System.Web.Security.Membership]::GeneratePassword(10, 3)
$NewPasswordProfile = @{}
$NewPasswordProfile["Password"]= $NewPassword
$NewPasswordProfile["ForceChangePasswordNextSignIn"] = $True
$DisplayName = "Jeff Atkinson"
$NewUser = New-MgUser -UserPrincipalName "Jeff.Atkinson@Office365ITPros.com" `
  -DisplayName "Jeff Atkinson (Information Technology)" `
  -PasswordProfile $NewPasswordProfile -AccountEnabled `
  -MailNickName Jeff.Atkinson -City NYC `
  -CompanyName "Office 365 for IT Pros" -Country "United States" `
  -Department "IT Operations" -JobTitle "GM Operations" `
  -BusinessPhones "+1 676 830 1101" -MobilePhone "+1 617 4466615" `
  -State "New York" -StreetAddress "1, Avenue of the Americas" `
  -Surname "Atkinson" -GivenName "Jeff" `
  -UsageLocation "US" -OfficeLocation "NYC"
If ($NewUser) { Write-Host ("Successfully added the {0} account" -f $NewUser.DisplayName) 
  } Else { Write-Host ("Failure adding the {0} account - exiting" -f $DisplayName); break }

# Add a license to the new account
$License = Set-MgUserLicense -UserId $NewUser.Id -AddLicenses @{SkuId = $Office365E3Sku } -RemoveLicenses @()
If ($License) { Write-Host ("Successfully assigned Office 365 E3 license to {0}" -f $NewUser.DisplayName)
  } Else { Write-Host ("Failed to assign Office 365 License to {0}" -f $NewUser.DisplayName) }

$HtmlBody = ("A new user account for {0} ({1}) has been added. The account password is <b>{2}</b>. An Office 365 E3 license has been assigned to the account. Please inform the user and inform them that they must reset their password when they sign into Office 365." -f $NewUser.DisplayName, $NewUser.UserPrincipalName, $NewPassword)
# Email the admins
$MsgFrom = "Administrator@redmondassociates.org"
$MsgSubject = ("New Azure AD account added for {0} ({1})" -f $NewUser.DisplayName, $NewUser.UserPrincipalName)
[array]$MsgToRecipients = Populate-MessageRecipient -ListOfAddresses $Admins.Mail
Write-Host "Sending notification message to tenant admins"
$HtmlHeaderUser = "<h2>New User " + $NewUser.DisplayName + "</h2>"    
$HtmlMsg = "</body></html>" + $HtmlHead + $htmlheaderuser + $htmlbody + "<p>"
# Construct the message body
$MsgBody = @{
     Content = "$($HtmlBody)"
     ContentType = 'html'   }
$Message =  @{subject           = $MsgSubject}
$Message += @{toRecipients      = $MsgToRecipients}  
$Message += @{body              = $MsgBody}
$Params   = @{'message'         = $Message}
$Params  += @{'saveToSentItems' = $True}
$Params  += @{'isDeliveryReceiptRequested' = $True}
Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params

#----------------------------------------#

# Example of how to add accounts after reading them in from a CSV file

$Accounts = Import-CSV c:\temp\Accounts.CSV
ForEach ($Account in $Accounts) {
  $NewPassword = [System.Web.Security.Membership]::GeneratePassword(10, 3)
  $NewPasswordProfile = @{}
  $NewPasswordProfile["Password"]= $NewPassword
  $NewPasswordProfile["ForceChangePasswordNextSignIn"] = $True
  $MailNickname = $Account.First + "." + $Account.Surname
  $DisplayName = $Account.First + " " + $Account.Surname
  Write-Host ("Processing the {0} account" -f $DisplayName)
  $NewUser = New-MgUser -UserPrincipalName $Account.UserPrincipalName `
  -DisplayName $DisplayName `
  -PasswordProfile $NewPasswordProfile `
  -MailNickName $MailNickName -City $Account.City `
  -CompanyName $Account.Company -Country $Account.Country `
  -Department $Account.Department -JobTitle $Account.Title `
  -BusinessPhones $Account.Phone -MobilePhone $Account.Mobile `
  -State $Account.State -StreetAddress $Account.Street `
  -Surname $Account.Surname -GivenName $Account.First `
  -UsageLocation $Account.Location -OfficeLocation $Account.Office `
  -AccountEnabled
 If ($NewUser) { Write-Host ("Successfully added the {0} account" -f $NewUser.DisplayName) 
  } Else { Write-Host ("Failure adding the {0} account - exiting" -f $DisplayName); break }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
