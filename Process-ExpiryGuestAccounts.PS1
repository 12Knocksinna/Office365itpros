# Process-ExpiryGuestAccounts.PS1
# Script to process all guest accounts in the tenant to:
# 1. Remove expired guest accounts.
# 2. Find guest accounts due to expire in the next 30 days
# 3. Email administrators with details of accounts to check
# Uses the Microsoft Graph PowerShell SDK 
# https://github.com/12Knocksinna/Office365itpros/blob/master/Process-ExpiryGuestAccounts.PS1

# Connect to Microsoft Graph in Azure Automation
Connect-AzAccount -Identity
$AccessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
Connect-MgGraph -AccessToken $AccessToken.Token -Scopes User.ReadWrite.All
Select-MgProfile Beta
# Connect to Exchange Online
# Connect-ExchangeOnline -ManagedIdentity -Organization redmondassociates.onmicrosoft.com
# Get display name for the tenant
$TenantName = (Get-MgOrganization).displayName
# Define variables
[datetime]$CheckDate = (Get-Date).AddDays(14)
[datetime]$Now = Get-Date
[datetime]$NewExpirationDate = (Get-Date).AddDays(120)
[datetime]$Check30 = (Get-Date).AddDays(-30)
[int]$i = 0
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report
[array]$Guests = Get-MgUser -Filter "userType eq 'Guest'" -All -Property SigninActivity

ForEach ($Guest in $Guests) {
  $i++
  #Write-Output ("Processing {0} {1}/{2}" -f $Guest.displayName, $i, $Guests.count)
  $UserLastSignInDate = $Null
  [datetime]$GuestExpirationDate = $Guest.onPremisesExtensionAttributes.extensionAttribute15
  $DaysExpired = $Null
  # Is account already marked as expired?
  If ($Guest.onPremisesExtensionAttributes.extensionAttribute14 -eq "Expired") { 
  # Guest account is expired and can be deleted 7 days after expiration 
     $DaysExpired = ($GuestExpirationDate | New-TimeSpan).Days
     If ($DaysExpired -ge 7) {
        Write-Output ("Removing guest account {0}" -f $Guest.displayname)
        Try {
          Remove-MgUser -UserId $Guest.Id 
          $ReportLine = [PSCustomObject]@{ 
            Timestamp       = Get-Date -format s
            Id              = $Guest.Id
            Action          = "Account deleted"
            Name            = $Guest.DisplayName
            Mail            = $Guest.Mail 
            Expiration      = $Guest.onPremisesExtensionAttributes.extensionAttribute15 
            LastSignIn      = "N/A" }
         $Report.Add($ReportLine)
         }
       Catch {
          Write-Output ("Error removing expired account{0} with expiration date of {1}" -f $Guest.displayName, $GuestExpirationDate) 
       }
     } # End days expired check
   }  # End processing section to remove expired accounts
   
   # Now check for accounts past their expiration date
   If (($Now -ge $GuestExpirationDate) -and ($Guest.onPremisesExtensionAttributes.extensionAttribute14 -ne "Expired")) {
    Write-Output ("Detected expired guest account {0} with an expiration date of {1}" -f $Guest.displayName, $Guest.onPremisesExtensionAttributes.extensionAttribute15)
    # Check last sign in date. If less than 30 days ago, extend the account by 120 days
    If ($Guest.SignInactivity.LastNonInteractiveSignInDateTime) {
      [datetime]$UserSignInDate = ($Guest.SignInactivity.LastNonInteractiveSignInDateTime) 
    } Else {
      [datetime]$UserSignInDate = ($Guest.createdDateTime) }   
    
    If ($UserSignInDate -ge $Check30) { 
      # Extend the expiration date 
      Update-MgUser -UserId $Guest.Id -OnPremisesExtensionAttributes @{'extensionAttribute15' = (Get-Date $NewExpirationDate -format s)}
      $ReportLine = [PSCustomObject]@{ 
        Timestamp            = Get-Date -format s
        Id                   = $Guest.Id
        Action               = "Account extended"
        Name                 = $Guest.DisplayName
        Mail                 = $Guest.Mail 
        Expiration           = Get-Date($NewExpirationDate) -format g
        LastSignIn           = $UserSignInDate }
      $Report.Add($ReportLine) 
     } Else { # Mark the account as expired so that it will be removed the next time this job runs
      Update-MgUser -UserId $Guest.Id -OnPremisesExtensionAttributes @{'extensionAttribute14' = 'Expired'}
      $ReportLine = [PSCustomObject]@{ 
        Timestamp            = Get-Date -format s
        Id                   = $Guest.Id
        Action               = "Account due to expire"
        Name                 = $Guest.DisplayName
        Mail                 = $Guest.Mail 
        Expiration           = $Guest.onPremisesExtensionAttributes.extensionAttribute15 
        LastSignIn           = "N.A" }
      $Report.Add($ReportLine) 
     } # End if sign-in data      
   } # End check to mark expired accounts 
   
} # End ForEach

$Report | Format-Table Id, Name, Action, Expiration, LastSignIn

# Define variables for the mailbox used to send the message, the recipient, and the message subject
$MsgFrom = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -Name "ExoAccountName" -AsPlainText
$ToAddress = "James.Smith@office365itpros.com"
$MsgSubject = "Guest Account Expiration Information for $($TenantName)"

# Define HTML header with styles
$htmlhead="<style>
	.UserTable {
		border:1px solid #C0C0C0;
		border-collapse:collapse;
		padding:5px;
	}
	.UserTable th {
		border:1px solid #C0C0C0;
		padding:5px;
		background:#F0F0F0;
	}
	.UserTable td {
		border:1px solid #C0C0C0;
		padding:5px;
	}
</style>"

# Build the message including the audit details in a table
$HtmlBody = "<body>
<p><font size='2' face='Segoe UI'>
<p><strong>Generated:</strong> $(Get-Date -Format g)</p>  
<h2><u>Please check these guest accounts</u></h2>
<p><b>If guest accounts have been deleted in error, please recover them using the Microsoft Entra admin center (or PowerShell) and assign the accounts a new expiration date.</b></p>
<p>Accounts marked as expired will be removed the next time the job runs. Accounts marked as extended have sign-in activity in the last 30 days.</p><p></p>
<table class='UserTable'>
	<caption><h2><font face='Segoe UI'>Guest Account Expiration Report</h2></font></caption>
	<thead>
	<tr>
	      <th>Id</th>
          <th>Account Name</th>
          <th>Email</th>
          <th>Expiration date</th>
          <th>Action</th>
          <th>LastSignIn</th>
	</tr>
	</thead>
	<tbody>"

$Report = $Report | Sort-Object Action 
ForEach ($A in $Report) {
      $HtmlBody += "<tr><td><font face='Segoe UI'>$($A.Id)</td><td><font face='Segoe UI'>$($A.Name)</td></font><td><font face='Segoe UI'>$($A.Mail)</td></font><td><font face='Segoe UI'>$($A.Expiration)</td><td><font face='Segoe UI'>$($A.Action)</td></font><td><font face='Segoe UI'>$($A.LastSignIn)</td></tr></font>"
    }
$HtmlBody += "</tbody></table><p>" 
$HtmlBody += '</body></html>'

$EmailAddress  = @{address = $ToAddress} 
$EmailRecipient = @{EmailAddress = $EmailAddress}  
    
$HtmlHeaderUser = "<h2>Guest Account Information</h2>"    
$HtmlMsg = "</html>" + $HtmlHead + $htmlbody + "<p>"
# Construct the message body
$MessageBody = @{
    content = "$($HtmlBody)"
    ContentType = 'html'  }

# Create a draft message in the mailbox used to send the message
$NewMessage = New-MgUserMessage -UserId $MsgFrom -Body $MessageBody -ToRecipients $EmailRecipient -Subject $MsgSubject 
# Send the message
Send-MgUserMessage -UserId $MsgFrom -MessageId $NewMessage.Id 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
