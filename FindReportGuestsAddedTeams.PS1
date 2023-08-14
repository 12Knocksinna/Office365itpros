# FindReportGuestsAddedTeams.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindReportGuestsAddedTeams.PS1
# Find and report guest users added to Teams in the last week
# Some tenants like to keep a close eye on the guest user accounts that are added to Teams and want to report those accounts. This script
# looks for audit records noting the guest additions over the last 7 days and extracts details of guest accounts if the accounts are created
# in the same period (older accounts are ignored because they are likely approved). For any account found, we send a polite email to the admin
# to tell them that they need to validate that the guest is OK.

# Updated 14-Aug-2023 for the Microsoft Graph PowerShell SDK 2.x

$Modules = Get-Module | Select-Object -ExpandProperty Name
# Connect to Exchange Online if not already connected
If ("ExchangeOnlineManagement" -notin $Modules) {
    Connect-ExchangeOnline 
}
# And connect to the Graph
If ('Microsoft.Graph.Authentication' -notin $Modules) {
  Connect-MgGraph -Scopes Directory.Read.All
}

$MsgFrom = "o365itprosrenewals@office365itpros.com" # Set your own address for email notifications here.
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 # Set TLS 1.2 for SMTP
[array]$Records = Search-UnifiedAuditLog -StartDate ((Get-Date).AddDays(-7)) -EndDate ((Get-Date).AddDays(1)) -ResultSize 5000 -Operations MemberAdded -RecordType MicrosoftTeams

If ($Records) {
   $Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report
   Write-Host "Processing" $Records.Count "audit records for addition of users to Microsoft Teams"
   ForEach ($Rec in $Records) {
     $AuditData = Convertfrom-Json $Rec.AuditData # Get payload
     ForEach ($M in $AuditData.Members) { # Examine users added to see if any are guests
      If (($M -Like "*#EXT#@*") -and ($AuditData.CommunicationType -eq "Team")) { # We have a guest user who's been added to a team rather than a group chat
         $GuestUser = Get-MgUser -UserId $M.UPN -Property Id, DisplayName, Mail, CreatedDateTime
         $AccountAge = ($GuestUser.CreatedDateTime | New-TimeSpan).Days
         If ($AccountAge -le 7) { # Guest created within last 7 days so write out details
            $ReportLine = [PSCustomObject]@{ 
               Guest            = $GuestUser.Mail   
               Name             = $GuestUser.DisplayName
               Created          = $CreationDate 
               AgeInDays        = $AccountAge
               DateAddedTeams   = Get-Date($AuditData.CreationTime) -format g
               TeamName         = $AuditData.TeamName
               AddedBy          = $AuditData.UserId
               GruupId          = $AuditData.AADGroupId} 
            $Report.Add($ReportLine) 
         } # End if (AccountAge)   
     } # End if (Guest user check)
   } # End Foreach (Members)
 } # End ForEach (Records)
} #End if (Records)

If ($Report) { # Some records have been created, so let's report them.
  Clear-Host
  $SmtpServer = "smtp.office365.com" ; $SmtpPort = '587'
  $ProgressDelta = 100/($Report.Count); $PercentComplete = 0; $UserNumber = 0
  # Define some variables for the message starting with HTML header with styles
  $htmlhead="<html>
     <style>
      BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    </style>"
  #Header for the message
  $HtmlBody = "<body>
     <h1>New Guest User Account Creation</h1>
     <p><strong>Date:</strong> $(Get-Date -Format g)</p>  
     <h2><u>New Update Available - Please Download Updated files for your eBook/u></h2>"
  # For each guest user found, create and send an email
  # First Remove any records added by 'Microsoft Teams Sync'
  $Report = $Report | Where-Object { $_.AddedBy -ne "Microsoft Teams Sync" }
  ForEach ($R in $Report) {
   $UserNumber++ 
   $UserStatus = $R.Name + "(" + $R.Guest + ") ["+ $UserNumber +"/" + $Users.Count + "]"
   Write-Progress -Activity "Processing user" -Status $UserStatus -PercentComplete $PercentComplete
   $PercentComplete += $ProgressDelta

   $BodyText = "<p><b><u>Details</u></b></p><p>Guest email address: $($R.Guest)</p><p>Guest name: $($R.Name)</p><p>Date added: $($R.DateAddedTeams)</p><p>Team added to: $($R.TeamName)</p><p>Recently you added a new guest user to our tenant by including them as a member in the $($R.TeamName) team. Please confirm that it is business-critical to allow this person access to tenant resources. If you do not confirm within the next week, we will automatically remove this guest account.</i></p><p>Best Regards</p><p><b>Your Friendly Admin</b></p>"
   $htmlHeaderUser = "<h2>A new guest user has been created in our tenant</h2>"; $htmlbody = $htmlheaderUser + $BodyText + "<p>"
   $HtmlMsg = "</body></html>" + $HtmlHead + $HtmlBody
  # Construct the message parameters and send it off...
    $MsgParam = @{
      To = $R.AddedBy
       From = $MsgFrom
       Subject = "New Guest User Added"
       Body = $HtmlMsg
       SmtpServer = $SmtpServer
       Port = $SmtpPort
       Credential = $O365Cred }
     Send-MailMessage @msgParam -UseSSL -BodyAsHTML 
  } # End ForEach
} # End if

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
