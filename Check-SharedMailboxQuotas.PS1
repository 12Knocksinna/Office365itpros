# Check-SharedMailboxQuotas.PS1
# Example of how to check the current quota usage of shared mailboxes and email
# the results to owners of each mailbox
# https://github.com/12Knocksinna/Office365itpros/blob/master/Check-SharedMailboxQuotas.PS1

Function Add-MessageRecipients {
    # Function to build an addressee list to send email   
     [cmdletbinding()]
         Param(
         [array]$ListOfAddresses )
          ForEach ($SMTPAddress in $ListOfAddresses) {
               @{
                  emailAddress = @{address = $SMTPAddress}
               }    
            }
    } 

# Need to connect to Exchange Online to get mailbox information
If ($Null -eq (Get-ConnectionInformation)) {
    Connect-ExchangeOnline
}
# Connect to the Graph to retrieve information about mailbox licenses and send messages. This 
# uses an Entra ID registered app authenticated using a client thumbprint. The app has consent for the
# necessary User.Read.All and Mail.Send.All permissions.
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3478" # Define your tenant identifier here
$AppId = "d86b1929-b818-411b-834a-206385bf5347" # Define the app identifier 
$CertThumbPrint = "F79286DB88C21491110109A0222348FACF694CBD" # Define the certificate thumbprint
# Get rid of any existing Graph sessions to make sure that we pick up the correct permissions
Disconnect-MgGraph | Out-Null
Connect-MgGraph -NoWelcome -TenantId $TenantId -AppId $AppId -CertificateThumbprint $CertThumbPrint

# Mailbox to receive notifications when a shared mailbox has no registered user with full access
$DefaultNotificationAddress = "Lotte.Vetler@Office365itpros.com"
# Address to send notification messages from
$MailSentFrom = "Azure.Management.Account@office365itpros.com"
# Service plan identifiers to check for licensing
[array]$ServicePlans = "efb87545-963c-4e0d-99df-69c6916d9eb0"

Write-Host "Searching for shared mailboxes..."
[array]$SharedMailboxes = Get-ExoMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -PropertySet Quota -Properties DisplayName
If (!($SharedMailboxes)) {
   Write-Host "Can't find any shared mailboxes to process" 
} Else {
   $QuotaReport = [System.Collections.Generic.List[Object]]::new() 
   ForEach ($Mbx in $SharedMailboxes) {
    $ExoP2 = $False; $NoOwnerFlag = $False
    Write-Host "Checking accounts with full access rights to shared mailbox..."
    [array]$Owners = Get-ExoMailboxPermission -Identity $Mbx.ExternalDirectoryObjectId | Where-Object {$_.User -ne "NT AUTHORITY\SELF"} | `
           Select-Object User, AccessRights | Where-Object {$_.AccessRights -Like "FullAccess"} | Select-Object -ExpandProperty User
    If (!($Owners)) { # No owners found - use default notification address
         $Owners += $DefaultNotificationAddress
         $NoOwnerFlag = $True
         Write-Host ("No owners found for mailbox {0}" -f $Mbx.DisplayName) -ForegroundColor Red
    } Else {
        [array]$ExpandedOwners = $Null
        ForEach ($Owner in $Owners) {
            $OwnerType = (Get-EXORecipient -Identity $Owner).RecipientTypeDetails
            Switch ($OwnerType) {
                "MailUniversalSecurityGroup" {
                   $ExpandedOwners += (Get-DistributionGroupMember -Identity $Owner).PrimarySmtpAddress }
                "UserMailbox" {
                   $ExpandedOwners += $Owner }
            }
        }
        [array]$Owners = $ExpandedOwners | Sort-Object -Unique 
        $OutputOwners = $Owners -join ", "
        Write-Host ("Found owners {0} for mailbox {1}" -f $OutputOwners, $Mbx.DisplayName )
    }
    # Check if the shared mailbox has an Exchange Online Plan 2 license to allow for a larger quota
    [array]$Licenses = Get-MgUserLicenseDetail -UserId $Mbx.ExternalDirectoryObjectId
    If ($Licenses) {
       [array]$ServicePlansMbx = $Licenses.ServicePlans | Where-Object {$_.ProvisioningStatus -eq "Success"} | Select-Object -ExpandProperty ServicePlanId
       ForEach ($SP in $ServicePlans) {
         If ($SP -in $ServicePlansMbx) {
           Write-Host ("Exchange Online Service Plan 2 found for mailbox {0}" -f $Mbx.displayName ) 
           $ExOP2 = $True
         }
       }
     }
     Write-Host "Fetching mailbox usage statistics"
     [array]$MbxStats = Get-ExoMailboxStatistics -Identity $Mbx.ExternalDirectoryObjectId
     [INT64]$QuotaUsed = [convert]::ToInt64(((($MbxStats.TotalItemSize.ToString().split("(")[-1]).split(")")[0]).split(" ")[0]-replace '[,]',''))
     # Byte count for mailbox quota
     [INT64]$MbxQuota = [convert]::ToInt64(((($Mbx.ProhibitSendReceiveQuota.ToString().split("(")[-1]).split(")")[0]).split(" ")[0]-replace '[,]',''))
     $MbxQuotaGB = [math]::Round(($MbxQuota/1GB),2)
     $QuotaPercentUsed = [math]::Round(($QuotaUsed/$MbxQuota)*100,2)
     $QuotaUsedGB = [math]::Round(($QuotaUsed/1GB),2)
     
     $ReportLine = [PSCustomObject]@{
        Mailbox    = $Mbx.Identity
        Name       = $Mbx.DisplayName
        Email      = $Mbx.PrimarySmtpAddress
        Quota      = ("{0} GB" -f $MbxQuotaGB)
        Items      = $MbxStats.ItemCount
        Size       = ("{0} GB" -f $QuotaUsedGB)
        "% Used"   = $QuotaPercentUsed
        Owners     = $Owners
        License    = $EXOP2 
        "No Owner" = $NoOwnerFlag}
     $QuotaReport.Add($ReportLine)    
   }
}

# Code to email the owners of each shared mailbox
[int]$i=0
ForEach ($Item in $QuotaReport) {
    Write-Host ("Sending notification message for mailbox {0}" -f $Item.Name)
    # Get the list of shared mailbox owners
    $EmailRecipients = @( $Item.Owners.Split(",").Trim() )
    # Add them as message recipients
    [array]$MsgToRecipients = Add-MessageRecipients -ListOfAddresses $EmailRecipients
    $MsgSubject = ("Shared Mailbox Quota Information for {0}" -f $Item.Mailbox)
    $HtmlHead = "<h2>Mailbox Quota Information</h2><p>The current quota information for the shared mailbox is as follows.</p>"
    $HtmlBody = $Item | ConvertTo-Html -Fragment 
    $HtmlMsg = "</body></html><p>" + $HtmlHead + $Htmlbody + "<p>"
    # Construct the message body
    $MsgBody = @{
      Content = "$($HtmlMsg)"
      ContentType = 'html'  
     }
    
    $Message =  @{subject           = $MsgSubject}
    $Message += @{toRecipients      = $MsgToRecipients}  
    $Message += @{body              = $MsgBody}	
    $Params   = @{'message'         = $Message}
    $Params  += @{'saveToSentItems' = $True}
    $Params  += @{'isDeliveryReceiptRequested' = $True}
    
    # And send the message using the parameters that we've filled in
    Send-MgUserMail -UserId $MailSentFrom -BodyParameter $Params
    $i++
}

Write-Host ("Processing complete and quota usage notifications sent for {0} mailboxes" -f $i)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
