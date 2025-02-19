# Report-MailboxesEnabledForEWS.PS1
# Report the set of mailboxes found to be still enabled for Exchange Web Services (EWS) and send email to 
# a nominated addressee to ask them to review the mailboxes and disable EWS for whatever mailboxes no longer need the 
# facility.

# Requires an Azure automation account and a Exchange Online HVE account. 

# V1.0 19-Feb-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-MailboxesEnabledForEWS.PS1

# Get the credentials for the HVE account
Connect-AzAccount -Identity
# Fetch the username and password credentials for the HVE account to use from an Azure Key Vault
# Change the vault name to the name used in the Azure account
$UserName = Get-AzKeyVaultSecret -VaultName "xxxx" -Name "UserName" -AsPlainText
$UserPassword = Get-AzKeyVaultSecret -VaultName "xxxx" -name "Password" -AsPlainText
# Create credentials object from the username and password
[securestring]$SecurePassword = ConvertTo-SecureString $UserPassword -AsPlainText -Force
[pscredential]$HVECredentials = New-Object System.Management.Automation.PSCredential ($UserName, $SecurePassword)

# Connect to your tenant with Exchange Online
Connect-ExchangeOnline -ManagedIdentity -Organization xxxxx

$EWSEnabled = (Get-OrganizationConfig).EWSEnabled
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
[array]$EWSEnabledMbx = Get-CasMailbox -filter "EWSEnabled -eq 1" -ResultSize Unlimited | Sort-Object DisplayName

# Process each mailbox and report what we find
ForEach ($Mbx in $EWSEnabledMbx) {

    $UserInfo = Get-User -Identity $Mbx.ExternalDirectoryObjectId

    $MbxReportLine = [PSCustomObject][Ordered]@{ 
        DisplayName         = $Mbx.DisplayName
        PrimarySMTPAddress  = $Mbx.PrimarySMTPAddress
        UPN                 = $UserInfo.UserPrincipalName
        Department          = $UserInfo.Department
        City                = $UserInfo.City
        Country             = $UserInfo.CountryOrRegion
        EntraIDObjectId     = $Mbx.ExternalDirectoryObjectId
        MailboxType         = $UserInfo.RecipientTypeDetails
    }
    $Report.Add($MbxReportLine)
}

# Create a HTML form of the report
$HtmlReport = $Report | ConvertTo-Html -Fragment

# Change this address to match your tenant
$DistributionListAddress = "Office365BookFans@office365itpros.com"
# Build some HTML content
$HTMLContent = ("<p>We found <b>{0}</b> mailboxes still enabled for Exchange Web Services.</p>" -f $Report.Count)
$HTMLContent = $HTMLContent + ("<p>The EWSEnabled setting in the organization confiuration is <b>{0}</b>.</p>" -f $EWSEnabled)
$HTMLContent = $HTMLContent + $HtmlReport
$HTMLContent = $HTMLContent + "<p>Please review these mailboxes and disable EWS for any mailbox that no longer needs this facility.</p>"
# Create the HVE parameters
$SendHVEMessageParams = @{}
$SendHVEMessageParams.Add('From', $UserName)
$SendHVEMessageParams.Add('To', $DistributionListAddress)  
# Change this address for your tenant
$SendHVEMessageParams.Add('Bcc', 'Customer.Services@office365itpros.com')
$SendHVEMessageParams.Add('Subject', "Mailboxes Enabled for Exchange Web Services")
$SendHVEMessageParams.Add('Body', $HTMLContent)
$SendHVEMessageParams.Add('UseSsl', $true)
$SendHVEMessageParams.Add('Credential', $HVECredentials)
$SendHVEMessageParams.Add('SmtpServer', 'smtp-hve.office365.com')
$SendHVEMessageParams.Add('Port', 587)
$SendHVEMessageParams.Add('BodyAsHtml', $True)

# And send the message
Try {
    Send-MailMessage @SendHVEMessageParams -ErrorAction Stop 
} Catch {
    Write-Output ("Failed to send email to {0} with error {1}" -f $Recipient, $_.Exception.Message)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.