# Send-HVEEmailRunbook.PS1
# An example of using an Azure Automation runbook to send email via the Exchange Online High Volume Email
# HVE service. See https://practical365.com/exchange-online-hve/ for a description of how HVE works.

# V1.0 27-Jan-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Send-HVEEmailRunbook.PS1

# Connect to Azure with a managed identity
Connect-AzAccount -Identity

# Fetch the username and password credentials for the HVE account to use from an Azure Key Vault
$UserName = Get-AzKeyVaultSecret -VaultName "Knocksinna12" -Name "UserName" -AsPlainText
$UserPassword = Get-AzKeyVaultSecret -VaultName "Knocksinna12" -name "Password" -AsPlainText

# Create credentials object from the username and password
[securestring]$SecurePassword = ConvertTo-SecureString $UserPassword -AsPlainText -Force
[pscredential]$HVECredentials = New-Object System.Management.Automation.PSCredential ($UserName, $SecurePassword)

# Define the last known date for an update of the Office 365 for IT Pros eBook
[datetime]$LastKnownVersion = "1-Dec-2024"
$DistributionListAddress = "Office365BookFans@office365itpros.com"

# Grab the content of the web page used for the Office 365 for IT Pros eBook
$WebPageURI = 'https://o365itpros.gumroad.com/l/O365IT'
$WebPage = Invoke-WebRequest -Uri $WebPageURI
If (!$WebPage) {
    Write-Output "Failed to retrieve the web page"
    Exit
}

# Extract the 20 characters after the string "latest updates were released on"
$Pattern = "latest updates were released on(.{20})"
If ($WebPage.RawContent -match $pattern) {
    $ExtractedString = $matches[1]
    [datetime]$DateFound = $ExtractedString.SubString(0,$ExtractedString.IndexOf("202")+4).Trim()
    [string]$DateNewVersion = Get-Date $DateFound -format "dd-MMM-yyyy"
    Write-Output ("Last update for Office 365 for IT Pros eBook was released on {0}" -f $DateNewVersion)
    If ($DateFound -gt $LastKnownVersion) {
        Write-Output "A new version of the eBook is available"
    } Else {
        Write-Output "No new version of the eBook is available"
        Exit
    }
} Else {
    Write-Output "Date for latest Office 365 for IT Pros eBook not found in the variable."
    Exit
}

# Build some HTML content
$HTMLContent = ("We detected that a new version of the <b>Office 365 for IT Pros ebook</b> was released on {0}" -f $DateNewVersion)
$HTMLContent = $HTMLContent + "<p> You can download the latest version from your Gumroad account or by using the link in the receipt received by email after you subcribed to the book."
$HTMLContent = $HTMLContent + "<p> More information about the book, including how to download updated files, is available from <a href='https://office365itpros.com/office-365-for-it-pros-faq/'>the FAQ</a></p>"

# Create the HVE parameters
$SendHVEMessageParams = @{}
$SendHVEMessageParams.Add('From', $UserName)
$SendHVEMessageParams.Add('To', $DistributionListAddress)  
$SendHVEMessageParams.Add('Bcc', 'Customer.Services@office365itpros.com')
$SendHVEMessageParams.Add('Subject', "New version of the Office 365 for IT Pros eBook is available")
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
    Write-Host ("Failed to send email to {0} with error {1}" -f $Recipient, $_.Exception.Message)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.