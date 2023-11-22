# Create-DDLBasedOnSMTPAddresses.PS1
# A script to update custom attribute 13 for mailboxes with the primary SMTP address for the mailbox
# and then create a dynamic distribution list based for each domain used in the organization
# https://github.com/12Knocksinna/Office365itpros/blob/master/Create-DDLBasedOnSMTPAddresses.PS1
# V1.0 22-Nov-2023

# Need a connection to Exchange Online
If (Get-ConnectionInformation) {
    Write-Host "Checking the target shared mailboxes and user accounts to process..." 
} Else {
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

[array]$Domains = "office365itpros.com", "office365exchangebook.com"
$DefaultDDLOwner = "Lotte.Vetler@Office365itpros.com"

[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited `
    -Filter {CustomAttribute13 -eq $Null}

If ($Mbx) {
# For each mailbox, find the domain part of the primary SMTP address and store it
    [int]$i = 0
    ForEach ($M in $Mbx) {
        $Domain = $M.PrimarySMTPAddress.Split("@")[1]
        If ($M.CustomAttribute13 -ne $Domain) {
            Write-Host ("Updating mailbox {0} ({1}/{2})" -f $M.DisplayName, $i, $Mbx.count)
            Set-Mailbox -Identity $M.ExternalDirectoryObjectId -CustomAttribute13 $Domain
        }
    }
} Else {
    Write-Host "Can't find any mailboxes - exiting"; break
}

Write-Host "Finished processing mailboxes - now creating dynamic distribution lists"
# Loop through each domain and create a dynamic distribution list (if one doesn't exist already)
ForEach ($D in $Domains) {
    $DDLAlias = ("Domain.{0}" -f $D)
    If ($Null -eq (Get-DynamicDistributionGroup -Identity $DDLAlias -ErrorAction SilentlyContinue)) {
        Write-Host ("Creating dynamic distribution list for the {0} domain..." -f $D) -ForegroundColor Yellow
        $Filter = "CustomAttribute13 -eq '$D' -and RecipientTypeDetails -eq 'UserMailbox'"
        $DDLName = ("{0} users" -f $D)
        $DDLDisplayName = ("People with {0} primary SMTP addresses" -f $D)
        $DDLPrimarySMTPAddress = ("{0}@office365itpros.com" -f $DDLAlias)
        $DDLMailTip = ("Use this distribution list to send to all users with a {0} email address" -f $D)
        # Create the new dynamic distribution list
        $DDL = New-DynamicDistributionGroup -Name $DDLName -DisplayName $DDLDisplayName -Alias $DDLAlias `
            -PrimarySmtpAddress $DDLPrimarySMTPAddress -RecipientFilter $Filter
        If ($DDL) {
            Write-Host ("Dynamic distribution list created for {0}" -f $DDLDisplayName)
            # Make sure that the dynamic distribution list has an owner
            Set-DynamicDistributionGroup -Identity $DDLAlias `
             -ManagedBy $DefaultDDLOwner -MailTip $DDLMailTip
        }
    } Else {
        Write-Host ("A dynamic distribution list already exists for {0} users" -f $D) -ForegroundColor Red
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
