# https://github.com/12Knocksinna/Office365itpros/blob/master/AzureAutomationAddMembersTeamChannels.PS1
# Example of using a managed identity with Azure Automation to do some processing with the
# Microsoft Graph PowerShell SDK and Microsoft Teams modules
# 
#Obtain AccessToken for Microsoft Graph via the managed identity
$ResourceURL = "https://graph.microsoft.com/" 
$Response = [System.Text.Encoding]::Default.GetString((Invoke-WebRequest -UseBasicParsing -Uri "$($env:IDENTITY_ENDPOINT)?resource=$resourceURL" -Method 'GET' -Headers @{'X-IDENTITY-HEADER' = "$env:IDENTITY_HEADER"; 'Metadata' = 'True'}).RawContentStream.ToArray()) | ConvertFrom-Json 
$AccessToken = $response.access_token 

#Connect to the Microsoft Graph using the aquired AccessToken
Connect-Graph -AccessToken $AccessToken
#Define the desired graph endpoint
Select-MgProfile Beta

Connect-MicrosoftTeams -Identity

[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All
If (!($Users)) { Write-Output "No user accounts found - exiting" ; break }
# Filter out any accounts marked that shouldn't be added to team membership
$FilteredUsers = $Users | ? {$_.OfficeLocation -ne "XXX"}

# Now check that each user actually has a Teams service plan
$UsersWithTeams =  [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $FilteredUsers) {
  $TeamsLicense =  Get-MgUserLicenseDetail -UserId $User.Id | Select-Object -ExpandProperty ServicePlans | ? {$_.ServicePlanId -eq "57ff2da0-773e-42df-b2af-ffb7a2317929"} | Select-Object -ExpandProperty ProvisioningStatus
 If ($TeamsLicense -eq "Success") {
     $UserData = [PSCustomObject][Ordered]@{  # Write out details of the user
       Id          = $User.Id
       DisplayName = $User.DisplayName }
     $UsersWithTeams.Add($UserData)
 } #End if
} #End ForEach
# These variables will differ depending on the host team and channel name you decide to use
$GroupId = (Get-Team -DisplayName "HR Questions and Answers").GroupId
$ChannelName = "Questions and Answers"
$ChannelId = (Get-TeamChannel -GroupId $GroupId | Where-Object {$_.DisplayName -eq $ChannelName} | Select-Object -ExpandProperty Id)

Write-Output ("Finding the membership of of the {0} channel" -f $ChannelName)
# Find current members and owners and add them to a hash table that we can lookup
$ChannelMembers = Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName -Role Member
$ChannelOwners = Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName -Role Owner

$CurrentMembers = @{}
ForEach ($Member in $ChannelMembers) {
   $CurrentMembers.Add($Member.UserId,$Member.User) }
ForEach ($Member in $ChannelOwners) {
   $CurrentMembers.Add($Member.UserId,$Member.User) }

$i = 0
# Check each user and add them if they are not found
$UsersAdded =  [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $UsersWithTeams) {
   If (!($CurrentMembers[$User.Id])) {
      Write-Output ("Adding {0} to the {1} channel" -f $User.DisplayName, $ChannelName)
      Add-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName -User $User.Id; $i++ 
      $UserData = [PSCustomObject][Ordered]@{  # Write out details of the user
         Id          = $User.Id
         DisplayName = $User.DisplayName }
     $UsersAdded.Add($UserData)}
}

# If we have added any users to the channel, generate a message that we can post to the channel
If ($UsersAdded) { # Generate a report and post it to Teams
  $Today = Get-Date -format dd-MMM-yyyy
$Body = '
<style>
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
</style>
<p><font size="2" face="Segoe UI">
<h3>Generated: ' + $Today + '</h3></font></p>
<table class="UserTable">
	<caption><h2><font face="Segoe UI">User Accounts Added to Shared Channel</h2></font></caption>
	<thead>
	<tr>
	    <th>Display Name</th>
		<th>Azure AD Account Identifier</th>
	</tr>
	</thead>
	<tbody>'

ForEach ($U in $UsersAdded) {
      $Body += "<tr><td><font face='Segoe UI'>$($U.DisplayName)</font></td><td><font face='Segoe UI'>$($U.Id)</td></tr></font>"
    }
$Body += "</tbody></table><p>" 
$Body += '</body></html>'
   
   Write-Output "Posting to Channel"
  # Get the user password, target site URL, and user name from Azure Key Vault
   Connect-AzAccount -Identity | Out-Null
   # Something like Joe.Doe@office365itpros.com - they must be a member of the target channel
   $UserName = Get-AzKeyVaultSecret -VaultName "MIKeyVault" -name "CredentialUser" -AsPlainText
   # Something like https://office365itpros.sharepoint.com/
   $SiteURL = Get-AzKeyVaultSecret -VaultName "MIKeyVault" -name "SPOSiteURL" -AsPlainText
   $UserPassword = Get-AzKeyVaultSecret -VaultName "MIKeyVault" -name "CredentialSecret" -AsPlainText
   [securestring]$SecurePassword = ConvertTo-SecureString $UserPassword -AsPlainText -Force
   [pscredential]$ChannelMemberCredentials = New-Object System.Management.Automation.PSCredential ($UserName, $SecurePassword)
   # Connect to PnP using the account credentials we just retrieved
   $PnpConnection = Connect-PnPOnline $SiteURL -Credentials $ChannelMemberCredentials -ReturnConnection
   Submit-PnPTeamsChannelMessage -Team $GroupId -Channel $ChannelId -Message $Body -ContentType Html -Important -Connection $PnpConnection
}

Write-Output ("{0} new members added to channel" -f $i)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
