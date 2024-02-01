# Analyze-TeamsExternalChats.PS1
# Analyze User chats to find chats with people outside the tenant
# Because the chat API uses delegated access in interactive mode, this code can only process
# chats for other users when run as an app or in an Azure Automation runbook
# V1.0 1 Nov 2023
# https://github.com/12Knocksinna/Office365itpros/blob/master/Analyze-TeamsExternalChats.PS1

# Permissions used:
# CrossTenantInformation.ReadBasic.All: Read tenant id and return tenant name
# User.Read.All: Read user information
# Chat.Read.All: Read user chats
# Directory.Read.All: Read tenant information from directory
# Mail.Send: Send email with results

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

# Start of processing
# Define the home tenant where the job will run
$HomeTenantId = 'a662313f-14fc-43a2-9a7a-d2e27f4f3478'
$CSVFile = "c:\temp\TeamsExternalAccessScan.csv"

# Define the settings used for certificate-based authentication for the registered app you want to use
# Thumbprint for the X.509 certificate uploaded to the registered app
$CertificateThumbPrint = "F79286DB88C21491110109A0222348FACF694CBD"
# Application (client) identifier for the registered app. Get this value from the app properties
$AppId = "8f005189-8c58-4fb5-a226-8851e13490cb"

# Connect to the Graph and Teams endpoints using certificate-based authentication
# Certificate-based authentication depends on the tenant id, app id, and certificate thumbprint. All must be
# defined and accurate as otherwise authentication will fail
Write-Host "Connecting to the Graph SDK endpoint..."
Connect-MgGraph -NoWelcome -ClientId $AppId -CertificateThumbprint $CertificateThumbPrint -TenantId $HomeTenantId
Write-Host "Now connecting to the Microsoft Teams endpoint..."
$Status = Connect-MicrosoftTeams -TenantId $HomeTenantId -ApplicationId $AppId -Certificate $CertificateThumbPrint

If (!($Status.Account -ne $AppId)) {
    Write-Host "Whoops... can't connect!"; break
} Else {
    Write-Host "All connected. Now setting things up..."
}
# Fetch organization information
$HomeTenant = Get-MgOrganization

# Find domains in the tenant so that we can discard chats with local people
[array]$Domains = Get-MgDomain | Select-Object -ExpandProperty Id
# Create hash table to store information about external tenants
$ExternalTenants = @{}
# Find user accounts to check
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
    -ConsistencyLevel eventual -CountVariable Records -All | Sort-Object displayName
Write-Output ("Found {0} user accounts to process..." -f $users.Count)
$ChatReport = [System.Collections.Generic.List[Object]]::new()

ForEach ($User in $Users) {
    Write-Output ("Looking for chats for user {0}..." -f $User.DisplayName)
    # Find the OneOnOne chats for this user
    [array]$Chats = Get-MgBetaUserChat -Userid $User.Id -All -Filter "ChatType eq 'oneOnOne'"
    If ($Chats) {
        Write-Output ("Found {0} one-on-one chats for {1}" -f $Chats.count, $User.DisplayName)
        ForEach ($Chat in $Chats) {
            # Find participants
            Try {
                [array]$Participants = Get-MgChatMember -ChatId $Chat.Id
                # Remove the user we're processing
                [array]$OtherParticipant = $Participants | `
                   Where-Object {$_.additionalProperties.email -ne $User.userPrincipalName}
            } Catch {
                Write-Output ("Can't fetch participant information for chat {0}" -f $Chat.Id)
            }
            If ($OtherParticipant.displayName) { 
               $TenantName = $Null 
               # See if we can get a domain name for the other participant 
               Try {
                    $EmailDomain = $OtherParticipant.additionalProperties.email.split('@')[1]
               }   Catch {
                    $EmailDomain = $Null
               }
            # Only process a chat if the other participant comes from a domain outside the home tenant
            If ($EmailDomain -notin $Domains) {
               # Check if the owning domain for the chat is the home or the external tenant
                $ExternalTenantId = $OtherParticipant.additionalProperties.tenantId.ToString()
                If ($ExternalTenantId -ne $HomeTenantId) {
                    Try {
                       $TenantName = $ExternalTenants[$ExternalTenantId] }
                    Catch {
                       Write-Output ("Resolving tenant id {0}" -f $ExternalTenantId) 
                    }  
                    $ChatCount = "N/A"
                    If (!($TenantName)) {
                        # Couldn't find the tenant, so resolve its identifier and record its details
                        $Uri = ("https://graph.microsoft.com/beta/tenantRelationships/findTenantInformationByTenantId(tenantId='{0}')" -f `
                        $ExternalTenantId)
                        $ExternalTenantData = Invoke-MgGraphRequest -Uri $Uri -Method Get
                        $TenantName = $ExternalTenantData.displayName
                        $ExternalTenants.Add($ExternalTenantId, $TenantName)
                    }
                }
                If ($ExternalTenantId -eq $HomeTenantId) {
                    $TenantName = $HomeTenant.DisplayName
                    [array]$ChatMessages = Get-MgChatMessage -ChatId $Chat.id -All
                    $ChatCount = $ChatMessages.Count
                }
                $Reportline = [PsCustomObject]@{
                        Chat                 = $Chat.Id 
                        User                 = $User.displayName
                        UPN                  = $User.userPrincipalName
                        Participant          = $OtherParticipant.displayName
                        'Participant Email'  = $OtherParticipant.additionalProperties.email
                        Domain               = $EmailDomain
                        'Owning Tenant Name' = $TenantName
                        Role                 = $OtherParticipant.Roles
                        ChatCreated          = $Chat.CreatedDateTime
                        LastUpdated          = $Chat.LastUpdatedDateTime
                        'Chat Messages'      = $ChatCount
                }
                $ChatReport.Add($ReportLine)
                # Slight pause before continuing... wouldn't want to be throttled
                Start-Sleep -Milliseconds 500 
                }
            }
        }
    }
}

Write-Output ""
Write-Output ("{0} chats found with external people in the {1} tenant" -f $ChatReport.count, $HomeTenant.displayName)
Write-Output ""

Write-Output "Checking the Teams External Access configuration for the tenant..."
# Find if we need to update the external access configuration
[array]$ExternalDomains = $ChatReport.Domain | Sort-Object -Unique
$DomainConfiguration  = Get-CsTenantFederationConfiguration  | Select-Object -ExpandProperty AllowedDomains
[array]$ConfiguredDomains = $DomainConfiguration.AllowedDomain.Domain
If (!($ConfiguredDomains)) {
    Write-Output ("The {0} tenant doesn't have any configured domains for external access" -f $HomeTenant.displayName)
} Else {  
    [array]$MissingDomains = $ExternalDomains | Where-Object {$_ -notin $ConfiguredDomains}
    If ($MissingDomains) {
        $MissingDomainsOutput = $MissingDomains -Join ", "
        Write-Output ("The following external domains are not configured for external access: {0}" -f $MissingDomainsOutput) 
    } Else {
        Write-Output ("External Access is configured for all external domains found in the scan")
    }
}

# Email the results
$EmailRecipient = "Tony.Redmond@office365itpros.com"
# Send a message from the shared mailbox
$MsgFrom = "Azure.Management.Account@office365itpros.com"
# Add your recipient address here
$ToRecipientList   = @( $EmailRecipient )
[array]$MsgToRecipients = Add-MessageRecipients -ListOfAddresses $ToRecipientList
$MsgSubject = "Teams external access review"
$HtmlHead = "<h2>Teams external chat participant report</h2><p>The following chats have been found to involve external participants.</p>"
$HtmlBody = $ChatReport | Select-Object User, UPN, 'Participant Email', Domain, ChatCreated, 'Chat Messages' | `
    ConvertTo-Html -Fragment 
# Add a line about missing domains if there are any in the external access configuration, else just send the
# data about external chats
If ($MissingDomains) {
    $HtmlBody2 = "<h2>The following domains are not in the tenant external access configuration: </h2>" + `
       "<p>" + $MissingDomainsOutput + "</p>" 
    $HtmlMsg = "</body></html><p>" + $HtmlHead + $Htmlbody + "<p>" + $HtmlBody2 + "<p>"
} Else {
    $HtmlMsg = "</body></html><p>" + $HtmlHead + $Htmlbody + "<p>"
}
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
Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params
Write-Output ("Message containing external chat information sent to {0}!" -f $EmailRecipient)

$ChatReport | Export-Csv -NoTypeInformation $CSVFile
Write-Output ("The complete data file is available in {0}" -f $CSVFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
