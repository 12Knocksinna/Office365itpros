# Analyze-TeamsExternalChats.PS1
# Analyze User chats to find chats with people outside the tenant
# Because the chat API uses delegated access in interactive mode, this code can only process
# chats for other users when run as an app or in an Azure Automation runbook

Connect-MgGraph -NoWelcome -Scopes User.Read.All, Chat.Read.All, Mail.Send
# Fetch organization information
$HomeTenant = Get-MgOrganization
$HomeTenantId = $HomeTenant.Id
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
    Write-Output ("Processing chats for user {0}..." -f $User.DisplayName)
    # Find the onOnOne chats for this user
    [array]$Chats = Get-MgBetaUserChat -Userid $User.Id -All -Filter "ChatType eq 'oneOnOne'"
    If ($Chats) {
        Write-Output ("Found {0} one-on-one chats for {1}" -f $Chats.count, $User.DisplayName)
        ForEach ($Chat in $Chats) {
            # Find participants
            [array]$Participants = Get-MgChatMember -ChatId $Chat.Id
            # Remove the user we're processing
            [array]$OtherParticipant = $Participants | `
               Where-Object {$_.additionalProperties.email -ne $User.userPrincipalName}
            If ($OtherParticipant.displayName) { 
               $TenantName = $Null 
               # See if we can get a domain name for the other participant 
               Try {
                    $EmailDomain = $OtherParticipant.additionalProperties.email.split('@')[1]
               }   Catch {
                    $EmailDomain = $Null
               }
               # Now check if the owning domain for the chat is the home or the external tenant
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
                # If the domain of the other participant doesn't belong to this tenant, record it
                If ($Domain -notin $Domains) {
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
                }
            }
        }
    }
}

Write-Output ("{0} chats found with external people in the {1} tenant" -f $ChatReport.count, $HomeTenant.displayName)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.