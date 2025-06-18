# Report-TeamsChatInfo.PS1
# An example to show how to report the Teams chat threads for the user signed into an
# interactive session with the Microsoft Graph PowerShell SDK

# V1.01 18-June-2025
# Github link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-TeamsChatInfo.PS1
# Referenced in article: https://practical365.com/delete-teams-chat-threads/

# Connect to the Graph with the necessary permissions
Connect-MgGraph -Scopes Chat.ReadWrite, Chat.ManageDeletion.All -NoWelcome
 # Find the object identifier for the signed-in user
$Account = (Get-MgContext).Account
$UserId = (Get-MgUser -UserId $Account).Id

# Get chats for the user
Write-Host ("Fetching chat threads for user {0}" -f $Account)
[array]$Chats = Get-MgUserChat -Userid $UserId -All -Filter "chattype eq 'group' or chattype eq 'oneonone' or chattype eq 'Meeting'" | Sort-Object LastUpdatedDateTime -Descending

# Extract the identifiers for the tenants that chat threads belong to so that we can resolve them to a tenant name
Write-Host "Finding tenant information for chats..."
$TenantList = [System.Collections.Generic.List[Object]]::new()
ForEach ($Chat in $Chats) {
   $TenantList.Add($Chat.TenantId)
}

$TenantsCount = @{}
$TenantList = $TenantList | Sort-Object
$TenantList | ForEach-Object {$TenantsCount[$_]++}
$TenantsCount = $TenantsCount.GetEnumerator() | Sort-Object -Property Value -Descending
$TenantsCount

# Now build a hash table of tenant identifiers and names that we can use to look up tenants
Write-Host "Resolving tenant identifiers..."
$TenantsHash = @{}
ForEach ($Tenant in $TenantsCount) {
    $Uri = ("https://graph.microsoft.com/V1.0/tenantRelationships/findTenantInformationByTenantId(tenantId='{0}')" -f $Tenant.Name.ToString())
    $ExternalTenantData = Invoke-MgGraphRequest -Uri $Uri -Method Get  
    $TenantsHash.Add([string]$Tenant.Name.toString(), [string]$ExternalTenantData.displayName)
}

# Create output report
$ChatReport = [System.Collections.Generic.List[Object]]::new()

# Extract information for each chat thread and write it out to the report file
Write-Host "Analyzing chat threads..."
ForEach($Chat in $Chats) {
    # Get members of the chat thread
    Try {
        [array]$Members = Get-MgUserChatMember -UserId $UserId -ChatId $Chat.Id -ErrorAction Stop
        # Extract the member display names and remove any blanks (accounts that no longer exist)
        [array]$MemberNames = $Members.displayName | Sort-Object -Unique
        $NumberOfParticipants = $MemberNames.Count
        If ($MemberNames.Count -gt 0) {
            $MemberNamesOutput = $MemberNames -Join ", "
        }
    } Catch {
        Write-Host ("Error getting members for chat {0}: {1}" -f $Chat.Id, $_.Exception.Message) -ForegroundColor Red
        $MemberNamesOutput = "Error retrieving members"
        $NumberOfParticipants = 0
        Continue
    }

    # Get tenant display name
    $TenantDisplayName = "Unknown"
    $TenantDisplayName = $TenantsHash[$Chat.TenantId]

    $ChatReportLine = [PSCustomObject][Ordered]@{
        ChatType             = $Chat.ChatType
        Topic                = $Chat.Topic
        Created              = Get-Date $Chat.CreatedDateTime -format 'dd-MMM-yyyy HH:mm:ss'
        LastUpdated          = Get-Date $Chat.LastUpdatedDateTime -format 'dd-MMM-yyyy HH:mm:ss'
        NumberOfParticipants = $NumberOfParticipants
        Participants         = $MemberNamesOutput
        TenantId             = $Chat.TenantId
        Tenant               = $TenantDisplayName
        Id                   = $Chat.Id
        WebUrl               = $Chat.WebUrl 
    }     
    $ChatReport.Add($ChatReportLine) 
} # End ForEach chats

# Tenant analysis
Write-Host ""
Write-Host "Analysis of chat threads by tenant"
Write-Host "----------------------------------"
$ChatReport | Group-Object Tenant -NoElement | Sort-Object Count -Descending | Format-Table @{expression={$_.Name}; label = "Host Tenant"}, Count
Write-Host ""
Write-Host "Analysis of chat threads by type"
Write-Host "--------------------------------"
$ChatReport | Group-Object ChatType -NoElement | Sort-Object Count -Descending | Format-Table @{expression={$_.Name}; label = "Host Tenant"}, Count
Write-Host ""

# And here's what we found sorted by the last updated date
$ChatReport | Select-Object ChatType, Tenant, Topic, Created, LastUpdated, NumberOfParticipants, Participants | Sort-Object LastUpdated -Descending | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 