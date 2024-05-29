# Report-TeamsChatInfo.PS1
# An example to show how to report the Teams chat threads for the user signed into an
# interactive session with the Microsoft Graph PowerShell SDK

# Connect to the Graph with the necessary permissions
Connect-MgGraph -Scopes Chat.ReadWrite, Chat.ManageDeletion.All -NoWelcome
 # Find the object identifier for the signed-in user
$Account = (Get-MgContext).Account
$UserId = (Get-MgUser -UserId $Account).Id

# Get chats for the user
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
$Report = [System.Collections.Generic.List[Object]]::new()

# Extract information for each chat thread and write it out to the report file
Write-Host "Analyzing chat threads..."
ForEach($Chat in $Chats) {
    # Get members of the chat thread
    [array]$Members = Get-MgUserChatMember -UserId $UserId -ChatId $Chat.Id

    # Extract the member display names and remove any blanks (accounts that no longer exist)
    [array]$MemberNames = $Members.displayName | Sort-Object -Unique
    $NumberOfParticipants = $MemberNames.Count
    If ($MemberNames.Count -gt 0) {
        $MemberNamesOutput = $MemberNames -Join ", "
    }

    # Get tenant display name
    $TenantDisplayName = "Unknown"
    $TenantDisplayName = $TenantsHash[$Chat.TenantId]

    $ReportLine = [PSCustomObject][Ordered]@{
      ChatType             = $Chat.ChatType
      Topic                = $Chat.Topic
      Created              = $Chat.CreatedDateTime
      LastUpdated          = $Chat.LastUpdatedDateTime
      NumberOfParticipants = $NumberOfParticipants
      Participants         = $MemberNamesOutput
      TenantId             = $Chat.TenantId
      Tenant               = $TenantDisplayName
      Id                   = $Chat.Id
      WebUrl               = $Chat.WebUrl }     
    $Report.Add($ReportLine) 
} # End ForEach chats

# And here's what we found sorted by the last updated date
$Report | Select-Object ChatType, Topic, Created, LastUpdated, NumberOfParticipants, Participants, Tenant | Sort-Object LastUpdated -Descending | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
