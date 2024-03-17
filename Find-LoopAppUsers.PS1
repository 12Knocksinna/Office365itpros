# Find-LoopAppUsers.PS1
# Example script to show how to find which licensed users are using the Loop App
# V1.0 17-Mar-2024
# https://github.com/12Knocksinna/Office365itpros/blob/master/Find-LoopAppUsers.PS1

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Connecting to Exchange Online" -ForegroundColor Red
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}
# To read user information from Entra ID
Connect-MgGraph -NoWelcome -Scopes User.Read.All

# This command finds user accounts with a license containing the Loop app. It specifies the SKU identifers for
# Microsoft 365 Business Standard, Microsoft 365 Business Premium, Microsoft 365 E3, amd Microsoft 365 E5. 
# There are other variants of these SKUs # for government and academic use, so it's important to pass the SKU 
# identifiers in use within your tenant. The service plan for Loop is MICROSOFT_LOOP (c4b8c31a-fb44-4c65-9837-a21f55fcabda)
Write-Host "Finding user accounts to check..."
[array]$Users = Get-MgUser -Filter "assignedLicenses/any(s:s/skuId eq cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46) `
    or assignedLicenses/any(s:s/skuid eq f245ecc8-75af-4f8e-b61f-27d8114de5f3) `
    or assignedLicenses/any(s:s/skuid eq 05e9a617-0261-4cee-bb44-138d3ef5d965) `
    or assignedLicenses/any(s:s/skuid eq 06ebc4ee-1bb5-47dd-8120-11324bc54e06)" `
    -ConsistencyLevel Eventual -CountVariable Licenses -All -Sort 'displayName' `
    -Property Id, displayName, signInActivity, userPrincipalName

If ($Users.count -eq 0) {
    Write-Host "No users can be found eligble to use the Loop app - exiting" -ForegroundColor Red
    break
}

Write-Host ("{0} user accounts are eligible to use the Loop app" -f $Users.count)

$StartDate = (Get-Date).AddDays(-30)
$EndDate = (Get-Date).AddDays(1)
# Define the set of SharePoint file audit records we are interested in
[array]$Operations = "FileModified", "FileModifiedExtended", "FileUploaded", "FileAccessed"
Write-Host "Search the unified audit log..."
[array]$Records = Search-UnifiedAuditLog -Operations $Operations -StartDate $StartDate -EndDate $EndDate -Formatted `
 -SessionCommand ReturnLargeSet -ResultSize 5000

If (!($Records)) {
    Write-Host "No audit records can be found - exiting"
    Break
}

Write-Host ("Analyzing {0} audit records..." -f $Records.count)
$LoopRecords = [System.Collections.Generic.List[Object]]::new()
$Records = $Records | Sort-Object {$_.CreationDate -as [datetime]}
ForEach ($Rec in $Records) {
   $AuditData = $Rec.AuditData | ConvertFrom-JSON
   If ($Auditdata.Applicationid -eq 'a187e399-0c36-4b98-8f04-1edc167a0996' -and $AuditData.Operation -ne 'UserLoggedIn') {
    $ReportLine = [PSCustomObject][Ordered]@{
        User             = $Rec.UserIds
        Timestamp        = $Rec.CreationDate
        fileName         = $AuditData.SourceFileName
        Operation        = $Rec.operations
        ObjectId         = $AuditData.ObjectId
    } 
    $LoopRecords.Add($ReportLine)      
   }
}

$LoopRecords = $LoopRecords | Sort-Object {$_.Timestamp -as [datetime]} -Descending
$LoopUsers = $LoopRecords | Sort-Object User -Unique

ForEach ($User in $Users) {
   If ($User.userPrincipalName -notin $LoopUsers.User) {
        Write-Host ("Can't find usage of the Loop app by {0}" -f $User.DisplayName)
   } Else {
        Write-Host ("User {0} last used the Loop app on {1}" -f $User.displayName, $LoopUsers.TimeStamp)
   }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.