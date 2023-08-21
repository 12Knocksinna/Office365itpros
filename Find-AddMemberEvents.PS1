# A script to show how to detect users with a specifc type of displayname can be detected if someone adds them to 
# a team


# Check if we can run an Exchange Online cmdlet. If we can, go on, else connect to Exchange Online
If ($Null -eq (Get-ConnectionInformation)) {
    Connect-ExchangeOnline
}
# Connect to the Graph
Connect-MgGraph
# Find users who have Project in their display name
Write-Host "Finding user details to check"
[array]$Users = Get-MgUser -Search "displayName:Project" -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel Eventual
If (!($Users)) { 
    Throw "No users found"
}

Write-Host ("Checking audit records for {0} users" -f $Users.count)
# Build hash table of users that we want to check
$UserLookup = @{}
ForEach ($User in $Users) {
   $UserLookup.Add($User.UserPrincipalName, $User.DisplayName)
}
$StartDate = (Get-Date).AddDays(-7)
$EndDate = (Get-Date).AddDays(1)
[array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Formatted -ResultSize 5000 `
   -RecordType MicrosoftTeams -Operations MemberAdded
If (!($Records)) {
    Throw "No records found"
}
$Report = [System.Collections.Generic.List[Object]]::new()  

ForEach ($Rec in $Records) {
    $AuditData = $Rec.AuditData | ConvertFrom-Json
    # Check the members noted as added to a group
    ForEach ($Member in $AuditData.Members) {
        If ($UserLookup[$Member.Upn]) {
           Write-Host ("User {0} added to team {1}" -f $Member.DisplayName, $AuditData.TeamName) 
           $ReportLine = [PSCustomObject]@{
             Date = $AuditData.CreationTime
             User = $Member.Upn   
             Name = $Member.DisplayName
             Team = $AuditData.TeamName
             AddedBy = $AuditData.UserId
           }
          $Report.Add($ReportLine)
        }
    }
}
