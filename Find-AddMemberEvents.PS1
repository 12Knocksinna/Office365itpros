
# Connect to the Graph
Connect-MgGraph
# Find users who have Project in their display name
[array]$Users = Get-MgUser -Search "displayName:Project" -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel Eventual
If (!($Users)) { 
    Throw "No users found"
}
# Build hash table of users that we want to check
$UserLookup = @{}
ForEach ($User in $Users) {
   $UserLookup.Add($User.UserPrincipalName, $User.DisplayName)
}