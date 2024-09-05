# Find-GroupsNoOwnersOrMembers.PS1
# example of a script to find and report groups with no owners or members
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-GroupsNoOwnersOrMembers.PS1
# V1.0 30-May-2024

$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}
# Connect to the Microsoft Graph
Connect-MgGraph -NoWelcome -Scopes Group.Read.All
$NoOwnersCSVFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\GroupsWithNoOwners.CSV"
# Find groups and expand owners. Only one property can be expanded at a time
[array]$NoOwners = Get-MgGroup -All -PageSize 999 -ExpandProperty "owners(`$select=id)" `
    -Property Owners, Id, displayName, groupTypes, CreatedDateTime, mailEnabled, Mail, SecurityEnabled | `
    Select-Object Owners, Id, displayName, groupTypes, CreatedDateTime, mailEnabled, Mail, SecurityEnabled | `
    Where-Object {($_.Owners.count) -eq 0} | Sort-Object DisplayName

If ($NoOwners.Count -eq 0) {
    Write-Host "No groups without owners found"
} Else {
    Write-Host ("Found {0} ownerless groups" -f $NoOwners.Count)
      
        $NoOwnersReport = [System.Collections.Generic.List[Object]]::new()
        ForEach ($G in $NoOwners) {
            $DistributionList = $false
            # Check if the group is a distribution list
            If ($G.GroupTypes.count -eq 0 -and $G.MailEnabled -eq $true) {
                $TestDL = Get-DistributionGroup -Identity $G.Id -ErrorAction SilentlyContinue
                If ($TestDL) {
                    $DistributionList = $true
                }
            }
            $ReportLine = [PSCustomObject][Ordered]@{
                ID                  = $G.Id  
                DisplayName         = $G.displayName
                Created             = Get-Date($G.CreatedDateTime).ToLocalTime()
                groupTypes          = ($G.groupTypes -join ", ")
                Mail                = $G.Mail
                mailEnabled         = $G.mailEnabled
                SecurityEnabled     = $G.SecurityEnabled
                'Distribution list' = $DistributionList
            }
            $NoOwnersReport.Add($ReportLine)
        }
        $NoOwnersReport | Export-Csv -Path $NoOwnersCSVFile -NoTypeInformation -Encoding utf8
        Write-Host ("Report for groups with no owners written to {0}" -f $NoOwnersCSVFile)
}

# Now check groups with no members...
$NoMembersCSVFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\GroupsWithNoMembers.CSV"
Write-Host "Checking for groups with no members..."
[array]$NoMembers = Get-MgGroup -All -PageSize 999 -ExpandProperty "members(`$select=id)" `
    -Property Members, Id, displayName, groupTypes, CreatedDateTime, mailEnabled, Mail, SecurityEnabled | `
    Select-Object Members, Id, displayName, groupTypes, CreatedDateTime, mailEnabled, Mail, SecurityEnabled | `
    Where-Object {($_.Members.count) -eq 0} | Sort-Object DisplayName

If ($NoMembers.Count -eq 0) {
    Write-Host "No groups without members found"
} Else {
    Write-Host ("Found {0} memberless groups" -f $NoMembers.Count)
      
        $NoMembersReport = [System.Collections.Generic.List[Object]]::new()
        ForEach ($G in $NoMembers) {
            $DistributionList = $false
            # Check if the group is a distribution list
            If ($G.GroupTypes.count -eq 0 -and $G.MailEnabled -eq $true) {
                $TestDL = Get-DistributionGroup -Identity $G.Id -ErrorAction SilentlyContinue
                If ($TestDL) {
                    $DistributionList = $true
                }
            }
            $ReportLine = [PSCustomObject][Ordered]@{
                ID                  = $G.Id  
                DisplayName         = $G.displayName
                Created             = Get-Date($G.CreatedDateTime).ToLocalTime()
                groupTypes          = ($G.groupTypes -join ", ")
                Mail                = $G.Mail
                mailEnabled         = $G.mailEnabled
                SecurityEnabled     = $G.SecurityEnabled
                'Distribution list' = $DistributionList
            }
            $NoMembersReport.Add($ReportLine)
        }
        $NoMembersReport | Export-Csv -Path $NoMembersCSVFile -NoTypeInformation -Encoding utf8
        Write-Host ("Report for groups with no members written to {0}" -f $NoMembersCSVFile)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.