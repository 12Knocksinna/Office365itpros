# Check-CopilotLicenseDetails.PS1
# Check the details of the Microsoft 365 Copilot license assigned to a user account and report same
# Just like the ExRCA Copilot license check

# V1.0 26-Aug-2025
# GitHub: https://github.com/12Knocksinna/Office365itpros/blob/master/Check-CopilotLicenseDetails.PS1

# Signed=in user needs permission to manage user accounts
Connect-MgGraph -NoWelcome -Scopes User.Read.All

$CopilotSKU = '639dec6b-bb19-468b-871c-c5c441c4b0cb'
$CopilotStudioPlan = 'fe6c28b3-d468-44ea-bbd0-a10a5167435c'
$CopilotSharePointPlan = '0aedf20c-091d-420b-aadf-30c042609612'
$CopilotGraphConnectorsPlan = '82d30987-df9b-4486-b146-198b21d164c7'
$CopilotConnectorsPlan = '89f1c4c8-0878-40f7-804d-869c9128ab5d'
$CopilotProductivityAppsPlan = 'a62f8878-de10-42f3-b68f-6149a25ceb97'
$CopilotTeamsPlan = 'b95945de-b3bd-46db-8437-f2beb6ea2347'
$CopilotChatPlan = '3f30311c-6b1e-48a4-ab79-725b469da960'
$CopilotIntelligentSearchPlan = '931e4a88-a67f-48b5-814f-16a5f1e6028d'

$User = Read-Host "Enter the UPN of the user to check"

Try {
    $UserDetails = Get-MgUser -UserId $User -ErrorAction Stop -Property AssignedLicenses, assignedPlans, DisplayName, UserPrincipalName, Id
}
Catch {
    Write-Host ("Error retrieving user {0}: $_") -f $User
    Break
}

[int]$FeatureCount = 0
If ($CopilotSKU -in $UserDetails.AssignedLicenses.SkuId) {
    Write-Host ("✔️ {0} has a valid Microsoft 365 Copilot license assigned." -f $User)
} else {
    Write-Host ("❌ {0} does not have the Microsoft 365 Copilot license assigned." -f $User)
    Break
}
If ($CopilotStudioPlan -in $UserDetails.AssignedPlans.ServicePlanId) {
    Write-Host ("✔️ Copilot Studio is enabled for {0}." -f $User)
    $FeatureCount++
} else {
    Write-Host ("❌ User {0} does not have the Copilot Studio plan assigned." -f $User)
}
If ($CopilotSharePointPlan -in $UserDetails.AssignedPlans.ServicePlanId) {
    Write-Host ("✔️ Copilot for SharePoint is enabled for {0}." -f $User)
    $FeatureCount++
} else {
    Write-Host ("❌ User {0} does not have the Copilot for SharePoint plan assigned." -f $User)
}
If ($CopilotGraphConnectorsPlan -in $UserDetails.AssignedPlans.ServicePlanId) {
    Write-Host ("✔️ Copilot Graph Connectors is enabled for {0}." -f $User)
    $FeatureCount++
} else {
    Write-Host ("❌ User {0} does not have the Copilot Graph Connectors plan assigned." -f $User)
}
If ($CopilotConnectorsPlan -in $UserDetails.AssignedPlans.ServicePlanId) {
    Write-Host ("✔️ Copilot Power Platform Connectors is enabled for {0}." -f $User)
    $FeatureCount++
} else {
    Write-Host ("User {0} does not have the Copilot Power Platform Connectors plan assigned." -f $User)
}
If ($CopilotProductivityAppsPlan -in $UserDetails.AssignedPlans.ServicePlanId) {
    Write-Host ("✔️ Copilot in Productivity Apps (Office) is enabled for {0}." -f $User)
    $FeatureCount++
} else {
    Write-Host ("❌ User {0} does not have the Copilot in Productivity Apps plan assigned." -f $User)
}
If ($CopilotTeamsPlan -in $UserDetails.AssignedPlans.ServicePlanId) {
    Write-Host ("✔️ Copilot for Teams is enabled for {0}." -f $User)
    $FeatureCount++ 
} else {
    Write-Host ("❌ User {0} does not have the Copilot for Teams plan assigned." -f $User)
}
If ($CopilotChatPlan -in $UserDetails.AssignedPlans.ServicePlanId) {
    Write-Host ("✔️ Copilot with Graph-grounded Chat is enabled for {0}." -f $User)
    $FeatureCount++
} else {
    Write-Host ("❌ User {0} does not have the Copilot with Graph-grounded Chat plan assigned." -f $User)
}
If ($CopilotIntelligentSearchPlan -in $UserDetails.AssignedPlans.ServicePlanId) {
    Write-Host ("✔️ Copilot with Intelligent Search is enabled for {0}." -f $User)
    $FeatureCount++
} else {
    Write-Host ("❌ User {0} does not have the Copilot with Intelligent Search plan assigned." -f $User)
}
Write-Host ("Microsoft 365 Copilot features enabled for {0}: {1} of 8" -f $User, $FeatureCount)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.