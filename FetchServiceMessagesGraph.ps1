# FetchServiceMessagesGraph.ps1
# Fetch Service Messages from the Microsoft Graph
# https://github.com/12Knocksinna/Office365itpros/blob/master/FetchServiceMessagesGraph.ps1
CLS
Write-Host "Setting things up for the Graph"
# Define the values applicable for the application used to connect to the Graph - change these details for your tenant
$AppId = "e716b32c-0edb-48be-9385-30a9cfd96155"
$TenantId = "c662313f-14fc-43a2-9a7a-d2e27f4f3478"
$AppSecret = 's_rkvIn1oZ1cNceUBvJ2or1lrrIsb*:='
$Body = @{
        grant_type="client_credentials";
        resource="https://manage.office.com";
        client_id=$AppId;
        client_secret=$AppSecret}

# Get OAuth 2.0 Token
$Uri = "https://login.microsoftonline.com/$($tenantId)/oauth2/token?api-version=1.0"
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
# Unpack Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token
# Base URL
$headers = @{Authorization = "Bearer $token"}

# Fetch information from Graph
Write-Host "Fetching Office 365 Message Center Notifications..."
$MessageCenterURI = "https://manage.office.com/api/v1.0/$($tenantid)/ServiceComms/Messages"
$ServiceData = (Invoke-RestMethod -Uri $MessageCenterURI -Headers $Headers -Method Get -ContentType "application/json") 
$Messages = $ServiceData.Value | ? {$_.MessageType -eq "MessageCenter"}

# And Report what we find
Write-Host "Parsing what we got and creating a report..."
$Report = [System.Collections.Generic.List[Object]]::new() 
ForEach ($M in $Messages) {
   If ([string]::IsNullOrEmpty($M.AffectedWorkloadDisplayNames)) {  # Parse out workloads
      $Workloads = "Office 365" }
   Else {
      $Workloads = $M.AffectedWorkloadDisplayNames;  $i = 0
        ForEach ($W in $Workloads) {
        $i++
        If ($i = 1) {$Workloads = $W}
           Else {$Workloads = $Workloads + "; " + $W}  } }
   If ([string]::IsNullOrEmpty($M.ActionRequiredByDate)) { # No action required date
       $ActionRequiredDate = $Null }
   Else {$ActionRequiredDate = Get-Date($M.ActionRequiredByDate) -format "dd MMM yyyy" }
   # Get age of update
   $Age = New-TimeSpan($M.LastUpdatedTime)
   # Trim the message text
   $MessageText = $M.Messages.MessageText -replace [regex]::Escape("["), "<br><b>" -replace [regex]::Escape("]"), "</b><br><br>"      
   $ReportLine  = [PSCustomObject] @{          
     MessageId          = $M.Id
     Title              = $M.Title
     Workloads          = $Workloads
     ActionType         = $M.ActionType
     StartTime          = Get-Date($M.StartTime) -format g
     EndTime            = Get-Date($M.EndTime) -format g
     LastUpdatedTime    = Get-Date($M.LastUpdatedTime) -format g
     MileStoneDate      = Get-Date($M.MileStoneDate) -format g
     ActionRequiredDate = $ActionRequiredDate
     MessageText        = $MessageText
     Category           = $M.Category
     ExternalLInk       = $M.ExternalLink
     Severity           = $M.Severity
     Age                = $Age.Days.ToString() + ":" + $Age.Hours.ToString()
     IsRead             = $M.IsRead
     IsDismissed        = $M.IsDismissed}
   $Report.Add($ReportLine) } 

$Output = $Report | Sort {$_.LastUpdatedTime -as [DateTime]} -Descending | Select Title, Category, LastUpdatedTime, ActionRequiredDate, MessageId, Age
$Output | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
