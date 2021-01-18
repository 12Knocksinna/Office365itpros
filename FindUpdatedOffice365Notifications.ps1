# Find Updated Service Notifications - FindUpdatedOffice365Notifications.ps1
# Download the Office 365 notifications posted in the message center in the Microsoft 365 admin center and process them with PowerShell.
# https://github.com/12Knocksinna/Office365itpros/blob/master/FetchServiceMessagesGraph.ps1
# Need to make sure that these values are correct for the target tenant and the app being used to access the data
#
$AppId = "d716b32c-0edb-48be-9385-30a9cfd96155"         # Registered App in Azure AD
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f347a"      # Renant GUID (use Get-AzureTenantDetail to get this information
$AppSecret = 's_rkvIn1oZ1cNceUBvJ2or1lrrIsb*:='         # App Secret for the registered app

$body = @{grant_type="client_credentials";resource="https://manage.office.com";client_id=$AppId;client_secret=$AppSecret }
$oauth = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($tenantId)/oauth2/token?api-version=1.0" -Body $body
$token = @{'Authorization' = "$($oauth.token_type) $($oauth.access_token)" }

$Uri = "https://manage.office.com/api/v1.0/b662313f-14fc-43a2-9a7a-d2e27f4f3478/ServiceComms/Messages?`$filter=MessageType eq 'MessageCenter'"
$Messages = Invoke-RestMethod -Uri $uri -Headers $token -Method Get

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Messages.Value) {

If ($M.LastUpdateTime -eq $Null) {
   $LastUpdate = "None" }
   Else { $LastUpdate = et-Date ($M.LastUpdateTime) -format g }

# Set flags to indicate affected workloads
$Dynamics = $False; $Exchange = $False; $EOP = $False; $Forms = $False; $Intune = $False; $Lync = $False; $ATP = $False; $Flow = $False; $Teams = $False
$PowerApps = $False; $OfficeOnline = $False; $OneDrive = $False; $Platform = $False; $Client = $False; $Planner = $False; $SharePoint = $False
$Stream = $False; $Yammer = $False; $Office365 = $False
Foreach ($Wl in $M.AffectedWorkloadnames) {
  Switch ($Wl) {
    "DynamicsCRM"             { $Dynamics = $True }
    "Exchange"                { $Exchange = $True }
    "Fope"                    { $EOP = $True }
    "Forms"                   { $Forms = $True }
    "Intune"                  { $Intune = $True }
    "Lync"                    { $Lync = $True }
    "MDATP"                   { $ATP = $True }
    "MicrosoftFlow"           { $Flow = $True }
    "MicrosoftFlowM365"       { $Flow = $True}
    "MicrosoftTeams"          { $Teams = $True }
    "MobileDeviceManagement"  { $Intune = $True }
    "PowerApps"               { $PowerApps = $True }
    "PowerAppsM365"           { $PowerApps = $True }
    "OfficeOnline"            { $OfficeOnline = $True }
    "OneDriveForBusiness"     { $OneDrive = $True }
    "OrgLiveId"               { $Platform = $True }
    "OSDPPlatform"            { $Platform = $True }
    "O365Client"              { $Client = $True }
    "Planner"                 { $Planner = $True }
    "SharePoint"              { $SharePoint = $True }  
    "Stream"                  { $Stream = $True }  
    "Yammer"                  { $Yammer = $True }
     default                  { $Office365 = $True }
  } #End Switch
} #End Foreach

# For notifications issued as updates, grab the update date from the text of the notification; otherwise just get the first 200 characters of the text
If ($M.Messages.MessageText -Like "Updated*") { 
     $UpdateText = $M.Messages.MessageText.SubString(0,200)
     $UpdateDate = $UpdateText.Substring(0,$Updatetext.IndexOf(":")) 
     $UpdateDate = $UpdateDate.SubString(8,($UpdateDate.length-8))
     [datetime]$StartPeriod = $M.StartTime
     [datetime]$EndPeriod   = $UpdateDate
     $DaysUpdate = (New-TimeSpan -Start $StartPeriod -End $EndPeriod).Days  }
  Else {
     $UpdateText = $M.Messages.MessageText.SubString(3,203) 
     $UpdateDate = $Null 
     $DaysUpdate = "N/A" }
     
# Generate output report line for the notification    
    $ReportLine = [PSCustomObject]@{  
     Id           = $M.Id
     Title        = $M.Title
     Category     = $M.Category
     ActionType   = $M.ActionType
     Text         = $UpdateText
     Updated      = $UpdateDate
     DaysUpdate   = $DaysUpdate
     Link         = $M.ExternalLink
     HelpLink     = $M.HelpLink
     Workloads    = $M.AffectedWorkloadDisplayNames
     StartDate    = Get-Date ($M.StartTime) -format g
     LastUpdate   = $LastUpdate
     EndDate      = Get-Date($M.Endtime) -format g
     Dynamics     = $Dynamics
     Exchange     = $Exchange
     EOP          = $EOP
     Forms        = $Forms
     Intune       = $Intune
     Lync         = $Lync
     ATP          = $ATP
     Flow         = $Flow
     Teams        = $Teams
     PowerApps    = $PowerApps
     OfficeOnline = $OfficeOnline
     OneDrive     = $OneDrive
     Platform     = $Platform
     Client       = $Client
     Planner      = $Planner
     SharePoint   = $SharePoint
     Stream       = $Stream
     Yammer       = $Yammer
     Office365    = $Office365  
    }
  $Report.Add($ReportLine) 

} # End ForEach
# Generate array of updated notifications 
[array]$Updates = $Report | ?{$_.Title -Like "*(Updated)*"}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
