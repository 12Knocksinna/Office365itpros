<#
.SYNOPSIS
Get-IPGeolocation.ps1 - Get IP address geolocation data
https://github.com/12Knocksinna/Office365itpros/blob/master/Get-IPGeolocation.ps1

.DESCRIPTION 
This PowerShell script performs a REST API query against the IP-API endpoint to retrieve geolocation information for an IP address. See 
https://ip-api.com/docs/unban#:~:text=IP%2DAPI%20endpoints%20are%20now,rate%20limit%20window%20is%20reset for information about rate 
throttling.

.OUTPUTS
Results are output to the console. Include the function in a script and you can use it to return information for
data such as the IP addresses in Microsoft 365 audit records.

.PARAMETER IPAddress
Specifies the IP address to lookup.

.EXAMPLE
.\Get-IPGeolocation.ps1 -IpAddress 1.1.1.1

$IPInfo = Get-IPGeoLocation $IPAddress

.NOTES
Written by: Tony Redmond
Uses the IP Geolocation API, which is free for commercial use and doesn't require an API key at the time of writing

V1.0 13 March 2023

#>

function Get-IPGeolocation {
 Param   ([string]$IPAddress)

  $IPInfo = Invoke-RestMethod -Method Get -Uri "http://ip-api.com/json/$IPAddress"

  [PSCustomObject]@{
     IP      = $IPInfo.Query
     City    = $IPInfo.City
     Country = $IPInfo.Country
     Region  = $IPInfo.Region
     Isp     = $IPInfo.Isp   }

}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
