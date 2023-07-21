# FindUserAuditActivities.PS1
# A script to demonstrate the principal of using the Microsoft 365 audit log to find information about user activities
# for the past week to help determine if the account has been comppromised by an attacker
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindUserAuditActivities.PS1

function Get-IPGeoLocation {
 Param   ([string]$IPAddress)

  $IPInfo = Invoke-RestMethod -Method Get -Uri "http://ip-api.com/json/$IPAddress"

  [PSCustomObject]@{
     IP      = $IPInfo.Query
     City    = $IPInfo.City
     Country = $IPInfo.Country
     Region  = $IPInfo.Region
     Isp     = $IPInfo.Isp   }

}

# Function to convert a CIDR IPv4 range to individual IP addresses 
# (from https://www.powershellgallery.com/packages/PoshFunctions/2.2.1.6/Content/Functions%5CGet-IpRange.ps1)
Function Get-IpRange {

    [CmdletBinding(ConfirmImpact = 'None')]
    Param(
        [Parameter(Mandatory, HelpMessage = 'Please enter a subnet in the form a.b.c.d/#', ValueFromPipeline, Position = 0)]
        [string[]] $Subnets
    )

    begin {
        Write-Verbose -Message "Starting [$($MyInvocation.Mycommand)]"
    }

    process {
        foreach ($subnet in $subnets) {
            if ($subnet -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/\d{1,2}$') {
                #Split IP and subnet
                $IP = ($Subnet -split '\/')[0]
                [int] $SubnetBits = ($Subnet -split '\/')[1]
                if ($SubnetBits -lt 7 -or $SubnetBits -gt 30) {
                    Write-Error -Message 'The number following the / must be between 7 and 30'
                    break
                }
                #Convert IP into binary
                #Split IP into different octects and for each one, figure out the binary with leading zeros and add to the total
                $Octets = $IP -split '\.'
                $IPInBinary = @()
                foreach ($Octet in $Octets) {
                    #convert to binary
                    $OctetInBinary = [convert]::ToString($Octet, 2)
                    #get length of binary string add leading zeros to make octet
                    $OctetInBinary = ('0' * (8 - ($OctetInBinary).Length) + $OctetInBinary)
                    $IPInBinary = $IPInBinary + $OctetInBinary
                }
                $IPInBinary = $IPInBinary -join ''
                #Get network ID by subtracting subnet mask
                $HostBits = 32 - $SubnetBits
                $NetworkIDInBinary = $IPInBinary.Substring(0, $SubnetBits)
                #Get host ID and get the first host ID by converting all 1s into 0s
                $HostIDInBinary = $IPInBinary.Substring($SubnetBits, $HostBits)
                $HostIDInBinary = $HostIDInBinary -replace '1', '0'
                #Work out all the host IDs in that subnet by cycling through $i from 1 up to max $HostIDInBinary (i.e. 1s stringed up to $HostBits)
                #Work out max $HostIDInBinary
                $imax = [convert]::ToInt32(('1' * $HostBits), 2) - 1
                $IPs = @()
                #Next ID is first network ID converted to decimal plus $i then converted to binary
                For ($i = 1 ; $i -le $imax ; $i++) {
                    #Convert to decimal and add $i
                    $NextHostIDInDecimal = ([convert]::ToInt32($HostIDInBinary, 2) + $i)
                    #Convert back to binary
                    $NextHostIDInBinary = [convert]::ToString($NextHostIDInDecimal, 2)
                    #Add leading zeros
                    #Number of zeros to add
                    $NoOfZerosToAdd = $HostIDInBinary.Length - $NextHostIDInBinary.Length
                    $NextHostIDInBinary = ('0' * $NoOfZerosToAdd) + $NextHostIDInBinary
                    #Work out next IP
                    #Add networkID to hostID
                    $NextIPInBinary = $NetworkIDInBinary + $NextHostIDInBinary
                    #Split into octets and separate by . then join
                    $IP = @()
                    For ($x = 1 ; $x -le 4 ; $x++) {
                        #Work out start character position
                        $StartCharNumber = ($x - 1) * 8
                        #Get octet in binary
                        $IPOctetInBinary = $NextIPInBinary.Substring($StartCharNumber, 8)
                        #Convert octet into decimal
                        $IPOctetInDecimal = [convert]::ToInt32($IPOctetInBinary, 2)
                        #Add octet to IP
                        $IP += $IPOctetInDecimal
                    }
                    #Separate by .
                    $IP = $IP -join '.'
                    $IPs += $IP
                }
                Write-Output -InputObject $IPs
            } else {
                Write-Error -Message "Subnet [$subnet] is not in a valid format"
            }
        }
    }

    end {
        Write-Verbose -Message "Ending [$($MyInvocation.Mycommand)]"
    }
}

# Start by connecting to the modules we need
Connect-MgGraph -Scopes Policy.Read.All
Connect-ExchangeOnline

[array]$IPAddressRanges = $Null 
[array]$IPAddresses = $Null
$Now = Get-Date
$StartTime = (Get-Date).AddDays(-7)
# Hash table for resolved IP addresses
$IPAddressHash = @{}

# This section attempts to load known IP locations from a CSV file. If it doesn't exist, we
# try and fetch IP locations from those defined for Conditional access policies.
$IPInfoFile = "C:\Temp\IPAddressData.txt"
If (Test-Path -Path $IPInfoFile -PathType Leaf) {
   # Import the data from the file
   [array]$IPAddresses = Get-Content $IPInfoFile
   Write-Host ("Found file containing internal IP addresses {0}" -f $IPInfoFile)
} Else {
  Write-Host "Checking conditional access IP locations"
# Find out if the tenant has any IP locations defined for conditional access policy
   [array]$CAKnownLocations = Get-MgIdentityConditionalAccessNamedLocation
   If ($CAKnownLocations) {
      ForEach ($Location in $CAKnownLocations) {
       $IPRanges = $Null
       $IPRanges = $Location.AdditionalProperties['ipRanges']
       If ($IPRanges) {
          ForEach ($Address in $IPRanges) {
            $IPAddressRanges += $Address['cidrAddress']
        } #End ForEach $IPRanges
       } # End if $IPRanges
    }  # End ForEach Location
  } # End CA Locations 

 # We don't handle IPV6 addresses for the purpose of this demo
 $IPAddressRanges = $IPAddressRanges | Where-Object {$_ -notlike "*::/*"}
 If ($IPAddressRanges) {
   # Resolve the CIDR used by conditional access into individual IP addresses
   [array]$IPAddresses = Get-IpRange -Subnets $IPAddressRanges }
 
   #  Add some addresses here if you want. For example
   $IPAddresses += "2001:bb6:5f1e:a900:f5fa:4963:a6a9:4128", "2001:bb6:5f1e:a900:57:9971:f615:e6bb", "2001:bb6:5f1e:a900:fcfa:981:71b7:f5c8", "2001:bb6:5f1e:a900:e592:65bb:b9d9:19b5", "2001:bb6:5f1e:a900:800f:c6d0:2c98:f11", "2001:bb6:5f1e:a900:219a:8a41:24c6:54cd", "2001:bb6:5f1e:a900:98cc:ccd7:b59:7b5c", "2001:bb6:5f1e:a900:2d77:d671:29b8:e13a"

  # Remove any duplicates that might have snuck in
  [array]$IPAdresses = $IPAddresses | Sort-Object -Unique   
  $IPAddresses | Out-File -FilePath $IPInfoFile
  Write-Host ("Saved file containing {0} IP addresses used for internal check in {1}" -f $IPAddresses.count, $IPInfoFile)

  # The $IPAddresses array now contains all the individual IP addresses in the CIDRs used by CA policies
}

$User = Read-Host "Enter name of user to search for"
[array]$Mbx = (Get-ExoMailbox -Identity $User -ErrorAction SilentlyContinue)
If (!($Mbx)) { 
    Write-Host ("Can't find the account for {0} - exiting" -f $User) ; break 
}

[array]$Operations = "UserLoggedIn", "FileAccessed", "FileDownloaded", "SendAs", "Set-InboxRule", "New-InboxRule"
Write-Host ("Searching for audit records for {0}..." -f $Mbx.UserPrincipalName)
[array]$Records = Search-UnifiedAuditLog -UserId $Mbx.UserPrincipalName -StartDate $StartTime -EndDate $Now -ResultSize 5000 -Formatted -Operations $Operations
Write-Host ("{0} records found." -f $Records.count)
If (!($Records)) { Write-Host "Exiting because no audit records can be found..." ; break }

$Records | Group operations -NoElement | Sort-Object Count -Descending | Format-Table Name, Count -AutoSize
$AuditInfo = [System.Collections.Generic.List[Object]]::new() 
[int]$IPLookups = 0

ForEach ($Rec in $Records) {
 $AuditData = $Rec.AuditData | ConvertFrom-Json

 # Check IP address against hash table. If it's not in the table, resolve the address and store the results.
 $IPInfo = $Null
 If (!($IPAddressHash[$AuditData.ClientIP])) {
   Write-Host "Querying IP Geolocation data for " $AuditData.ClientIP -foregroundcolor Red
   $IPLookups++
   $IPInfo = Get-IPGeoLocation -IPAddress $AuditData.ClientIP
   Try {
      $Status = $IPAddressHash.Add([string]$IPInfo.IP,$IPInfo) 
   } Catch {
      Write-Host ("Unable to add IP information for {0} to the hash table" -f $AuditData.ClientIP) 
   }
   # Sleep to avoid any throttling issues with the web service
   Start-Sleep -Seconds 1
 } Else {
   # Get the IP information from the hash table
   $IPInfo = $IpAddressHash[$AuditData.ClientIP]
 }
 # Brief pause to avoid any geolocation service throttling
 If ($IPLookups -eq 44) {
    Start-Sleep -Seconds 15
    $IpLookups = 0 }

 # Is this an internal IP address?
 If ($AuditData.ClientIP -in $IPAddresses) {
   $InternalFlag = $True 
 } Else {
  $InternalFlag = $False }

$ClientInfo = $Null; $SendAsUser = $Null; $Mailbox = $Null; $RuleId = $Null; $RuleName = $Null; $RedirectTo = $Null
$OS = $Null; $DeviceName = $Null; $CompliantDevice = $Null; $UserAgent = $Null; $SPOSite = $Null; $SPOLibrary = $Null; $SPODocument = $Null

Switch ($Rec.Operations) {
   "UserLoggedIn"  {
     $OS              = $AuditData.deviceproperties | Where-Object {$_.Name -eq "OS"} | Select-Object -ExpandProperty Value
     $DeviceName      = $AuditData.deviceproperties | Where-Object {$_.Name -eq "DisplayName"} | Select-Object -ExpandProperty Value
     $CompliantDevice = $AuditData.deviceproperties | Where-Object {$_.Name -eq "IsCompliantAndManaged"} | Select-Object -ExpandProperty Value
  }
  "FileAccessed" {
     $SPOSite         = $AuditData.SiteURL
     $SPODocument     = $AuditData.SourceFileName
     $SPOLibrary      = $AuditData.SourceRelativeURL
     $UserAgent       = $AuditData.UserAgent
  }
  "FileDownloaded" {
     $SPOSite         = $AuditData.SiteURL
     $SPODocument     = $AuditData.SourceFileName
     $SPOLibrary      = $AuditData.SourceRelativeURL
     $UserAgent       = $AuditData.UserAgent  
  }
   "SendAs"   {
     $UserAgent       = $AuditData.UserAgent
     $ClientInfo      = $AuditData.ClientInfoString
     $Mailbox         = $AuditData.MailboxOwnerUPN
     $SendAsUser      = $AuditData.SendAsUserSmtp
  }
   "New-InboxRule" {
     $RuleId          = $Null
     $RuleName        = $AuditData.Parameters | Where-Object {$_.Name -eq "Identity"} | Select-Object -ExpandProperty Value
     $RedirectTo      = $AuditData.Parameters | Where-Object {$_.Name -eq "RedirectTo"} | Select-Object -ExpandProperty Value
   }
   "Set-InboxRule" {
     $RuleId          = $AuditData.ObjectId
     $RuleName        = $AuditData.Parameters | Where-Object {$_.Name -eq "Identity"} | Select-Object -ExpandProperty Value
     $RedirectTo      = $AuditData.Parameters | Where-Object {$_.Name -eq "RedirectTo"} | Select-Object -ExpandProperty Value
  }
}

 $DataLine  = [PSCustomObject] @{
           Timestamp    = $Rec.CreationDate
           User         = $Rec.UserIds
           Operation    = $Rec.Operations
           Device       = $DeviceName
           OS           = $OS
           Compliant    = $CompliantDevice
           ClientInfo   = $ClientInfo
           IP           = $AuditData.ClientIP
           City         = $IPInfo.City
           Country      = $IPInfo.Country
           ISP          = $IPInfo.ISP
           Internal     = $InternalFlag
           Site         = $SPOSite
           Library      = $SPOLibrary
           Document     = $SPODocument
           Mailbox      = $Mailbox
           SendAsUser   = $SendAsUser
           RuleId       = $RuleId
           RuleName     = $RuleName
           RedirectTo   = $RedirectTo
         }
$AuditInfo.Add($DataLine)

} # End of processing audit records

Write-Host
Write-Host "Audit records found originating in these cities:"
Write-Host ""
$AuditInfo | Group-Object City -NoElement | Sort-Object Count -Descending | Format-Table Count, Name

[array]$ExternalIPAccess = $AuditInfo | Where-Object {$_.Internal -eq $False}
Write-Host ""
Write-Host ("{0} records found from external IP addresses" -f $ExternalIPAccess.count)
$ExternalIpAccess | Sort-Object IP | Format-Table IP, City, ISP
$ExternalIPAccess | Format-Table Timestamp, Operation, City, Country, ISP, IP

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
