# Test-BatchProcessing.PS1
# A script to demonstrate how to use batch processing with Microsoft Graph API  by using mailboxes fetched from Exchange
# to get user account information.
# https://github.com/12Knocksinna/Office365itpros/blob/master/Test-BatchProcessing.PS1
# 15-May-2024

# Requires PowerShell V7 or later

#Requires -Version 7.4

$CSVOutputFile = "C:\Temp\UserReport.csv"
$SkuDataPath = "C:\temp\SkuDataComplete.csv"
If ((Test-Path $SkuDataPath) -eq $False) {
    Write-Host ("Can't find the product data file ({0}). Exiting..." -f $SkuDataPath) ; break 
}
$ImportSkus = Import-CSV $SkuDataPath
$SkuHashTable = @{}
ForEach ($Line in $ImportSkus) { 
    $SkuHashTable.Add([string]$Line.SkuId, [string]$Line.DisplayName) 
}

# Connects to the Graph (to read user data) and Exchange Online (to get mailboxes)
Connect-MgGraph -Scopes "User.Read.All" -NoWelcome
Connect-ExchangeOnline -SkipLoadingCmdletHelp

# Fetch all mailboxes
Write-Host "Fetching mailboxes..."
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited
If ($Mbx.Count -eq 0) { 
    Write-Host "No mailboxes found" 
    Break
} Else {
    Write-Host ("Total of {0} mailboxes found" -f $Mbx.Count)
}

# Define batch size (maxiumum 20)
$BatchSize = 20
# Define output list (concurrent bag to store user details)
$UserDetails = [System.Collections.Concurrent.ConcurrentBag[System.Object]]::new()
$Batches = For($i = 0; $i -lt $Mbx.count; $i += $BatchSize) {
    $End = $i + $BatchSize - 1
    If ($end -ge $Mbx.count) { 
        $end = $Mbx.count 
    }
    $Index = $i
    # Create requests from the next batch of mailboxes
    $Index = $i
    $Requests = $Mbx[$i..($end)] 
    # For each mailbox in the batch, create a request to get user details
    $RequestData = [System.Collections.Generic.List[Object]]::new()
    ForEach ($Request in $Requests) {
        # the URL is the Graph lookup request to fetch user account details
        $Url = "users/{0}?`$select=id,displayname,assignedLicenses,country,city,jobtitle,officelocation,userprincipalname,businessphones,employeeid,employeehiredate" -f $Request.ExternalDirectoryObjectId
        $ReportLine = [PSCustomObject]@{
            Id     = $Index++
            Method = 'GET'
            Url    = $Url
        }
        $RequestData.Add($ReportLine)
    }

    # Create a batch request
    @{
        'Method'      = 'Post'
        'Uri'         = 'https://graph.microsoft.com/v1.0/$batch'
        'ContentType' = 'application/json'
        'Body'        = @{
            'requests' = @($RequestData)
        } | ConvertTo-Json
    }
}

$Batches | ForEach-Object -Parallel {
    $Responses = $using:UserDetails
    $RequestSubmission = Invoke-MgGraphRequest @PSItem
     # Invoke-MgGraphRequest deserializes request to a hashtable
    $RequestSubmission.responses | ForEach-Object { $responses.Add([pscustomobject]$PSItem.Body) }
}
If ($UserDetails.Count -ne $Mbx.Count) { 
        throw [System.Exception]::new() 
}
Write-Host ("Total of {0} user accounts processed" -f $UserDetails.count)

# Create a hash table containing the user data
$UserHash = @{}
ForEach ($User in $UserDetails) {
    $UserHash.Add($User.Id,$User)
}
# Create the report by combining mailbox and user data
$UserReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Mailbox in $Mbx) {
    $User = $UserHash[$Mailbox.ExternalDirectoryObjectId]
    # Resolve product license names from the SKU identifiers    
    [array]$AssignedLicenses = $null
    ForEach ($License in $User.assignedLicenses.SkuId) {
        $SKUName = $SkuHashTable[$License]
        $AssignedLicenses += $SKUName
    }
    $UserReport.Add([PSCustomObject]@{
        DisplayName             = $User.DisplayName
        UserPrincipalName       = $User.UserPrincipalName
        SMTPAddress             = $Mailbox.PrimarySmtpAddress
        City                    = $User.City
        Country                 = $User.Country
        'Job Title'             = $User.JobTitle
        'Office Location'       = $User.OfficeLocation
        'Business Phones'       = $User.BusinessPhones -join ', '
        'Assigned Licenses'     = $AssignedLicenses -join ', '
        EmployeeId              = $User.EmployeeId
        'Employee Hire Date'    = $User.EmployeeHireDate
    })
}
# Generate a CSV file and display the report in a grid view
$UserReport | Export-CSV -Path $CSVOutputFile -NoTypeInformation -Encoding UTF8
Write-Host ("CSV file available in {0}" -f $CSVOutputFile)
$UserReport | Sort-Object DisplayName | Out-GridView -Title "User Account Information" 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
