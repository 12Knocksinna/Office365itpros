# Test-BatchProcessing.PS1
# A script to demonstrate how to use batch processing with Microsoft Graph API  by using mailboxes fetched from Exchange
# to get user account information.
# Requires PowerShell V7 or later

#Requires -Version 7.4

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
$Batches = For($i = 0; $i -lt $mbx.count; $i += $BatchSize) {
    $End = $i + $BatchSize - 1
    If ($end -ge $mbx.count) { 
        $end = $mbx.count 
    }
    $Index = $i
    # Create requests from the next batch of mailboxes
    $Index = $i
    $Requests = $Mbx[$i..($end)] 
    # For each mailbox in the batch, create a request to get user details
    $RequestData = [System.Collections.Generic.List[Object]]::new()
    ForEach ($Request in $Requests) {
        # the URL is the Graph lookup request to fetch user account details
        $Url = "users/{0}?`$select=id,displayname,assignedLicenses,country,jobtitle,officelocation,userprincipalname,businessphones,employeeid,employeehiredate " -f $Request.ExternalDirectoryObjectId
        $ReportLine = [PSCustomObject]@{
            Id     = $Index++
            Method = 'GET'
            Url    = $Url
        }
        $RequestData.Add($ReportLine)
    }

    # This is the original method described in https://nicolasuter.medium.com/optimising-microsoft-graph-powershell-scripts-3eb76a280077
    #$Requests = $Mbx[$i..($end)] | ForEach-Object {
    #    [PSCustomObject]@{
    #        'Id'     = ++$index
    #        'Method' = 'GET'
    #        'Url'    = "users/{0}?`$select=id,displayname,assignedLicenses,country,jobtitle,officelocation,userprincipalname,businessphones,employeeid,employeehiredate " -f $PSItem.ExternalDirectoryObjectId
    #    }
    #}

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
[array]$UserData = $UserDetails | Sort-Object displayName 
Write-Host ("Total of {0} user accounts processed" -f $UserData.count)
