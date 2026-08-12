# Get-AuditLogSignInData.PS1
# A script to demonstrate how to fetch Entra ID audit sign-in data using the Microsoft Graph PowerShell SDK one day at a time
# up to the maximum of 30 days. By breaking the data into one-day chunks, we can avoid HTTP client failures that can occur when fetching large amounts of data at once.

# V1.0 12-Aug-2026
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Get-AuditLogSignInData.PS1

Function Get-DateRangeData {
    param (
        [string]$Filter
    )
    
    # Function to fetch audit sign-in data based on the provided filter.
    [array]$Logs = $null

    Write-Host ("Collecting audit log sign-in data between {0} and {1}..." -f $RangeStart, $RangeEnd)

    Try {
        [array]$Logs = Get-MgAuditLogSignIn -Filter $Filter -All -ErrorAction Stop
 
    } Catch {
        Write-Host ("Error fetching audit log sign-in data: {0}" -f $_.Exception.Message)
        Return
    }
    
    return $Logs
}

Connect-MgGraph -Scopes "AuditLog.Read.All"
# Signed-in account must have an Entra P1 license to fetch audit log via an API

# This code creates an array of one-day ranges to fetch the audit log data.
$StartDate = (Get-Date).AddDays(-30).ToUniversalTime()
$EndDate = (Get-Date).ToUniversalTime().AddMinutes(10)

$iterations = 30 # 30 days of data to fetch
[int]$i = 0

[array]$DateRanges = $null

For ($i=0; $i -lt $iterations; $i++) {
    $RangeStart = $StartDate.AddDays($i * 1)
    $RangeEnd = $StartDate.AddDays(($i + 1) * 1)
    If ($RangeEnd -gt $EndDate) {
        $RangeEnd = $EndDate
    }

    $DateRanges += @{Start=$RangeStart; End=$RangeEnd}
}

[array]$Data = $null
[array]$Logs = $null

# Fetch each chunk of audit sign-in data and add it to the Logs array
ForEach ($DateRange in $DateRanges) {
    $RangeStart = $DateRange.Start
    $RangeEnd = $DateRange.End
    # Make sure that the start and end dates are UTC DateTimeOffset strings in ISO 8601 format (yyyy-MM-ddTHH:mm:ssZ) for the Microsoft Graph API
    $Start = $RangeStart.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $End = $RangeEnd.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    # Any valid filter can be used. This example fetches only successful sign-in events that have used conditional access.
    # You can modify the filter to suit your needs.
    $Filter = "createdDateTime gt $start and createdDateTime lt $end and ConditionalAccessStatus eq 'Success'"
    [array]$Data = Get-DateRangeData -Filter $Filter
    # Process the data for this date range by adding it to the Logs array
    Write-Host ("Fetched {0} audit sign-in logs between {1} and {2}" -f $Data.count, $DateRange.Start, $DateRange.End)
    $Logs += $Data
}

# All done - tell them how many logs were fetched in total
Write-Host ("Fetched a total of {0} audit sign-in logs over the last {1} days." -f $Logs.count, $iterations)

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.