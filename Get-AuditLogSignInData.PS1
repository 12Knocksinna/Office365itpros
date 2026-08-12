# Get-AuditLogSignInData.PS1
# A script to demonstrate how to fetch Entra ID audit sign-in data using the Microsoft Graph PowerShell SDK one day at a time
# up to the maximum of 30 days. By breaking the data into one-day chunks, we can avoid HTTP client failures that can occur when fetching large amounts of data at once.

# V1.0 12-Aug-2026
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Get-AuditLogSignInData.PS1

Function Get-DateRangeData {
    param (
        [datetime]$RangeStart,
        [datetime]$RangeEnd
    )
    
    # Function to fetch message trace data for expanded events over a 10-day period.
    [array]$Logs = $null

    # Make sure that the start and end dates are UTC DateTimeOffset strings in ISO 8601 format (yyyy-MM-ddTHH:mm:ssZ) for the Microsoft Graph API
    $Start = $RangeStart.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $End = $RangeEnd.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    Write-Host ("Collecting audit log sign-in data between {0} and {1}..." -f $RangeStart, $RangeEnd)

    # The example is to fetch successful sign-in events that have used conditional access, but the same technique
    # can be used to fetch other types of sign-in events by changing the filter criteria.
    Try {
        [array]$Logs = Get-MgAuditLogSignIn -Filter "createdDateTime gt $start and createdDateTime lt $end and ConditionalAccessStatus eq 'Success'" -All `
            -ErrorAction Stop
 
    } Catch {
        Write-Host ("Error fetching audit log sign-in data: {0}" -f $_.Exception.Message)
        Return
    }
    
    return $Logs
}

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
    $Data = Get-DateRangeData -RangeStart $RangeStart -RangeEnd $RangeEnd
    # Process the data for this date range, for example by adding it to a report or exporting it to a file
    Write-Host ("Fetched {0} audit sign-in logs between {1} and {2}" -f $Data.count, $DateRange.Start, $DateRange.End)
    $Logs += $Data
}

# All done - tell them how many logs were fetched in total
Write-Host ("Fetched a total of {0} audit sign-in logs over the last 30 days." -f $Logs.count)

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.