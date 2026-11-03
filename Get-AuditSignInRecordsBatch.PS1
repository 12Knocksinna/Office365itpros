<#
.SYNOPSIS
Fetches Microsoft Entra sign-in records by using Microsoft Graph JSON batching.

.DESCRIPTION
Divides a UTC date range into nonoverlapping time windows and submits up to 20 requests in each call to the Microsoft Graph batch endpoint.

The script follows every @odata.nextLink returned for an individual request until the Graph returns no further pages.

"All records" means all sign-in records available through the Graph API for the specified period. Entra sign-in log retention periods and the
caller's permissions still apply.

.REQUIREMENTS
Microsoft.Graph.Authentication
Delegated permission: AuditLog.Read.All

The default is to fetch audit sign-in reciords for the last 7 days. You can specify a different date range with the -StartDate and -EndDate parameters.
Batches are generated to fetch records for 24 hours. This is controlled by the WindowsHour parameter. 
By default, the Graph returns 1000 records per request. Update the PageSize parameter to change the number of records returned per request. 

V1.0 3-Nov-2026
GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Get-AuditSignInRecordsBatch.PS1

.EXAMPLE
.\Get-AuditSignInRecordsBatch.ps1 -StartDate "2026-10-01T00:00:00Z" -EndDate "2026-11-01T00:00:00Z" -WindowHours 6 
#>

[CmdletBinding()]
param (
    [Parameter()]
    [datetime]$StartDate = (Get-Date).ToUniversalTime().AddDays(-7),

    [Parameter()]
    [datetime]$EndDate = (Get-Date).ToUniversalTime(),

    [Parameter()]
    [ValidateRange(1, 168)]
    [int]$WindowHours = 24,

    [Parameter()]
    [ValidateRange(1, 20)]
    [int]$BatchSize = 20,

    [Parameter()]
    [ValidateRange(0, 10)]
    [int]$MaximumRetries = 5,

    [Parameter()]
    [ValidateRange(1, 1000)]
    [int]$PageSize = 1000
)

function ConvertTo-GraphRelativeUri {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Uri
    )

    return $Uri -replace '^https://graph\.microsoft\.com/(v1\.0|beta)', ''
}

function New-SignInRequestQueue {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [datetime]$RangeStart,

        [Parameter(Mandatory)]
        [datetime]$RangeEnd,

        [Parameter(Mandatory)]
        [int]$Hours,

        [Parameter(Mandatory)]
        [int]$Top 
    )

    # Create a queue of requests to fetch sign-in records for the specified date range. Each request will cover a time slice of $Hours hours.
    $Queue = [System.Collections.Generic.Queue[object]]::new()
    $SliceStart = $RangeStart.ToUniversalTime()
    $FinalEnd = $RangeEnd.ToUniversalTime()
    $RequestNumber = 0

    While ($SliceStart -lt $FinalEnd) {
        $SliceEnd = $SliceStart.AddHours($Hours)

        If ($SliceEnd -gt $FinalEnd) {
            $SliceEnd = $FinalEnd
        }

        # Create batch details for this time slice.
        $RequestNumber++
        $StartText = $SliceStart.ToString("yyyy-MM-ddTHH:mm:ss.fffZ",[System.Globalization.CultureInfo]::InvariantCulture)
        $EndText = $SliceEnd.ToString("yyyy-MM-ddTHH:mm:ss.fffZ",[System.Globalization.CultureInfo]::InvariantCulture)
     
        # Using ge/lt prevents adjacent windows from overlapping.
        $Filter = "createdDateTime ge $StartText and createdDateTime lt $EndText"
        $EncodedFilter = [System.Uri]::EscapeDataString($Filter)

        # Add batch job details to the queue.
        $Queue.Enqueue([PSCustomObject]@{
                Id        = $RequestNumber.ToString()
                Method    = "GET"
                Url       = "/auditLogs/signIns?`$filter=$EncodedFilter&`$top=$Top"
                Retry     = 0
                NotBefore = [datetime]::MinValue
            }
        )

        $SliceStart = $SliceEnd
    }

      Write-Output -NoEnumerate $Queue
}

function Get-RetryAfterSeconds {
    [CmdletBinding()]
    param (
        [Parameter()]
        $Headers,

        [Parameter()]
        [int]$DefaultSeconds = 5
    )

    # Figure out when to retry based on the Retry-After header. If the header is missing or invalid, return the default value.
    If (-not $Headers) {
        return $DefaultSeconds
    }

    $RetryAfter = $null

    If ($Headers -is [System.Collections.IDictionary]) {
        ForEach ($Key in $Headers.Keys) {
            If ($Key -ieq "Retry-After") {
                $RetryAfter = $Headers[$Key]
                break
            }
        }
    } Else {
        $RetryAfter = $Headers.'Retry-After'

        If (-not $RetryAfter) {
            $RetryAfter = $Headers.retryAfter
        }
    }

    If (-not $RetryAfter) {
        return $DefaultSeconds
    }

    $RetryAfterSeconds = 0

    If ([int]::TryParse([string]$RetryAfter,[ref]$RetryAfterSeconds)) {
        return [Math]::Max(1, $RetryAfterSeconds)
    }

    return $DefaultSeconds
}

# Basic date checking to ensure the user didn't accidentally swap the start and end dates.
$StartUtc = $StartDate.ToUniversalTime()
$EndUtc = $EndDate.ToUniversalTime()

If ($StartUtc -ge $EndUtc) {
    Throw "StartDate must be earlier than EndDate."
}

# Check that we have the right permissions
$RequiredScope = "AuditLog.Read.All"
$Context = Get-MgContext -ErrorAction SilentlyContinue
    
If (-not $Context -or $Context.Scopes -notcontains $RequiredScope) {
    Connect-MgGraph -Scopes $RequiredScope -NoWelcome
}

# Number of seconds to delay between batch submissions. This is to avoid hitting the Microsoft Graph service too quickly and getting throttled.
$DelayBetweenBatchSubmission = 1

# Compute how many batches we need to process to find all sign-in records in the specified date range.
$PendingRequests = New-SignInRequestQueue -RangeStart $StartUtc -RangeEnd $EndUtc -Hours $WindowHours -Top $PageSize

# Check that we got a queue back from New-SignInRequestQueue. If not, throw an error.
If ($PendingRequests -isnot [System.Collections.Generic.Queue[object]]) {
    Throw "New-SignInRequestQueue returned $($PendingRequests.GetType().FullName) instead of a Queue."
}

# Create a list to hold all the sign-in records we retrieve. We'll deduplicate them later in case Microsoft Graph returns overlapping records.
$SignIns = [System.Collections.Generic.List[object]]::new()
$NextRequestId = $PendingRequests.Count
$BatchNumber = 0
$Now = Get-Date
Write-Output ("Starting to submit {0} batch jobs at {1}" -f $PendingRequests.count, (Get-Date $Now -format 'dd-MMM-yyyy HH:mm:ss'))

# Start the loop to submit the batch requests. We'll keep going until there are no more pending requests.
While ($PendingRequests.Count -gt 0) {
    $BatchNumber++

    $RequestsForBatch = [System.Collections.Generic.List[object]]::new()
    $DeferredRequests = [System.Collections.Generic.List[object]]::new()
    $Now = Get-Date

    # Fetch a batch job from the queue. If the request is not ready to be submitted, defer it until later.
    While ($PendingRequests.Count -gt 0 -and $RequestsForBatch.Count -lt $BatchSize) {
        $Request = $PendingRequests.Dequeue()

        If ($Request.NotBefore -le $Now) {
            $RequestsForBatch.Add($Request)
        } Else {
            $DeferredRequests.Add($Request)
        }
    }

    ForEach ($Request in $DeferredRequests) {
        $PendingRequests.Enqueue($Request)
    }

    # No requests are ready because every queued request is waiting for a retry.
    If ($RequestsForBatch.Count -eq 0) {
        $EarliestRetry = @($PendingRequests.ToArray() | Sort-Object NotBefore | Select-Object -First 1)[0].NotBefore

        $SecondsUntilRetry = ($EarliestRetry - (Get-Date)).TotalSeconds

        $SleepSeconds = [int][math]::Max(1,[math]::Ceiling($SecondsUntilRetry))

        Write-Verbose "Waiting $SleepSeconds seconds before retrying requests."
        Start-Sleep -Seconds $SleepSeconds
        continue
    }

    $GraphBatchRequests = @(
        Foreach ($Request in $RequestsForBatch) {
            [ordered]@{
                id     = $Request.Id
                method = $Request.Method
                url    = $Request.Url
            }
        }
    )

    $BatchBody = @{
        Requests = $GraphBatchRequests
    }

    Write-Output ("Submitting batch {0} with {1} requests; {2} requests remain queued." -f $BatchNumber, $GraphBatchRequests.Count, $PendingRequests.Count )
    # Let's submit the batch request to Microsoft Graph. If it fails, throw an error and stop processing.
    Try {
        $BatchResponse = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/`$batch" `
            -Body ($BatchBody | ConvertTo-Json -Depth 10) -ContentType "application/json" -ErrorAction Stop
        Write-Output ("Submitted batch number {0}" -f $BatchNumber)
        # Pause to avoid hitting Microsoft Graph with too many batch requeststoo quickly and getting throttled.
        Start-Sleep -Seconds $DelayBetweenBatchSubmission
    } Catch {
        Throw "Batch $BatchNumber failed: $($_.Exception.Message)"
    }

    # Check what the response was
    ForEach ($Response in ($BatchResponse.responses)) {
        $OriginalRequest = $RequestsForBatch | Where-Object Id -eq ([string]$Response.id) | Select-Object -First 1

        If (-not $OriginalRequest) {
            throw "No original request was found for response $($Response.id)."
        }

        $StatusCode = [int]$Response.status

        If ($StatusCode -ge 200 -and $StatusCode -lt 300) {
            ForEach ($SignIn in @($Response.body.value)) {
                if ($null -ne $SignIn) {
                    $SignIns.Add($SignIn)
                }
            }

            $NextLink = $Response.body.'@odata.nextLink'
            # If Microsoft Graph returned a nextLink, we need to queue up another request to fetch the next page of results.
            If ($NextLink) {
                $NextRequestId++

                $PendingRequests.Enqueue(
                    [PSCustomObject]@{
                        Id        = $NextRequestId.ToString()
                        Method    = "GET"
                        Url       = ConvertTo-GraphRelativeUri -Uri $NextLink
                        Retry     = 0
                        NotBefore = [datetime]::MinValue
                    }
                )
            }
        } Elseif ( $StatusCode -eq 429 -or $StatusCode -eq 503 -or $StatusCode -eq 504 ) {
            $RetryCount = $OriginalRequest.Retry + 1

            If ($RetryCount -gt $MaximumRetries) {
                Throw ("Request {0} exceeded the retry limit after HTTP {1}." -f $Response.id, $StatusCode)
            }

            $RetryAfter = Get-RetryAfterSeconds -Headers $Response.headers -DefaultSeconds ([math]::Min(60, [math]::Pow(2, $RetryCount)))

            $NextRequestId++

            $PendingRequests.Enqueue(
                [PSCustomObject]@{
                    Id        = $NextRequestId.ToString()
                    Method    = $OriginalRequest.Method
                    Url       = $OriginalRequest.Url
                    Retry     = $RetryCount
                    NotBefore = (Get-Date).AddSeconds($RetryAfter)
                }
            )

            Write-Warning ("Request {0} returned HTTP {1}. Retry {2} is scheduled in {3} seconds." -f $Response.id, $StatusCode, $RetryCount, $RetryAfter)
        } Else {
            # Error out if we get a non-successful response that is not a retryable status code.
            $ErrorCode = $Response.body.error.code
            $ErrorMessage = $Response.body.error.message

            Throw ("Subrequest {0} failed with HTTP {1}. {2}: {3}" -f $Response.id, $StatusCode, $ErrorCode, $ErrorMessage)
        }
    }

    Write-Output ( "{0:N0} sign-in records retrieved so far." -f $SignIns.Count )
}

# Deduplicate records just in case Microsoft Graph returns an unexpected overlap.
[array]$SignIns = $Signins | Sort-Object Id -Unique 

$AuditReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($AuditRecord in $SignIns) {

   If ($null -ne $AuditRecord.Status) { 
       If ($AuditRecord.Status.ErrorCode -eq 0) { 
           $SignInStatus = "Success" 
       } Else { 
           $SignInStatus = ("Failure: {0} {1}" -f $AuditRecord.Status.ErrorCode, $AuditRecord.Status.FailureReason)
       }
   } 

    $AuditReportLine = [PSCustomObject]@{
        CreatedDateTime     = (Get-Date $AuditRecord.CreatedDateTime -format 'dd-MMM-yyyy HH:mm:ss')
        UserPrincipalName   = $AuditRecord.UserPrincipalName
        UserDisplayName     = $AuditRecord.UserDisplayName
        UserId              = $AuditRecord.UserId
        AppDisplayName      = $AuditRecord.AppDisplayName
        AppId               = $AuditRecord.AppId
        ClientAppUsed       = $AuditRecord.ClientAppUsed
        ResourceId          = $AuditRecord.ResourceId
        ResourceDisplayName = $AuditRecord.ResourceDisplayName
        HomeTenantId        = $AuditRecord.HomeTenantId
        ResourceTenantId    = $AuditRecord.ResourceTenantId
        IPAddress           = $AuditRecord.IPAddress
        'CA Status'         = $AuditRecord.ConditionalAccessStatus
        RiskState           = $AuditRecord.RiskState
        RiskLevelAggregated = $AuditRecord.RiskLevelAggregated
        RiskDetail          = $AuditRecord.RiskDetail
        Location            = if ($null -ne $AuditRecord.Location) { "$($AuditRecord.Location.City), $($AuditRecord.Location.State), $($AuditRecord.Location.CountryOrRegion)" } else { "" }
        Status              = $SignInStatus  
        Id                  = $AuditRecord.Id
        IsInteractive       = $AuditRecord.IsInteractive
    }
    $AuditReport.Add($AuditReportLine)
        
}

$AuditReport = $AuditReport | Sort-Object {$_.CreatedDateTime -as [datetime]} -Descending

# Export the data to an Excel worksheet or a CSV file
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $true
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\EntraAuditSignIns.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    } 
    $AuditReport | Export-Excel -Path $ExcelOutputFile -WorksheetName "Entra Audit Sign-ins" -Title ("Entra Audit Sign-ins {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "EntraAuditSignIns" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\EntraAuditSignIns.CSV"
    $AuditReport | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}

If ($ExcelGenerated) {
    Write-Output ("{0:N0} unique sign-in records written to {1}" -f $UniqueSignInRecords.Count, $ExcelOutputFile)
} Else {
    Write-Output ("{0:N0} unique sign-in records written to {1}" -f $UniqueSignInRecords.Count, $CSVOutputFile)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 