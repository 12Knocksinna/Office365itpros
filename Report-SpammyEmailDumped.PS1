# Report-SpammyEmailDumped.PS1
# A script to report the volume of spammy email intercepted by a mail flow rule and dumped before
# it reaches user mailboxes. The script is designed to run on a weekly basis as an Azure Automation runbook
# and record its data in a CSV file stored in a SharePoint Online document library. The data in the CSV
# file can then be used to generate statistics over a long-term basis.

# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-SpammyEmailDumped.PS1
# V1.0  13-Oct-2026

# Interactive login to Exchange Online
Connect-ExchangeOnline -ShowBanner:$false
# The Azure Automation sign in using a managed identity uses the service domain for your tenant, which you will
# Have to change here.
# Connect-ExchangeOnline -ManagedIdentity -Organization office365itpros.onmicrosoft.com

# Connect to the Microsoft Graph to access SharePoint Online
Try {
    Connect-MgGraph -Identity -ErrorAction Stop -NoWelcome
} Catch {
    Write-Error "Failed to connect to Microsoft Graph. Please ensure you have the necessary permissions."
    Exit
}

 $RunDate = (Get-Date).toUniversalTime().ToString("yyyy-MM-dd HH:mm") 
# We fetch every seven days using a scheduled job
[datetime]$StartDate = (Get-Date).AddDays(-7)
[datetime]$EndDate = Get-Date

# Update the variable with the site identifier of the target SharePoint Online site where the CSV file is stored
$SiteId = "office365itpros.sharepoint.com,acfe74d8-edfb-436d-924b-e018666605ee,a2aba197-5f1d-4864-a2c7-4daf0ff6379b"

# For Temporary files
$LocalDownloadPath = "C:\Temp"

# Make sure that the local folder exists
If (-Not (Test-Path -Path $LocalDownloadPath -PathType Container)) {
    Try {
        New-Item -Path $LocalDownloadPath -ItemType Directory -ErrorAction Stop | Out-Null
    } Catch {
        Write-Error "Failed to create $LocalDownloadPath directory: $_"
        Return
    }
}

# Define the path for the exported spam data CSV file
$ExportedSpamDataFile = $LocalDownloadPath + "\SpammyEmailData.csv"

# Fetch data from the mail transport service
[int]$Page = 1
$Data = [System.Collections.Generic.List[Object]]::new()

Do {
   $PageResults = Get-MailDetailTransportRuleReport -Direction inbound -Action DeleteMessage -Page $Page -StartDate $StartDate -EndDate $EndDate -ErrorAction SilentlyContinue
   If ($PageResults) {
        $Data += $PageResults
    }
    $Page++
}
While ($PageResults.Count -gt 0)

If ($Data) {
    Write-Output ""
    Write-Output "Total messages deleted in the last seven days: $($Data.Count)"
}
Else {
    Write-Output ""
    Write-Output "No messages deleted in the last seven days."
    Return
}

Write-Output ""
Write-Output "Mail flow rules that have deleted messages in the last seven days:"
$Data | Group-Object 'TransportRule' -NoElement | Format-Table Name, Count

$TransportRule = 'Quarantine Traffic from Junk Email Domains'
[array]$MData = $Data | Where-Object {$_.'TransportRule' -eq  $TransportRule}

Write-Output ""
Write-Output "The blocked spam was sent to the following recipients:"
$MData | Group-Object -NoElement RecipientAddress | Sort-Object Count -Descending | Format-Table Name, Count

$SpamData = [System.Collections.Generic.List[Object]]::new()
[array]$SpammyDomains = (Get-TransportRule -Identity $TransportRule).SenderDomainIs
ForEach ($Domain in $SpammyDomains) {
  $CheckDomain = "*"+$Domain+"*"
  [array]$SpamRecords = $MData | Where-Object {$_.SenderAddress -like $CheckDomain} 
  If ($SpamRecords.Count -gt 0) {
    Write-Output ("Domain {0} sent {1} spammy messages" -f $Domain, $SpamRecords.Count)
    $SpamDataRecord = [PSCustomObject]@{
        RunDate = $RunDate
        Domain = $Domain
        Count = $SpamRecords.Count
    }
    $SpamData.Add($SpamDataRecord)
  }
}

# Create CSV file with values for this run
Remove-Item -Path $ExportedSpamDataFile -ErrorAction SilentlyContinue
$SpamData | Export-CSV $ExportedSpamDataFile -Encoding utf8

Write-Output "Updating historical spammy email data in SharePoint Online..."
# Define path to the folder where the item is stored and the name of the CSV file that contains the historical data. The folder must already exist in the document library.
$FolderPath = "The Base of All Good in Office 365"
$SpammyEmailHistoricalFile = "HistoricalSpammyEmail.csv"

# Get site and default document library
Try {
    $Site = Get-MgSite -SiteId $SiteId -ErrorAction Stop
} Catch {
    Write-Error "Failed to retrieve site information: $_"
    Return
}
# Get the default document library for the site
Try {
    $Drive = Get-MgSiteDefaultDrive -SiteId $Site.Id -ErrorAction Stop
} Catch {
    Write-Error "Failed to retrieve default document library: $_"
    Return
}

$UpdateItemRequired = $false
# Get properties of the historical spam tracking file from SharePoint. We need the item identifiers to download and upload
# files later
$Uri = ("https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}/{2}" -f $Drive.Id, $FolderPath, $SpammyEmailHistoricalFile)
Try {
    $SpamFileItem = Invoke-MgGraphRequest -Uri $Uri -Method GET -ErrorAction Stop
    $UpdateItemRequired = $true
} Catch {
    Write-Error "Failed to retrieve Spammy Email Historical file: $_"
    Write-Output "Creating new file..."
    $TargetLocation = "root:/" + $FolderPath + "/" + $SpammyEmailHistoricalFile + ":"
    Try {
        $SpamFileItem = Set-MgDriveItemContent -DriveId $Drive.Id -DriveItemId $TargetLocation -InFile $ExportedSpamDataFile -ErrorAction Stop
    } Catch {
        Write-Error "Failed to create new Spammy Email Historical file: $_"
        Return
    }
}

# If we need to update the historical file, download it first, append the new data, and then upload it back to SharePoint
If ($UpdateItemRequired) {  
    # Download the existing historical file
    $DownloadPath = Join-Path -Path $LocalDownloadPath -ChildPath $SpammyEmailHistoricalFile
    Try {
        Remove-Item -Path $DownloadPath -ErrorAction SilentlyContinue
        Get-MgDriveItemContent -DriveId $Drive.Id -DriveItemId $SpamFileItem.id -OutFile $DownloadPath -ErrorAction Stop
    } Catch {
        Write-Error "Failed to download existing Spammy Email Historical file: $_"
        Return
    }
    # Import the existing historical data
        $ExistingData = Import-Csv -Path $DownloadPath
        # Append the new data to the existing data
        $CombinedData = $ExistingData + $SpamData
        # Export the combined data back to CSV
        $CombinedData | Export-Csv -Path $ExportedSpamDataFile -NoTypeInformation -Encoding utf8
        # Upload the updated file back to SharePoint
        Try {
            $SpamFileItem = Set-MgDriveItemContent -DriveId $Drive.Id -DriveItemId $SpamFileItem.id -InFile $ExportedSpamDataFile -ErrorAction Stop
        } Catch {
            Write-Error "Failed to upload updated Spammy Email Historical file: $_"
            Return
        }
}

$Runs = $CombinedData | Group-Object -Property RunDate -NoElement | Sort-Object RunDate -Descending
Write-Output "Spammy email historical data updated successfully in SharePoint Online."
Write-Output "The historical data now contains $($CombinedData.Count) records from $($Runs.Count) runs."
