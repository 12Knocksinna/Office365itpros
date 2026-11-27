# Get-ArchiveMailboxContentGraph.PS1
# Example of using the Mailbox Import-Export API to access Archive mailbox content via Microsoft Graph. This script is intended for 
# demonstration purposes and may require modifications to fit specific use cases.

# V1.0 26-Nov-2026
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Get-ArchiveMailboxContentGraph.PS1

# Needs Mail.Read application permission
# MailboxItem.Read.All to read mail items

$AppId = 'aec124c5-b9d1-4fd4-b079-005ab1183773'
$TenantId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
$CertificateThumbprint = 'your-certificate-thumbprint'

Function Get-ArchiveMailFolders {
    param (
        [string]$UserId,
        [string]$FolderId
    )

    # Return the set of folders in the archive mailbox for the specified user. 
    # If a folder ID is provided, return the child folders of that folder; otherwise, return the top-level folders.

    If ($FolderId) {
        $Uri = ("https://graph.microsoft.com/beta/users/{0}/mailFolders/{1}/childFolders" -f $UserId, $FolderId)
    } Else {
        $Uri = ("https://graph.microsoft.com/beta/users/{0}/mailFolders" -f $UserId)
    }

    Do {
        $Result = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
        ForEach ($Folder in $Result.value) {
            # Output the folder
            $Folder
            # Follow the folder tree if this folder has children
            If ($Folder.childFolderCount -gt 0) {
                Get-ArchiveMailFolders -UserId $UserId -FolderId $Folder.id
            }
        }
        # Deal with paging at the current level
        $Uri = $Result.'@odata.nextLink'
    } While ($Uri)
}


Connect-ExchangeOnline -ShowBanner:$false
Connect-MgGraph -AppId $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint 

Write-Host "Checking for mailboxes with an archive mailbox..."
[array]$Users = Get-Exomailbox -RecipientTypeDetails UserMailbox -PropertySets Archive -Filter {ArchiveStatus -ne "None"} | Sort-Object DisplayName
If ($Users.Count -eq 0) {
    Write-Host "No mailboxes with an archive mailbox found."
    return
} Else {
    Write-Host ("Found {0:N0} mailboxes with an archive mailbox." -f $Users.Count)
}

# Output report
$Report = [System.Collections.Generic.List[Object]]::new()
$PropertyFilter = "id eq 'String 0x0037' or id eq 'String 0x1035' or id eq 'String 0x1000' or id eq 'Binary 0x1013' or id eq 'Boolean 0x0E1B' or id eq 'Binary 0x3013' or id eq 'String 0x0C1A' or id eq 'String 0x0C1F'"
$EncodedPropertyFilter = [uri]::EscapeDataString($PropertyFilter)

ForEach ($User in $Users) {
    Write-Host ("Processing mailbox: {0}" -f $User.UserPrincipalName)
    $OverallFolderSize = 0
    # Get Exchange settings for the user to find the archive mailbox ID
    Try {
        $Uri = ("https://graph.microsoft.com/beta/users/{0}/settings/exchange" -f ($User.UserPrincipalName))
        $Settings = Invoke-MgGraphRequest -Uri $Uri -Method Get -OutputType PSObject -ErrorAction Stop
        $ArchiveMailboxId = $Settings.InPlaceArchiveMailboxId
        
    } Catch {
        Write-Warning ("Failed to retrieve archive settings for {0}: {1}" -f $User.UserPrincipalName, $_.Exception.Message)
        Continue
    }
    # Get set of folders in the archive mailbox
    Try {
        [array]$Folders = Get-ArchiveMailFolders -UserId $ArchiveMailboxId
        Write-Host ("Found {0:N0} folders in archive mailbox for {1}" -f $Folders.Count, $User.UserPrincipalName)
    } Catch {
        Write-Host ("Failed to retrieve archive mailbox folders for {0}: {1}" -f $User.UserPrincipalName, $_.Exception.Message)
    }
    # Try and find an Inbox folder (which should be in the archive mailbox) and retrieve items from it
    $MailboxFolder = $Folders | Where-Object {$_.displayName -eq "Inbox"}
    If (-not $MailboxFolder) {
        Write-Host ("Inbox folder not found in archive mailbox for {0}" -f $User.UserPrincipalName)
        Continue
    }
    # Construct the request URI to get items from the Inbox folder in the archive mailbox, including specific extended properties
    $MailboxItemsUri = ("https://graph.microsoft.com/v1.0/admin/exchange/mailboxes/{0}/folders/{1}/items?`$top=500&`$expand=singleValueExtendedProperties(`$filter=$EncodedPropertyFilter)" -f $ArchiveMailboxId, $MailboxFolder.Id)

    # Fetch all inbox items from the archive mailbox, handling paging with nextLink and the redirection that might occur if
    # folder items are held in an auxilary archive (shard). When this happens, Exchange responds with 308 and we can get the 
    # redirect URI from the response headers.
    Try {   
        $ItemsFromFolder = [System.Collections.Generic.List[Object]]::new() 
        [int]$RedirectCount = 0
        [int]$RedirectLimit = 5
        Do {
            $Data = Invoke-MgGraphRequest -Uri $MailboxItemsUri -Method GET -OutputType PsObject `
	            -SkipHttpErrorCheck -StatusCodeVariable StatusCode -ResponseHeadersVariable ResponseHeaders -ErrorAction Stop
            # Check for an auxiliary archive redirect
            If ($StatusCode -eq 308) {
                $RedirectCount++
                If ($RedirectCount -gt $RedirectLimit) {
                    Throw "Too many redirects while accessing archive mailbox folder."
                }
                $RedirectUri = $ResponseHeaders.Location
                $RedirectTarget = [System.Uri]$RedirectUri                
                # Don't forward the access token anywhere except Microsoft Graph
                If (($RedirectTarget.Scheme -ne 'https') -or
                    ($RedirectTarget.Host -ne 'graph.microsoft.com')) {
                    Throw "Unsafe archive mailbox redirect: $RedirectUri"
                }
                $MailboxItemsUri = $RedirectUri
                Continue
            }
               # Throw for any other unsuccessful HTTP response
            If ($StatusCode -lt 200 -or $StatusCode -ge 300) {
                Throw "Archive mailbox request failed with HTTP status $StatusCode."
            }
            # Add items returned on this page
            If ($Data.Value) {
                $ItemsFromFolder.AddRange([object[]]$Data.Value)
            }
            # Follow Graph pagination
            $MailboxItemsUri = $Data.'@odata.nextLink'

        } While ($MailboxItemsUri)
        Write-Host ("Retrieved {0:N0} items from archive mailbox Inbox for {1}" -f $ItemsFromFolder.Count, $User.UserPrincipalName)
    } Catch {
        Write-Host ("Failed to retrieve items from archive mailbox Inbox for {0}: {1}" -f $User.UserPrincipalName, $_.Exception.Message)
        Continue
    }

    # Compute the total size of items in the mailbox
    $OverallFolderSize = ($ItemsFromFolder | Measure-Object -Property size -Sum).Sum
    $OverallFolderSizeMB = [math]::Round($OverallFolderSize / 1MB, 2)
    Write-Host ("Total size of items in archive mailbox Inbox for {0}: {1:N0} MB" -f $User.UserPrincipalName, $OverallFolderSizeMB)
    # Find the largest item in the Inbox based on size and output its details
    If ($FolderItems.Count -gt 0) {
        $LargestItem = $ItemsFromFolder | Sort-Object -Property size -Descending | Select-Object -First 1
    } Else {
        Write-Host ("No items found in archive mailbox Inbox for {0}" -f $User.UserPrincipalName)
    }
    # Convert largest item size to MB for reporting
    If ($LargestItem) { 
        $LargestItemSizeMB = [math]::Round($LargestItem.size / 1MB, 2) 
        $LargestItemSubject = $LargestItem.singleValueExtendedProperties | Where-Object {$_.id -eq 'String 0x37'} | Select-Object -ExpandProperty Value
        Write-Host ("Largest item in archive mailbox Inbox for {0}: Subject: {1}, Size: {2:N0} bytes" -f $User.UserPrincipalName, $LargestItemSubject, $LargestItem.size)
    } Else {
        $LargestItemSizeMB = 0
        $LargestItemSubject = $null
    }
    # And report everything
    $ReportLine = [PSCustomObject]@{
        UserPrincipalName   = $User.UserPrincipalName
        ArchiveMailboxId    = $ArchiveMailboxId
        FolderCount         = $Folders.Count
        InboxItemCount      = $ItemsFromFolder.Count
        LargestItemSubject  = $LargestItemSubject
        LargestItemSize     = If ($LargestItem) { "{0} MB" -f $LargestItemSizeMB + " MB" } else { 0 }
    }
    $Report.Add($ReportLine)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 