# Report-RecoverableItems.PS1

function FormatFileSize {
    # Format File Size nicely
    param (
            [parameter(Mandatory = $true)]
            $InFileSize
        ) 
    
     If ($InFileSize -lt 1KB) { # Format the size of a document
            $FileSize = $InFileSize.ToString() + " B" } 
          ElseIf ($InFileSize -lt 1MB) {
            $FileSize = $InFileSize / 1KB
            $FileSize = ("{0:n2}" -f $FileSize) + " KB"} 
          Elseif ($InFileSize -lt 1GB) {
            $FileSize = $InFileSize / 1MB
            $FileSize = ("{0:n2}" -f $FileSize) + " MB" }
         Elseif ($InFileSize -ge 1GB) {
            $FileSize = $InFileSize / 1GB
            $FileSize = ("{0:n2}" -f $FileSize) + " GB" }
      Return $FileSize    
} 

# Connect to Exchange Online, if we're not already connected
$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}
# Connect to the Graph with an Entra ID app that has the Mail.Read permission
Connect-MgGraph -AppId $AppId -Scopes $Scopes -TenantId $TenantId -CertificateThumbprint $Thumbprint

Write-Host "Scanning for mailboxes..."
[array]$NoReportFolders = @("Audits", "Calendar Logging", "SubstrateHolds")
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited
If (!($Mbx)) {
    Write-Host "No mailboxes found - exiting!"; break
}  

Write-Host ("Processing {0} mailboxes..." -f $Mbx.Count)
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {
    Write-Host ("Processing mailbox {0}" -f $M.UserPrincipalName)
   # Graph request 
   # $Uri = ("https://graph.microsoft.com/v1.0/users/{0}/MailFolders/RecoverableItemsRoot/childfolders" -f $M.ExternalDirectoryObjectId)

    Try {
       $RecoverableItemsRoot = (Get-MgUserMailFolder -Userid $M.ExternalDirectoryObjectId -MailFolderId 'RecoverableItemsRoot').id
       [array]$Folders = Get-MgUserMailFolderChildFolder -MailFolderId $RecoverableItemsRoot -UserId $M.ExternalDirectoryObjectId
    }
    Catch {
        Write-Host ("Error {0} when processing mailbox {1}" -f $_.Exception.Message, $M.UserPrincipalName)
        Continue
    }
  
    Write-Host ("Found {0} folders in  Recoverable Items" -f $Folders.Count)
    
    ForEach ($Folder in $Folders){
        # Ignore the folders used for mailbox auditing and calendar logging
        If ($Folder.displayName -in $NoReportFolders) {
        # Write-Host "Ignored folder" $Folder.displayName
            Continue
        }
        If ($Folder.TotalItemCount -gt 0) {
            Write-Host ("Processing folder {0} in mailbox" -f $Folder.displayName, $M.displayName)
        # Graph API request 
        # $Uri = ("https://graph.microsoft.com/v1.0/users/{0}/mailfolders/{1}/Messages/?`$select=sender,createdDateTime,subject&`$expand=singleValueExtendedProperties(`$filter=Id%20eq%20'LONG%200x0E08')" -f $M.ExternalDirectoryObjectId, $Folder.id)
       
            [array]$Items = Get-MgUserMailFolderMessage -UserId $M.ExternalDirectoryObjectId -MailFolderId $Folder.id -All `
                -Property sender,createdDateTime,subject`
                -ExpandProperty "singleValueExtendedProperties(`$filter=Id eq 'LONG 0x0E08')"
            # If some items are returned, report them
            Write-Host ("Found {0} items in folder {1}" -f $Items.Count, $Folder.displayName)
            ForEach ($Item in $Items) {
                [long]$ItemFileSize = $Item.singleValueExtendedProperties.value
                $ReportLine = [PSCustomObject][Ordered]@{
                    Mailbox     = $M.UserPrincipalName
                    Folder      = $Folder.displayName
                    Subject     = $Item.Subject
                    Sender      = $Item.sender.emailAddress.address
                    Created     = Get-Date($Item.createdDateTime).ToLocalTime()
                    Size        = FormatFileSize -InFileSize $ItemFileSize
                }
                $Report.Add($ReportLine)   
            }
        }
    }
}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.