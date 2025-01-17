# Update-M365GroupOwners.PS1
# Script showing how to remove a user as the owner of Microsoft 365 groups and replace them with another user

# Github link: https://github.com/12Knocksinna/Office365itpros/blob/master/Update-M365GroupOwners.PS1
# V1.0 17-Jan-2025

# Connect to the Graph - Directory.ReadWrite.All allows the script to update group owners
Connect-MgGraph -Scopes Directory.ReadWrite.All
[array]$Modules = Get-Modules | Select-Object -ExpandProperty Name
If ("ExchangeOnlineModule" -notin $Modules) {
    Write-Host "Connecting to Exchange Online"
    Connect-ExchangeOnline -showBanner:$false
}

$OldOwner = Read-Host "Enter user to remove as a group owner"
$Mbx = (Get-ExoMailbox -Identity $OldOwner -RecipientTypeDetails UserMailbox -ErrorAction SilentlyContinue)
If (!($Mbx)) { 
   Write-Host ("Can't find a mailbox for {0}" -f $OldOwner) 
   break
}

Write-Host "Checking Microsoft 365 Groups membership for" $Mbx.DisplayName
[array]$Groups = Get-MgUserOwnedObject -UserId $Mbx.ExternalDirectoryObjectId -All | `
    Where-Object {$_.additionalProperties.groupTypes -eq "unified"} 
If (!($Groups)) { 
   Write-Host ("No Microsoft 365 groups found owned by {0}" -f $Mbx.DisplayName)
   break 
}

Write-Host ("{0} Microsoft 365 groups found owned by {1}" -f $Groups.Count, $Mbx.DisplayName)
$Groups.additionalProperties.displayName

$NewOwner = Read-Host "Enter the new owner for the groups"
$Mbx2 = (Get-ExoMailbox -Identity $NewOwner -RecipientTypeDetails UserMailbox -ErrorAction SilentlyContinue)
If (!($Mbx2)) { 
   Write-Host ("Can't find a mailbox for {0}" -f $NewOwner) 
   break
}

$OKtoProceed = Read-Host "Do you want to proceed with changing the ownership of the groups (Y/N)?"
If ($OKtoProceed.toUpper() -ne "Y") { 
    break 
} Else {
    Write-Host ("Changing ownership of the {0} groups" -f $Groups.Count)
}

$Report = [System.Collections.Generic.List[Object]]::new()
$OldOwnerId = $Mbx.ExternalDirectoryObjectId
$NewOwnerId = $Mbx2.ExternalDirectoryObjectId

ForEach ($Group in $Groups) {
    Write-Host ("Processing group {0}" -f $Group.additionalProperties.displayName) -ForegroundColor Yellow
    # Add the new owner as a member
    Try {
        New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $NewOwnerId -ErrorAction Stop
    } Catch {
        Write-Host ("Failed to add {0} as a member of {1}" -f $Mbx2.DisplayName, $Group.additionalProperties.displayName)
        Continue
    }
    # And now as an owner. We do this to avoid any problems with attempting to remove the last owner
    $NewOwnerURL = ("https://graph.microsoft.com/v1.0/users/{0}" -f $NewOwnerId)
    $NewOwnerParameters = @{"@odata.id"=$NewOwnerURL}
    Try {
        New-MgGroupOwnerByRef -GroupId $Group.Id -BodyParameter $NewOwnerParameters -ErrorAction Stop
    } Catch {
        Write-Host ("Failed to add {0} as an owner of {1}" -f $Mbx2.DisplayName, $Group.additionalProperties.displayName)
        Continue
    }

    # Remove the user from the group membership. This cmdlet removes the user as both a member and owner
    Try {
    Remove-MgGroupOwnerDirectoryObjectByRef -DirectoryObjectId $OldOwnerId -GroupId $Group.Id -ErrorAction Stop
    } Catch {
        Write-Host ("Failed to remove {0} as an owner of {1}" -f $Mbx.DisplayName, $Group.additionalProperties.displayName)
        Continue
    }
    $ReportLine = [PSCustomObject][Ordered]@{
        Timestamp = Get-Date -format 'dd-MMM-yyyy HH:mm:ss'
        Action    = 'Ownership switch'
        GroupName = $Group.additionalProperties.displayName
        OldOwner  = $Mbx.DisplayName
        NewOwner  = $Mbx2.DisplayName
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.