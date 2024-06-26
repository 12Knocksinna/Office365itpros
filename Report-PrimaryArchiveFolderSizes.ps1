# Report-PrimaryArchiveFolderSizes.ps1
#
$CheckUser = Read-Host "Enter User to check" 
If ($null -eq $CheckUser.Length -eq 0) { Write-Host "Please enter a valid user to check" ; break } 
$UserGuid = (Get-ExoMailbox -Identity $CheckUser -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ExternalDirectoryObjectId) 
If (!($UserGuid)) {  
   Write-Host ("Can't find the mailbox for {0} - exiting" -f $CheckUser); break  
} 
$MailboxLocation = (Get-MailboxLocation -User $UserGuid | Select-Object MailboxLocationType, MailboxGuid | Where-Object {$_.MailboxLocationType -eq "MainArchive"}) 
If( $null -eq $MailboxLocation) { Throw( 'Mailbox not archive-enabled') } 
[array]$Folders = (Get-ExoMailboxFolderStatistics -Identity $MailboxLocation.MailboxGuid.Guid | Select FolderPath, Movable, FolderType, Name, ItemsinFolder, FolderSize)  
$NumFolders = 0; $TotalSize = 0 
$Folders | Add-Member -MemberType ScriptProperty -Name FolderSizeInBytes -Value {$this.FolderSize -replace "(.*\()|,| [a-z]*\)", ""} 
ForEach ($F in $Folders) { 
    If ($F.FolderType -eq "DeletedItems" -or $F.FolderType -eq "RecoverableItems" -or $F.Movable -eq $True) { 
       $TotalSize = ($TotalSize + $F.FolderSizeInBytes) 
       $NumFolders++  
    } 
} 
$TotalSizeGB = [math]::round($TotalSize/1GB, 3) 
$ThresholdPercent = ($TotalSizeGB/99).ToString("p") 
Write-Host $NumFolders "movable folders found. Occupied space" $TotalSize "bytes or" $TotalSizeGB "GB." "At" $ThresholdPercent "of 99 GB transition threshold" 
