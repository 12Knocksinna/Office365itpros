# Restore-DeletedEntraGroups.PS1
# A script to show how to restore deleted Entra ID groups (Microsoft 365 and security groups) using the Microsoft Graph PowerShell SDK
# The script connects to Microsoft Graph, retrieves the list of deleted groups, and displays them in a grid view for easy selection and restoration.

# V1.0 7-Nov-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Restore-DeletedEntraGroups.PS1

If ([Environment]::UserInteractive) { 
    Clear-Host
    Connect-MgGraph -Scopes Directory.Read.All, Group.ReadWrite.All -NoWelcome
    Write-Host "Script running interactively... connecting to the Graph" -ForegroundColor Yellow
    $Interactive = $true
}   Else { 
    Write-Host "Script can't be run non-interactively. Exiting..." -ForegroundColor Red
    Break
}   

# Check that we have the correct scopes
If ($Interactive) {
    [string[]]$CurrentScopes = (Get-MgContext).Scopes
    [string[]]$RequiredScopes = @('Directory.Read.All', 'Group.ReadWrite.All')

    $CheckScopes =[object[]][Linq.Enumerable]::Intersect($RequiredScopes,$CurrentScopes)
    If ($CheckScopes.Count -ne 2) { 
        Write-Host ("To run this script, you need to connect to Microsoft Graph with the following scopes: {0}" -f $RequiredScopes) -ForegroundColor Red
        Break
    }
}

[array]$DeletedGroups = Get-MgDirectoryDeletedItemAsGroup -All
If ($DeletedGroups.count -eq 0) { 
    Write-Host "No recoverable groups can be found - exiting"; break
}

$Report = [System.Collections.Generic.List[Object]]::new(); $Now = Get-Date
ForEach ($Group in $DeletedGroups) {    
    [datetime]$DeletedDate = $Group.deletedDateTime
    $PermanentRemovalDue = Get-Date($DeletedDate).AddDays(30)
    $TimeTillRemoval = $PermanentRemovalDue - $Now
    If ("unified" -notin $Group.GroupTypes) {   
        $GroupType = "Security"
    }   Else {   
        $GroupType = "Microsoft 365"
    }

    $ReportLine = [PSCustomObject]@{ 
      Group                      = $Group.displayName
      Id                         = $Group.id
      GroupType                  = $GroupType
      Created                    = Get-Date ($Group.createdDatetime) 
      Deleted                    = Get-Date($Group.deletedDateTime) 
      'Permanent Deletion due'   = Get-Date($PermanentRemovalDue) 
      DaysRemaining              = $TimeTillRemoval.Days 
    } 
    $Report.Add($ReportLine) 
}

$Report | Sort-Object DaysRemaining -Descending | Out-GridView -Title "Select groups to restore" -PassThru | ForEach-Object { 
    Write-Host "Restoring group $($_.Group) Id $($_.Id)" 
    Try {
        $Status = Restore-MgDirectoryDeletedItem -DirectoryObjectId $_.Id 
    } Catch {
        Write-Host "Failed to restore group $($_.Group) Id $($_.Id). Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    If ($Status) {
        Write-Host "Successfully restored group $($_.Group) Id $($_.Id)" -ForegroundColor Green
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.