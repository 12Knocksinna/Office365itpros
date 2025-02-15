# Get=ServiceAlertsGraph.ps1
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Get-ServiceAlertsGraph.ps1
# Define the values applicable for the application used to connect to the Graph

# Update to use the Graph SDK 15-Feb-2025

Connect-MgGraph -Scopes SecurityEvents.Read.All -NoWelcome

[array]$Alerts = Get-MgSecurityAlert -All -PageSize 999
If (!$Alerts) {
    Write-Host "No security found"
    Break
}   

$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Alert in $Alerts) {
    $ExtraInfo = $Null
    Switch ($Alert.Title) {
        "Email messages containing phish URLs removed after delivery" {
             $User = $Alert.UserStates.UserPrincipalName[1]  
        }
        "User restricted from sending email" {
            $User = $Alert.UserStates.UserPrincipalName 
        }
        "Data Governance Activity Policy" {
            $User = "N/A" 
        }
        "Admin Submission Result Completed" {
            $User = $Alert.UserStates.UserPrincipalName[0] 
            $ExtraInfo = "Email from " + $Alert.UserStates.UserPrincipalName[1] + " reported for " + $Alert.UserStates.UserPrincipalName[2]
        }   
        Default {
            $User = $Alert.UserStates.userPrincipalName 
        } 
    } # End Switch

    If ([string]::IsNullOrEmpty($Alert.Description)) { 
        $AlertDescription = "Office 365 alert" 
    } Else { 
        $AlertDescription = $Alert.Description 
    } 

    # Unpack comments
    [String]$AlertComments = $Null; $i = 0
    ForEach ($Comment in $Alert.Comments) {
        If ($i -eq 0) { 
            $AlertComments = $Comment; $i++ 
        } Else { 
            $AlertComments = $AlertComments + "; " + $Comment 
        }
    }

    Switch ($Alert.Status) {
        "newAlert"   { $Color = "ff0000"  }
        "inProgress" { $Color = "ffff00"  }
        "Default"    { $Color = "00cc00"  }
    }

    $ReportLine = [PSCustomObject][Ordered]@{
        Title       = $Alert.Title
        Category    = $Alert.Category
        User        = $User
        Description = $AlertDescription
        Date        = Get-Date($Alert.EventDateTime) -format g
        Status      = $Alert.Status
        Severity    = $Alert.Severity
        ViewAlert   = $Alert.SourceMaterials[0]
        Comments    = $AlertComments
        ExtraInfo   = $ExtraInfo
        Color       = $Color 
    }
    $Report.Add($ReportLine)   

} # End ForEach

$Report | Out-GridView -Title "Office 365 Security Alerts" -PassThru
   
# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
