# Post-NewMicrosoft365RoadmapItemstoTeamsChannel.ps1
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Post-NewMicrosoft365RoadmapItemstoTeamsChannel.ps1
# A script to grab roadmap items from the Microsoft 365 roadmap, store them in a list that's written to a CSV file after parsing
# the items to figure out what they relate to. The second part of the script looks for recent roadmap items and posts them to a 
# Teams channel using a Power Automate workflow.

# V1.0 7 Jan 2020 - Original article at https://office365itpros.com/2020/01/08/webhook-connector-roadmap-items/
# V1.1 7 Sept 2026 Updated to use Power Automate workflow to post to Teams channel instead of using the webhook directly. This is because the webhook connector is being deprecated by Microsoft and will be removed in 2027. The Power Automate flow is a simple flow that takes the JSON payload and posts it to the Teams channel.

# URI pointing to the webhook connector for the target Teams channel - this will be different in your tenant!
$Uri = "https://defaultb662313f14fc43a29a7ad2e27f4f34.00.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/10/workflows/29bf5cf4a6b04069bb855b7a188eac7e/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=p8_vC2OTmeYOmAg0DYBE6mODjG__ktb8RaecsCELbCU"

$RoadmapItems = 'https://www.microsoft.com/en-us/microsoft-365/RoadmapFeatureRSS'

[int]$DaysToCheck = 7 # Number of days to check for recent roadmap items that are posted to Teams

#Fetch current set of roadmap items and parse the information
[array]$Updates = (Invoke-RestMethod -Uri $RoadmapItems  -Method Get)

If ($Updates.count -eq 0) {
    Write-Host "No updates found in the Microsoft 365 Roadmap RSS feed. Exiting script."
    Return
}

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Item in $Updates) {               
    # Figure out the categories

    $i = $Item.Category.Count
    $Categories = $Item.Category[0..$i] -join ", "

    #Set the color line of the card according to the Status of the environment
    If ($Item.Category.Contains("In development")) { 
        $Color = "ff0000"  
        } Elseif ($Item.Category.Contains("Rolling out")) { 
            $Color = "ffff00"  
        } Else { 
            $Color = "00cc00" 
    }
    # Now process the categories to identify favorite products
    $Outlook = $False; $OneDrive = $False; $Clipchamp = $False; $Exchange = $False; $SharePoint = $False; $Windows = $False; `
        $InTune = $False; $VivaEngage = $False; $EntraID = $False; $Forms = $False; $iOS = $False; $Android = $False; $O365 = $False; `
        $Project = $False; $Planner = $False; $Teams = $False; $GCC = $False; $Education = $False; $Mac = $False; $Excel = $False; `
        $Developer = $False; $AllEnv = $False; $StandardMT = $False; $MCAS = $False; $Dod = $False; $MIP = $False; $Visio = $False; `
        $Technology = $Null; $Purview = $false; $Viva = $false; $Color = $Null; $Copilot = $false; $Availability = $Null
        
    If ($Item.Category.Contains("Outlook")) { 
        $Outlook = $True
        $O365 = $true
        $Technology = "Outlook"
    }
    If ($Item.Category.Contains("Microsoft Copilot")) { 
        $Copilot = $True
        $O365 = $True
        $Technology = "Microsoft Copilot for Microsoft 365" 
    }
    If ($Item.Category.Contains("Exchange")) { 
        $Exchange = $True
        $O365 = $True
        $Technology = "Exchange Online" 
    }
    If ($Item.Category.Contains("SharePoint")) { 
        $SharePoint = $True
        $O365 = $True
        $Technology = "SharePoint Online"
    }
    If ($Item.Category.Contains("OneDrive")) { 
        $OneDrive = $True
        $O365 = $True
        $Technology = "OneDrive for Business"
    }
    If ($Item.Category.Contains("Microsoft Viva")) { 
        $Viva = $True
        $O365 = $True
        $Technology = "Microsoft Viva" 
    }
    If ($Item.Category.Contains("Microsoft Clipchamp")) { 
        $Clipchamp = $True
        $O365 = $True
        $Technology = "Microsoft Clipchamp" 
    }
    If ($Item.Category.Contains("Microsoft Purview")) { 
        $Purview = $True
        $O365 = $True
        $Technology = "Microsoft Purview" 
    }
    If ($Item.Category.Contains("Windows Desktop") -or $Item.Category.Contains("Windows")) { 
        $Windows = $True
        $Technology = "Windows" 
    }
    If ($Item.Category.Contains("Microsoft Intune")) { 
        $Intune = $True
        $Technology = "Intune" 
    }
    If ($Item.Category.Contains("Viva Engage")) { 
        $VivaEngage = $True
        $O365 = $True
        $Technology = "Viva Engage" 
    }
    If ($Item.Category.Contains("Entra ID")) { 
        $EntraID = $True
        $Technology = "Entra ID" 
    }
    If ($Item.Category.Contains("Microsoft Forms")) { 
        $Forms = $True
        $O365 = $True
        $Technology = "Forms" 
    }
    If ($Item.Category.Contains("iOS")) { 
        $iOS = $True
        $Technology = "Clients" 
    }
    If ($Item.Category.Contains("Android")) { 
        $Android = $True
        $Technology = "Clients" 
    }
    If ($Item.Category.Contains("Mac")) { 
        $Mac = $True
        $Technology = "Clients" 
    }
    If ($Item.Category.Contains("Visio")) { 
        $Visio = $True
        $Technology = "Desktop App" 
    }
    If ($Item.Category.Contains("Excel")) { 
        $Excel = $True
        $Technology = "Desktop App" 
    }
    If ($Item.Category.Contains("Microsoft Information Protection") -or $Item.Category.Contains("Azure Information Protection")) { 
        $MIP = $True
        $Technology = "Information Protection" 
    }
    If ($Item.Category.Contains("Project")) { 
        $Project = $True
        $Technology = "Project" 
    }
    If ($Item.Category.Contains("Planner")) { 
        $Planner = $True
        $O365 = $True
        $Technology = "Planner" 
    }
    If ($Item.Category.Contains("Microsoft Teams")) { 
        $Teams = $True
        $O365 = $True
        $Technology = "Teams" 
    }
    If ($Item.Category.Contains("O365") -or $Item.Category.Contains("Office 365")) { 
            $O365 = $True 
    }
    If ($Item.Category.Contains("Microsoft Cloud App Security")) { 
        $MCAS = $True
        $Technology = "Cloud App Security" 
    }
    If ($Item.Category.Contains("GCC")) { 
        $GCC = $True
        $Technology = "GCC" 
    }
    If ($Item.Category.Contains("Dod")) { 
        $Dod = $True
        $Technology = "DoD"
    }
    If ($Item.Category.Contains("Education")) { 
        $Education = $True
        $Technology = "Education" 
    }
    If ($Item.Category.Contains("Developer")) { 
        $Developer = $True 
    }
    If ($Item.Category.Contains("All Environments")) { 
        $AllEnv = $True 
    }
    If ($Item.Category.Contains("Standard Multi-Tenant")) { 
        $StandardMT = $True 
    }

    $ItemAge = ($Item.PubDate | New-TimeSpan).Days
    If ($ItemAge -lt $DaysToCheck -and $O365 -eq $True ) {
        #Extract FeatureId from Link in the update
        $FeatureId = $Item.Link.Split("=")[1]
     # If there's an availability date in the description, extract it
        If ($Item.Description.Contains("date:")) {
            $i = $Item.Description.LastIndexOf(":")
            $Availability = $Item.Description.SubString($i+2) 
        } Else { 
            $Availability = "Not defined" 
        }

     # Generate report line and update the list
        $ReportLine = [PSCustomObject]@{      
            FeatureId     = $FeatureId
            Title         = $Item.Title
            Technology    = $Technology
            Availability  = $Availability
            Status        = $Item.Category[0]
            Date          = Get-Date($Item.PubDate) -format 'dd-MMM-yyyy HH:mm'
            LastUpdated   = Get-Date($Item.Updated) -format 'dd-MMM-yyyy HH:mm'
            Categories    = $Categories
            Description   = $Item.Description 
            O365          = $O365
            Copilot       = $Copilot
            EntraID       = $EntraID
            Excel         = $Excel
            Exchange      = $Exchange
            Forms         = $Forms
            Intune        = $Intune
            MCAS          = $MCAS
            Clipchamp     = $Clipchamp
            MIP           = $MIP
            Planner       = $Planner
            Project       = $Project
            OneDrive      = $OneDrive
            Outlook       = $Outlook
            SharePoint    = $SharePoint
            'Microsoft Viva' = $Viva
            'Microsoft Purview' = $Purview
            Teams         = $Teams
            Visio         = $Visio
            'Viva Engage' = $VivaEngage
            Android       = $Android
            IOS           = $iOS
            Mac           = $Mac
            Windows       = $Windows
            AllEnv        = $AllEnv
            Developer     = $Developer
            Education     = $Education
            DoD           = $DoD
            GCC           = $GCC
            StandardMT    = $StandardMT     
            Link          = $Item.Link
            Color         = $Color   }
        $Report.Add($ReportLine)
    }
}

# Now we have processed the list, we can export it to CSV and then loop through it to generate cards for recent items
Write-Host $Report.Count "Microsoft 365 Roadmap Items stored in c:\temp\RoadmapItems.csv"
$Report | Sort-Object FeatureId | Export-CSV -NoTypeInformation c:\temp\RoadmapItems.csv

If ($Report.Count -eq 0) {
    Write-Host "No new Microsoft 365 Roadmap items found in the last" $DaysToCheck "days. Exiting script."
    Return
} Else {
    Write-Host $Report.Count "Microsoft 365 Roadmap items found in the last" $DaysToCheck "days. Posting to Teams channel."
}

ForEach ($Item in $Report) { 
    
    # Generate payload(s)          
    $PayloadObject = @{
    '@context' = 'https://schema.org/extensions'
    '@type'    = 'MessageCard'
    potentialAction = @(
        @{
            '@type' = 'OpenUri'
            name    = 'More info'
            targets = @(
                @{
                    os  = 'default'
                    uri = $Item.Link
                }
            )
        }
    )
    sections = @(
        @{
            facts = @(
                @{
                    name  = 'Status:'
                    value = $Item.Status
                }
                @{
                    name  = 'Category:'
                    value = $Item.Categories
                }
                @{
                    name  = 'Date:'
                    value = $Item.Date
                }
            )
            text = $Item.Description
        }
    )
    summary    = $Item.Title
    themeColor = $Item.Color
    title      = "Feature ID: $($Item.FeatureId) - $($Item.Title)"
}

    $Payload = $PayloadObject | ConvertTo-Json -Depth 10 -Compress

    # If we have an update, post details to Teams
    Write-Host "Posting details of feature" $Item.FeatureID "to Teams."
    Invoke-RestMethod -uri $URI -Method Post -body $Payload -ContentType 'application/json; charset=utf-8'
}

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com. See our post about the Microsoft 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.