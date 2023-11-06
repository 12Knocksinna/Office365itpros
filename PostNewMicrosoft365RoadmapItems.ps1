# Adapted from the original https://github.com/einast/PS_M365_scripts/blob/master/M365RoadmapUpdates.ps1
# See https://thingsinthe.cloud/Teams-message-cards-Getting-Office-365-roadmap-into-channel/ for details
£ https://github.com/12Knocksinna/Office365itpros/blob/master/PostNewMicrosoft365RoadmapItems.ps1
# A script to grab roadmap items from the Microsoft 365 roadmap, store them in a list that's written to a CSV file after parsing
# the items to figure out what they relate to. The second part of the script looks for recent roadmap items and posts them to a 
# Teams channel using the Office 365 incoming webhook connector.

# V1.0 7 Jan 2020

# URI pointing to the Office 365 webhook connector for the target Teams channel - this will be different in your tenant!
$URI = "https://outlook.office.com/webhook/7aa49aa6-7840-443d-806c-08ebe8f59966@b662313f-14fc-43a2-9a7a-d2e27f4f3478/IncomingWebhook/8592f62b50cf41b9b93ba0c0a00a0b88/eff4cd58-1bb8-4899-94de-795f656b4a18"

$RoadmapItems = 'https://www.microsoft.com/en-us/microsoft-365/RoadmapFeatureRSS'
$Report = [System.Collections.Generic.List[Object]]::new()
$DaysToCheck = 7 # Number of days to check for recent roadmap items that are posted to Teams

#Fetch current set of roadmap items and parse the information
[array]$Updates = (Invoke-RestMethod -Uri $RoadmapItems  -Method Get)
ForEach ($Item in $Updates) {               
      # Figure out the categories
      $i = $Item.Category.Count
      $Categories = $Item.Category[0..$i] -join ", "

      #Set the color line of the card according to the Status of the environment
      if ($Item.Category.Contains("In development")) { $Color = "ff0000"  }
        elseif ($Item.Category.Contains("Rolling out")) { $color = "ffff00"  }
        else { $color = "00cc00"  }
     # Now process the categories to identify favorite products
     $Outlook = $False; $OneDrive = $False; $Stream = $False; $Exchange = $False; $SharePoint = $False; $Windows = $False; $InTune = $False; $Yammer = $False; $AAD = $False; $Forms = $False; $iOS = $False; $Android = $False; $O365 = $False; $Project = $False; $Planner = $False; $Teams = $False; $GCC = $False; $Education = $False; $Mac = $False; $Excel = $False; $Developer = $False; $AllEnv = $False; $StandardMT = $False; $MCAS = $False; $Dod = $False; $MIP = $False; $Visio = $False; $Technology = $Null
     If ($Item.Category.Contains("Outlook")) { $Outlook = $True; $Technology = "Outlook"}
     If ($Item.Category.Contains("Exchange")) { $Exchange = $True; $O365 = $True; $Technology = "Exchange" }
     If ($Item.Category.Contains("SharePoint")) { $SharePoint = $True; $O365 = $True; $Technology = "SharePoint"}
     If ($Item.Category.Contains("OneDrive for Business")) { $OneDrive = $True; $O365 = $True; $Technology = "OneDrive"}
     If ($Item.Category.Contains("Microsoft Stream")) { $Stream = $True; $O365 = $True; $Technology = "Stream" }
     If ($Item.Category.Contains("Windows Desktop") -or $Item.Category.Contains("Windows")) { $Windows = $True; $Technology = "Windows" }
     If ($Item.Category.Contains("Microsoft Intune")) { $Intune = $True; $Technology = "Intune" }
     If ($Item.Category.Contains("Yammer")) { $Yammer = $True; $O365 = $True; $Technology = "Yammer" }
     If ($Item.Category.Contains("Azure Active Directory")) { $AAD = $True; $Technology = "Azure Active Directory" }
     If ($Item.Category.Contains("Microsoft Forms")) { $Forms = $True; $O365 = $True; $Technology = "Forms" }
     If ($Item.Category.Contains("iOS")) { $iOS = $True; $Technology = "Clients" }
     If ($Item.Category.Contains("Android")) { $Android = $True ; $Technology = "Clients" }
     If ($Item.Category.Contains("Mac")) { $Mac = $True; $Technology = "Clients" }
     If ($Item.Category.Contains("Visio")) { $Visio = $True; $Technology = "Desktop App" }
     If ($Item.Category.Contains("Excel")) { $Excel = $True; $Technology = "Desktop App" }
     If ($Item.Category.Contains("Microsoft Information Protection") -or $Item.Category.Contains("Azure Information Protection")) { $MIP = $True; $Technology = "Information Protection" }
     If ($Item.Category.Contains("Project")) { $Project = $True; $Technology = "Project" }
     If ($Item.Category.Contains("Planner")) { $Planner = $True; $O365 = $True; $Technology = "Planner" }
     If ($Item.Category.Contains("Microsoft Teams")) { $Teams = $True; $O365 = $True; $Technology = "Teams"}
     If ($Item.Category.Contains("O365") -or $Item.Category.Contains("Office 365")) { $O365 = $True }
     If ($Item.Category.Contains("Microsoft Cloud App Security")) { $MCAS = $True; $Technology = "Cloud App Security" }
     If ($Item.Category.Contains("GCC")) { $GCC = $True; $Technology = "GCC" }
     If ($Item.Category.Contains("Dod")) { $Dod = $True; $Technology = "DoD"}
     If ($Item.Category.Contains("Education")) { $Education = $True; $Technology = "Education" }
     If ($Item.Category.Contains("Developer")) { $Developer = $True }
     If ($Item.Category.Contains("All Environments")) { $AllEnv = $True 
     If ($Item.Category.Contains("Standard Multi-Tenant")) { $StandardMT = $True }}

     #Extract FeatureId from Link in the update
     $FeatureId = $Item.Link.Split("=")[1]
     # If there's an availability date in the description, extract it
     If ($Item.Description.Contains("date:")) {
          $i = $Item.Description.LastIndexOf(":")
          $Availability = $Item.Description.SubString($i+2) }
     Else { $Availability = "Not defined" }

     # Generate report line and update the list
     $ReportLine = [PSCustomObject]@{      
         FeatureId     = $FeatureId
         Title         = $Item.Title
         Technology    = $Technology
         Availability  = $Availability
         Status        = $Item.Category[0]
         Date          = Get-Date($Item.PubDate) -format g
         LastUpdated   = Get-Date($Item.Updated) -format g
         Categories    = $Categories
         Description   = $Item.Description 
         O365          = $O365
         AAD           = $AAD
         Excel         = $Excel
         Exchange      = $Exchange
         Forms         = $Forms
         Intune        = $Intune
         MCAS          = $MCAS
         MIP           = $MIP
         Planner       = $Planner
         Project       = $Project
         OneDrive      = $OneDrive
         Outlook       = $Outlook
         SharePoint    = $SharePoint
         Stream        = $Stream
         Teams         = $Teams
         Visio         = $Visio
         Yammer        = $Yammer
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

# Now we have processed the list, we can export it to CSV and then loop through it to generate cards for recent items
Write-Host $Report.Count "Microsoft 365 Roadmap Items stored in c:\temp\RoadmapItems.csv"
$Report | Sort-Object FeatureId | Export-CSV -NoTypeInformation c:\temp\RoadmapItems.csv

Write-Host "Checking the Roadmap for Office 365 updates in the last" $DaysToCheck "days..."
ForEach ($Item in $Report) {
     $ItemAge = ($Item.Date | New-TimeSpan).Days
     If ($ItemAge -lt $DaysToCheck -and $Item.O365 -eq $True ) {
       # Convert MessageText to JSON beforehand, if not the payload will fail.
       $MessageText = ConvertTo-Json $Item.Description
       # Generate payload(s)          
       $Payload = @"
{
    "@context": "https://schema.org/extensions",
    "@type": "MessageCard",
    "potentialAction": [
            {
            "@type": "OpenUri",
            "name": "More info",
            "targets": [
                {
                    "os": "default",
                    "uri": "$($Item.Link)"
                }
            ]
        },
     ],
    "sections": [
        {
            "facts": [
                {
                    "name": "Status:",
                    "value": "$($Item.Status)"
                },
                {
                    "name": "Category:",
                    "value": "$($Item.Categories)"
                },
            {
                    "name": "Date:",
                    "value": "$($Item.Date)"
                }
            ],
            "text": $($MessageText)
        }
    ],
    "summary": "$($Item.Title)",
    "themeColor": "$($Item.Color)",
    "title": "Feature ID: $($Item.FeatureId) - $($Item.Title)"
}
"@
    # If we have an update, post details to Teams
    Write-Host "Posting details of feature" $Item.FeatureID "to Teams."
    $Status = (Invoke-RestMethod -uri $URI -Method Post -body $Payload -ContentType 'application/json; charset=utf-8')
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.