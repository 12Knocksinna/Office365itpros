#Report-TeamsEmojis.PS1
# A quick and simple script to report the usage of emojis as reactions to Teams chat and channel conversation messages

# V1.0 6-Nov-2025
# GitHub Link:

# Connect to Microsoft Graph to fetch user information
Connect-MgGraph -Scopes "User.ReadBasic.All"

$UserInfoAvailable = $false

# Fetching user information
[array]$Users = Get-MgUser -All -PageSize 999 -Property DisplayName, UserPrincipalName

If ($Users) {
    Write-Output ("Fetched {0} users from the tenant." -f $Users.Count)
    $UserHash = @{}
    ForEach ($User in $Users) {
        $UserHash.Add($User.UserPrincipalName, $User.DisplayName)
    }
    $UserInfoAvailable = $true

}   Else {
    Write-Output "No users found in the tenant."
    Break
}

# Check if Exchange Online is already connected - otherwise connect
If ($Null -eq (Get-ConnectionInformation)) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
}

# Build hash table to resolve reaction name to the emoji shown in Teams
$Emojis = @{}
$Emojis.Add("head shaking vertically", "🤦‍♂️")
$Emojis.Add("head shaking horizontally", "🤦‍♀️")
$Emojis.Add("thumbs up", "👍")
$Emojis.Add("thumbs down", "👎")
$Emojis.Add("like", "👍")
$Emojis.Add("heart", "❤️")
$Emojis.Add("laugh", "😂")
$Emojis.Add("peekingeye", "🙈")
$Emojis.Add("surprised", "😮")
$Emojis.Add("sad", "😢")
$Emojis.Add("saddog", "🐶")
$Emojis.Add("angry", "😠")
$Emojis.Add("yes-tone1", "👍🏻")
$Emojis.Add("blankface", "😐")
$Emojis.Add("bowing", "🙇")
$Emojis.Add("call", "📞")
$Emojis.Add("cool", "😎")
$Emojis.Add("cry", "😢")
$Emojis.Add("coolkoala", "🦘")
$Emojis.Add("eyeinspeechbubble", "👁️‍🗨️")
$Emojis.Add("fire", "🔥")
$Emojis.Add("fistbump", "👊")
$Emojis.Add("follow", "👣")
$Emojis.Add("grinningfacewithsmilingeyes", "😁")
$Emojis.Add("handsinair", "🙌")
$Emojis.Add("heartblue", "💙")
$Emojis.Add("hearteyes", "😍")
$Emojis.Add("happyface", "😊")
$Emojis.Add("loudlycrying", "😭")
$Emojis.Add("personmotorwheelchairright2tone1", "🧑🏻‍🦼")
$Emojis.Add("rainbowsmileyface", "🌈")
$Emojis.Add("clap", "👏")
$Emojis.Add("smilingfacewithtear", "🥲")
$Emojis.Add("squintingfacewithtongue", "😜")
$Emojis.Add("2714_heavycheckmark", "✔️")
$Emojis.Add("mickeymouse;0-neu-d1-63cc7d1c9164d9a1b23ae85714dd20fa", "🐭")
$Emojis.Add("think", "🤔")
$Emojis.Add("1f51c_soon", "🔜")
$Emojis.Add("1f6a2_ship", "🚢")
$Emojis.Add("1f4ce_paperclip", "📎")
$Emojis.Add("relieved", "😌")
$Emojis.Add("rofl", "🤣")
$Emojis.Add("peekingeyes", "🙈")
$Emojis.Add("1f3a8_artistpalette", "🎨")
$Emojis.Add("1f458_kimono", "👘")
$Emojis.Add("1f4a4_zzz", "💤")
$Emojis.Add("1f4af_hundredpointssymbol", "💯")
$Emojis.Add("1f603_grinningfacewithbigeyes", "😃")
$Emojis.Add("1f920_facewithcowboyhat", "🤠")
$Emojis.Add("1f627_anguishedface", "😧")
$Emojis.Add("1f911_moneymouthface", "🤑")
$Emojis.Add("1f915_facewithheadbandage", "🤕")
$Emojis.Add("1f4a1_electriclightbulb", "💡")
$Emojis.Add("1f917_huggingface", "🤗")
$Emojis.Add("1f97a_pleadingface", "🥺")
$Emojis.Add("1f92a_zanyface", "🤪")
$Emojis.Add("1f9b7_tooth", "🦷")
$Emojis.Add("203c_doubleexclamationmark", "‼️")
$Emojis.Add("1f440_eyes", "👀")
$Emojis.Add("2795_heavyplussign", "➕")
$Emojis.Add("2796_heavyminussign", "➖")   
$Emojis.Add("1f389_partypopper", "🎉") 
$Emojis.Add("screamingfear", "😱")
$Emojis.Add("womanraisinghand", "🙋‍♀️ ")
$Emojis.Add("support", "🆘")
$Emojis.Add("callme", "🧘‍♂️")
$Emojis.Add("Vampire-tone3", "🧛🏽‍♂️")
$Emojis.Add("sweatgrinning", "😅")
$Emojis.Add("ttm", "🈶")
$Emojis.Add("oreoyum", "🍪")
$Emojis.Add("victory", "✌️")
$Emojis.Add("vulcansalute", "🖖")
$Emojis.Add("rainbowsmile", "🌈")
$Emojis.Add("smile", "😄")
$Emojis.Add("womandeveloper", "👩‍💻")
$Emojis.Add("whew", "😮‍💨")

$StartDate = (Get-Date).AddDays(-360) 
$EndDate = (Get-Date).AddDays(1)

# Find audit records for Teams emoji reactions
Write-Output "Looking for audit records..."
[array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations ReactedToMessage -Formatted -SessionCommand ReturnLargeSet -ResultSize 5000
If ($Records.Count -eq 0) {
    Write-Output "No Teams emoji reaction records found"
    Break
}   Else {
    $Records = $Records | Sort-Object Identity -Unique
    Write-Output ("Found {0} Teams emoji reaction records to process..." -f $Records.Count)
    $Report = [System.Collections.Generic.List[Object]]::new()  
}

# Genreate the report
ForEach ($Rec in $Records) {

    $AuditData = $Rec.AuditData | ConvertFrom-Json
    Try {
        $Emoji = $Emojis[$AuditData.MessageReactionType]
    } Catch {
        $Emoji = $null
    }
    If ($Null -eq $Emoji) {
        $Emoji = $AuditData.MessageReactionType
    }
    # Is the message from a chat or a channel conversation?
    $Channel = $AuditData.ChannelName
    $Team = $AuditData.TeamName
    If ($Null -eq $Channel) {
        $Channel = "Chat"
        $Team = $null
    }

    $Reaction = $AuditData.MessageReactionType
    If ($null -eq $Reaction) {
       $Reaction = $AuditData.OldValue + " (Removed)"
    }

    # Get user displayname if user info is available
    If ($UserInfoAvailable) {
        $UserDisplayName = $UserHash[$Rec.UserIds]
        If ($Null -eq $UserDisplayName) {
            $UserDisplayName = $Rec.UserIds
        }
    }   Else {
        $UserDisplayName = $Rec.UserIds
    }

    $ReportLine = [PSCustomObject][Ordered]@{ 
        User            = $UserDisplayName
        Reaction        = $Reaction
        Emoji           = $Emoji
        ReactionTime    = $Rec.CreationDate
        Channel         = $Channel
        Team            = $Team
        ParentMessageId = $AuditData.ParentMessageId
        UserId           = $Rec.UserIds
    }
    $Report.Add($ReportLine)
}

$Report = $Report | Sort-Object {$_.ReactionTime -as [datetime]}
$Report | Out-GridView -Title "Teams Emoji Reactions Report"

Write-Output ""
Write-Output "Top ten Emoji reactions used in Teams"    
$Report | Group-Object Reaction -NoElement | Sort-Object Count -Descending | Select-Object -First 10 | Format-Table Name, Count
Write-Output ""
Write-Output "Top ten Users who reacted with Emojis in Teams"
$Report | Group-Object User -NoElement | Sort-Object Count -Descending | Select-Object -First 10 | Format-Table Name, Count

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
