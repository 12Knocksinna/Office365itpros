# Create-NewsItemsFromRSSFeed.PS1
# An example of using the SharePoint Pages API to create and publish news items from an RSS feed
# V1.0 19-Jan-2025

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Create-NewsItemsFromRSSFeed.PS1

# Requires the Graph Sites.ReadWrite.All permission to read and create pages in the target site
# Needs to run in app-only mode to use application permissions to access any site. The best way to do this
# is to create a registered app, consent for the app to use the Sites.ReadWrite.All permission, and 
# use a certificate to authenticate the app. You can then sign into the Graph as follows:
# Disconnect-MgGraph
# Connect-MgGraph -CertificateThumbprint $Thumbprint -ApplicationId $ApplicationId -TenantId $TenantId

# News items created when in app-only mode are shown as created by the SharePoint app

# If you're using delegated permissions, this script runs in the context of the signed-in user and can 
# create news items in any site that the user can create files in. Items created when using delegated
# permissions are shown as created by the signed-in user.

# Set up the script parameters
Param (
    [Parameter(Mandatory = $true)]
    [string]$SiteURL,  # URL of the SharePoint Online site to process - like https://office365itpros.sharepoint.com/sites/IndustryNews/
    [Parameter(Mandatory = $true)]
    [string]$RSSFeed  # URL of the RSS feed to process - like https://practical365.com/feed/
)

# Date threshold to check for recent posts
$RSSDateThreshold = (Get-Date).AddDays(-7)

Write-Host ("Processing RSS feed {0} to find recent posts" -f $RSSFeed)
Try {
    [xml]$Content = Invoke-WebRequest -Uri $RSSFeed -ErrorAction Stop
} Catch {
    Write-Host ("Failed to download RSS feed {0}" -f $RSSFeed)
    Break
}

[array]$Feed = $Content.RSS.channel.item

If ($Feed.Count -eq 0) {
    Write-Host "No items found in the RSS feed"
    Break
}

# Build an input list based on the posts from the RSS feed
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Item in $Feed){

    [datetime]$LastUpdate = $Item.pubdate.split('+')[0].trim()
    If ($LastUpdate -lt $RSSDateThreshold) { Continue }
    $ReportLine = [PSCustomObject][Ordered]@{
        Title        = $Item.title
        LastUpdated  = Get-Date ($LastUpdate) -format 'dd-MMM-yyyy HH:mm:ss'
        Description  = $Item.description.'#cdata-section'
        Category     = $Item.category.'#cdata-section' -join ", "
        Author       = $Item.creator.'#cdata-section'
        Link         = $Item.link
    }
    $Report.Add($ReportLine)
}

# Sort the report so that we create the news items in the correct order
$Report = $Report | Sort-Object LastUpdated

# Now that we have a set of data from the RSS feed, we can create news items in SharePoint
# First, connect to the SharePoint site
# $Uri = "https://redmondassociates.sharepoint.com/sites/IndustryNews"
$SiteId = $SiteURL.Split('//')[1].split("/")[0] + ":/sites/" + $SiteURL.Split('//')[1].split("/")[2]

$Site = Get-MgSite -SiteId $SiteId
If ($Site) { 
    Write-Host "Connected to site $($Site.DisplayName)"
} Else {
    Write-Host "Target site not found using $Uri"
    Break
}

# Get the current set of site pages so that we can check if a news item already exists
[array]$SitePages = Get-MgSitePage -SiteId $Site.Id -Top 100

$PostLog = [System.Collections.Generic.List[Object]]::new()
ForEach ($NewPost in $Report) {

    If ($SitePages.Title -contains $NewPost.Title) {
        Write-Host ("News item {0} already exists in {1}" -f $NewPost.Title, $Site.DisplayName)
        Continue
    }

    $PostSuccess = $true
    $Link = ('<a href="{0}" target="_blank"> Read full article</a>' -f $NewPost.Link)
    $PostTitle = $NewPost.Title
    $PostName = ("News Post {0}.aspx" -f (Get-Date -format 'MMddyyy-HHmmss'))
    #$PostImage = "https://i0.wp.com/office365itpros.com/wp-content/uploads/2025/01/Top-Five-SharePoint-Features.png"
    $PostContent = $NewPost.Description + "<p><p>" + "by: " + $NewPost.Author + "</p><p>" + $Link + "</p>"

    # The title area
    $TitleArea = @{}
    $TitleArea.Add("enableGradientEffect", $true)
    $TitleArea.Add("imageWebUrl", $PostImage)
    $TitleArea.Add("layout", "imageAndTitle")
    $TitleArea.Add("showAuthor",$true)
    $TitleArea.Add("showPublishedDate", $true)
    $TitleArea.Add("showTextBlockAboveTitle", $true)
    $TitleArea.Add("textAboveTitle", $PostTitle)
    $TitleArea.Add("textAlignment", "center")
    $TitleArea.Add("imageSourceType", $null)
    $TitleArea.Add.("title", "News Post")

    # A news item only needs one web part to publish the content
    $WebPart1 = @{}
    $WebPart1.Add("id", "6f9230af-2a98-4952-b205-9ede4f9ef548")
    $WebPart1.Add("innerHtml", $PostContent)
    $WebParts = @($WebPart1)

    # The webpart is in a single column
    $Column1 = @{}
    $Column1.Add("id", "1")
    $Column1.Add("width", "12")
    $Column1.Add("webparts", $webparts)
    $Columns = @($Column1)

    $Section1 = @{}
    $Section1.Add("layout", "oneColumn") 
    $Section1.Add("id", "1")
    $Section1.Add("emphasis", "none")
    $Section1.Add("columns", $Columns)

    $HorizontalSections = @($Section1)
    $CanvasLayout = @{}
    $CanvasLayout.Add("horizontalSections", $HorizontalSections)

    # Bringing all the creation parameters together
    $Params = @{}
    $Params.Add("@odata.type", "#microsoft.graph.sitePage")
    $Params.Add("name", $PostName)
    $Params.Add("title", $PostTitle)
    $Params.Add("pagelayout", "article")
    $Params.Add("showComments", $true)
    $Params.Add("showRecommendedPages", $false)
    $Params.Add("titlearea", $TitleArea)
    $Params.Add("canvasLayout", $CanvasLayout)

    $Post = New-MgSitePage -SiteId $site.Id -BodyParameter $Params
    If ($Post) { 
        Write-Host ("Successfully created new page {0}" -f $PostTitle) 
    } Else {
        Write-Host ("Post {0} failed" -f $PostTitle) 
        $PostSuccess = $false
        Continue
    }

    If ($PostSuccess) {
    # We have a page created, so promote it to be a news post
        $Description = $NewPost.Description
        $UpdateBody = @{}
        $UpdateBody.Add("@odata.type", "#microsoft.graph.sitePage")
        $UpdateBody.Add("promotionKind", "newsPost")
        $UpdateBody.Add("description", $Description)
        $Uri = ("https://graph.microsoft.com/V1.0/sites/{0}/pages/{1}/microsoft.graph.sitePage" -f $Site.Id, $Post.Id)
        $Status = Invoke-MgGraphRequest -Uri $Uri -Method Patch -Body $UpdateBody
        If ($Status) { 
            Write-Host 'Post Promoted to News Item'
        } Else {
            Write-Host 'Post Update Failed'
            $PostSuccess = $false
            Continue
        }
    }

    If ($PostSuccess) {
        # Publish our new news item
        $Uri = ("https://graph.microsoft.com/V1.0/sites/{0}/pages/{1}/microsoft.graph.sitePage/publish" -f $Site.Id, $Post.Id)
        Invoke-MgGraphRequest -Uri $Uri -Method Post
    }
    $PostLine = [PSCustomObject][Ordered]@{
        Timestamp       = Get-Date -format 'dd-MMM-yyyy HH:mm:ss'
        Title           = $PostTitle
        'File name'     = $PostName
        Description     = $NewPost.Description
        Link            = $Link
        Site            = $Site.DisplayName
    }
    $PostLog.Add($PostLine)
}

Write-Host ""
Write-Host ("{0} posts processed and posted to {1}" -f $PostLog.Count, $Site.DisplayName)
$PostLog | Format-Table -AutoSize Timestamp, Title, Site

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.