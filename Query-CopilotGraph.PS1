# Query-CopilotGraph.PS1
# V1.0 23-Feb-2026
# An example of using the Copilot retrieval Graph API

# Github Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Query-CopilotGraph.PS1

# Delegated access only - executes in the context of a signed-in user
Connect-MgGraph -NoWelcome -Scopes Files.Read.All, Sites.Read.All

$Uri = 'https://graph.microsoft.com/v1.0/copilot/retrieval'

$Query = Read-Host "Enter your query for Copilot retrieval"
Write-Host ("OK. I'll ask Copilot to retrieve information about {0}:" -f $Query) -ForegroundColor Green

$CurrentDate = Get-Date -format 'yyyy-MM-dd'

$LookBackDate = Read-Host "How far back should I look for information? (Enter a date in the format YYYY-MM-DD, or leave blank for 30 days)"
If ( ([String]::IsNullOrEmpty($LookBackDate))) {
    $LookBackDate = (Get-Date).AddDays(-30).ToString('yyyy-MM-dd')
} Else {    
    $ParsedDate = Get-Date $LookBackDate -ErrorAction SilentlyContinue
    If (([String]::IsNullOrEmpty($ParsedDate))) {
        Write-Host "Invalid date format. Please enter a date in the format YYYY-MM-DD." -ForegroundColor Red
        break
    } ElseIf ($ParsedDate -gt (Get-Date)) {
        Write-Host "The date cannot be in the future. Please enter a valid date." -ForegroundColor Red
        break
    } Else {
        $LookBackDate = $ParsedDate.ToString('yyyy-MM-dd')
    }
}

$BodyRequest = @{}
$BodyRequest.Add("datasource", "sharepoint")
$BodyRequest.Add("queryString", $Query)
$BodyRequest.Add("maximumNumberOfResults", "5")

# Add the date limiter
$BodyRequest.add("filterexpression",'LastModifiedTime>=' + $LookBackDate + ' AND LastModifiedTime<=' + $CurrentDate)

# Add the SharePoint metadata to return for each hit
$Metadata = "title", "author", "size", "Created", "LastModifiedTime"

$BodyRequest.add("resourceMetadata", $metadata)

Write-Host "Sending request to Copilot retrieval API..." -ForegroundColor Green
[array]$Data = Invoke-MgGraphRequest -Uri $Uri -Method Post -Body $BodyRequest -OutputType PsObject

If ($Data.RetrievalHits.Count -gt 0) {
    Write-Host ("Copilot retrieval found {0} hits for your query." -f $Data.RetrievalHits.Count) -ForegroundColor Green
} Else {
    Write-Host "No results found for your query." -ForegroundColor Yellow
    break
}

Write-Host ""
Write-Host ("Copilot retrieval results for the query: {0}" -f $Query) -ForegroundColor Green
Write-Host "----------------------------------------------------------------------------"

ForEach ($Hit in $Data.RetrievalHits) {
    Write-Host ("Title: {0}" -f $Hit.ResourceMetadata.title) -ForegroundColor Cyan
    Write-Host ("Author: {0}" -f $Hit.ResourceMetadata.author) -ForegroundColor Cyan
    Write-Host ("Size: {0} bytes" -f $Hit.ResourceMetadata.size) -ForegroundColor Cyan
    Write-Host ("Created: {0}" -f $Hit.ResourceMetadata.Created) -ForegroundColor Cyan
    Write-Host ("Last Modified: {0}" -f $Hit.ResourceMetadata.LastModifiedTime) -ForegroundColor Cyan
    If ($Hit.webUrl) {
         Write-Host ("Web URL: {0}" -f $Hit.webUrl) -ForegroundColor Blue
    } Else {
        Write-Host ("Resource extract: {0}" -f $Hit.extracts.text.substring(0,120)) -ForegroundColor Blue
    }
    If ($Hit.SensitivityLabel.DisplayName) {
        Write-Host ("Sensitivity Label: {0}" -f $Hit.SensitivityLabel.DisplayName) -ForegroundColor Magenta
    }
    Write-Host ("Relevance: {0}" -f $Hit.extracts.relevanceScore)
    Write-Host ""
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.