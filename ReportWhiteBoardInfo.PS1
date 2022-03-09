# ReportWhiteBoardInfo.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportWhiteBoardInfo.PS1
# An example of how to use the Whiteboard Admin module to create a report of all whiteboards and their owners in a tenant. The original code
# is in https://pnp.github.io/script-samples/whiteboard-report-usage/README.html?tabs=whiteboard
# Import WhiteboardAdmin module
# Warning - the current version of the WhiteboardAdmin module does not retrieve details of whiteboards stored in OneDrive for Business.
Import-Module WhiteboardAdmin
# Connect to the Microsoft Graph
Connect-MgGraph -TenantId $TenantId -Scope "Directory.Read.All, User.Read.All"

try {
	$dateTime = (Get-Date).toString("dd-MM-yyyy")
	$fileName = "WhiteboardReport-" + $dateTime + ".csv"
	$outputView = "c:\temp\" + $fileName
	
	# The geography to look for board owners in. Accepted values are: Europe, Australia, or Worldwide (all boards not in australia or europe).
	$supportedGeographies = @("Europe", "Australia", "Worldwide")
	
	# Array to hold Whiteboard owners
	$WhiteboardOwners = [System.Collections.Generic.List[Object]]::new(); $i=0

	foreach ($geography in $supportedGeographies) {
		Write-Host "Getting Whiteboard owners for geography: $($geography)..."
		$GeographyOwners = Get-WhiteboardOwners -Geography $Geography		
		
		foreach ($UserId in $GeographyOwners.items) {	
              $User = Get-MgUser -UserId $UserId
              $i++
              $ReportLine  = [PSCustomObject][Ordered]@{
                DisplayName     = $User.DisplayName
                UPN             = $User.UserPrincipalName 
                Geography       = $Geography
                UserId          = $UserId
               }
            $WhiteboardOwners.Add($ReportLine) 

		} # End ForEach Owner
		
		Write-Host "Total whiteboard owners found so far $($i)"
	} # EndForEach Geography
	
	# Array to hold Whiteboard details
	$Whiteboards = [System.Collections.Generic.List[Object]]::new()
	
	# Get whiteboards from the Microsoft Whiteboard service by owners
	foreach ($Owner in $WhiteboardOwners) {
		Write-Host "Getting Whiteboards for owner: $($Owner.UPN) ..."
		$whiteboardInfo = Get-Whiteboard -UserId $Owner.UserID
		
		foreach ($whiteboardInstance in $whiteboardInfo) {   
              $ReportLine  = [PSCustomObject][Ordered]@{
                User            = $Owner.DisplayName
                UPN             = $Owner.UPN
                WhiteboardId    = $whiteboardInstance.Id
                Title           = $whiteboardInstance.Title
                IsShared        = $whiteboardInstance.IsShared
                Created         = Get-Date($whiteboardInstance.CreatedTime) -format g
                Modified        = Get-Date($whiteboardInstance.LastModifiedTime) -format g
                Geography       = $Owner.Geography
                UserId          = $Owner.UserId
               }
           $Whiteboards.Add($ReportLine)             
       } #End Foreach Whiteboards
    	
	    Write-Host "Found $($whiteboards.Count) Whiteboards owned by: $($Owner.UPN)"
	} # End Foreach Whiteboard owners
	
	Write-Host "Found $($whiteboards.Count) Whiteboards in the tenant."

	# Export the results to a CSV file and Out-GridView
	$Whiteboards | Export-CSV -Path $outputView -Force -NoTypeInformation
  $Whiteboards | Out-GridView
	Write-Host "Finished"
}
catch {
    Write-Host -f Red "Error:" $_.Exception.Message
}
# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
