# Update-TenantCustomBranding.PS1
# A runbook to update the tenant's sign-in text and background image when the year changes
# 
Connect-MgGraph -Scopes Organization.ReadWrite.All -NoWelcome 
# If running in Azure Automation, use Connect-MgGraph -Scopes Organization.ReadWrite.All -NoWelcome -Identity

$TenantId = (Get-MgOrganization).Id
# Get current sign-in text
[string]$SignInText = (Get-MgOrganizationBranding -OrganizationId $TenantId -ErrorAction SilentlyContinue).SignInPageText
If ($SignInText.Length -eq 0) {
    Write-Host "Branding information not found - exiting"
    break
}
[string]$CurrentYear = Get-Date -format yyyy
$DefaultYearImage = "c:\temp\DefaultYearImage.jpg"
$YearPresent = $SignInText.IndexOf($CurrentYear)
If ($YearPresent -gt 0) {
    Write-Output ("Year found in sign in text is {0}. No update necessary" -f $CurrentYear)
} Else {
    Write-Output ("Updating copyright date for tenant to {0}" -f $CurrentYear )
    $YearPosition = $SignInText.IndexOf('202')
    $NewSIT = $SignInText.SubString(0, ($YearPosition)) + $CurrentYear
    # Create hash table for updated parameters
    $BrandingParams = @{}
    $BrandingParams.Add("signInPageText",$NewSIT)
    Update-MgOrganizationBranding -OrganizationId $TenantId -BodyParameter $BrandingParams
    If (Test-Path $DefaultYearImage) {
        Write-Output "Updating background image..."
        $Uri = ("https://graph.microsoft.com/v1.0/organization/{0}/branding/localizations/0/backgroundImage" -f $TenantId)
        Invoke-MgGraphRequest -Method PUT -Uri $Uri -InputFilePath $DefaultYearImage -ContentType "image/jpg"
    } Else {
        Write-Output "No new background image available to update"
    }   
}