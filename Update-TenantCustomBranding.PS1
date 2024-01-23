# Update-TenantCustomBranding.PS1
# A runbook to update the tenant's sign-in text and background image when the year changes
# 
Connect-MgGraph -Scopes Organization.ReadWrite.All -NoWelcome -Identity

$TenantId = (Get-MgOrganization).TenantId
# Get current sign-in text
[string]$SignInText = (Get-MgOrganizationBranding -OrganizationId $TenantId).SignInPageText
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