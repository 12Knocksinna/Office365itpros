# Update-BannedPasswordsList.ps1
# This script updates the list of banned passwords in Microsoft Entra ID (Azure AD) using the Microsoft Graph PowerShell SDK.
# See https://learn.microsoft.com/en-us/entra/identity/authentication/concept-password-ban-bad

# V1.0 17-Jun-2025
# GitHub link: 

Connect-MgGraph -Scopes Directory.ReadWrite.All -NoWelcome

$Scopes = Get-MgContext | Select-Object -ExpandProperty Scopes
If ($Scopes -notcontains "Directory.ReadWrite.All") {
    Write-Host "You need to consent to the Directory.ReadWrite.All scope to run this script" -ForegroundColor Red
    Break
}

# Import the list of banned passwords from a CSV file. The CSV file should have a single column named 'Banned Passwords'.
$ImportPasswordFile = "C:\temp\BannedPasswords.csv"
[array]$NewBannedPasswordsInput = Import-CSV -Path $ImportPasswordFile | Select-Object -ExpandProperty 'Banned Passwords'
[array]$NewBannedPasswords = $null

If ($NewBannedPasswordsInput -eq $null -or $NewBannedPasswordsInput.Count -eq 0) {
    Write-Host "No banned passwords found in the input file. Please check the file and try again." -ForegroundColor Red
    Break
}

# Check that we have valid passwords to add
ForEach ($Password in $NewBannedPasswordsInput) {
    If ($Password.Length -lt 4 -or $Password.Length -gt 16) {
        Write-Host ("Password {0} is not between 4 and 16 characters long. It won't be added to the list of banned passwords." -f $Password) -ForegroundColor Red
        Continue
    } Else  {
        $NewBannedPasswords += $Password.Trim()
    }
}

# Get directory setting object for password rules
$Policy = (Get-MgBetaDirectorySetting | Where-Object {$_.TemplateId -eq "5cf42378-d67d-4f36-ba46-e8b86229381d"})
If ($null -eq $Policy) {
    # No custom directory setting found for password rules, so create a new one
    Write-Host "No directory setting policy found for password rules - creating new policy"
    [array]$NewBannedPasswords = $NewBannedPasswords -join ([char]9)
    [int32]$LockOutDuration = 60
    [int32]$LockOutThreshold = 10

    $Value1 = @{}
    $Value1.Add("Name", "BannedPasswordList")   
    $Value1.Add("Value", $NewBannedPasswords)
    $Value2 = @{}
    $Value2.Add("Name", "BannedPasswordCheckOnPremisesMode")
    $Value2.Add("Value", "Enforce")
    $Value3 = @{}
    $Value3.Add("Name", "EnableBannedPasswordCheckOnPremises")    
    $Value3.Add("Value", 'false')
    $Value4 = @{}
    $Value4.Add("Name", "EnableBannedPasswordCheck")
    $Value4.Add("Value", 'true')
    $Value5 = @{}
    $Value5.Add("Name", "LockoutDurationInSeconds")
    $Value5.Add("Value", $LockoutDuration -as [int32])
    $Value6 = @{}
    $Value6.Add("Name", "LockoutThreshold")
    $Value6.Add("Value", $LockOutThreshold)

    $NewBannedListParameters = @{}
    $NewBannedListParameters.Add("templateId", "5cf42378-d67d-4f36-ba46-e8b86229381d")
    $NewBannedListParameters.Add("values", ($Value1, $Value2, $Value3, $Value4, $Value5, $Value6))
    Try {
        $Policy = New-MgBetaDirectorySetting -BodyParameter $NewBannedListParameters -ErrorAction Stop
        Write-Host "New directory setting created with ID {0}" -f $Policy.Id
    } Catch {
        Write-Host "Error creating directory setting for password rules"
        Write-Host $_.Exception.Message
        Break
    }
} Else {
    # The tenant already has a directory setting for password rules, so update it
    Write-Host ("Directory setting found with ID {0}. Updating current policy" -f $Policy.Id)
    [array]$PolicyValues =  Get-MgBetaDirectorySetting -DirectorySettingId $Policy.Id | Select-Object -ExpandProperty Values
    [array]$CurrentBannedList = $PolicyValues | Where-Object {$_.Name -eq "BannedPasswordList"} | Select-Object -ExpandProperty Value
    [array]$CurrentBannedList = $CurrentBannedList -Split([char]9)
   
    [array]$NewBannedPasswords = $NewBannedPasswords + $CurrentBannedList | Sort-Object -Unique
    If ($NewBannedPasswords.count -gt 1000) {
        Write-Host "Banned password list has more than 1000 entries. Trimming to first 1000 entries"
        $NewBannedPasswords = $NewBannedPasswords | Select-Object -First 1000
    }
    [array]$NewBannedPasswords = $NewBannedPasswords -join ([char]9)
    ($PolicyValues | Where-Object {$_.Name -eq "BannedPasswordList"}).Value = $NewBannedPasswords
    Try {
        Update-MgBetaDirectorySetting -DirectorySettingId $Policy.id -Values $PolicyValues -ErrorAction Stop
        Write-Host "Password rules updated successfully"
    } Catch {
        Write-Host "Error updating directory setting for password rules"
        Write-Host $_.Exception.Message
        Break
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.