# AssignLocalLanguageStringsSensitivityLabels.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/AssignLocalLanguageStringsSensitivityLabels.PS1
# A script showing how to set local language values for sensitivity labels using the Microsoft Translator service
# to translate the English lamguage values from the labels.

Function Translate-String {
  [cmdletbinding()]
    Param(
        [string]$InputString,
        [string]$LanguageCode )

$textJson = @{
        "Text" = $InputString
    } | ConvertTo-Json
$Body = "[$textJson]"

$Uri = "$($global:BaseUri)&to=$($LanguageCode)&textType=html"
$Status = Invoke-RestMethod -Method Post -Uri $uri -Headers $global:headers -Body $Body
$Translation = $Status[0].translations[0].text

Return $Translation
}

$TranslatorKey = "20be21011ba2747a0996c4134492bfc97" # You need a valid key for Microsoft Translator

$global:BaseUri = "https://api.cognitive.microsofttranslator.com/translate?api-version=3.0"
$global:Headers = @{
    'Ocp-Apim-Subscription-Key' = $translatorKey
    'Ocp-Apim-Subscription-Region' = 'northeurope' # added (change if needed)
    'Content-type' = 'application/json; charset=utf-8'
}

# Make sure that we're connected to Exchange Online and the compliance endopoint
Connect-ExchangeOnline
Connect-IPPSSession

# Declare the set of languages that we're going to translate for
$Languages = @("fr-fr","it-it","de-de", "ar-ar")

# Find set of labels that support LLV strings
Write-Host "Finding sensitivity labels to process..."
[array]$FileLabels = Get-Label | Where-Object {$_.ContentType -Like "*File*"}

ForEach ($Label in $FileLabels) {
   $FrenchDisplayName = $Null; $GermanDisplayName = $Null; $ItalianDisplayName = $Null; $ArabicDisplayName = $Null
   $FrenchTooltip = $Null; $GermanToolTip = $Null; $FrenchToolTip = $Null; $ArabicToolTip = $Null
   Write-Host ("Setting language values for the {0} label" -f $Label.displayname)
   [string]$FrenchDisplayName = Translate-String -InputString $Label.DisplayName -LanguageCode "fr"
   [string]$FrenchTooltip = Translate-String -InputString $Label.ToolTip -LanguageCode "fr"
   [string]$ItalianDisplayName = Translate-String -InputString $Label.DisplayName -LanguageCode "it"
   [string]$ItalianTooltip = Translate-String -InputString $Label.ToolTip -LanguageCode "it"
   [string]$GermanDisplayName = Translate-String -InputString $Label.DisplayName -LanguageCode "de"
   [string]$GermanTooltip = Translate-String -InputString $Label.ToolTip -LanguageCode "de"
   [string]$ArabicDisplayName = Translate-String -InputString $Label.DisplayName -LanguageCode "ar"
   [string]$ArabicTooltip = Translate-String -InputString $Label.ToolTip -LanguageCode "ar"
   
   $DisplayNameLocaleSettings = [PSCustomObject]@{LocaleKey='DisplayName';
    Settings=@(
       @{key=$Languages[0];Value=$FrenchDisplayName;}
       @{key=$Languages[1];Value=$ItalianDisplayName;}   
       @{key=$Languages[2];Value=$GermanDisplayName;}
       @{key=$Languages[3];Value=$ArabicDisplayName;})}
  $TooltipLocaleSettings = [PSCustomObject]@{LocaleKey='Tooltip';
   Settings=@(
       @{key=$Languages[0];Value=$FrenchToolTip;}
       @{key=$Languages[1];Value=$ItalianToolTip;}
       @{key=$Languages[2];Value=$GermanToolTip;}
       @{key=$Languages[3];Value=$ArabicTooltip;})}
 Set-Label -Identity $Label.ImmutableId -LocaleSettings (ConvertTo-Json $DisplayNameLocaleSettings -Depth 3 -Compress),(ConvertTo-Json $TooltipLocaleSettings -Depth 3 -Compress)
}

rite-Host "All labels updated with local language values..."

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
