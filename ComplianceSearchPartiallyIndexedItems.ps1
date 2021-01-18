# ComplianceSearchPartiallyIndexedItems.ps1
# Original version written by Microsoft and published at https://docs.microsoft.com/en-us/microsoft-365/compliance/investigating-partially-indexed-items-in-ediscovery?view=o365-worldwide
# Updated by Tony Redmond to add what I hope is better analysis of the partial items
# Cited in Chapter 20 of the Office 365 for IT Pros eBook
# https://github.com/12Knocksinna/Office365itpros/blob/master/ComplianceSearchPartiallyIndexedItems.ps1
 
  write-host "**************************************************"
  write-host "     Microsoft 365 Compliance Center      " -foregroundColor yellow -backgroundcolor darkgreen
  write-host "   eDiscovery Partially Indexed Item Statistics   " -foregroundColor yellow -backgroundcolor darkgreen
  write-host "**************************************************"
  " " 
  # Create a search with Error Tags Refinders enabled
  Remove-ComplianceSearch "RefinerTest" -Confirm:$false -ErrorAction 'SilentlyContinue'
  New-ComplianceSearch -Name "RefinerTest" -ContentMatchQuery "size>0" -RefinerNames ErrorTags -ExchangeLocation ALL
  Start-ComplianceSearch "RefinerTest"
  # Loop while search is in progress
  do{
      Write-host "Waiting for search to complete..."
      Start-Sleep -s 5
      $complianceSearch = Get-ComplianceSearch "RefinerTest"
  }while ($complianceSearch.Status -ne 'Completed')
  $refiners = $complianceSearch.Refiners | ConvertFrom-Json
  $errorTagProperties = $refiners.Entries | Get-Member -MemberType NoteProperty
  $partiallyIndexedRatio = $complianceSearch.UnindexedItems / $complianceSearch.Items
  $partiallyIndexedSizeRatio = $complianceSearch.UnindexedSize / $complianceSearch.Size
  " "
  "===== Partially indexed items ====="
  "         Total          Ratio"
  "Count    {0:N0}{1:P2}" -f $complianceSearch.Items.ToString("N0").PadRight(15, " "), $partiallyIndexedRatio
  "Size(GB) {0:N2}{1:P2}" -f ($complianceSearch.Size / 1GB).ToString("N2").PadRight(15, " "), $partiallyIndexedSizeRatio
  " "
  Write-Host ===== Reasons for partially indexed items =====
  $Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
  foreach($errorTagProperty in $errorTagProperties)
  {
      $name = $refiners.Entries.($errorTagProperty.Name).Name
      $count = $refiners.Entries.($errorTagProperty.Name).TotalCount
      $frag = $name.Split("{_}")
      $errorTag = $frag[0]
      $fileType = $frag[1]
      $ErrorTagText = $Null; $FileTypeText = $Null
      Switch ($ErrorTag) {
       "attachmentrms"         { $ErrorTagText = "Rights Management Encrypted Attachment" }
       "parserencrypted"       { $ErrorTagText = "Parser couldn't decrypt item" }
       "parsererror"           { $ErrorTagText = "Parser encountered an error" }
       "parserinputsize"       { $ErrorTagText = "Parser error on input (maximum) size" }
       "parsermalformed"       { $ErrorTagText = "Parser encountered malformed item" }
       "parseroutputsize"      { $ErrorTagText = "Parser error on output size" }
       "parserunknowntype"     { $ErrorTagText = "Parser encountered unknown format" }
       "parserunsupportedtype" { $ErrorTagText = "Parser encountered unsupported format" }
       "retrieverrms"          { $ErrorTagText = "Rights Management Encrypted Item" }   
       "default"               { $ErrorTagText = "Unknown error" } 
      } #End switch
      Switch ($FileType) {
        "bmp"         { $FileTypeText = "Bitmap graphic file" }
        "doc"         { $FileTypeText = "Word (DOC) document" }
        "docm"        { $FileTypeText = "Word (DOCM) template" }
        "docx"        { $FileTypeText = "Word (DOCX) document" }
        "eml"         { $FileTypeText = "Email item" }
        "encoffmetro" { $FileTypeText = "Password protected PowerPoint PPTX" }
        "gzip"        { $FileTypeText = "GZIP file" }
        "json"        { $FileTypeText = "JSON-formatted data" }
        "mhtml"       { $FileTypeText = "Web archive" }
        "mp3"         { $FileTypeText = "MP3 audio file" }
        "mp4"         { $FileTypeText = "MP4 audio/video file" }
        "mpp"         { $FileTypeText = "Microsoft Project file" }
        "pdf"         { $FileTypeText = "PDF file" }
        "noformat"    { $FileTypeText = "Unknown/no format" }
        "png"         { $FileTypeText = "PNG graphic file" }
        "ppt"         { $FileTypeText = "PowerPoint (PPT) presentation" }
        "pptx"        { $FileTypeText = "PowerPoint (PPTX) presentation" }
        "ps"          { $FileTypeText = "PostScript file" }
        "tiff"        { $FileTypeText = "TIFF graphic file" }
        "txt"         { $FileTypeText = "Text file" }
        "wav"         { $FileTypeText = "Waveform audio file" }
        "xls"         { $FileTypeText = "Excel (XLS) worksheet" }
        "xlsx"        { $FileTypeText = "Excel (XLSX) worksheet" }
        "xml"         { $FileTypeText = "XML format file" }
        "zip"         { $FileTypeText = "ZIP file" }
        "default"     { $FileTypeText = "Unknown file format" }
      } #End switch
      $ReportLine = [PSCustomObject] @{
         ErrorType   = $ErrorTag
         ErrorText   = $ErrorTagText
         FileExt     = $FileType
         FileType    = $FileTypeText
         Count       = $Count }
       $Report.Add($ReportLine) 
}
$Report | Sort Count -desc | Format-Table FileType,Count, ErrorText 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
