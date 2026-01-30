<#
.SYNOPSIS
    Cleans DOCX fonts, exports via Word Online, and generates a print‑optimized PDF.

.REQUIREMENTS
    - Microsoft Graph PowerShell SDK
    - A registered AAD app with Files.ReadWrite / Sites.ReadWrite permissions
    - DOCX must be in OneDrive or SharePoint for the cloud export


.DESCRIPTION    This script automates the process of preparing a DOCX document for Kindle Direct Publishing (KDP) by:
    1. Cleaning up fonts inside the DOCX (removing Lucida* fonts)
    2. Uploading the cleaned DOCX to OneDrive and exporting it to PDF using Word Online via Microsoft Graph
    3. Optimizing the resulting PDF for print by removing hyperlinks and annotations

    Example usage:
.\Build-KDPReadyPDF.ps1 `
    -DocxPath "C:\Book\Automating365.docx" `
    -OneDrivePath "Books/PowerShell" `
    -OutputFolder "C:\Book\Build"

.NOTES
    Author: Tony Redmond’s automated workflow builder (Copilot)
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$DocxPath,          # Local path to your master DOCX
    [Parameter(Mandatory=$true)]
    [string]$OneDrivePath,      # Path in OneDrive where the DOCX will be uploaded
    [Parameter(Mandatory=$true)]
    [string]$OutputFolder       # Local folder for final PDFs
)

# -----------------------------
# 0. Setup
# -----------------------------
Write-Host "Starting KDP document pipeline..." -ForegroundColor Cyan

$CleanDocx    = Join-Path $OutputFolder "Cleaned.docx"
$CloudPdf     = Join-Path $OutputFolder "CloudExport.pdf"
$PrintPdf     = Join-Path $OutputFolder "KDP_PrintReady.pdf"

# -----------------------------
# 1. CLEAN FONTS IN DOCX
# -----------------------------
Write-Host "Cleaning fonts inside DOCX (removing Lucida*)" -ForegroundColor Green

$temp = Join-Path $env:TEMP ("DOCX_" + (New-Guid))
New-Item -ItemType Directory -Path $temp | Out-Null

# Unzip DOCX
Expand-Archive -Path $DocxPath -DestinationPath $temp -Force

# Target XML files
$xmlTargets = @(
    Join-Path $temp "word\styles.xml",
    Join-Path $temp "word\document.xml",
    Join-Path $temp "word\fontTable.xml"
)

foreach ($file in $xmlTargets) {
    if (Test-Path $file) {
        (Get-Content $file) |
            ForEach-Object { $_ -replace "Lucida[^<""]*", "Consolas" } |
            Set-Content $file -Encoding UTF8
    }
}

# Rebuild cleaned DOCX
if (Test-Path $CleanDocx) { Remove-Item $CleanDocx -Force }
Compress-Archive -Path (Join-Path $temp "*") -DestinationPath $CleanDocx -Force

# Cleanup
Remove-Item $temp -Recurse -Force

Write-Host "DOCX cleaned and rebuilt → $CleanDocx" -ForegroundColor Green

# -----------------------------
# 2. EXPORT DOCX → PDF USING WORD ONLINE (Graph)
# -----------------------------
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Files.ReadWrite", "Sites.ReadWrite.All" -NoWelcome

# Upload file to OneDrive
Write-Host "Uploading DOCX to OneDrive..." -ForegroundColor Green
$drive = Get-MgUserDrive -UserId (Get-MgContext).Account
$driveId = $drive.Id

$filename = Split-Path $CleanDocx -Leaf
$cloudItem = New-MgDriveItemContentUploadSession `
    -DriveId $driveId `
    -DriveItemId "root:/$($OneDrivePath)/$filename:" `
    -AdditionalProperties @{}

# Upload in one shot
Invoke-RestMethod -Uri $cloudItem.UploadUrl -Method PUT -InFile $CleanDocx -ContentType "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

Write-Host "Exporting to PDF via Word Online..." -ForegroundColor Green

$encodedPath = [System.Web.HttpUtility]::UrlEncode("$OneDrivePath/$filename")
$pdfBytes = Invoke-MgGraphRequest `
    -Uri "https://graph.microsoft.com/v1.0/me/drive/root:/$encodedPath:/content?format=pdf" `
    -Method GET `
    -OutputType Stream

[IO.File]::WriteAllBytes($CloudPdf, $pdfBytes)

Write-Host "PDF export complete → $CloudPdf" -ForegroundColor Green

# -----------------------------
# 3. OPTIMIZE PDF FOR PRINT
# -----------------------------
Write-Host "Optimizing PDF for print (removing hyperlinks & annotations)..." -ForegroundColor Cyan

Add-Type -AssemblyName System.IO.Compression.FileSystem

try {
    $reader = New-Object iTextSharp.text.pdf.PdfReader($CloudPdf)
    $n = $reader.NumberOfPages
    $out = New-Object iTextSharp.text.pdf.PdfStamper($reader, [System.IO.File]::Create($PrintPdf))

    for ($i=1; $i -le $n; $i++) {
        $page = $reader.GetPageN($i)
        $page.Remove(PdfName.ANNOTS)
    }

    $out.Close()
    $reader.Close()
}
catch {
    Write-Host "iTextSharp optimisation failed — using fallback scrubber." -ForegroundColor Yellow
    
    # Fallback method using PyPDF2-like cleaning via .NET
    $pdf = [System.IO.File]::ReadAllText($CloudPdf)
    $pdf = $pdf -replace "/Annots\s*\[.*?\]", ""
    [System.IO.File]::WriteAllText($PrintPdf, $pdf)
}

Write-Host "Print‑optimized PDF created → $PrintPdf" -ForegroundColor Green

Write-Host "`nPipeline complete. KDP-ready PDF is available." -ForegroundColor Cyan