param(
    [switch]$SkipInstall,
    [switch]$SkipPdf
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

Write-Host "ProtocolDesign build root: $Root"

if (-not $SkipInstall) {
    Write-Host "Installing Python dependencies..."
    python -m pip install -r requirements.txt
}

Write-Host "Checking Python syntax..."
python -m py_compile `
    "$Root\src\ProtocolDesign.py" `
    "$Root\src\protocol_renderer.py" `
    "$Root\src\build_help_docx.py"

Write-Host "Generating help DOCX..."
python "$Root\src\build_help_docx.py"

$HelpDocx = Get-ChildItem -LiteralPath $Root -File -Filter "ProtocolDesign_*.docx" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if (-not $HelpDocx) {
    throw "Help DOCX was not generated under $Root."
}

if (-not $SkipPdf) {
    $Soffice = Get-Command soffice -ErrorAction SilentlyContinue
    if ($Soffice) {
        Write-Host "Generating help PDF with LibreOffice..."
        & $Soffice.Source --headless --convert-to pdf --outdir "$Root" $HelpDocx.FullName
    }
    else {
        Write-Host "LibreOffice soffice command was not found; skipping PDF generation."
    }
}

Write-Host "Encrypting resources..."
python "$Root\src\encrypt_resources.py"

Write-Host "Building ProtocolDesign.exe..."
pyinstaller `
    --noconfirm `
    --onefile `
    --windowed `
    --name ProtocolDesign `
    --distpath "$Root\APP" `
    --workpath "$Root\build" `
    --specpath "$Root\build" `
    --add-data "$Root\resources\resources.enc;resources" `
    "$Root\src\ProtocolDesign.py"

Remove-Item "$Root\resources\resources.enc" -ErrorAction SilentlyContinue

$HelpPdf = Get-ChildItem -LiteralPath $Root -File -Filter "ProtocolDesign_*.pdf" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

Write-Host "Copying distribution files to APP..."
Copy-Item $HelpDocx.FullName "$Root\APP\" -Force
if ($HelpPdf) {
    Copy-Item $HelpPdf.FullName "$Root\APP\" -Force
}
Copy-Item "$Root\samples\*.docx" "$Root\APP\" -Force

Write-Host "Build complete: $Root\APP\ProtocolDesign.exe"
