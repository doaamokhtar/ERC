<#
.SYNOPSIS
    Rebuilds the WB documents in docs/ from the scripts in docs/user-guide-build/.

.DESCRIPTION
    Installs what the build needs the first time (npm package docx, Python openpyxl + Pillow),
    then writes the user guide and SLA template (Word) and the requirements workbook (Excel).
    With -Pdf, Word also exports each .docx to PDF next to it.

    Needs Node.js 18+ and Python 3.10+ on PATH. -Pdf needs Microsoft Word.
    A document that is open in Word or Excel cannot be overwritten - close it first.

.EXAMPLE
    .\tools\build-docs.ps1
    .\tools\build-docs.ps1 -Pdf
#>
param(
    [switch]$Pdf,
    [string]$Guide        = "WB-User-Guide-v1.1-DRAFT.docx",
    [string]$Sla          = "WB-SLA-Requirements-Template-v1.1.docx",
    [string]$Requirements = "WB-Requirements-and-SLA-Template-v1.4.xlsx"
)
$ErrorActionPreference = "Stop"
$docsDir  = Resolve-Path "$PSScriptRoot\..\docs"
$buildDir = Join-Path $docsDir "user-guide-build"

foreach ($tool in "node", "npm", "python") {
    if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) { throw "$tool is not on PATH - install it first (see README, 'Building the documents')." }
}

# --- dependencies, first run only
if (-not (Test-Path "$buildDir\node_modules\docx")) {
    Write-Host "Installing npm packages..."
    Push-Location $buildDir; try { npm install --no-audit --no-fund | Out-Host } finally { Pop-Location }
}
python -c "import openpyxl, PIL" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Installing Python packages..."
    python -m pip install --user -r "$buildDir\requirements.txt" | Out-Host
}

# --- a file open in Office is locked; say so plainly instead of failing half-way
$targets = @($Guide, $Sla, $Requirements) | ForEach-Object { Join-Path $docsDir $_ }
foreach ($t in $targets) {
    if (Test-Path $t) {
        try { [IO.File]::Open($t, 'Open', 'ReadWrite', 'None').Close() }
        catch { throw "$(Split-Path $t -Leaf) is open in another program - close it and run again." }
    }
}

# --- build
Push-Location $buildDir
try {
    node build.js        (Join-Path $docsDir $Guide)
    node sla-template.js (Join-Path $docsDir $Sla)
    python requirements-template.py (Join-Path $docsDir $Requirements)
} finally { Pop-Location }

# --- optional PDFs (also fills in the guide's table of contents)
if ($Pdf) {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false; $word.DisplayAlerts = 0
    try {
        foreach ($name in $Guide, $Sla) {
            $path = Join-Path $docsDir $name
            $wdoc = $word.Documents.Open($path, $false, $false)
            foreach ($toc in $wdoc.TablesOfContents) { $toc.Update() }
            $wdoc.Save()
            $wdoc.SaveAs2([IO.Path]::ChangeExtension($path, ".pdf"), 17)
            $wdoc.Close(0)
            Write-Host "pdf     $([IO.Path]::ChangeExtension($name, '.pdf'))"
        }
    } finally { $word.Quit() }
} else {
    Write-Host "Open the guide in Word and press F9 on the contents to fill in page numbers, or run again with -Pdf."
}
