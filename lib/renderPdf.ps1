param(
    [Parameter(Mandatory=$true)][string]$PptxPath,
    [Parameter(Mandatory=$true)][string]$PdfPath
)

$ErrorActionPreference = "Stop"

if (Test-Path $PdfPath) { Remove-Item $PdfPath -Force }

$app = New-Object -ComObject PowerPoint.Application
try {
    $pres = $app.Presentations.Open($PptxPath, $true, $false, $false)
    try {
        $pres.SaveAs($PdfPath, 32) # ppSaveAsPDF
    } finally {
        $pres.Close()
    }
} finally {
    $app.Quit()
}

if (-not (Test-Path $PdfPath)) {
    throw "PDF was not created at $PdfPath"
}

Write-Output "OK"
