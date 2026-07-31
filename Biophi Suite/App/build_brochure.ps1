<#
.SYNOPSIS
    Windows PowerPoint COM-automation brochure builder for Biophi Suite.

.DESCRIPTION
    NOTE: This script uses Windows COM automation and requires Microsoft
    PowerPoint to be installed.  It is preserved for reference and for
    Windows-desktop workflows.

    For a server-hosted or Linux environment, use build_deck.py instead,
    which relies on python-pptx and has no OS-level dependencies.

.PARAMETER OutputPath
    Destination path for the generated .pptx file.

.PARAMETER TemplatePath
    Path to the branded template .pptx file.

.PARAMETER DataJson
    Path to a JSON file containing analysis results produced by analysis.py.

.EXAMPLE
    .\build_brochure.ps1 -DataJson ".\output\results.json" -OutputPath ".\output\brochure.pptx"
#>
param(
    [string]$OutputPath   = ".\output\brochure.pptx",
    [string]$TemplatePath = ".\template.pptx",
    [string]$DataJson     = ".\output\results.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---- Validate inputs -------------------------------------------------------
if (-not (Test-Path $TemplatePath)) {
    Write-Error "Template not found: $TemplatePath"
    exit 1
}
if (-not (Test-Path $DataJson)) {
    Write-Error "Data file not found: $DataJson.  Run run_biophi_suite.py first."
    exit 1
}

$results = Get-Content $DataJson | ConvertFrom-Json

# ---- Open PowerPoint via COM -----------------------------------------------
$pptApp = New-Object -ComObject PowerPoint.Application
$pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

$deck = $pptApp.Presentations.Open((Resolve-Path $TemplatePath).Path)

foreach ($result in $results) {
    $fields      = $result.fields
    $predictions = $result.predictions
    $titleText   = if ($fields.Name) { $fields.Name } elseif ($fields.Title) { $fields.Title } else { $result.record_id }

    # Title slide
    $titleLayout = $deck.SlideMaster.CustomLayouts.Item(1)
    $slide = $deck.Slides.AddSlide($deck.Slides.Count + 1, $titleLayout)
    $slide.Shapes.Title.TextFrame.TextRange.Text = $titleText

    foreach ($pred in $predictions) {
        $contentLayout = $deck.SlideMaster.CustomLayouts.Item(2)
        $contentSlide  = $deck.Slides.AddSlide($deck.Slides.Count + 1, $contentLayout)
        $contentSlide.Shapes.Title.TextFrame.TextRange.Text = $pred.filename

        $lines = @()
        foreach ($p in $pred.result.predictions) {
            $pct   = [math]::Round($p.confidence * 100)
            $lines += "- $($p.class): $pct% confidence"
        }
        if ($lines.Count -eq 0) { $lines = @("No detections above threshold.") }
        $contentSlide.Shapes.Item(2).TextFrame.TextRange.Text = ($lines -join "`n")

        $imgPath = Join-Path (Split-Path $DataJson) "..\Assets\image_cache\$($pred.filename)"
        if (Test-Path $imgPath) {
            $contentSlide.Shapes.AddPicture(
                (Resolve-Path $imgPath).Path,
                [Microsoft.Office.Core.MsoTriState]::msoFalse,
                [Microsoft.Office.Core.MsoTriState]::msoTrue,
                360, 108, 288, 0
            ) | Out-Null
        }
    }
}

$outDir = Split-Path $OutputPath
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

$deck.SaveAs((Resolve-Path -LiteralPath (New-Item -Force $OutputPath -ItemType File)).Path)
$deck.Close()
$pptApp.Quit()

Write-Host "Brochure saved to: $OutputPath"
