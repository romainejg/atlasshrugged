$ErrorActionPreference = 'Stop'
$data = Get-Content -Raw -LiteralPath $env:BROCHURE_MANIFEST | ConvertFrom-Json
$rows = @($data.rows)
if ($rows.Count -eq 0) { throw 'No analysis varieties were available for the overview.' }

$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = -1
$pres = $pp.Presentations.Open($env:BROCHURE_TEMPLATE, $false, $false, $false)
$source = $pres.Slides.Item(1)
trap {
  try { if ($pres) { $pres.Close() } } catch {}
  try { if ($pp) { $pp.Quit() } } catch {}
  throw
}

function FindShape($slide, [string]$name) {
  foreach ($shape in @($slide.Shapes)) {
    if ($shape.Name -eq $name) { return $shape }
  }
  return $null
}

function ReplaceText($shape, [string]$oldText, $newText) {
  if (-not $shape -or -not $oldText) { return }
  $replacement = if ($null -eq $newText) { '' } else { [string]$newText }
  try { $shape.TextFrame.TextRange.Replace($oldText, $replacement) | Out-Null } catch {}
}

function FormatNumber($value, [string]$format) {
  if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { return '' }
  return ([double]$value).ToString($format, [Globalization.CultureInfo]::InvariantCulture)
}

function FormatDate($value) {
  if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { return '' }
  try { return ([datetime]::Parse([string]$value)).ToString('MMMM d, yyyy') } catch { return [string]$value }
}

function ReplacePicture($slide, [string]$shapeName, $path) {
  $old = FindShape $slide $shapeName
  if (-not $old) { return }
  if (-not $path -or -not (Test-Path -LiteralPath ([string]$path))) {
    $old.Delete()
    return
  }
  $frameLeft = $old.Left; $frameTop = $old.Top
  $frameWidth = $old.Width; $frameHeight = $old.Height
  $rotation = $old.Rotation
  $z = $old.ZOrderPosition
  $old.Delete()
  $pic = $slide.Shapes.AddPicture([string]$path, 0, -1, 0, 0, -1, -1)
  $pic.LockAspectRatio = -1
  $scale = [Math]::Min($frameWidth / $pic.Width, $frameHeight / $pic.Height)
  $pic.Width = $pic.Width * $scale
  $pic.Height = $pic.Height * $scale
  $pic.Left = $frameLeft + (($frameWidth - $pic.Width) / 2)
  $pic.Top = $frameTop + (($frameHeight - $pic.Height) / 2)
  $pic.Rotation = $rotation
  while ($pic.ZOrderPosition -gt $z) { $pic.ZOrder(3) }
}

function FindLeafForPoint($leaves, $point) {
  if (-not $point) { return $null }
  return @($leaves | Where-Object { [int]$_.dtm -eq [int]$point.dtm } | Select-Object -First 1)[0]
}

function PlaceScaledLeaf($slide, [string]$shapeName, $leaf, [double]$pointsPerCm, [double]$baseline) {
  $old = FindShape $slide $shapeName
  if (-not $old) { return }
  $center = $old.Left + ($old.Width / 2)
  $z = $old.ZOrderPosition
  $old.Delete()
  if (
    -not $leaf -or -not $leaf.path -or
    -not (Test-Path -LiteralPath ([string]$leaf.path)) -or
    $null -eq $leaf.height_cm
  ) { return }
  $pic = $slide.Shapes.AddPicture([string]$leaf.path, 0, -1, 0, 0, -1, -1)
  $pic.LockAspectRatio = -1
  $pic.Height = [double]$leaf.height_cm * $pointsPerCm
  $pic.Left = $center - ($pic.Width / 2)
  $pic.Top = $baseline - $pic.Height
  while ($pic.ZOrderPosition -gt $z) { $pic.ZOrder(3) }
}

function AddRuler($slide, [double]$x, [double]$baseline, [double]$pointsPerCm, [int]$maxCm) {
  $black = 0
  $line = $slide.Shapes.AddLine($x, $baseline, $x, $baseline - ($maxCm * $pointsPerCm))
  $line.Line.ForeColor.RGB = $black
  $line.Line.Weight = 1
  for ($cm = 0; $cm -le $maxCm; $cm++) {
    $y = $baseline - ($cm * $pointsPerCm)
    $major = ($cm % 5 -eq 0)
    $tick = if ($major) { 6 } else { 3 }
    $mark = $slide.Shapes.AddLine($x, $y, $x + $tick, $y)
    $mark.Line.ForeColor.RGB = $black
    if ($major) { $mark.Line.Weight = 1.0 } else { $mark.Line.Weight = 0.5 }
    if ($major) {
      $label = $slide.Shapes.AddTextbox(1, $x - 24, $y - 6, 20, 12)
      $label.TextFrame.TextRange.Text = [string]$cm
      $label.TextFrame.TextRange.Font.Name = 'Rubik'
      $label.TextFrame.TextRange.Font.Size = 7
      $label.TextFrame.TextRange.ParagraphFormat.Alignment = 3
      $label.TextFrame.MarginLeft = 0
      $label.TextFrame.MarginRight = 0
      $label.TextFrame.MarginTop = 0
      $label.TextFrame.MarginBottom = 0
      $label.Line.Visible = 0
      $label.Fill.Visible = 0
    }
  }
  $unit = $slide.Shapes.AddTextbox(1, $x - 7, $baseline - ($maxCm * $pointsPerCm) - 14, 25, 12)
  $unit.TextFrame.TextRange.Text = 'cm'
  $unit.TextFrame.TextRange.Font.Name = 'Rubik'
  $unit.TextFrame.TextRange.Font.Size = 7
  $unit.TextFrame.MarginLeft = 0
  $unit.TextFrame.MarginRight = 0
  $unit.TextFrame.MarginTop = 0
  $unit.TextFrame.MarginBottom = 0
  $unit.Line.Visible = 0
  $unit.Fill.Visible = 0
}

function SetHarvestLabel($shape, [string]$sampleDays, [string]$sampleSize, [string]$sampleEfficiency, $point, $leaf) {
  $days = if ($point -and $null -ne $point.dtm) { ([string]$point.dtm) + ' Days' } else { '' }
  $height = if ($leaf -and $null -ne $leaf.height_cm) { '(' + (FormatNumber $leaf.height_cm '0.#') + ' cm)' } elseif ($point -and $null -ne $point.height_cm) { '(' + (FormatNumber $point.height_cm '0.#') + ' cm)' } else { '' }
  $efficiency = if ($point -and $null -ne $point.gmol) { (FormatNumber $point.gmol '0.00') + ' g/mol' } else { '' }
  ReplaceText $shape ($sampleDays + ' Days') $days
  ReplaceText $shape $sampleSize $height
  ReplaceText $shape $sampleEfficiency $efficiency
}

# Use one deck-wide physical scale so leaf sizes are comparable across varieties.
$maxLeafHeight = 0.0
$maxLeafWidth = 0.0
foreach ($row in $rows) {
  foreach ($leaf in @($row.trial_data.leaves)) {
    if ($null -ne $leaf.height_cm) {
      $maxLeafHeight = [Math]::Max($maxLeafHeight, [double]$leaf.height_cm)
      $leafWidth = if ($null -ne $leaf.width_cm) {
        [double]$leaf.width_cm
      } elseif ($leaf.pixel_width -and $leaf.pixel_height) {
        [double]$leaf.height_cm * [double]$leaf.pixel_width / [double]$leaf.pixel_height
      } else { 0.0 }
      $maxLeafWidth = [Math]::Max($maxLeafWidth, $leafWidth)
    }
  }
}
$leafPointsPerCm = 10.0
if ($maxLeafHeight -gt 0) { $leafPointsPerCm = [Math]::Min($leafPointsPerCm, 180.0 / $maxLeafHeight) }
if ($maxLeafWidth -gt 0) { $leafPointsPerCm = [Math]::Min($leafPointsPerCm, 100.0 / $maxLeafWidth) }
$leafBaseline = 396.3
$rulerMaxCm = if ($maxLeafHeight -gt 0) { [int]([Math]::Ceiling($maxLeafHeight / 5.0) * 5) } else { 5 }

# Create every output slide from the untouched template before replacing anything.
$slides = @($source)
for ($i = 1; $i -lt $rows.Count; $i++) {
  $clone = $source.Duplicate().Item(1)
  $clone.MoveTo($i + 1)
  $slides += $clone
}

for ($i = 0; $i -lt $rows.Count; $i++) {
  $slide = $slides[$i]
  $r = $rows[$i]
  $trial = $r.trial_data
  $name = if ($r.variety) { $r.variety } elseif ($r.name) { $r.name } else { $r.id }

  $identity = FindShape $slide 'TextBox 7'
  ReplaceText $identity 'C127' $name
  $segmentText = if ($r.segment) { 'Segment: ' + [string]$r.segment } else { 'Segment: ' }
  ReplaceText $identity 'Segment: Crispy' $segmentText
  ReplaceText $identity 'Nr:0' $r.hr
  ReplaceText $identity 'LMV:1' $r.ir

  $description = FindShape $slide 'TextBox 17'
  ReplaceText $description 'C127 is an experimental Crispyano type with high yield, tipburn tolerance, and slower growth. It puts on weight early so may benefit from earlier harvest. It is recommended for high density production.' $r.description

  $points = @()
  $leaves = @()
  if ($trial) {
    $points = @($trial.yield | Sort-Object dtm | Select-Object -First 3)
    $leaves = @($trial.leaves | Sort-Object dtm | Select-Object -First 3)
  }
  $point1 = if ($points.Count -gt 0) { $points[0] } else { $null }
  $point2 = if ($points.Count -gt 1) { $points[1] } else { $null }
  $point3 = if ($points.Count -gt 2) { $points[2] } else { $null }
  $leaf1 = FindLeafForPoint $leaves $point1
  $leaf2 = FindLeafForPoint $leaves $point2
  $leaf3 = FindLeafForPoint $leaves $point3
  SetHarvestLabel (FindShape $slide 'TextBox 9') '19' '(10cm)' '10 g/mol' $point1 $leaf1
  SetHarvestLabel (FindShape $slide 'TextBox 12') '21' '(13 cm)' '10 g/mol' $point2 $leaf2
  SetHarvestLabel (FindShape $slide 'TextBox 15') '21' '(14 cm)' '10/mol' $point3 $leaf3

  $details = FindShape $slide 'TextBox 29'
  $env = if ($trial) { $trial.environment } else { $null }
  $dliText = if ($env) { FormatNumber $env.dli '0.00' } else { '' }
  $tempText = if ($env) { FormatNumber $env.temperature '0.00' } else { '' }
  $rhText = if ($env -and $null -ne $env.rh) { 'RH: ' + (FormatNumber $env.rh '0.00') + '%' } else { 'RH:' }
  $dateText = if ($env) { FormatDate $env.harvest_date } else { '' }
  $strengthText = if ($trial) { $trial.strengths } else { '' }
  $weaknessText = if ($trial) { $trial.weaknesses } else { '' }
  ReplaceText $details '23.16' $dliText
  ReplaceText $details '20.65' $tempText
  ReplaceText $details 'RH:' $rhText
  ReplaceText $details 'May 25, 2026' $dateText
  ReplaceText $details 'higher yield; strong overall value' $strengthText
  ReplaceText $details 'weak uniformity' $weaknessText

  ReplaceText (FindShape $slide 'TextBox 34') '560 plants/m2' ''

  ReplacePicture $slide 'Picture 36' $r.hero_path
  PlaceScaledLeaf $slide 'Picture 8' $leaf1 $leafPointsPerCm $leafBaseline
  PlaceScaledLeaf $slide 'Picture 11' $leaf2 $leafPointsPerCm $leafBaseline
  PlaceScaledLeaf $slide 'Picture 14' $leaf3 $leafPointsPerCm $leafBaseline
  AddRuler $slide 58 $leafBaseline $leafPointsPerCm $rulerMaxCm

  $chartShape = FindShape $slide 'Chart 13'
  if ($chartShape -and $points.Count -gt 0) {
    try {
      $categories = @($points | ForEach-Object { [string]$_.dtm })
      $values = @($points | ForEach-Object { if ($null -ne $_.value) { [double]$_.value } else { 0.0 } })
      $series = $chartShape.Chart.SeriesCollection(1)
      $series.XValues = [object[]]$categories
      $series.Values = [object[]]$values
      if ($series.HasDataLabels) {
        $series.DataLabels().NumberFormat = '0.00'
      }
    } catch {
      throw "Could not update the growth curve for $name. $($_.Exception.Message)"
    }
  }
}

if (Test-Path -LiteralPath $env:BROCHURE_OUTPUT) {
  Remove-Item -LiteralPath $env:BROCHURE_OUTPUT -Force
}
$pres.SaveAs($env:BROCHURE_OUTPUT, 24)
$pres.Close()
$pp.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
$env:BROCHURE_OUTPUT
