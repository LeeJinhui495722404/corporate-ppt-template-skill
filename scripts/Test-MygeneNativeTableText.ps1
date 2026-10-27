param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath,

    [int[]]$Page = @(),

    [ValidateRange(0, 10)]
    [double]$TolerancePt = 0.5,

    [string]$ReportPath
)

# Read-only PowerPoint measurement. Does not save the presentation or emit cell text.
$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $PptxPath -PathType Leaf)) {
    throw "PPTX does not exist: $PptxPath"
}
if (@($Page | Where-Object { $_ -lt 1 }).Count -gt 0) {
    throw 'Page numbers must be positive.'
}
$resolvedPptx = (Resolve-Path -LiteralPath $PptxPath).Path
$resolvedReport = $null
if ($ReportPath) {
    $resolvedReport = [System.IO.Path]::GetFullPath($ReportPath)
    if ([string]::Equals($resolvedReport, $resolvedPptx, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'ReportPath must not overwrite the input presentation.'
    }
}
$powerPoint = $null
$presentation = $null
$checks = [System.Collections.Generic.List[object]]::new()
$measurementErrors = [System.Collections.Generic.List[object]]::new()
$tableCount = 0
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($resolvedPptx, -1, 0, 0)
    $selected = if ($Page.Count -eq 0) { @(1..$presentation.Slides.Count) } else { @($Page | Sort-Object -Unique) }
    if (@($selected | Where-Object { $_ -gt $presentation.Slides.Count }).Count -gt 0) {
        throw 'A requested page exceeds the presentation slide count.'
    }
    foreach ($pageNumber in $selected) {
        $slide = $presentation.Slides.Item($pageNumber)
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTable -ne -1) { continue }
            $tableCount++
            $table = $shape.Table
            for ($row = 1; $row -le $table.Rows.Count; $row++) {
                for ($column = 1; $column -le $table.Columns.Count; $column++) {
                    try {
                        $cell = $table.Cell($row, $column).Shape
                        $frame = $cell.TextFrame2
                        $available = [double]$cell.Height - [double]$frame.MarginTop - [double]$frame.MarginBottom
                        $bound = [double]$frame.TextRange.BoundHeight
                        $checks.Add([ordered]@{
                            page = $pageNumber
                            table_shape_id = [int]$shape.Id
                            row = $row
                            column = $column
                            text_height_pt = [Math]::Round($bound, 3)
                            available_height_pt = [Math]::Round($available, 3)
                            fits_height = ($bound -le $available + $TolerancePt)
                        })
                    }
                    catch {
                        $measurementErrors.Add([ordered]@{
                            page = $pageNumber
                            table_shape_id = [int]$shape.Id
                            row = $row
                            column = $column
                            error = $_.Exception.Message
                        })
                    }
                }
            }
        }
    }
}
finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    if ($null -ne $powerPoint) {
        # Do not close other presentations already open in the user's application.
        try { if ($powerPoint.Presentations.Count -eq 0) { $powerPoint.Quit() } } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
$overflow = @($checks | Where-Object { -not $_.fits_height })
$report = [ordered]@{
    passed = ($tableCount -gt 0 -and $measurementErrors.Count -eq 0 -and $overflow.Count -eq 0)
    pptx = $resolvedPptx
    selected_pages = $selected
    table_count = $tableCount
    checked_cells = $checks.Count
    overflowing_cells = $overflow.Count
    tolerance_pt = $TolerancePt
    measurement_errors = $measurementErrors.ToArray()
    checks = $checks.ToArray()
    claim_boundary = 'Native cell text height only. Review PNGs for width, readability, images, overlap and page layout.'
}
$json = $report | ConvertTo-Json -Depth 6
if ($ReportPath) {
    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($resolvedReport)) | Out-Null
    [System.IO.File]::WriteAllText($resolvedReport, $json, [System.Text.UTF8Encoding]::new($false))
}
$json
if (-not $report.passed) { exit 1 }
