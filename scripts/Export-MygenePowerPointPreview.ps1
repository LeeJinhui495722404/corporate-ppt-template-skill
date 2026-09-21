param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,

    [int]$Width = 1280,

    [int]$Height = 720
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $PptxPath -PathType Leaf)) {
    throw "PPTX does not exist: $PptxPath"
}

$resolvedPptx = (Resolve-Path -LiteralPath $PptxPath).Path
$resolvedOutput = [System.IO.Path]::GetFullPath($OutputDirectory)
[System.IO.Directory]::CreateDirectory($resolvedOutput) | Out-Null

$powerPoint = $null
$presentation = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($resolvedPptx, -1, 0, 0)
    $slideCount = $presentation.Slides.Count
    $presentation.Export($resolvedOutput, 'PNG', $Width, $Height)
    $presentation.Close()
    $presentation = $null
    $powerPoint.Quit()
    $powerPoint = $null

    $exportedFiles = @(Get-ChildItem -LiteralPath $resolvedOutput -Filter '*.PNG' -File)
    [ordered]@{
        pptx = $resolvedPptx
        output_directory = $resolvedOutput
        slide_count = $slideCount
        exported_png_count = $exportedFiles.Count
        complete = ($exportedFiles.Count -eq $slideCount)
    } | ConvertTo-Json -Depth 3
}
finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    if ($null -ne $powerPoint) {
        try { $powerPoint.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
