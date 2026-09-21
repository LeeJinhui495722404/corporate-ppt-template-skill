param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath,

    [string]$ReferencePath = (Join-Path $PSScriptRoot '..\assets\reference.pptx'),

    [int]$ExpectedSlideCount = 0
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression

function Resolve-ExistingFile {
    param([string]$Path, [string]$Label)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label does not exist: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

function Get-ProtectedHashes {
    param([string]$Path)

    $result = @{}
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Read, $false)
        try {
            foreach ($entry in $zip.Entries) {
                if ($entry.FullName -notmatch '^ppt/(slideMasters|slideLayouts|theme|media)/') {
                    continue
                }
                $hashAlgorithm = [System.Security.Cryptography.SHA256]::Create()
                try {
                    $entryStream = $entry.Open()
                    try {
                        $hash = $hashAlgorithm.ComputeHash($entryStream)
                        $result[$entry.FullName] = [Convert]::ToHexString($hash)
                    }
                    finally {
                        $entryStream.Dispose()
                    }
                }
                finally {
                    $hashAlgorithm.Dispose()
                }
            }
        }
        finally {
            $zip.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
    return $result
}

function Read-ZipText {
    param([string]$Path, [string]$EntryName)

    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Read, $false)
        try {
            $entry = $zip.GetEntry($EntryName)
            if ($null -eq $entry) {
                throw "Missing OOXML part: $EntryName"
            }
            $entryStream = $entry.Open()
            try {
                $reader = [System.IO.StreamReader]::new($entryStream, [System.Text.Encoding]::UTF8, $true)
                try {
                    return $reader.ReadToEnd()
                }
                finally {
                    $reader.Dispose()
                }
            }
            finally {
                $entryStream.Dispose()
            }
        }
        finally {
            $zip.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

$resolvedPptx = Resolve-ExistingFile -Path $PptxPath -Label 'PPTX'
$resolvedReference = Resolve-ExistingFile -Path $ReferencePath -Label 'Reference PPTX'

$referenceHashes = Get-ProtectedHashes -Path $resolvedReference
$outputHashes = Get-ProtectedHashes -Path $resolvedPptx
$missingProtectedParts = @()
$changedProtectedParts = @()

foreach ($name in ($referenceHashes.Keys | Sort-Object)) {
    if (-not $outputHashes.ContainsKey($name)) {
        $missingProtectedParts += $name
    }
    elseif ($referenceHashes[$name] -ne $outputHashes[$name]) {
        $changedProtectedParts += $name
    }
}

$presentationXml = Read-ZipText -Path $resolvedPptx -EntryName 'ppt/presentation.xml'
$slideCount = ([regex]::Matches($presentationXml, '<p:sldId\b')).Count
$countMatches = ($ExpectedSlideCount -le 0 -or $slideCount -eq $ExpectedSlideCount)

$titleWarnings = @()
$stream = [System.IO.File]::Open($resolvedPptx, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
try {
    $zip = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Read, $false)
    try {
        $slideEntries = @($zip.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide\d+\.xml$' } | Sort-Object FullName)
        foreach ($entry in $slideEntries) {
            $entryStream = $entry.Open()
            try {
                $reader = [System.IO.StreamReader]::new($entryStream, [System.Text.Encoding]::UTF8, $true)
                try {
                    [xml]$slideXml = $reader.ReadToEnd()
                }
                finally {
                    $reader.Dispose()
                }
            }
            finally {
                $entryStream.Dispose()
            }

            $titleShapes = @($slideXml.SelectNodes("//*[local-name()='sp'][.//*[local-name()='ph' and (@type='title' or @type='ctrTitle')]]"))
            foreach ($shape in $titleShapes) {
                $spPr = $shape.SelectSingleNode("./*[local-name()='spPr']")
                $lstStyle = $shape.SelectSingleNode("./*[local-name()='txBody']/*[local-name()='lstStyle']")
                $hasLocalGeometry = ($null -ne $spPr -and $spPr.HasChildNodes)
                $hasLocalListStyle = ($null -ne $lstStyle -and $lstStyle.HasChildNodes)
                if ($hasLocalGeometry -or $hasLocalListStyle) {
                    $titleWarnings += [ordered]@{
                        slide = $entry.FullName
                        local_geometry = $hasLocalGeometry
                        local_list_style = $hasLocalListStyle
                    }
                }
            }
        }
    }
    finally {
        $zip.Dispose()
    }
}
finally {
    $stream.Dispose()
}

$passed = ($missingProtectedParts.Count -eq 0 -and $changedProtectedParts.Count -eq 0 -and $countMatches)
$report = [ordered]@{
    passed = $passed
    pptx = $resolvedPptx
    reference = $resolvedReference
    slide_count = $slideCount
    expected_slide_count = if ($ExpectedSlideCount -gt 0) { $ExpectedSlideCount } else { $null }
    protected_part_count = $referenceHashes.Count
    missing_protected_parts = $missingProtectedParts
    changed_protected_parts = $changedProtectedParts
    title_placeholder_warnings = $titleWarnings
}

$report | ConvertTo-Json -Depth 6
if (-not $passed) {
    exit 1
}
