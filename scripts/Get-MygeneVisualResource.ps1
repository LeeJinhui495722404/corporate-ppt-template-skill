param(
 [ValidateSet('method','icon','all')][string]$Kind='method',
 [string]$Id='',
 [string]$Query='',
 [string]$Category='',
 [string]$Style='',
 [switch]$EditableOnly,
 [switch]$IncludeBrands,
 [ValidateRange(1,500)][int]$Limit=16,
 [string]$IndexPath=(Join-Path $PSScriptRoot '..\references\visual-resources.json')
)
$ErrorActionPreference='Stop'
$index=Get-Content -Raw -Encoding UTF8 -LiteralPath $IndexPath|ConvertFrom-Json
$items=@();if($Kind -in @('method','all')){$items+=@($index.Methods)};if($Kind -in @('icon','all')){$items+=@($index.Icons)}
$matches=@($items|Where-Object{
 ($IncludeBrands -or !$_.Brand) -and
 (!$Id -or $_.Id -eq $Id) -and
 (!$Category -or $_.Category -eq $Category) -and
 (!$Style -or $_.Style -eq $Style) -and
 (!$EditableOnly -or ($_.Kind -eq 'icon' -and $_.EditableType -eq '可编辑形状')) -and
 (!$Query -or (($_.Name,$_.Category,$_.Inputs,$_.Rule,$_.Id -join ' ').IndexOf($Query,[StringComparison]::OrdinalIgnoreCase) -ge 0))
})
if($Kind -eq 'icon'){
 $matches=@($matches|Sort-Object @{Expression={if($_.EditableType -eq '可编辑形状'){0}else{1}}},Id)
}
[ordered]@{Reference=$index.Reference;MatchedCount=$matches.Count;ReturnedCount=[math]::Min($Limit,$matches.Count);Resources=@($matches|Select-Object -First $Limit)}|ConvertTo-Json -Depth 12
