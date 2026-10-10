param(
 [Parameter(Mandatory=$true)][int[]]$Page,
 [Parameter(Mandatory=$true)][string]$OutputPath,
 [string]$ReferencePath=(Join-Path $PSScriptRoot '..\assets\reference.pptx')
)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$source=(Resolve-Path -LiteralPath $ReferencePath).Path
$output=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $output){throw 'Output exists; choose a new draft path.'}
if(($Page|Sort-Object -Unique).Count -ne $Page.Count){throw 'Repeated pages require cloning XML parts; use unique pages with this selector.'}
[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($output))|Out-Null
$zip=[IO.Compression.ZipFile]::OpenRead($source)
try{$reader=[IO.StreamReader]::new($zip.GetEntry('ppt/presentation.xml').Open(),[Text.Encoding]::UTF8,$true);try{[xml]$presentation=$reader.ReadToEnd()}finally{$reader.Dispose()}}finally{$zip.Dispose()}
$p='http://schemas.openxmlformats.org/presentationml/2006/main';$r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$ns=[Xml.XmlNamespaceManager]::new($presentation.NameTable);$ns.AddNamespace('p',$p)
$list=$presentation.SelectSingleNode('/p:presentation/p:sldIdLst',$ns);$slides=@($list.ChildNodes)
foreach($number in $Page){if($number -lt 1 -or $number -gt $slides.Count){throw "Invalid reference page $number"}}
$selected=@($Page|ForEach-Object{$slides[$_-1].CloneNode($true)});$list.RemoveAll();foreach($sid in $selected){[void]$list.AppendChild($sid)}
# Resource sections belong to the reference, not to the generated report.
foreach($sections in @($presentation.SelectNodes("//*[local-name()='sectionLst']"))){$parent=$sections.ParentNode;[void]$parent.RemoveChild($sections);if(!$parent.HasChildNodes){[void]$parent.ParentNode.RemoveChild($parent)}}
Copy-Item -LiteralPath $source -Destination $output
$zip=[IO.Compression.ZipFile]::Open($output,[IO.Compression.ZipArchiveMode]::Update)
function WriteXml($name,$xml){$zip.GetEntry($name).Delete();$entry=$zip.CreateEntry($name);$writer=[IO.StreamWriter]::new($entry.Open(),[Text.UTF8Encoding]::new($false));try{$writer.Write($xml.OuterXml)}finally{$writer.Dispose()}}
try{
 WriteXml 'ppt/presentation.xml' $presentation
 $reader=[IO.StreamReader]::new($zip.GetEntry('ppt/_rels/presentation.xml.rels').Open(),[Text.Encoding]::UTF8,$true);try{[xml]$rels=$reader.ReadToEnd()}finally{$reader.Dispose()}
 $ids=@($selected|ForEach-Object{$_.GetAttribute('id',$r)})
 foreach($rel in @($rels.DocumentElement.ChildNodes)){if($rel.Type -match '/slide$' -and $rel.Id -notin $ids){[void]$rels.DocumentElement.RemoveChild($rel)}}
 WriteXml 'ppt/_rels/presentation.xml.rels' $rels
 $reader=[IO.StreamReader]::new($zip.GetEntry('docProps/app.xml').Open(),[Text.Encoding]::UTF8,$true);try{[xml]$app=$reader.ReadToEnd()}finally{$reader.Dispose()}
 $app.SelectSingleNode("//*[local-name()='Slides']").InnerText=[string]$Page.Count;WriteXml 'docProps/app.xml' $app
}finally{$zip.Dispose()}
[ordered]@{Draft=$output;SourcePages=@($Page);SlideCount=$Page.Count;SourceChanged=$false}|ConvertTo-Json -Depth 4
