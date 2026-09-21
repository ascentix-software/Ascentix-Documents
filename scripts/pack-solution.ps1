[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PluginAssemblyPath,
    [ValidateSet('Development','Official')][string]$SigningMode = 'Official'
)
$ErrorActionPreference = 'Stop'
$productRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$source = Join-Path $productRoot 'solution/AscentixDocuments/src'
$PluginAssemblyPath = (Resolve-Path -LiteralPath $PluginAssemblyPath).Path
$identity = [Reflection.AssemblyName]::GetAssemblyName($PluginAssemblyPath)
$metadata = @(Get-ChildItem -LiteralPath (Join-Path $source 'PluginAssemblies') -Filter '*.dll.data.xml' -Recurse)
if ($metadata.Count -ne 1) { throw 'Expected one solution plugin assembly.' }
[xml]$plugin = Get-Content -LiteralPath $metadata[0].FullName -Raw
$originalIdentity = $plugin.PluginAssembly.FullName
if ($identity.Name -ne 'Ascentix.Documents.Plugins' -or $identity.GetPublicKeyToken().Length -eq 0) { throw 'Expected the signed Documents plugin assembly.' }
if ($SigningMode -eq 'Official' -and $identity.FullName -ne $originalIdentity) { throw 'Official assembly identity differs from the solution. A development build cannot replace an official assembly.' }
[xml]$manifest = Get-Content -LiteralPath (Join-Path $source 'Other/Solution.xml') -Raw
$solution = $manifest.ImportExportXml.SolutionManifest
if ($solution.UniqueName -ne 'AscentixDocuments' -or $solution.Publisher.UniqueName -ne 'ascentix' -or $solution.Publisher.CustomizationPrefix -ne 'asx') { throw 'Unexpected solution or publisher identity.' }
$version = [string]$solution.Version
if ($version -notmatch '^\d+\.\d+\.\d+\.\d+$') { throw 'Invalid solution version.' }
$output = Join-Path $productRoot ('artifacts/solutions/' + $SigningMode.ToLowerInvariant())
# Solution Packager uses .NET Framework path limits; nested checkouts need a short staging path.
$stage = Join-Path ([IO.Path]::GetTempPath()) ('documents-pack-' + [Guid]::NewGuid().ToString('N').Substring(0,12))
[IO.Directory]::CreateDirectory($stage) | Out-Null
[IO.Directory]::CreateDirectory($output) | Out-Null
Get-ChildItem -LiteralPath $source -Force | Copy-Item -Destination $stage -Recurse
$relativeMetadata = $metadata[0].FullName.Substring($source.TrimEnd('\','/').Length + 1)
$stagedMetadata = Join-Path $stage $relativeMetadata
$assemblyRelativePath = $plugin.PluginAssembly.FileName.TrimStart('/','\')
Copy-Item -LiteralPath $PluginAssemblyPath -Destination (Join-Path $stage $assemblyRelativePath) -Force
if ($SigningMode -eq 'Development') {
    $text = [IO.File]::ReadAllText($stagedMetadata).Replace($originalIdentity, $identity.FullName)
    [IO.File]::WriteAllText($stagedMetadata, $text, [Text.UTF8Encoding]::new($false))
}
$assets = @('index.html','admin.css','admin.js','sites-access.js')
foreach ($name in $assets) {
    Copy-Item -LiteralPath (Join-Path $productRoot "client/admin/$name") -Destination (Join-Path $stage "WebResources/asx_admin/$name") -Force
}
& node (Join-Path $PSScriptRoot 'verify-flows.cjs') (Join-Path $stage 'Workflows')
if ($LASTEXITCODE -ne 0) { throw 'Solution flow validation failed.' }
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Read-Entry([IO.Compression.ZipArchiveEntry]$Entry) {
    $reader = [IO.StreamReader]::new($Entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}
function Entry-Hash([IO.Compression.ZipArchiveEntry]$Entry) {
    $stream = $Entry.Open(); $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-','') }
    finally { $stream.Dispose(); $sha.Dispose() }
}
$packages = foreach ($type in @('Unmanaged','Managed')) {
    $suffix = if ($SigningMode -eq 'Development') { '_development' } else { '' }
    $file = 'AscentixDocuments_' + $version + $suffix + '_' + $type.ToLowerInvariant() + '.zip'
    $zipPath = Join-Path $output $file
    & pac solution pack --zipfile $zipPath --folder $stage --packagetype $type --errorlevel Warning | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "Solution packing failed: $type" }
    $archive = [IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        [xml]$packed = Read-Entry $archive.GetEntry('solution.xml')
        $expectedManaged = if ($type -eq 'Managed') { '1' } else { '0' }
        if ($packed.ImportExportXml.SolutionManifest.Version -ne $version -or [string]$packed.ImportExportXml.SolutionManifest.Managed -ne $expectedManaged) { throw 'Packed solution version/type mismatch.' }
        $dlls = @($archive.Entries | Where-Object { $_.FullName -like '*.dll' })
        if ($dlls.Count -ne 1 -or (Entry-Hash $dlls[0]) -ne (Get-FileHash -LiteralPath $PluginAssemblyPath).Hash) { throw 'Packed plugin differs from supplied assembly.' }
        foreach ($name in $assets) {
            [xml]$resource = Get-Content -LiteralPath (Join-Path $source "WebResources/asx_admin/$name.data.xml") -Raw
            $entry = $archive.GetEntry($resource.WebResource.FileName.TrimStart('/'))
            if ($null -eq $entry -or (Entry-Hash $entry) -ne (Get-FileHash -LiteralPath (Join-Path $productRoot "client/admin/$name")).Hash) { throw "Packed web resource differs: $name" }
        }
        foreach ($flow in Get-ChildItem -LiteralPath (Join-Path $source 'Workflows') -Filter '*.json') {
            $entry = $archive.GetEntry('Workflows/' + $flow.Name)
            if ($null -eq $entry -or (Entry-Hash $entry) -ne (Get-FileHash -LiteralPath $flow.FullName).Hash) { throw "Packed flow differs: $($flow.Name)" }
        }
    } finally { $archive.Dispose() }
    [ordered]@{ file=$file; sha256=(Get-FileHash -LiteralPath $zipPath).Hash }
}
[ordered]@{ version=$version; signingMode=$SigningMode; pluginIdentity=$identity.FullName; pluginSha256=(Get-FileHash -LiteralPath $PluginAssemblyPath).Hash; packages=@($packages) } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'build-manifest.json')
$packages | ForEach-Object { $_.sha256.ToLowerInvariant() + '  ' + $_.file } | Set-Content -LiteralPath (Join-Path $output 'SHA256SUMS')
Write-Output "Verified $SigningMode managed and unmanaged solution ZIPs in $output. No deployment performed."
