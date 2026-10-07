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
$assets = @('index.html','admin.css','shell.js','admin.js','sites-access.js','operations.js')
foreach ($name in $assets) {
    Copy-Item -LiteralPath (Join-Path $productRoot "client/admin/$name") -Destination (Join-Path $stage "WebResources/asx_admin/$name") -Force
}
& node (Join-Path $PSScriptRoot 'verify-flows.cjs') (Join-Path $stage 'Workflows')
if ($LASTEXITCODE -ne 0) { throw 'Solution flow validation failed.' }
foreach ($stepFile in Get-ChildItem -LiteralPath (Join-Path $stage 'SdkMessageProcessingSteps') -Filter '*.xml') {
    [xml]$step = Get-Content -LiteralPath $stepFile.FullName -Raw
    $node = $step.SdkMessageProcessingStep
    if ($node.Name -notlike 'Ascentix Documents: guard *' -or $node.ImpersonatingUserIdName) {
        throw "Package contract: only guard steps without impersonation may ship ($($node.Name))."
    }
    if ($node.PrimaryEntity -eq 'asx_runtime') {
        throw "Package contract: asx_runtime is protected by role privileges, not guard steps ($($node.Name))."
    }
}
foreach ($roleFile in Get-ChildItem -LiteralPath (Join-Path $stage 'Roles') -Filter '*.xml') {
    [xml]$role = Get-Content -LiteralPath $roleFile.FullName -Raw
    $runtimeWrites = @($role.Role.RolePrivileges.RolePrivilege | Where-Object { $_.name -match '^prv(Create|Write|Delete|Append|AppendTo)asx_runtime$' })
    if ($runtimeWrites.Count -ne 0) {
        throw "Package contract: only System Administrator may change asx_runtime ($($roleFile.Name): $($runtimeWrites.name -join ', '))."
    }
}
# The catalog and security APIs run as their caller, and the worker writes as itself, so each
# role must hold what its commands write. Depth: Basic < Local < Deep < Global.
$depths = @{ Basic = 1; Local = 2; Deep = 3; Global = 4 }
$needs = @(
    @('Documents Security Administrator', 'prvDeleteasx_library', 'Global', 'Remove deletes a library nothing refers to'),
    @('Documents Security Administrator', 'prvDeleteasx_site', 'Global', 'Remove deletes a site nothing refers to'),
    @('Documents Security Administrator', 'prvWriteasx_operation', 'Global', 'Remove and Apply access with the inheritance acknowledgement cancel an idle access run; Approve, Next libraries and site identity upgrade save their probe'),
    @('Documents Security Administrator', 'prvCreateasx_attempt', 'Global', 'Cancel releases an expired run claim and records it'),
    @('Documents Security Administrator', 'prvWriteasx_claim', 'Global', 'Suspend, Remove, Approve and Apply access save the site writer row'),
    @('Documents Security Administrator', 'prvCreateasx_claim', 'Global', 'Suspend, Remove, Approve and Apply access create the site writer row'),
    @('Documents Security Administrator', 'prvCreateasx_operation', 'Global', 'Add, Re-point and Create library queue their work'),
    @('Documents Security Administrator', 'prvWriteasx_library', 'Global', 'Suspend, Remove, Approve and Apply access update the library'),
    @('Documents Security Administrator', 'prvWriteasx_site', 'Global', 'Suspend, Remove and Approve update the site'),
    @('Documents Security Administrator', 'prvWriteasx_policy', 'Global', 'Save, Apply and Remove update the access policy'),
    @('Documents Security Administrator', 'prvWriteasx_policyentry', 'Global', 'Apply and Remove update the team references'),
    @('Documents Security Administrator', 'prvCreateSharePointDocumentLocation', 'Basic', 'Approve of an added library creates its Dataverse document location'),
    @('Documents Security Administrator', 'prvAppendSharePointDocumentLocation', 'Basic', 'that location is linked to its SharePoint site'),
    @('Documents Worker', 'prvWriteSharePointDocumentLocation', 'Global', 'Re-point updates the library document location chain, whoever created it'),
    @('Documents Worker', 'prvAppendSharePointDocumentLocation', 'Global', 'Re-point can move a location under its SharePoint site'),
    @('Documents Worker', 'prvWriteasx_operation', 'Global', 'every worker step saves its operation'),
    @('Documents Worker', 'prvCreateasx_attempt', 'Global', 'every worker step records its attempts')
)
foreach ($need in $needs) {
    $roleFile = Join-Path $stage ('Roles/' + $need[0] + '.xml')
    [xml]$role = Get-Content -LiteralPath $roleFile -Raw
    $held = @($role.Role.RolePrivileges.RolePrivilege | Where-Object { $_.name -eq $need[1] })
    if ($held.Count -ne 1 -or $depths[[string]$held[0].level] -lt $depths[$need[2]]) {
        throw "Package contract: $($need[0]) needs $($need[1]) at $($need[2]) or wider: $($need[3])."
    }
}
foreach ($flow in Get-ChildItem -LiteralPath (Join-Path $stage 'Workflows') -Filter '*.data.xml') {
    [xml]$data = Get-Content -LiteralPath $flow.FullName -Raw
    if ($data.Workflow.StateCode -ne '0') { throw "Package contract: flows must ship Off ($($flow.Name))." }
}
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
