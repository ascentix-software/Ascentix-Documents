[CmdletBinding()]
param(
    [ValidateSet('Development','Official')][string]$SigningMode = 'Development',
    [string]$SigningKeyPath,
    [string]$ExpectedPublicKeyToken
)
$ErrorActionPreference = 'Stop'
$productRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
function Invoke-Checked([string]$program, [string[]]$arguments) { & $program @arguments; if ($LASTEXITCODE -ne 0) { throw "$program failed with exit code $LASTEXITCODE" } }
if ($SigningMode -eq 'Official') {
    if ([string]::IsNullOrWhiteSpace($SigningKeyPath) -or -not (Test-Path -LiteralPath $SigningKeyPath -PathType Leaf) -or $ExpectedPublicKeyToken -notmatch '\A[0-9a-fA-F]{16}\z') {
        throw 'Official builds require an explicit protected signing key and the independently configured expected public key token. No development-key fallback is permitted.'
    }
} elseif ([string]::IsNullOrWhiteSpace($SigningKeyPath)) {
    $SigningKeyPath = Join-Path $productRoot '.local/development.snk'
    if (-not (Test-Path -LiteralPath $SigningKeyPath)) {
        [IO.Directory]::CreateDirectory((Split-Path $SigningKeyPath)) | Out-Null
        $signer = [Security.Cryptography.RSACryptoServiceProvider]::new(2048)
        try { $signer.PersistKeyInCsp = $false; [IO.File]::WriteAllBytes($SigningKeyPath, $signer.ExportCspBlob($true)) }
        finally { $signer.Dispose() }
    }
}
$SigningKeyPath = (Resolve-Path -LiteralPath $SigningKeyPath).Path
& (Join-Path $PSScriptRoot '../tests/build/test-source-boundary.ps1')
& (Join-Path $PSScriptRoot 'verify-source-boundary.ps1')
$property = '-p:DocumentsSigningKeyPath=' + $SigningKeyPath
Invoke-Checked 'dotnet' @('build',(Join-Path $productRoot 'Ascentix.Documents.sln'),'-c','Release',$property)
$assemblyPath = Join-Path $productRoot 'src/Ascentix.Documents.Plugins/bin/Release/net462/Ascentix.Documents.Plugins.dll'
$identity = [Reflection.AssemblyName]::GetAssemblyName($assemblyPath)
$token = ([BitConverter]::ToString($identity.GetPublicKeyToken())).Replace('-','').ToLowerInvariant()
if ($SigningMode -eq 'Official' -and $token -ne $ExpectedPublicKeyToken.ToLowerInvariant()) { throw 'Built assembly does not match the approved official public key token.' }
Invoke-Checked 'dotnet' @('test',(Join-Path $productRoot 'Ascentix.Documents.sln'),'-c','Release','--no-build',$property,'--logger','trx;LogFileName=documents.trx')
$output = Join-Path $productRoot 'artifacts/plugins'
[IO.Directory]::CreateDirectory($output) | Out-Null
Copy-Item -LiteralPath $assemblyPath -Destination (Join-Path $output 'Ascentix.Documents.Plugins.dll')
[ordered]@{
    signingMode=$SigningMode; assemblyName=$identity.Name; publicKeyToken=$token
    sha256=(Get-FileHash -LiteralPath $assemblyPath -Algorithm SHA256).Hash
    releaseAuthenticity='Strong-name identity only; release artifact signature is a separate release gate.'
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'build-identity.json')
Write-Output "Plugin artifact built with $SigningMode signing. No solution packaging or deployment performed."
