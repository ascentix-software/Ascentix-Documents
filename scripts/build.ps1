[CmdletBinding()]
param(
    [ValidateSet('Development','Official')][string]$SigningMode = 'Development',
    [string]$SigningKeyPath,
    [string]$ExpectedPublicKeyToken
)
$ErrorActionPreference = 'Stop'
$productRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
& (Join-Path $PSScriptRoot 'build-plugins.ps1') -SigningMode $SigningMode -SigningKeyPath $SigningKeyPath -ExpectedPublicKeyToken $ExpectedPublicKeyToken
& (Join-Path $PSScriptRoot 'build-webresources.ps1')
& (Join-Path $PSScriptRoot 'pack-solution.ps1') -PluginAssemblyPath (Join-Path $productRoot 'artifacts/plugins/Ascentix.Documents.Plugins.dll') -SigningMode $SigningMode
