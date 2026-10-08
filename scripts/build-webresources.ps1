[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$productRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
function Invoke-Checked([string[]]$arguments) { & node @arguments; if($LASTEXITCODE -ne 0){throw 'Web resource validation failed'} }
foreach($name in @('admin/shell.js','admin/admin.js','admin/sites-access.js','admin/operations.js','form/documents-tab.js')){Invoke-Checked @('--check',(Join-Path $productRoot "client/$name"))}
foreach($name in @('verify-admin.cjs','test-fake-dom.cjs','test-shell.cjs','test-admin-state.cjs','test-sites-access.cjs','test-monitor-settings.cjs','test-documents-tab.cjs')){Invoke-Checked @((Join-Path $productRoot ('tests/web/' + $name)))}
$output=Join-Path $productRoot 'artifacts/webresources'
[IO.Directory]::CreateDirectory($output) | Out-Null
$files=@('admin/index.html','admin/admin.css','admin/shell.js','admin/admin.js','admin/sites-access.js','admin/operations.js','form/documents-tab.js')
Compress-Archive -LiteralPath @($files | ForEach-Object {Join-Path $productRoot "client/$_"}) -DestinationPath (Join-Path $output 'Documents.Admin.zip') -Force
$manifest=foreach($name in $files){[ordered]@{name=(Split-Path $name -Leaf);sha256=(Get-FileHash -LiteralPath (Join-Path $productRoot "client/$name") -Algorithm SHA256).Hash}}
@($manifest) | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'manifest.json')
Write-Output 'Web resource artifact built. No .NET SDK, signing key, Power Platform CLI or environment connection used.'
