[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$productRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
function Invoke-Checked([string[]]$arguments) { & node @arguments; if($LASTEXITCODE -ne 0){throw 'Web resource validation failed'} }
foreach($name in @('shell.js','admin.js','sites-access.js','operations.js')){Invoke-Checked @('--check',(Join-Path $productRoot "client/admin/$name"))}
foreach($name in @('verify-admin.cjs','test-fake-dom.cjs','test-shell.cjs','test-admin-state.cjs','test-sites-access.cjs','test-monitor-settings.cjs')){Invoke-Checked @((Join-Path $productRoot ('tests/web/' + $name)))}
$output=Join-Path $productRoot 'artifacts/webresources'
[IO.Directory]::CreateDirectory($output) | Out-Null
$files=@('index.html','admin.css','shell.js','admin.js','sites-access.js','operations.js')
Compress-Archive -LiteralPath @($files | ForEach-Object {Join-Path $productRoot "client/admin/$_"}) -DestinationPath (Join-Path $output 'Documents.Admin.zip') -Force
$manifest=foreach($name in $files){[ordered]@{name=$name;sha256=(Get-FileHash -LiteralPath (Join-Path $productRoot "client/admin/$name") -Algorithm SHA256).Hash}}
@($manifest) | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'manifest.json')
Write-Output 'Web resource artifact built. No .NET SDK, signing key, Power Platform CLI or environment connection used.'
