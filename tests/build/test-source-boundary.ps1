[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$verify = Join-Path $PSScriptRoot '../../scripts/verify-source-boundary.ps1'
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('documents-boundary-' + [Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($fixture) | Out-Null
$project = Join-Path $fixture 'Fixture.csproj'

function Assert-Boundary([string]$Definition, [string]$ExpectedError = '') {
    [IO.File]::WriteAllText($project, $Definition)
    $failure = $null
    try { & $verify -ProjectRoot $fixture | Out-Null }
    catch { $failure = $_.Exception.Message }

    if ($ExpectedError) {
        if (-not $failure -or $failure -notlike "*$ExpectedError*") {
            throw "Expected '$ExpectedError', received '$failure'."
        }
    } elseif ($failure) {
        throw $failure
    }
}

try {
    Assert-Boundary @'
<Project>
  <PropertyGroup>
    <WebResourceDirectory Condition="'$(WebResourceDirectory)' == ''">$(MSBuildThisFileDirectory)client/admin</WebResourceDirectory>
  </PropertyGroup>
  <ItemGroup>
    <Compile Include="One.cs;Two.cs" />
    <EmbeddedResource Include="$(WebResourceDirectory)/index.html" />
  </ItemGroup>
</Project>
'@
    Assert-Boundary '<Project><ItemGroup><Compile Include="One.cs;../Outside.cs" /></ItemGroup></Project>' 'External source reference'
    Assert-Boundary '<Project><ItemGroup><Compile Include="$(Unknown)/One.cs" /></ItemGroup></Project>' 'Review dynamic source reference'
    Assert-Boundary @'
<Project>
  <PropertyGroup>
    <WebResourceDirectory Condition="'$(WebResourceDirectory)' == ''">$(MSBuildThisFileDirectory)../outside</WebResourceDirectory>
  </PropertyGroup>
  <ItemGroup><EmbeddedResource Include="$(WebResourceDirectory)/index.html" /></ItemGroup>
</Project>
'@ 'External source reference'
    Assert-Boundary '<Project><ItemGroup><EmbeddedResource Include="$(WebResourceDirectory)/index.html" /></ItemGroup></Project>' 'Ambiguous web-resource source directory'
    Write-Output 'PASS source boundaries: default resources, path lists, external paths, and unresolved properties.'
} finally {
    Remove-Item -LiteralPath $project -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $fixture
}
