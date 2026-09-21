[CmdletBinding()]
param([string]$ProjectRoot = (Join-Path $PSScriptRoot '..'))
$ErrorActionPreference = 'Stop'
$productRoot = [IO.Path]::GetFullPath($ProjectRoot)
$boundary = $productRoot.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
$checked = 0
$projects = @(Get-ChildItem -LiteralPath $productRoot -Filter *.csproj -Recurse | Where-Object { $_.FullName.Substring($boundary.Length) -notmatch '(^|[\\/])(bin|obj|artifacts|node_modules)[\\/]' })
if ($projects.Count -eq 0) { throw 'No Documents projects found; boundary verification cannot pass.' }
foreach ($project in $projects) {
    [xml]$definition = Get-Content -LiteralPath $project.FullName -Raw
    foreach ($node in $definition.SelectNodes('//ProjectReference|//Compile|//Content|//EmbeddedResource|//PdSolution|//HintPath|//AssemblyOriginatorKeyFile')) {
        $value = if ($node.HasAttribute('Include')) { $node.GetAttribute('Include') } else { $node.InnerText }
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        if ($node.Name -eq 'AssemblyOriginatorKeyFile' -and $value -eq '$(DocumentsSigningKeyPath)') {
            continue
        }

        # Resolves the repository's default web-resource directory before checking each resource path.
        if ($value.Contains('$(WebResourceDirectory)')) {
            $defaults = @($definition.SelectNodes('//PropertyGroup/WebResourceDirectory'))
            if ($defaults.Count -ne 1 -or $defaults[0].GetAttribute('Condition') -ne "'`$(WebResourceDirectory)' == ''") {
                throw "Ambiguous web-resource source directory in $($project.Name)"
            }
            $value = $value.Replace('$(WebResourceDirectory)', $defaults[0].InnerText)
        }
        $value = $value.Replace('$(MSBuildThisFileDirectory)', $project.DirectoryName + [IO.Path]::DirectorySeparatorChar)
        if ($value.Contains('$(')) {
            throw "Review dynamic source reference in $($project.Name): $value"
        }

        # MSBuild Include attributes may contain multiple semicolon-separated paths.
        foreach ($sourcePath in $value.Split(';')) {
            if ([string]::IsNullOrWhiteSpace($sourcePath)) { continue }
            $resolved = [IO.Path]::GetFullPath([IO.Path]::Combine($project.DirectoryName, $sourcePath))
            if (-not $resolved.StartsWith($boundary, [StringComparison]::OrdinalIgnoreCase)) {
                throw "External source reference in $($project.Name): $sourcePath"
            }
            $checked++
        }
    }
    foreach ($reference in $definition.SelectNodes('//PackageReference|//Reference')) {
        if ($reference.GetAttribute('Include') -match '(?i)Ascentix[.-]RulesEngine|Ascentix[.-]SharePointDemo') { throw "Rules Engine dependency in $($project.Name)" }
    }
}
Write-Output "PASS: $($projects.Count) projects and $checked explicit source, resource and package-asset references stay inside Documents; no Rules Engine assembly/package references."
