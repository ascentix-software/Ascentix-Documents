# Contributing to Ascentix Documents

Local builds require no company signing key or live Dataverse/SharePoint environment.

## Build and test

Run from PowerShell with .NET SDK 10.0.302, .NET Framework 4.6.2 support, Node.js, and Power Platform CLI available:

```powershell
./scripts/build.ps1
```

This builds/tests the plug-in, checks web resources and solution flows, and produces managed and unmanaged development ZIPs in `artifacts/solutions/development`. Local builds generate `.local/development.snk`; retain it for a stable development assembly identity. Development packages are for separate sandboxes and cannot upgrade officially signed installations.

For independent component checks:

```powershell
./scripts/build-plugins.ps1
./scripts/build-webresources.ps1
node scripts/verify-flows.cjs
```

Plug-in and web-resource artifacts are written to `artifacts/plugins` and `artifacts/webresources`. Runtime tests live under `tests/Ascentix.Documents.*`; UI handler and optional browser tests live under `tests/web`.

## Solution source and release packages

`solution/AscentixDocuments/src` is the complete unpacked solution: schema, roles, app, plug-in/API registrations, web resources, and Power Automate flows. Edit flow definitions there; there is no separate worker generator or bootstrap installer.

The packer copies the solution to a temporary build directory, inserts the compiled plug-in and current `client/admin` and `client/form` resources, then creates and verifies both ZIPs. It does not modify tracked solution files. Compiled plug-in DLLs are build outputs and are not tracked.

Maintainers with the official signing key can run:

```powershell
./scripts/build.ps1 -SigningMode Official -SigningKeyPath <protected-key-path> -ExpectedPublicKeyToken 50214c64469649b5
```

To package an already-built official plug-in:

```powershell
./scripts/pack-solution.ps1 -PluginAssemblyPath <official-plugin-dll> -SigningMode Official
```

Official ZIPs and checksums are written to `artifacts/solutions/official`. The packer verifies the assembly identity against the solution metadata and checks the packaged DLL, web resources, and flows against its inputs. Packaging alone does not establish successful installation; validate the resulting managed ZIP in the target environment.

Azure Pipelines definitions under `pipelines/` cover independent component CI and protected-key plug-in builds. See [builds and signing](docs/pipelines-and-signing.md). Install using the [managed installation guide](docs/customer-installation.md).

## Code readability

Use descriptive names and separate validation, data loading, state changes, and result construction with blank lines. Expand multi-step branches and nested conditionals. Keep comments factual: describe a function's contract or a complicated block's behavior. Keep design rationale and historical notes in documentation. Do not add comments that repeat an assignment or a self-explanatory function name.

Write C# method descriptions as `///` XML documentation with a `<summary>` and informative `<param>` entries. Include `<typeparam>` and `<returns>` when applicable. Keep explanations inside complicated method bodies as ordinary `//` comments. XML method documentation is available in IntelliSense when navigating source references.

The repository pins CSharpier for C# and project XML, and Prettier for web resources and JavaScript tools. Install the development tools once:

```powershell
dotnet tool restore
npm.cmd ci
```

Format code before review:

```powershell
dotnet csharpier format src tests Directory.Build.props
npm.cmd run format
```

Check formatting without modifying files:

```powershell
dotnet csharpier check src tests Directory.Build.props
npm.cmd run format:check
```

The web formatter covers `client`, JavaScript checks under `scripts`, and `tests/web`. Exported solution files and architecture documents are outside the formatting scope. Formatting does not build, test, or deploy the product; run the relevant component build afterward.
