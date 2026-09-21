# Ascentix Documents

Ascentix Documents provisions SharePoint folder structures from Dataverse records and manages library access through Dataverse teams. Define folder templates, publish them, and let the worker apply the resulting folder and access changes. It runs independently of Ascentix Rules Engine.

**Version 0.1.0.3 is a public preview.** Read the [release notes](docs/public-preview.md) for features and known limitations. The managed solution is the intended evaluation download. Source is licensed under [Apache 2.0](LICENSE).

## What it does

- Creates conditional folder structures with names drawn from record data and supported lookups.
- Supports multiple document destinations and a draft, preview, and publication workflow for templates.
- Reuses existing managed folders when work is replayed or replanned.
- Applies library Read and Contribute policies for registered Dataverse teams and reconciles membership changes.
- Provides work inspection, retry, cancellation, and replan operations.
- Offers a system-wide record-update switch, Off by default. Disabling it disables the Update and UpdateMultiple plugin steps.
- Retains document folders when their business record is deleted and marks the record for decommission review.

## Evaluate

Use a dedicated Dataverse and SharePoint evaluation environment. Follow the [managed installation guide](docs/customer-installation.md), including the application identity, connector connections, worker roles, event registrations, and disposable-record check. Credentials stay in the platform.

The [operations guide](docs/operations.md) covers work recovery, access changes, upgrades, and retention. Read [record-update processing](docs/record-update-processing.md) before enabling update monitoring.

This preview has bounded live validation. Sustained high-volume capacity, advanced recovery, and a fresh installation of the corrected package remain unverified. See the release notes for the remaining limitations.

## Build from source

Prerequisites: .NET SDK 10.0.302 or compatible, .NET Framework 4.6.2 reference/runtime support, Node.js, and Power Platform CLI for the complete package build. Run from PowerShell:

```powershell
./scripts/build.ps1
```

The build runs the .NET tests, web-resource and solution-flow checks, source-boundary checks, and produces managed and unmanaged development solution ZIPs under `artifacts/solutions/development`. It does not deploy to an environment. Local builds use a generated development signing key; official packages use a separate signing identity. A development-signed DLL cannot replace an officially signed assembly in place.

See [CONTRIBUTING.md](CONTRIBUTING.md) for independent component builds and formatting. The complete unpacked solution is under `solution/AscentixDocuments/src`. Customer evaluation uses the managed release package.

## Understand the system

- [Illustrated architecture guide](docs/architecture-guide.html) and its [Markdown source](docs/architecture-guide.md)
- [Worker setup](docs/worker-installation.md)
- [Documentation index](docs/README.md)
- [Signing and pipelines](docs/pipelines-and-signing.md)
- [Source attribution and notices](NOTICE)
