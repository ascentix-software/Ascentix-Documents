# Ascentix Documents

Ascentix Documents provisions SharePoint folder structures from Dataverse records and manages library access through Dataverse teams. Define folder templates, publish them, and let the worker apply the resulting folder and access changes. It runs independently of Ascentix Rules Engine.

**Version 0.1.0.4 is a public preview.** Read the [release notes](docs/public-preview.md) for features and known limitations. The managed solution is the intended evaluation download. Source is licensed under [Apache 2.0](LICENSE).

## What it does

- Creates conditional folder structures with names drawn from record data and supported lookups.
- Supports multiple document destinations and a draft, preview, and publication workflow for templates.
- Reuses existing managed folders when work is replayed or re-run.
- Applies library Read and Contribute policies for registered Dataverse teams and reconciles membership changes.
- Provides work inspection in Monitor, with retry, cancellation, and re-runs.
- Offers a system-wide record-update switch, Off by default. Turning it off deactivates the Update event steps Documents registers.
- Retains document folders when their business record is deleted and marks the record for decommission review.

## New in 0.1.0.4

- No customer save can be blocked by Documents. Capture is asynchronous.
- No table-count limit. Event registration is managed in the application.
- Temporary failures retry automatically.
- Admin stops always work.
- B2B guests and Entra or Microsoft 365 group teams are supported (pending verification).
- Re-point and Remove for destinations.
- Monitor lists blocked jobs and library setups, records waiting for data, and jobs retrying automatically, with **Retry**, **Cancel job** or **Re-run**; a stuck access run has **Retry access run** and **Cancel access run** in Sites & access.
- Team access can be edited while it is being applied, and customized Read and Contribute permission levels are accepted.
- SharePoint hosts in every Microsoft cloud, including GCC High, DoD and 21Vianet (pending verification outside the worldwide cloud).

Upgrading from 0.1.0.3? Read the [upgrade notes](docs/upgrade-0.1.0.4.md).

## Evaluate

Use a dedicated Dataverse and SharePoint evaluation environment. Follow the [managed installation guide](docs/customer-installation.md), including the application identity, connector connections, worker roles, tables, and disposable-record check. Credentials stay in the platform.

The [operations guide](docs/operations.md) covers work recovery, access changes, upgrades, and retention. Read [record-update processing](docs/record-update-processing.md) before enabling update monitoring.

This preview has bounded live validation. Sustained high-volume capacity, advanced recovery, a fresh installation, and uninstall remain unverified. Entra and Microsoft 365 group teams and B2B guest access are pending verification. See the release notes for the remaining limitations.

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
