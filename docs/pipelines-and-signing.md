# Builds and signing

## Contributor builds

The [contributor guide](../CONTRIBUTING.md) describes local builds and tests. Plug-in and web-resource builds can run independently. The complete build packs and verifies managed and unmanaged ZIPs from the unpacked solution. See the contributor guide for development and official builds.

`build-plugins.ps1` and `build.ps1` default to `SigningMode=Development`. They create or reuse `.local/development.snk` when no key is supplied. No company key or environment credentials are required.

## Official assembly identity

`SigningMode=Official` requires an existing `SigningKeyPath` and an `ExpectedPublicKeyToken`. Missing inputs or a mismatched compiled identity fail the build. There is no fallback to development signing.

Official packages use public key token `50214c64469649b5`. Keep the signing identity stable across upgrades: a development-signed assembly has a different identity and cannot replace an officially signed installed assembly in place. Private signing keys are not included in this repository.

## Pipeline source

The `pipelines/` directory contains Azure Pipelines definitions:

| Definition | Purpose |
| --- | --- |
| `plugins-ci.yml` | Build and test plug-ins using a generated development key |
| `webresources-ci.yml` | Check and package the administration web resources |
| `plugins-official.yml` | Build plug-ins using a separately configured protected signing key |

Official signing requires separately configured maintainer infrastructure. Use the managed package and [installation guide](customer-installation.md) for evaluation. Fork builds should use contributor CI without protected keys or deployment credentials.

## Download verification

Compare downloaded artifacts with the release's `SHA256SUMS`, for example using PowerShell `Get-FileHash -Algorithm SHA256`. The release manifest also records the source snapshot and assembly identity.

Checksums detect differences from the published files. A .NET strong name identifies the assembly; it does not authenticate the release publisher. This preview does not provide a signed release attestation. Authentication certificates used by Dataverse and SharePoint are separate from assembly signing.
