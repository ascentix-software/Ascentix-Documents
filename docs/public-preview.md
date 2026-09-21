# Public preview — 0.1.0.3

Ascentix Documents provisions SharePoint folders from Dataverse records and manages library access through Dataverse teams. Source is available under Apache 2.0.

## Included in this preview

- Conditional folder templates with draft, preview, and publication workflows.
- Multiple document destinations and reuse of existing managed folders during replay or replanning.
- Team-based Read and Contribute library policies with membership reconciliation.
- Work inspection, retry, cancellation, and replanning.
- A system-wide record-update setting, Off by default, which disables the corresponding Update and UpdateMultiple plug-in steps when Off.
- A fix for registered-team deletion failing during the worker's retirement update.

## Installation

Use `AscentixDocuments_0.1.0.3_managed.zip` for evaluation. The unmanaged solution and source archive are provided for development. Follow the [installation guide](customer-installation.md) to configure application identities, connections, worker roles, runtime, and event registrations.

Official packages use a stable assembly signing identity. Local builds use a development key and cannot replace an officially signed assembly in place. The release manifest and SHA-256 checksums identify the supplied artifacts; checksums are not a signed release attestation.

## Preview limitations

- Sustained high-volume capacity has not been established. Start with a small template and disposable records in an evaluation environment.
- Managed import and upgrade have been exercised; a fresh installation of this version in an isolated empty environment remains unverified.
- Recovery from lost acknowledgements and ambiguous external writes, certificate rotation, and disaster recovery have not been rehearsed live.
- Large membership pagination and team deletion across multiple active library grants require further validation.
- Changes to event registrations can take time to propagate. Enable update monitoring during a quiet interval and verify a disposable write. Use explicit replanning to catch up on changes made while monitoring was Off.

See [operations](operations.md) for ongoing administration and [record-update processing](record-update-processing.md) for update behavior.
