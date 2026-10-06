# Public preview — 0.1.0.4

Ascentix Documents provisions SharePoint folders from Dataverse records and manages library access through Dataverse teams. Source is available under Apache 2.0.

## Included in this preview

- Conditional folder templates with draft, preview, and publication workflows.
- Multiple document destinations and reuse of existing managed folders during replay or replanning.
- Team-based Read and Contribute library policies with membership reconciliation.
- Work inspection, retry, cancellation, and replanning.
- A system-wide record-update setting, Off by default, which disables the corresponding Update and UpdateMultiple plug-in steps when Off.
- A fix for registered-team deletion failing during the worker's retirement update.

## What is new in 0.1.0.4

- Capture is asynchronous and cannot block a save. A failed capture leaves the save intact and appears in a recovery list.
- Event registration is managed in the application. Runtime administration registers and verifies the event steps.
- There is no table-count limit. Enable tables in the template workspace.
- Temporary failures retry automatically, with a visible next-attempt time.
- Admin stops always work: pause, suspend, approve, and Remove take effect while folder jobs run.
- B2B guests and Entra or Microsoft 365 group teams are supported.
- Re-point follows a moved or renamed library. Remove retires a destination that no template uses.
- Blocked records and failed capture jobs can be retried or replanned.

To upgrade from 0.1.0.3, follow the [upgrade notes](upgrade-0.1.0.4.md). The upgrade has a breaking change: the `asx_PublishTemplate` result is now JSON.

## Installation

Use `AscentixDocuments_0.1.0.4_managed.zip` for evaluation. The unmanaged solution and source archive are provided for development. Follow the [installation guide](customer-installation.md) to configure application identities, connections, worker roles, runtime, and tables.

Official packages use a stable assembly signing identity. Local builds use a development key and cannot replace an officially signed assembly in place. The release manifest and SHA-256 checksums identify the supplied artifacts; checksums are not a signed release attestation.

## Preview limitations

- Sustained high-volume capacity has not been established. Start with a small template and disposable records in an evaluation environment.
- Managed import and upgrade have been exercised; a fresh installation of this version in an isolated empty environment remains unverified.
- Uninstall, including removal of the application-created event steps, is not verified.
- Entra and Microsoft 365 group claims and guest sign-in are verified in a test environment, not across tenants. Verification in a test environment is pending for this release.
- Recovery from lost acknowledgements and ambiguous external writes, certificate rotation, and disaster recovery have not been rehearsed live.
- Large membership pagination and team deletion across multiple active library grants require further validation.
- Event registration changes can take time to propagate. Enable update monitoring during a quiet interval and verify a disposable write. Use explicit replanning to catch up on changes made while monitoring was Off.

See [operations](operations.md) for ongoing administration and [record-update processing](record-update-processing.md) for update behavior.
