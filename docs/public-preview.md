# Public preview — 0.1.0.4

Ascentix Documents provisions SharePoint folders from Dataverse records and manages library access through Dataverse teams. Source is available under Apache 2.0.

## Included in this preview

- Conditional folder templates with draft, preview, and publication workflows.
- Multiple document destinations and reuse of existing managed folders during replay or a re-run.
- Team-based Read and Contribute library policies with membership reconciliation.
- Work inspection in Monitor, with retry, cancellation, and re-runs.
- A fix for registered-team deletion failing during the worker's retirement update.

## What is new in 0.1.0.4

- Capture is asynchronous and cannot block a save. A failed capture leaves the save intact and appears in Monitor under **Not captured**.
- Event registration (change tracking) is managed in the application. Settings registers and verifies the event steps.
- There is no table-count limit. Enable tables in Settings › Tables (**＋ Add table**).
- Temporary failures retry automatically, with a visible next-attempt time.
- Admin stops always work: turning automation off, suspending, approving, and Remove take effect while folder jobs run.
- A redesigned admin page. Folder templates opens on an overview and edits in three steps (Destinations, Folders, Review and publish) with autosave and a list of changes before you publish. Monitor shows every problem in one filtered list. Sites & access edits team access in a side panel. Settings holds automation, tables and change tracking in one place.
- Support for B2B guests, Entra and Microsoft 365 group teams, and government clouds is built in. Further testing is ongoing, as these are more complex areas.
- Re-point follows a moved or renamed library. Remove retires a destination that no template uses.
- Blocked records can be retried, and changes not captured can be re-run.
- A library whose access run stopped or waits shows **Needs attention** with **Retry access run** and **Cancel access run**, and a library setup that needs attention has **Retry** and **Cancel setup**. Blocked jobs have **Cancel job** next to **Retry**, and jobs waiting after a temporary error are listed in Monitor under **Retrying**.
- Folders that wait for a value, a usable name or a shorter path are listed in Monitor under **Waiting for data**, with **Re-run**. Folder paths up to SharePoint's 400-character limit are created, with a notice above 300 characters.
- Teams can be edited while their access is being applied; the newer change replaces the queued run or waits for it.
- Customized Read and Contribute permission levels are accepted unless they carry administrative rights.
- SharePoint hosts may be on any Microsoft cloud: worldwide and GCC, GCC High, DoD and 21Vianet. There is no host-count limit.
- Update folders when records change is not in this release. An environment that had it on stops capturing record updates after Repair all.

To upgrade from 0.1.0.3, follow the [upgrade notes](upgrade-0.1.0.4.md). The upgrade has a breaking change: the `asx_PublishTemplate` result is now JSON.

## Installation

Use `AscentixDocuments_0.1.0.4_managed.zip` for evaluation. The unmanaged solution and source archive are provided for development. Follow the [installation guide](customer-installation.md) to configure application identities, connections, worker roles, runtime, and tables.

Official packages use a stable assembly signing identity. Local builds use a development key and cannot replace an officially signed assembly in place. The release manifest and SHA-256 checksums identify the supplied artifacts; checksums are not a signed release attestation.

## Preview limitations

- Sustained high-volume capacity has not been established. Start with a small template and disposable records in an evaluation environment.
- Managed import and upgrade have been exercised; a fresh installation of this version in an isolated empty environment remains unverified.
- Uninstall, including removal of the application-created event steps, is not verified.
- Support for B2B guests, Entra and Microsoft 365 group teams, and government clouds (GCC High, DoD and 21Vianet) is built in. Further testing is ongoing, as these are more complex areas.
- Recovery from lost acknowledgements and ambiguous external writes, certificate rotation, and disaster recovery have not been rehearsed live.
- Large membership pagination and team deletion across multiple active library grants require further validation. A team with more people than one access run can store (roughly 2,000) keeps its group members as they are, with a notice; use an Entra or Microsoft 365 group team for larger teams.

## Known issues

- After upgrading, reload the Documents admin page without the browser cache (Ctrl+Shift+R) until the left menu shows Folder templates, Sites & access, Monitor and Settings. It can take more than one reload.
- Leaving the template editor through the app's left menu or the browser's Back button does not ask about unsaved edits. Edits autosave about 1.5 seconds after you stop typing; edits that cannot be saved yet (an incomplete condition, or a new template without a name) are lost.
- Monitor refuses a whole list when you cannot read the table of one of its records, and **Not captured** shows 0 unless you can read System Jobs at organization level.
- Operators see approximate record counts for template re-runs, from Dataverse's daily row count, even without read access to that table.
- The re-run option when publishing is not blocked for a template that is off or past its end date; the re-run then stops at once and Monitor says why.
- Bookmarks to the old `#access`, `#runtime` and `#operations` addresses open the default page.
- Opening a record's Documents tab before its folders exist makes Dynamics create its own folder, unless the form has the Documents form script. See [Hide the Documents tab until folders exist](customer-installation.md#hide-the-documents-tab-until-folders-exist).

See [operations](operations.md) for ongoing administration and [record changes](record-update-processing.md) for applying a template to existing records.
