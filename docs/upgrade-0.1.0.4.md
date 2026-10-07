# Upgrade to 0.1.0.4

Version 0.1.0.4 moves event capture into the application. Documents now registers and verifies its own event steps from Runtime administration. Read this page before you upgrade from 0.1.0.3.

Requires a System Administrator for step 4.

## Steps

1. Let the work queue drain with the runtime on. Then turn the runtime off and turn both flows off. A paused runtime does not drain the queue.
2. Import 0.1.0.4 managed with `--stage-and-upgrade`. This removes the old packaged event steps.
3. Publish all customizations (Power Apps › Solutions › Publish all customizations), then reload the Documents admin page. The import's `--publish-changes` does not always refresh the admin page's scripts: until everything is published, the app can keep serving the previous version.
4. As a System Administrator, open Runtime administration. Check the table list and click **Save**. Every table should show Ready.
5. Check the solution layers of both flows (**Dispatch durable work** and **Provision requested record**). If either has an active unmanaged layer, for example because it was edited in the environment, remove that active customization. Otherwise it stays on top of the upgraded flow and hides its retries, failure handling and write guard.
6. Turn the flows and runtime back on.
7. Records saved between steps 2 and 4 have no event. Replan them (filter by modified date).

Use stage-and-upgrade. A plain update leaves the removed components behind.

## Before you upgrade

- The Documents Worker role now includes System Jobs read (`prvReadAsyncOperation`). Capture runs as a background job owned by the worker, so the worker needs this privilege.
- The Documents Worker role now includes organization-level Write on Document Locations (`prvWriteSharePointDocumentLocation`). **Re-point** updates the library's Dataverse document location, which may belong to whoever created it.
- The Documents Security Administrator role now includes:
  - organization-level Delete on Sites and Libraries (`prvDeleteasx_site`, `prvDeleteasx_library`), so **Remove** can delete a site or library row nothing refers to;
  - organization-level Write on Operations and Create on Attempts (`prvWriteasx_operation`, `prvCreateasx_attempt`), so **Remove** and **Apply access** can cancel an idle access run;
  - user-level Create and Append on Document Locations (`prvCreateSharePointDocumentLocation`, `prvAppendSharePointDocumentLocation`), so approving an added library can create its document location.

  The guard plug-ins still allow writes to Documents tables only through the product APIs. If you copied these roles, add the same privileges to your copies.

- The worker application user needs organization-level Read on every enabled table. Registration refuses any table the worker cannot read and names it.
- The worker must be an application user.

## Behaviour changes

These change what Documents does with existing data. Review them before you upgrade.

- **Documents reverts hand changes to its own groups.** If someone changes the permission level of a Documents group on a library, or deletes the group, the next access run puts it back and records a notice. Documents never changes sharing links, Limited Access, or people and groups added by hand.
- **Whole-library permission drift is no longer checked.** An access run reads only the Documents group's own assignment. Changes to other entries on the library no longer stop access runs or folder work.
- **Customized permission levels are accepted.** The site's Read and Contribute levels may be customized, for example Contribute without "Delete items". A level is refused only when it carries administrative rights (ManageLists, ManageSubwebs, CreateGroups, ManagePermissions, ManageWeb, EnumeratePermissions or Full Control), and the notice names them. Every access run reads the level again, so a later customization no longer stops it.
- **Folder names from record data are cleaned.** Characters SharePoint forbids become "-", trailing dots and spaces are trimmed, and reserved names get "_". Records that were Blocked before the upgrade get cleaned names when you retry them.
- **Folders wait instead of blocking the record.** A blank or unusable name, a name a sibling already took, or a path longer than SharePoint's 400 characters makes that folder and the folders below it wait, with a notice. The rest of the record is planned. A path between 300 and 400 characters is created, with a notice that files inside need short names.
- **Pause takes effect at once.** It no longer waits for active writers. A flow run that is working when you pause makes no new SharePoint call and the job continues after resume. The answer to a write already sent is still recorded.
- **Team edits are allowed while access is being applied.** Apply access replaces a queued access run. If a flow is working on the run, or SharePoint has not answered one of its writes, the change waits and is applied right after.
- **Large teams no longer stop the access run.** A team or Documents group with more people than one run can store (roughly 2,000) keeps its group members as they are, with a notice, and still gets its access. Each scheduled refresh tries again. Team membership was not synced: people removed from the team keep access, and people added get none, until this is resolved. The library shows **Needs attention** until membership syncs; use an Entra or Microsoft 365 group team for a team this large.
- **A team deleted in Dataverse loses its library access without blocking the library.** The library shows it as "Deleted team" with its last known name. The next access run removes Documents' grant for the team's Documents group: an admin's **Apply access**, also one that changes only other teams, or the scheduled refresh. Once no library refers to the team, its registration is finished, with a notice. The Documents group stays in SharePoint; Documents does not delete SharePoint groups.
- **The daily access review covers every library.** It reviews every due library on each run, within half of the 2-minute custom API limit, and continues with the rest on the next run.
- **SharePoint hosts in every Microsoft cloud.** Hosts may be on `sharepoint.com`, `sharepoint.us` (GCC High), `sharepoint-mil.us` or `dps.mil` (DoD), or `sharepoint.cn` (21Vianet), and there is no host-count limit.
- **Pre-upgrade work carries on.** Folder jobs, access runs and library setups saved by 0.1.0.3 continue after the upgrade. A setup whose create may have reached SharePoint goes to `RecoveryRequired` (see the [operations guide](operations.md#recovery-and-retries)).

## What changes

- **Breaking change:** the `asx_PublishTemplate` result is now JSON, `{Status, Notices}`, instead of the text "Published". Update any caller that compares the result to "Published".
- The Documents Security Administrator role no longer has write access to the runtime settings row. Runtime changes stay System Administrator only.
- The runtime-row guard plug-in is removed. The solution has 57 guard steps.
- The outbox has a new column, `asx_nextattempt`. It holds the time of the next automatic attempt.
- Tables are enabled in the Tables panel of the template workspace, not in a fixed list. There is no table-count limit.
- **Remove** hides a site or library from Documents: it leaves the pickers, planning and access sync, and its unfinished work is cancelled. If anything refers to it, Documents keeps its catalog row for history. For a library that means template revisions, access settings (including the inheritance confirmation) or record folders. For a site it means libraries or library setups. Otherwise the row is deleted. Either way nothing in SharePoint changes, and adding it again works: a kept row is reactivated, and a deleted one is created again.

## After you upgrade

- Open Runtime administration and click **Save**, even if the table list looks right. This registers the event steps.
- Records that were Blocked before the upgrade can be retried from **Blocked records**.
- Folder jobs queued before the upgrade carry on by themselves.
- The flow-details page can show a notice about dispatcher concurrency. This is expected. See the [operations guide](operations.md#runtime-and-pausing).

## Uninstalling

Event steps created by the application are not part of the managed solution. Before you uninstall, click **Remove all event registrations** in Runtime administration. See the [operations guide](operations.md). Uninstall has not been verified in this release.
