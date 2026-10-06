# Upgrade to 0.1.0.4

Version 0.1.0.4 moves event capture into the application. Documents now registers and verifies its own event steps from Runtime administration. Read this page before you upgrade from 0.1.0.3.

Requires a System Administrator for step 3.

## Steps

1. Turn the runtime off and stop the flows. Let the work queue empty.
2. Import 0.1.0.4 managed with `--stage-and-upgrade`. This removes the old packaged event steps.
3. As a System Administrator, open Runtime administration. Check the table list and click **Save**. Every table should show Ready.
4. Turn the flows and runtime back on.
5. Records saved between steps 2 and 3 have no event. Replan them (filter by modified date).

Use stage-and-upgrade. A plain update leaves the removed components behind.

## Before you upgrade

- The Documents Worker role now includes System Jobs read (`prvReadAsyncOperation`). Capture runs as a background job owned by the worker, so the worker needs this privilege.
- The Documents Worker role now includes organization-level Write on Document Locations (`prvWriteSharePointDocumentLocation`). **Re-point** updates the library's Dataverse document location, which may belong to whoever created it.
- The Documents Security Administrator role now includes:
  - organization-level Delete on Sites and Libraries (`prvDeleteasx_site`, `prvDeleteasx_library`), so **Remove** can delete a site or library nothing refers to;
  - organization-level Write on Operations and Create on Attempts (`prvWriteasx_operation`, `prvCreateasx_attempt`), so **Remove** and **Apply access** can cancel an idle access run;
  - user-level Create and Append on Document Locations (`prvCreateSharePointDocumentLocation`, `prvAppendSharePointDocumentLocation`), so approving an added library can create its document location.

  The guard plug-ins still allow writes to Documents tables only through the product APIs. If you copied these roles, add the same privileges to your copies.
- The worker application user needs organization-level Read on every enabled table. Registration refuses any table the worker cannot read and names it.
- The worker must be an application user.

## What changes

- **Breaking change:** the `asx_PublishTemplate` result is now JSON, `{Status, Notices}`, instead of the text "Published". Update any caller that compares the result to "Published".
- The Documents Security Administrator role no longer has write access to the runtime settings row. Runtime changes stay System Administrator only.
- The runtime-row guard plug-in is removed. The solution has 57 guard steps.
- The outbox has a new column, `asx_nextattempt`. It holds the time of the next automatic attempt.
- Tables are enabled in the Tables panel of the template workspace, not in a fixed list. There is no table-count limit.

## After you upgrade

- Open Runtime administration and click **Save**, even if the table list looks right. This registers the event steps.
- Records that were Blocked before the upgrade can be retried from **Blocked records**.
- Folder jobs queued before the upgrade carry on by themselves.
- The flow-details page can show a notice about dispatcher concurrency. This is expected. See the [operations guide](operations.md#runtime-and-pausing).

## Uninstalling

Event steps created by the application are not part of the managed solution. Before you uninstall, click **Remove all event registrations** in Runtime administration. See the [operations guide](operations.md). Uninstall has not been verified in this release.
