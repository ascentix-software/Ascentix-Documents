# Operations and retention runbook

Keep runtime and flows disabled during installation. Follow the [installation procedure](customer-installation.md). Review the [preview limitations](public-preview.md#preview-limitations), including recovery scenarios that have not yet been rehearsed live.

## Enable after verification

Install and read back own component identities, active keys, plugin messages/step identity, roles and connection references. Assign the six roles only to reviewed identities; give the worker Read and Append To for its configured business tables. Prepare approved libraries, native navigation, unique permission boundaries and the optional writable AsxdWorkKey text column. Configure an exact worker, source tables and SharePoint-site allowlist. Obtain fresh GET-only catalog captures and separately review library policy. Verify connector ownership and read-only exact-site requests before enabling the serialized dispatcher.

## Runtime and pausing

Turning the runtime off pauses processing. Events keep queueing while it is off, so nothing is lost. Turning it back on resumes the queue.

Pause and resume take effect at once, even while work is running. Pause and resume are System Administrator only. A flow run that is driving a job when you pause ends at its next step without calling SharePoint, and shows as Succeeded; the job stays queued and continues after resume, once its 5-minute claim lapses.

Capture of record and team events runs as background jobs, so a save is never blocked by Documents. If capture fails, the save still succeeds and the failed job is kept for recovery.

The flow-details page shows a notice about dispatcher concurrency (trigger concurrency throttling). It is expected. The dispatcher runs one at a time on purpose, and the notice means a poll was skipped. No action is needed.

Run **Remove all event registrations** (Unregister) in Runtime administration before you uninstall. See [Upgrade, rollback and decommission](#upgrade-rollback-and-decommission).

**System Customizer is an administrator role for Documents.** Dataverse gives System Customizer full access to new custom tables, so a System Customizer can edit the runtime settings row directly, including the worker identity. Assign System Customizer only to people you would make System Administrator.

## Recovery and retries

Runtime administration has three recovery lists.

- **Failed capture jobs** lists record or team events whose background job failed. Fix the cause, select the records, and click **Replan selected records**. Replan asks Documents to evaluate the records again. The list finds the jobs through the Documents event handlers, so it works however many tables are enabled.
- **Blocked records** lists outbox rows that were blocked. Each has **Retry**. A record blocked for a reason that still holds blocks again. Records blocked before the 0.1.0.4 upgrade can be retried here. Its **Waiting** section lists records with folders that wait (see [Folder names](#folder-names)), each with its notice and **Replan**.
- **Blocked jobs** lists blocked operations and library setups in `RecoveryRequired`. Each has **Retry** and **Cancel**. Cancel asks in the page first and never undoes or deletes anything in SharePoint. A setup in `RecoveryRequired` also has **Open recovery**, which opens the recovery panel filled in for it. Its **Waiting to retry** section lists jobs waiting after a temporary error, those that have waited longest first, with their notice and next attempt.

Temporary failures wait and retry on their own. There is no attempt cap. A waiting job shows a notice and the time of its next attempt. **Cancel** stops a waiting job.

Unknown outcomes are handled by type. Reads retry. For folder and access writes, after the 5-minute claim expires, Documents re-checks the result and then finishes the work. Only library creation still needs the evidence-based recovery panel (`RecoveryRequired`), described in the next section. The panel inspects any operation by its pasted key, not only the latest 50.

## Admin actions

Suspend, approve, and queue or apply access always work while folder jobs run. Writes that were already submitted finish.

- **Suspend** holds the library's folder jobs, its queued access run and, for a site, its library setups. Each one waits with a notice and resumes by itself once the library and its site are approved again. The exception is a library setup whose create may already have reached SharePoint: it goes to `RecoveryRequired` and needs the original create response, or **Cancel** (see below). Approving again is refused only while the access run has a write whose result is not known yet. The daily access refresh and team events skip a suspended library with a notice on its access policy; other libraries are not affected.
- **Removing a table** from the Tables panel cancels its unsent folder jobs and its queued records, with a notice. Nothing in SharePoint is deleted. Enable the table again and replan the records to plan them.
- A library setup interrupted before its create was permitted (for example by a pause) reads SharePoint again and continues by itself. **Retry** and **Cancel** also work on a setup in `RecoveryRequired`. Retry of a create that may have reached SharePoint still needs the original response; setups saved before 0.1.0.4 count as possibly sent. **Cancel** never deletes anything; creating the same library again starts over with fresh reads. In Sites, a setup card that needs attention (`RecoveryRequired`, blocked or waiting) has **Retry** and **Cancel setup**; a Documents Security Administrator can use them without the Operator role.

- **Re-point** follows a library that was renamed, or a site that moved within the same tenant. If the site URL changes, first update the SharePoint host in Runtime administration and the Dataverse SharePoint site record. Re-point names exactly what to change. Existing folders keep working.
- **Remove** is always accepted. A destination used by a Draft or Published template is refused, and the refusal lists the templates. A destination still referenced by history is hidden rather than deleted. Nothing in SharePoint is ever deleted.

## Access management

Documents manages only its own SharePoint groups and their grants on the library. Sharing links, Limited Access, and people or groups added by hand are left alone.

- A hand edit to the permission level of a Documents group is reverted, with a notice. A Documents group deleted by hand is recreated.
- A team member that SharePoint rejects is skipped with a notice and retried on every scheduled refresh. A 401 or 403 Blocks the job, because it means a permission was lost.
- A library whose access run stopped or waits shows **Needs attention** with the run's notice, and **Retry access run** and **Cancel access run**. Cancel undoes nothing in SharePoint.
- Teams can be added, changed or removed while an access run is queued or running. **Apply access** replaces the queued run with the new access. If a flow is working on the run, or SharePoint has not answered one of its writes, the change is saved and the library says it waits; the next access review, a minute later, applies it as soon as the run can be replaced. A run is never replaced while one of its writes is unanswered.
- The site's Read and Contribute permission levels may be customized, for example Contribute without "Delete items". Documents refuses a level only when it carries administrative rights: ManageLists, ManageSubwebs, CreateGroups, ManagePermissions, ManageWeb, EnumeratePermissions, or Full Control. The refusal names the rights to remove. Every access run reads the level again, so a change made after approval is picked up.
- A team whose people do not fit in one access run keeps the members of its Documents group as they are, with a notice, and still gets its access; every scheduled refresh tries again. The bound comes from the 500,000-character Dataverse row that stores a team's people twice (as the team has them and as SharePoint lists them), which is roughly 2,000 people with typical sign-in names. Use an Entra or Microsoft 365 group team for larger teams: it is granted as one group.
- B2B guests are supported.
- Entra security-group and Microsoft 365 group teams are granted as the group itself.
  - "Members" teams also give the group's guests access.
  - Security-group "Owners" teams give all members of the group access.
  - Both require the admin to confirm a warning.
  - "Guests only" teams cannot be supported, because SharePoint has no claim for them.
- When you register an existing library that inherits permissions, Documents asks you to confirm. It then stops the inheritance and keeps a copy of the current permissions. If an admin later resets the library to inherit, access sync stops with a notice until **Apply access** is confirmed again.

## Folder names

- Characters SharePoint forbids are replaced with "-". Trailing dots and spaces are trimmed. Reserved names get "_".
- "Forms" is reserved only at the library root.
- A folder whose full path, from the site root and including the library, is longer than 300 characters is created with a notice: SharePoint allows 400 characters including file names, so files inside need short names. A folder whose path would pass 400 characters waits, with the folders below it, and the rest of the record is planned. Shorten the record's value or the template's folder names, then replan the record.
- A blank naming value makes that folder, and the folders below it, wait. The record is listed under **Waiting** in Blocked records until a later plan includes the folder. With record updates on, filling in a field of the record itself creates the folder at the record's next update. With record updates off, or for a field of a related record, fill in the field and then **Replan** the record. The notice says which applies.
- If sibling folders get the same name, the first one (by order) is kept and the rest wait, listed the same way, until a change to the record makes the names differ and the record is updated or replanned.
- **Risk:** Documents does not follow a record's folder if it is renamed or deleted in SharePoint. Document locations point at the old folder. Do not rename or delete record folders in SharePoint.

## Limits and pacing

- SharePoint calls are paced at about 1 per second across the environment. This follows the documented limit of the HTTP connector: 100 calls per 60 seconds per connection. Throughput does not grow with the number of sites.
- There is one writer per SharePoint site at a time. There is no environment-wide writer limit.
- There is no limit on the number of enabled tables.
- There is no limit on the number of SharePoint hosts. Hosts can be on any Microsoft SharePoint Online domain: `sharepoint.com` (worldwide and GCC), `sharepoint.us` (GCC High), `sharepoint-mil.us` or `dps.mil` (DoD), and `sharepoint.cn` (21Vianet). The list is stored in one column of 5,000 characters, about 190 typical hosts; a longer list is refused with that reason. Every call also needs an active SharePoint site record with the exact address.
- The daily access review reviews every library that is due, oldest first, for up to one minute per dispatcher run (half of Dataverse's 2-minute limit for one custom API call). Libraries left over are reviewed on the next run, a minute later.
- Removing a site or library reads every template destination and library that refers to it, however many there are.
- Templates can use lookup columns from tables that are not enabled. Changes to those related records do not update folders until the main record changes or is replanned.
- Field-secured columns are refused as naming sources, because folder names are not secured.
- The worker must be an application user.

## Blocked or ambiguous operation

Inspect the operation by its key in the recovery panel: its status, notice, claim RunId, Token and claim expiry. Read its notice, then check the exact resource in SharePoint. Repair the cause of a Blocked job before you **Retry** it; a cause that still holds blocks it again.

A write whose answer was lost is never sent again blindly:

- **Folder and access writes:** each job holds a 5-minute claim. Once the claim expires, the next dispatcher run takes the job over and reads SharePoint first. A folder that is there is adopted, and a group, member or grant change that took effect is recorded; only what is missing is written, once. No evidence is needed.
- **Library creation:** a create that may have reached SharePoint goes to `RecoveryRequired`. Find the flow run that sent it, and record its original create response in the recovery panel (**Open recovery** in Blocked jobs fills it in). Without that response, **Cancel** the setup; it deletes nothing, and creating the library again starts over with fresh reads. A setup interrupted before its create was permitted, for example by a pause, continues by itself.

The recovery panel records evidence; it does not stop a flow run that is still working.

Cancellation stops idle work and work that is waiting to retry. Preserve all observed physical IDs/nonces and partial grant receipts. Batch replan requires a fresh impact review if source or publication changed. A deletion tombstone keeps documents and prevents new work; completed physical work can retain a decommission-review binding without creating native navigation to a deleted record.

## Membership and access changes

Opt out a supported team or queue an approved None policy, then wait for independent managed grant/membership readback. Group membership remains when needed by other libraries. Check ordinary-user effective access and separately inventory other grants, shares, links, owners and site administrators. Applied means the product's reviewed scope converged; it is not an assertion that every alternative access path was removed. Repair incomplete identity/pagination or foreign grant conflicts rather than adopting or deleting them.

## Connection or certificate rotation

Drain the writer and pending unknown requests before changing runtime or connection configuration. Turn off schedules, verify the claim is idle, and disable the profile with its current row version. A tenant administrator rotates credentials in the managed connection through the approved platform procedure; never copy credentials into this repository or flow inputs. Revalidate the exact reference/worker/site binding and GET-only identity/denial probes before re-enabling.

## Upgrade, rollback and decommission

Preserve durable state, source provenance, the stable local assembly signing identity, and environment backups. Do not roll the solution back while an external writer is active or restore stale claims as permission to issue new calls. Keep flows disabled during solution import and configuration changes. To upgrade from 0.1.0.3, follow the [0.1.0.4 upgrade notes](upgrade-0.1.0.4.md). Verify all guards/keys before resuming. Verify retained configuration, connections, roles, and document identities after an upgrade.

Uninstall is not a cleanup command. First, in Runtime administration, click **Remove all event registrations**. Event steps created by the application are not part of the managed solution, and uninstall does not remove them. Uninstall is not verified in this release. Then explicitly revoke managed access while site approval and credentials remain available; verify residual access; export and retain binding/group/grant/attempt evidence. Remove platform components only through a separately reviewed decommission plan. Uninstall does not automatically delete SharePoint documents, libraries, groups, or historical bindings.
