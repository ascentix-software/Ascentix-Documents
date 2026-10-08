# Operations and retention runbook

Keep Automation and flows off during installation. Follow the [installation procedure](customer-installation.md). Review the [preview limitations](public-preview.md#preview-limitations), including recovery scenarios that have not yet been rehearsed live.

## Enable after verification

Install and read back own component identities, active keys, plugin messages/step identity, roles and connection references. Assign the six roles only to reviewed identities; give the worker Read and Append To for its configured business tables. Prepare approved libraries, native navigation, unique permission boundaries and the optional writable AsxdWorkKey text column. Configure an exact worker, source tables and SharePoint-site allowlist. Obtain fresh GET-only catalog captures and separately review library policy. Verify connector ownership and read-only exact-site requests before enabling the serialized dispatcher.

## Automation and pausing

Turning automation off (the switch on the **Automation** card in Settings) pauses processing. The pill in the Monitor header says whether automation runs, with **Change in Settings**. Events keep queueing while it is off, so nothing is lost. Turning it back on resumes the queue.

Pause and resume take effect at once, even while work is running. Pause and resume are System Administrator only. A flow run that is driving a job when you pause makes no new SharePoint call: it ends at its next SharePoint request and shows as Succeeded. The job stays queued and continues after resume, once its 5-minute claim lapses. Pause never drops an answer: a SharePoint write already sent when you pause has its answer recorded, and a job that only needs to record a verified result completes. A library create sent just before a pause therefore continues after resume and does not need recovery.

Capture of record and team events runs as background jobs, so a save is never blocked by Documents. If capture fails, the save still succeeds and the failed job is kept for recovery.

The flow-details page shows a notice about dispatcher concurrency (trigger concurrency throttling). It is expected. The dispatcher runs one at a time on purpose, and the notice means a poll was skipped. No action is needed.

Before you uninstall, use **Settings › Before uninstalling › Stop tracking changes**. See [Upgrade, rollback and decommission](#upgrade-rollback-and-decommission).

**System Customizer is an administrator role for Documents.** Dataverse gives System Customizer full access to new custom tables, so a System Customizer can edit the runtime settings row directly, including the worker identity. Assign System Customizer only to people you would make System Administrator.

## Monitor and retries

Monitor is one table of everything that needs attention, newest first by **Since** (select it to reverse the order). The filter chips above it show one list each, with its count: **All**, **Blocked records**, **Waiting for data**, **Blocked jobs**, **Retrying** and **Not captured**. Each row shows its type, the item, the problem and when it started, with one action as a button and the others in its **⋯** menu, which ends with **Copy ID**. Monitor shows names and local times; IDs appear only through Copy ID or in Look up an operation. **Refresh** reads every list again; the header says when Monitor last checked.

Re-runs of templates show as cards above the chips while they run or for a day after they end.

- **Template re-runs** shows background re-runs started from Folder templates (the overview's **⋯** menu › Re-run for existing records, or Publish with its re-run), with progress, an estimated finish once a page of records is done, and who started each. A re-run queues one page of 19 records at a time and only after the previous page is planned, so record changes and other work go first: expect about 1,100 records an hour when nothing else is queued. Totals read "about N" until the run ends, because they come from Dataverse's daily row count. **Pause** stops it after the page in progress; **Resume** continues; **Cancel re-run** skips the records not yet planned. Folder work already queued continues, and nothing in SharePoint is undone.
- **Not captured** lists record or team events whose background job failed. **Re-run** queues the record for every published, active template of its table; under the Not captured chip, **Re-run selected** does it for the checked rows; **Dismiss** removes the failed job. The count and the list show the background jobs the person using Monitor can read: with user-level read on System Jobs, only their own. Monitor users need organization-level read on System Jobs to see every missed change.
- **Blocked records** lists records whose planning stopped, with **Retry**. A record blocked for a reason that still holds blocks again.
- **Waiting for data** lists records with folders that wait for a value or a shorter path (see [Folder names](#folder-names)), with **Re-run** and **Open record**.
- **Blocked jobs** lists folder, access and library jobs that stopped, each with what stopped it and how to fix it, and **Retry** and **Cancel job**. Cancel asks in the page first and never undoes or deletes anything in SharePoint. A library setup whose create may have reached SharePoint shows what Documents found in SharePoint and its choices (see [Blocked or ambiguous operation](#blocked-or-ambiguous-operation)).
- **Retrying** lists jobs waiting after a temporary error, with their attempt and, once it is set, their next attempt. There is no attempt cap; **Retry now** runs one at once.

Counts above 5,000 show as 5,000+.

**Tools ▾** opens two side panels. **Check a record…** shows what Documents did for one record and template, and re-runs it with the published template. **Look up an operation…** lists recent operations and looks one up by its ID.

Temporary failures wait and retry on their own. There is no attempt cap. A waiting job shows a notice and the time of its next attempt. **Cancel job** stops a waiting job.

Unknown outcomes are handled by type. Reads retry. For folder and access writes, after the 5-minute claim expires, Documents re-checks the result and then finishes the work. Library creation is looked up in SharePoint instead, described in [Blocked or ambiguous operation](#blocked-or-ambiguous-operation).

## Settings

Settings is one column of cards. **Automation** has the switch that pauses and resumes processing. **Automation settings** holds who Documents runs as and the SharePoint host names. **Tables** lists the tables Documents tracks with their change tracking: **＋ Add table**, **Repair** on a table whose tracking needs it, **Repair all**, and **Remove**. **Connections** lists the connection references, and **Before uninstalling** has **Stop tracking changes**. Changes to Automation settings wait in a bar at the bottom of the page until **Save settings**, or **Discard**. People who are not System Administrators see the values read-only.

## Sites & access

The rail lists the sites with their library count; **＋ Add** adds a site. A site's libraries are a table: each library, the templates that use it (**Used by**), its teams and its access state. **Add existing library** and **＋ Create library** are in the header, and the site's **⋯** menu has Re-point and Remove. Selecting a library opens its access drawer: its teams with their access, **＋ Add team**, and the library's **⋯** menu. Changes wait in the drawer's footer until **Apply access**, or **Discard**. Escape or ✕ closes the drawer.

## Folder templates

The templates list groups templates by table with their state (Live, Draft or Off); below 1000 pixels it becomes a **Template** select above the overview, under **Templates** and **＋ New**. The overview shows one template: its destinations, its folders with the rule of each, its schedule and runs, and its versions. **Edit template** (or **Continue Draft**) opens the editor in three steps: **1 Destinations**, **2 Folders** and **3 Review and publish**. Edits save by themselves; see [Draft save lifecycle](draft-save-lifecycle.md). A step with changes since the published version has a dot, and the footer counts them.

- **Folders** shows each destination's folder tree beside the selected folder's panel: its name with **＋ Field**, **Create this folder** (Always or Only when…) with its conditions, and **Test records**. A test record is previewed with the draft as it is after each pause in editing, and says whether the selected folder is Created or Skipped.
- **Review and publish** lists every change since the published version, each with a link to its step; shows the folders a test record would get (**Result for**); and says what publishing does. **Re-run existing records after publishing** is on by default; it needs the Documents Operator role, and it is off when **Starts** is **Later…**, because a re-run of a template that has not started stops at once. **Publish** publishes the saved draft, writes the start time, and then starts the re-run, and returns to the overview. If the re-run is refused, the version stays published and the overview says why, with **Re-run existing records…**.

## Admin actions

Suspending or approving a site or library, and applying access, always work while folder jobs run. Writes that were already submitted finish.

- **Suspending** a site or library is a catalog API command (`SuspendSite` or `SuspendLibrary` on `asx_CatalogAdmin`); Sites & access has no Suspend control, and shows a suspended site or library as **Needs attention**. Suspending holds the library's folder jobs, its queued access run and, for a site, its library setups. Each one waits with a notice and resumes by itself once the library and its site are approved again. The exception is a library setup whose create may already have reached SharePoint: Documents looks it up in SharePoint and offers the choices that fit (see below). Approving again is refused only while the access run has a write whose result is not known yet. The daily access refresh and team events skip a suspended library with a notice on its access policy; other libraries are not affected.
- **Removing a table** from Settings › Tables cancels its unsent folder jobs and its queued records, with a notice. Nothing in SharePoint is deleted. Enable the table again and re-run the records to plan them.
- A library setup interrupted before its create was permitted (for example by a pause) reads SharePoint again and continues by itself. **Retry** and **Cancel setup** also work on a setup in `RecoveryRequired`. A create that may have reached SharePoint is looked up instead of retried; setups saved before 0.1.0.4 count as possibly sent. **Cancel setup** never deletes anything; creating the same library again starts over with fresh reads. In Sites & access, a setup card that needs attention (`RecoveryRequired`, blocked or waiting) has **Retry**, **Cancel setup** and **Open in Monitor**; a Documents Security Administrator can use Retry and Cancel setup without the Operator role. A setup whose create may have reached SharePoint shows Documents' SharePoint check and its choices instead, as in Monitor.

- **Add existing library** checks Documents' access to the library and creates its navigation entry. A library that inherits the site's permissions is added only after you confirm, in the page, that Documents stops the inheritance. It does not offer a library Documents already has. It matches by the SharePoint list ID, so a library renamed in SharePoint is not offered again under its new name; use **Re-point** to follow the rename.
- **Re-point** follows a library that was renamed, or a site that moved within the same tenant. If the site URL changes, first update the SharePoint host in Settings › Automation settings and the Dataverse SharePoint site record. Re-point names exactly what to change. Existing folders keep working.
- **Remove** hides a site or library from Documents: it leaves the pickers, planning and access sync, and its unfinished work is cancelled. If anything refers to it, Documents keeps its catalog row for history. For a library that means template revisions, access settings (including the inheritance confirmation) or record folders. For a site it means libraries or library setups. Otherwise the row is deleted. Either way nothing in SharePoint changes, and adding it again works: a kept row is reactivated, and a deleted one is created again. A destination that a Draft or Published template uses is refused, and the refusal lists the templates. Remove a site's libraries before the site: **Remove site** stays unavailable, with that reason, while the site has libraries. Re-point and Remove are in the **⋯** menu of the site or the library, and ask in the page before they run.

## Access management

Documents manages only its own SharePoint groups and their grants on the library. Sharing links, Limited Access, and people or groups added by hand are left alone.

- Site owners keep their administrative access; only the teams you choose get access to the library.
- Team access applies to every folder in the library. To restrict a folder, use a separate library.
- **Remove** on a team row strikes it through, with **Undo**, until you apply. Removed teams lose the access Documents gave them. Access given another way, such as sharing links or site membership, is not changed.
- A hand edit to the permission level of a Documents group is reverted, with a notice. A Documents group deleted by hand is recreated.
- A team member that SharePoint rejects is skipped with a notice and retried on every scheduled refresh. A 401 or 403 Blocks the job, because it means a permission was lost.
- A library whose access run stopped or waits shows **Needs attention** with the run's notice, and **Retry access run** and **Cancel access run**. Cancel undoes nothing in SharePoint. A change applied while the run is stopped waits for it: Retry or Cancel the run, then the change applies.
- Teams can be added, changed or removed while an access run is queued or running. **Apply access** replaces the queued run with the new access. If a flow is working on the run, or SharePoint has not answered one of its writes, the change is saved and the library says it waits; the next access review, a minute later, applies it as soon as the run can be replaced. A run is never replaced while one of its writes is unanswered.
- The site's Read and Contribute permission levels may be customized, for example Contribute without "Delete items". Documents refuses a level only when it carries administrative rights: ManageLists, ManageSubwebs, CreateGroups, ManagePermissions, ManageWeb, EnumeratePermissions, or Full Control. The refusal names the rights to remove. Every access run reads the level again, so a change made after approval is picked up.
- A team whose people do not fit in one access run keeps the members of its Documents group as they are, with a notice, and still gets its access; every scheduled refresh tries again. **Team membership was not synced: people removed from the team keep access, and people added get none, until this is resolved.** The library shows **Needs attention** instead of "Access and team membership confirmed." whenever the last run left membership unsynced, including when SharePoint refused a member. The bound comes from the 500,000-character Dataverse row that stores a team's people twice (as the team has them and as SharePoint lists them), which is roughly 2,000 people with typical sign-in names. Use an Entra or Microsoft 365 group team for larger teams: it is granted as one group.
- B2B guests are supported. Further testing is ongoing.
- Entra security-group and Microsoft 365 group teams are granted as the group itself.
  - "Members" teams also give the group's guests access.
  - Security-group "Owners" teams give all members of the group access.
  - Both require the admin to confirm a warning.
  - "Guests only" teams cannot be supported, because SharePoint has no claim for them.
- When you register an existing library that inherits permissions, Documents asks you to confirm. It then stops the inheritance and keeps a copy of the current permissions. If an admin later resets the library to inherit, its access run stops and the library shows **Needs attention** with the notice, **Retry access run** and **Cancel access run**. **Apply access**, confirmed again, stops the inheritance and applies the access.

## Folder names

- Characters SharePoint forbids are replaced with "-". Trailing dots and spaces are trimmed. Reserved names get "_".
- "Forms" is reserved only at the library root.
- A folder whose full path, from the site root and including the library, is longer than 300 characters is created with a notice: SharePoint allows 400 characters including file names, so files inside need short names. A folder whose path would pass 400 characters waits, with the folders below it, and the rest of the record is planned. Shorten the record's value or the template's folder names, then re-run the record.
- A folder Documents already created or found, and every parent folder, is read by its ID, so its path never appears in the request. A new folder's path and a new library's name are sent in the query string of the request, not in its URL path: the HTTP connector refuses a URL path longer than its `maxUrlLength` setting. The query string itself is limited to 2,048 characters, and an escaped character can take up to 9 (most non-Latin letters do), so a folder whose escaped path would pass that limit waits with a notice, like a path over 400 characters, and a library name that would pass it is refused when you create the library. If the connector still refuses a request for its URL length, the job is Blocked with `RequestUrlTooLong` (a library setup or catalog check shows the same reason in words). SharePoint did not receive the request. Shorten the record's value or the template's folder names, then re-run the record.
- A blank naming value makes that folder, and the folders below it, wait. The record is listed in Monitor under **Waiting for data** until a later plan includes the folder. Fill in the field and then **Re-run** the record.
- If sibling folders get the same name, the first one (by order) is kept and the rest wait, listed the same way, until a change to the record makes the names differ and the record is re-run.
- **Risk:** Documents does not follow a record's folder if it is renamed or deleted in SharePoint. Document locations point at the old folder. Do not rename or delete record folders in SharePoint.

**Path limits.** SharePoint refuses a folder whose decoded path, file names included, is longer than 400 characters, so such a folder waits with a notice like 'needs a shorter path: 404 characters; the limit is 400'. The HTTP connector refuses a request whose query string passes 2,048 characters; Documents reads folders by their path in the query string, so a path that would pass it waits the same way.

**Preview and folder creation.** Preview evaluates names and conditions with your access; existing folders and name collisions are found when the folder is created.

## Limits and pacing

- SharePoint calls are paced at about 1 per second across the environment. This follows the documented limit of the HTTP connector: 100 calls per 60 seconds per connection. Throughput does not grow with the number of sites.
- There is one writer per SharePoint site at a time. There is no environment-wide writer limit.
- There is no limit on the number of enabled tables.
- There is no limit on the number of SharePoint hosts. Hosts can be on any Microsoft SharePoint Online domain: `sharepoint.com` (worldwide and GCC), `sharepoint.us` (GCC High), `sharepoint-mil.us` or `dps.mil` (DoD), and `sharepoint.cn` (21Vianet). The list is stored in one column of 5,000 characters, about 190 typical hosts; a longer list is refused with that reason. Every call also needs an active SharePoint site record with the exact address.
- The daily access review reviews every library that is due, oldest first, for up to one minute per dispatcher run (half of Dataverse's 2-minute limit for one custom API call). Libraries left over are reviewed on the next run, a minute later.
- Removing a site or library reads every template destination and library that refers to it, however many there are.
- Templates can use lookup columns from tables that are not enabled. Changes to those related records do not update folders until the main record changes or is re-run.
- Field-secured columns are refused as naming sources, because folder names are not secured.
- The worker must be an application user.

## Blocked or ambiguous operation

Each blocked job in Monitor shows what stopped it and how to fix it. Repair the cause before you **Retry**; a cause that still holds blocks it again. Monitor › Tools › Look up an operation looks any operation up by its ID.

A write whose answer was lost is never sent again blindly:

- **Folder and access writes:** each job holds a 5-minute claim. Once the claim expires, the next dispatcher run takes the job over and reads SharePoint first. A folder that is there is adopted, and a group, member or grant change that took effect is recorded; only what is missing is written, once.
- **Library creation:** Documents waits until the claim of the run that sent the create has ended, then looks SharePoint up with reads only: the library by its title, the library at the address SharePoint gives that title, and the Documents catalog. While it looks, the setup shows "Checking SharePoint…". It does not hold its site. What it finds decides the choices:
  - **Found** (one library matches the name, the address, the type, and was created after the request): **Use the library** in Monitor (**Use the library that was created** on the setup card in Sites & access). The setup continues with that library and never creates one.
  - **Not found** (no library with the title or at the address, and no catalog entry with the name): **Create it again**. The setup reads the title again first; a library that appeared in between stops it with "A library with this name already exists. Add the existing library instead."
  - **Ambiguous** (anything else): the reason and each candidate with the checks it failed. **Use this one** on a candidate, after an in-page confirmation. **Create it again** is offered only while no candidate has the requested name.
  - Every finding offers **Check again** and **Cancel setup**. Cancel deletes nothing in SharePoint.

  If you added the created library as an existing library meanwhile, using it adopts that catalog entry; access you set there is kept, and with none the setup's own access is applied. A setup interrupted before its create was permitted, for example by a pause, reads SharePoint again and continues by itself. Documents never creates a library twice.

Cancellation stops idle work and work that is waiting to retry. Preserve all observed physical IDs/nonces and partial grant receipts. Re-running records from a preview (Re-run these records) requires a fresh impact review if source or publication changed. A deletion tombstone keeps documents and prevents new work; completed physical work can retain a decommission-review binding without creating native navigation to a deleted record.

## Membership and access changes

Opt out a supported team or queue an approved None policy, then wait for independent managed grant/membership readback. Group membership remains when needed by other libraries. Check ordinary-user effective access and separately inventory other grants, shares, links, owners and site administrators. Applied means the product's reviewed scope converged; it is not an assertion that every alternative access path was removed. Repair incomplete identity/pagination or foreign grant conflicts rather than adopting or deleting them.

A team deleted in Dataverse needs no action. The library shows it as "Deleted team" with its last known name, or its ID if none was stored. The next access run of each library it reached removes Documents' grant for the team's Documents group and its members: **Apply access**, including one that changes only other teams, or the scheduled refresh. A library that was removed from Documents and added again is reviewed at once. When no library refers to the team any more, its registration is finished (`Revoked`) and the library's notices say so. The Documents group stays in SharePoint, because Documents does not delete SharePoint groups. Delete it there if it is no longer needed.

## Connection or certificate rotation

Drain the writer and pending unknown requests before changing Automation or connection configuration. Turn off schedules, verify the claim is idle, and disable the profile with its current row version. A tenant administrator rotates credentials in the managed connection through the approved platform procedure; never copy credentials into this repository or flow inputs. Revalidate the exact reference/worker/site binding and GET-only identity/denial probes before re-enabling.

## Upgrade, rollback and decommission

Preserve durable state, source provenance, the stable local assembly signing identity, and environment backups. Do not roll the solution back while an external writer is active or restore stale claims as permission to issue new calls. Keep flows disabled during solution import and configuration changes. To upgrade from 0.1.0.3, follow the [0.1.0.4 upgrade notes](upgrade-0.1.0.4.md). Update folders when records change is not in this release. An environment that had it on stops capturing record updates after Repair all. Verify all guards/keys before resuming. Verify retained configuration, connections, roles, and document identities after an upgrade.

Uninstall is not a cleanup command. First, use **Settings › Before uninstalling › Stop tracking changes**. Event steps created by the application are not part of the managed solution, and uninstall does not remove them. Uninstall is not verified in this release. Then explicitly revoke managed access while site approval and credentials remain available; verify residual access; export and retain binding/group/grant/attempt evidence. Remove platform components only through a separately reviewed decommission plan. Uninstall does not automatically delete SharePoint documents, libraries, groups, or historical bindings.

## Terms

| Term | Meaning |
|---|---|
| Automation | The runtime that runs folder and access work. |
| Change tracking | The event steps that capture record and team changes. |
| Re-run | Plan a record again with the published template. |
| Template re-run | A background re-run of every record of a template's table. |
| Missed change | An event whose capture job failed. |
| Destination | A library a template creates folders in. |
| Check SharePoint | Documents' read-only lookup of a library creation whose answer was lost. |
