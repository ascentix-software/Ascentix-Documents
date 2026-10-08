# Managed customer installation

Use the managed solution package for evaluation. Review the [preview limitations](public-preview.md#preview-limitations) before setup.

## Prepare and import

1. Select the Dataverse target, application identity, SharePoint site, and document-enabled business tables. Retain the approved package's version and hash alongside the deployment record.
2. Create the Dataverse application user and the Dataverse application connection in the target environment. One application connection is used across authorized sites. Create the SharePoint certificate HTTP connection in Power Automate; the **Connections** card in the Documents app's Settings shows whether each connection is connected. Enter certificates, passwords, and secrets directly in the platform.
3. Generate deployment settings from the managed ZIP with `pac solution create-settings`. Map all packaged connection references to the target connections. Use all reference names from the package, including any additional references carried by packaged flows.
4. Import the managed solution using those settings. For an existing installation from 0.1.0.3, follow the [0.1.0.4 upgrade notes](upgrade-0.1.0.4.md). Turn **Automation** off, pause the flows, and preserve the runtime, roles, registrations, connection mappings, and retained document identities before applying the upgrade. For a managed upgrade, the package has been exercised with `--stage-and-upgrade --publish-changes --async`.
5. Publish all customizations (Power Apps › Solutions › Publish all customizations), then reload the Documents admin page without the browser cache (Ctrl+Shift+R, or Ctrl+F5). Until then the app's left menu can keep the previous names. Do this after every import, including an upgrade: `--publish-changes` does not always refresh the admin page's scripts, so the app can keep serving the previous version until everything is published.
6. Confirm the installed solution is managed, its version and signed assembly match the intended package, alternate keys are Active, custom APIs and guard steps exist, and connection references point to the target connections. Use application-user mappings from the destination environment.

## Configure and enable

1. Create the worker as an application user. Assign Documents Worker to it, plus a role granting organization-level Read and Append To on every enabled business table. Documents Worker includes Global Read Attribute for field metadata and System Jobs read (`prvReadAsyncOperation`). Event capture runs as background jobs owned by the worker, so the worker needs both System Jobs read and organization-level Read on each enabled table. Verify the effective role grant; System Administrator is not required for the worker.
2. Configure **Settings › Automation settings** with the exact application user (Run as) and SharePoint host name. Source tables are enabled in Settings › Tables.
3. As a System Administrator, open **Settings**, select the worker application user (**Run as**) on the **Automation settings** card and click **Save settings** in the bar at the bottom of the page. Documents registers and verifies its event steps; every table on the **Tables** card must show Ready. Repair any that does not with **Repair all** (or **Repair** on its row). Tables are enabled on the same card (**＋ Add table**). Registration refuses a table the worker cannot read and names it. Verify with a disposable record.
4. Set up the target SharePoint site in Dataverse document management first: **＋ Add site** in Sites & access lists the SharePoint sites set up there. Then add the site, add or create the document library, and apply the intended team access. Confirm actual library role assignments and managed group membership. Keep the installation's Owners access intact.
5. Publish a small template (Folder templates › **Edit template** › **Review and publish**), turn **Automation** on (the switch on the **Automation** card in Settings) and enable the worker flows, and provision a disposable business record. Verify physical folder IDs and paths. Replay the request to confirm reuse. Confirm that the worker's installed product and business-table roles are sufficient.
6. Capture the resulting configuration and enable only the event paths needed by the installation. Use the operations runbook for backlog inspection, retry, cancellation, and recovery. Configure document navigation in the business app and forms used by the installation, and [hide the Documents tab until folders exist](#hide-the-documents-tab-until-folders-exist).

## Hide the Documents tab until folders exist

When a record's Documents tab opens before the record has a document location, Dynamics creates its own folder in the table's default library (named `<record name>_<id>`) and a location called "Documents on <site>". There is no setting to turn this off. Documents creates the record's location only after its SharePoint folder exists, so a user who opens the tab too early gets a stray folder.

The solution includes a form script that hides Documents until the record has a location. On each form of a document-enabled table:

1. Open the form in the form designer and add the web resource `asx_form/documents-tab.js` to the form libraries.
2. Add an **On Load** event handler: library `asx_form/documents-tab.js`, function `AscentixDocuments.Form.onLoad`. Check **Pass execution context as first parameter**.
3. Optional: if the form shows documents elsewhere, such as a tab with a Documents subgrid, list those tabs or sections in the handler's parameters, quoted and comma-separated: `"tab_documents", "tab_summary.section_files"` (a section is `tab.section`). The Related › Documents item is always handled. A name the form does not have is ignored.
4. Save and publish the form.

A new record keeps Documents hidden until it is saved. An existing record without a location keeps it hidden and shows "Folders are being created in SharePoint. The Documents tab will be visible when complete." The form checks every 5 seconds and shows Documents as soon as the location exists. After 2 minutes it stops checking and says "Folders aren't ready yet. Reopen this record later, or ask an administrator to check Monitor in Ascentix Documents." If a user cannot read document locations, Documents is shown as usual. Users do not need a Documents role for the script.

## Recovery and upgrade record

Retain the package hash, import operation, worker application/user IDs, connection-reference mapping, enabled-event list, configuration, and sample physical folder IDs. Keep credentials in the platform. Before an upgrade or connection rotation, drain active writers and pause runtime/flows; verify the same mappings and retained documents afterward.

Do not use solution uninstall to clear document state. See the [operations guide](operations.md) for recovery and decommissioning procedures.
