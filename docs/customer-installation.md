# Managed customer installation

Use the managed solution package for evaluation. Review the [preview limitations](public-preview.md#preview-limitations) before setup.

## Prepare and import

1. Select the Dataverse target, application identity, SharePoint site, and document-enabled business tables. Retain the approved package's version and hash alongside the deployment record.
2. Create the Dataverse application user and the Dataverse application connection in the target environment. Create the SharePoint certificate HTTP connection in Power Automate. Enter certificates, passwords, and secrets directly in the platform.
3. Generate deployment settings from the managed ZIP with `pac solution create-settings`. Map all packaged connection references to the target connections. Use all reference names from the package, including any additional references carried by packaged flows.
4. Import the managed solution using those settings. For an existing installation, pause runtime and flows and preserve the runtime, roles, registrations, connection mappings, and retained document identities before applying the upgrade. For a managed upgrade, the package has been exercised with `--stage-and-upgrade --publish-changes --async`.
5. Confirm the installed solution is managed, its version and signed assembly match the intended package, alternate keys are Active, custom APIs and guard steps exist, and connection references point to the target connections. Use application-user mappings from the destination environment.

## Configure and enable

1. Assign Documents Worker to the application user, plus a role granting Read and Append To on the configured business tables. Documents Worker includes Global Read Attribute for field metadata. Verify the effective role grant; System Administrator is not required.
2. Configure Runtime administration with the exact application user, business-table allowlist, and SharePoint hostname. Keep record updates Off unless the installation needs that behavior.
3. Configure the event registrations for the selected tables with the target worker identity. An allowlist entry or template does not install missing steps. The packaged business-record registrations cover Account; additional tables require operator registration of the product event plugin and its supported messages. The Account steps in `solution/AscentixDocuments/src/SdkMessageProcessingSteps` provide the registration pattern; use the target worker identity for each selected table. Verify Create/CreateMultiple and any required Delete/team events. Update and UpdateMultiple registrations must follow the system-wide update setting. Allow registration changes to propagate before declaring activation complete; verify with a disposable record.
4. Register the target SharePoint site, configure or create the document library, and apply the intended team policy. Confirm actual library role assignments and managed group membership. Keep the installation's Owners access intact.
5. Publish a small template, enable runtime and the worker flows, and provision a disposable business record. Verify physical folder IDs and paths. Replay the request to confirm reuse. Confirm that the worker's installed product and business-table roles are sufficient.
6. Capture the resulting configuration and enable only the event paths needed by the installation. Use the operations runbook for backlog inspection, retry, cancellation, and recovery. Configure document navigation in the business app and forms used by the installation.

## Recovery and upgrade record

Retain the package hash, import operation, worker application/user IDs, connection-reference mapping, enabled-event list, configuration, and sample physical folder IDs. Keep credentials in the platform. Before an upgrade or connection rotation, drain active writers and pause runtime/flows; verify the same mappings and retained documents afterward.


Do not use solution uninstall to clear document state. See the [operations guide](operations.md) for recovery and decommissioning procedures.
