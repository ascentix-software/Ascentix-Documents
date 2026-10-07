# Worker identity and connections

The worker uses an application user in the target Dataverse environment. Assign **Documents Worker**, including Global **Read Attribute** (`prvReadAttribute`) for field metadata, plus Read and Append To for the configured business tables. System Administrator is not required.

Application users, credentials, and connector connections are environment-specific. A managed solution does not supply ready-to-use identities. Follow the [managed installation procedure](customer-installation.md) before enabling processing.

## Configuration inputs

| Input | Meaning |
| --- | --- |
| WorkerApplicationId | Entra application client ID selected for the installation |
| WorkerUserId | That application's enabled Dataverse systemuser ID in the target environment |
| DataverseSourceReference | Reference bound to a Dataverse connection authenticated as that application |
| CertificateSourceReference | HTTP with Microsoft Entra ID (preauthorized) reference for the configured SharePoint tenant hosts |
| SharePointHosts | Lowercase tenant hostnames on a Microsoft SharePoint Online domain: `sharepoint.com` (worldwide and GCC), `sharepoint.us` (GCC High), `sharepoint-mil.us` or `dps.mil` (DoD), or `sharepoint.cn` (21Vianet), such as `tenant.sharepoint.com`. No URL paths or wildcards. Any number of hosts, stored in one 5,000-character column |
| AllowedTables | Template (root) table logical names; add a lookup-source table only when its changes should update folders |
| EventTables | Root and related-source tables requiring events, within AllowedTables |

## Connections and permissions

1. Create the application user in the target environment and assign its product and business-table roles.
2. In Power Automate, create the Dataverse application connection and SharePoint certificate HTTP connection. Enter credentials directly in the platform and bind the packaged connection references.
3. Verify the Dataverse connection's authenticated UserId and OrganizationId against the intended worker and environment. Connection ownership or display name alone does not identify its authenticated caller.
4. Verify the SharePoint application's site grant and a read request to the intended site. Product site approval does not grant the service principal access to SharePoint.
5. Configure the worker identity (**Run as**) and SharePoint host names in **Settings › Automation settings**, and enable tables in Folder templates › Tables (**＋ Add table**). Keep runtime and flows disabled until setup is complete.

Use an enabled application user for the configured worker and event-step impersonation. Role assignment remains an administrator task.

## Sites, tables, and activation

Enable document management for the selected business tables. A template can use lookup columns from tables that are not enabled; the worker reads them with its Read privilege, and changes to them update folders only when that table is enabled too. Each target site requires an active native SharePoint Site registration within the configured hosts. Use Add site to validate the site, then create or register libraries and configure team access.

Publish a small template. Documents registers each enabled table's event steps for the configured worker when you click **Save settings** in Settings › Automation settings or add the table with **＋ Add table** in Folder templates › Tables. Settings › Change tracking shows whether every table is Ready; **Repair all** (or **Repair** on a row) fixes one that is not. The [record-update setting](record-update-processing.md) turns the registered Update steps on or off and defaults to Off.

Enable runtime and worker flows, then provision a disposable record and verify the resulting folders. Confirm replay reuses them and that the worker roles are sufficient. Consult [preview limitations](public-preview.md#preview-limitations) before expanding the workload.
