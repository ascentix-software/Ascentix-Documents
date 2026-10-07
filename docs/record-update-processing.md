# System-wide record-update processing

**Settings › Automation settings** includes **Update folders when records change**. It is off by default and is stored as `asx_runtime.asx_processrecordupdates`; the runtime API field is `ProcessRecordUpdates`.

Saving the setting changes the state of the Update event steps that Documents registers. With the setting off, those steps are disabled, avoiding their invocation on high-volume record updates. With it on, update steps for the enabled source tables are enabled. Create, Delete, team events, integrity guards, and Power Automate flow activation are unaffected.

The setting covers root business records and related records used as template sources. An upsert that updates an existing record follows the update setting; an upsert that creates a new record can still trigger creation processing. There is no separate Upsert registration or Power Automate record-update trigger. The dispatcher processes queued work on its recurrence schedule.

When updates are enabled, the event handler filters each published template against the source columns or lookups present in the request. Event steps run asynchronously as background jobs, so capture cannot block a save. The handler queues evaluation and makes no SharePoint calls. When disabled, a defensive check exits before template, dependency, or outbox queries if an update invocation was already in flight or a step was manually re-enabled.

## Saving and installation

Load the runtime profile, change the setting, and save. Saving uses the runtime row version. When a save changes only this setting and/or **Enabled** (same worker and SharePoint hosts), it takes effect at once: it does not wait for active writers or check the worker, hosts or tables first, and it only turns the existing Update event steps on or off. Pausing therefore always succeeds. Resuming also succeeds and lists any worker, host or table problem in the readiness result. A save that changes the worker or hosts, or changes nothing, requires idle external writers, validates the complete owned registration inventory before changing it, and reads changed step states back; a registration mismatch, missing update setup, insufficient privileges, or readback failure prevents that save. The API runs inside a Dataverse transaction. Allow for registration propagation before relying on the new setting.

The administrator saving this setting needs Dataverse privileges to read the relevant plug-in metadata and update plug-in steps. The API uses the caller's service and does not elevate privileges or grant a role. Runtime permission alone does not grant step-management privileges.

The solution includes the runtime Boolean column used by the plug-in and admin resources. Fresh installations leave update monitoring off. **Save settings** in Settings › Automation settings creates, corrects, and verifies the event steps for every enabled table; Settings › Change tracking shows each table Ready when its steps are in place, and **Repair all** (or **Repair** on a row) fixes one that is not. See the [upgrade notes](upgrade-0.1.0.4.md).

Turning updates off does not cancel already queued work. Turning them on does not replay updates missed while monitoring was off. Use a reviewed re-run for catch-up: Folder templates › **Re-run for existing records…**, or **Re-run** in Monitor. Creation processing remains available, subject to its own installed/enabled Create steps and the worker being enabled.

Enable monitoring during a quiet interval and verify a disposable write before depending on it. Registration changes can take time to propagate; there is no guaranteed propagation interval.

## Worker queue creation

Dataverse authorizes the original business-record write. Record and membership event plug-ins validate their own current execution, configured worker, and source scope, then append using the ordinary organization service.

For Create of `asx_outbox`, `StateGuard` accepts the configured worker identity without reconstructing event ancestry. The new `WorkerOutboxAppend` check verifies the JSON schema version, nonempty key/status, matching indexed key/status, and deterministic row identity. Direct worker creates are intentionally trusted. It does not require an event tag, SharedVariables marker, correlation match, or registered ancestor. Semantic template, source, destination, and policy checks remain in the planner and worker.

This exception does not authorize ordinary-user direct writes, outbox Update/Delete, or writes to other state tables. Existing product API paths and receipt/history protections remain. Operator and Security Administrator APIs retain their existing guarded write paths. No new API, credential, role grant, or Power Automate trigger is introduced.
