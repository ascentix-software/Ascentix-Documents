# Operations and retention runbook

Keep runtime and flows disabled during installation. Follow the [installation procedure](customer-installation.md). Review the [preview limitations](public-preview.md#preview-limitations), including recovery scenarios that have not yet been rehearsed live.

## Enable after verification

Install and read back own component identities, active keys, plugin messages/step identity, roles and connection references. Assign the six roles only to reviewed identities; give the worker Read and Append To for its configured business tables. Prepare approved libraries, native navigation, unique permission boundaries and the optional writable AsxdWorkKey text column. Configure an exact worker, table allowlist and SharePoint-site allowlist. Obtain fresh GET-only catalog captures and separately review library policy. Verify connector ownership and read-only exact-site requests before enabling the serialized dispatcher.

## Blocked or ambiguous operation

Inspect the operation key and actual claim RunId, Token and LeaseUntilUtc in the operator panel. Read its receipt/error and independently inspect the exact resource. A known idle failure can be retried only after its cause is repaired. An ambiguous write retains the global claim: stop the originating flow run, establish that outstanding external requests have settled, wait for the lease to expire, and record specific evidence against the exact old identity. The recovery API grants permission to reread and reconcile; it is not proof that a remote writer was stopped. Never take over on expiry alone or repeatedly submit an unknown POST.

Cancellation is limited to idle unsubmitted work. Preserve all observed physical IDs/nonces and partial grant receipts. Batch replan requires a fresh impact review if source or publication changed. A deletion tombstone keeps documents and prevents new work; completed physical work can retain a decommission-review binding without creating native navigation to a deleted record.

## Membership and access changes

Opt out a supported team or queue an approved None policy, then wait for independent managed grant/membership readback. Group membership remains when needed by other libraries. Check ordinary-user effective access and separately inventory other grants, shares, links, owners and site administrators. Applied means the product's reviewed scope converged; it is not an assertion that every alternative access path was removed. Repair incomplete identity/pagination or foreign grant conflicts rather than adopting or deleting them.

## Connection or certificate rotation

Drain the writer and pending unknown requests before changing runtime or connection configuration. Turn off schedules, verify the claim is idle, and disable the profile with its current row version. A tenant administrator rotates credentials in the managed connection through the approved platform procedure; never copy credentials into this repository or flow inputs. Revalidate the exact reference/worker/site binding and GET-only identity/denial probes before re-enabling. Reapproval or policy review is required if observed catalog/ACL state changed.

## Upgrade, rollback and decommission

Preserve durable state, source provenance, the stable local assembly signing identity, and environment backups. Do not roll the solution back while an external writer is active or restore stale claims as permission to issue new calls. Keep flows disabled during solution import and configuration changes. Verify all guards/keys before resuming. Verify retained configuration, connections, roles, and document identities after an upgrade.

Uninstall is not a cleanup command. First explicitly revoke managed access while site approval and credentials remain available; verify residual access; export and retain binding/group/grant/attempt evidence. Remove platform components only through a separately reviewed decommission plan. Uninstall does not automatically delete SharePoint documents, libraries, groups, or historical bindings.
