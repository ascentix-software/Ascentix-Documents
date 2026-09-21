# Author-controlled template lifecycle

Authors can deactivate/reactivate a template, set optional start and end dates, or delete it including published versions. These are ordinary Dataverse Update/Delete operations using the author's table privileges, not another approval API.

The admin page offers Active, Start, End, Save availability and Delete template. Dates are entered in the browser's local time and stored as UTC instants. Start is inclusive; end is exclusive. End must be later than start. Empty dates mean no time restriction. Publication does not override availability.

Deleting removes the template, its revisions and their owned configuration through Dataverse cascading relationships. It does not remove SharePoint folders, documents, permissions, native document locations or historical worker receipts. The deletion confirmation states this clearly. See Microsoft's [relationship behavior documentation](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/configure-entity-relationship-cascading-behavior).

Events and the scheduled record sweep skip unavailable templates. Queue/replan and pending planning enforce availability. An operation checks availability before claiming and before each processing step; unsubmitted work becomes Cancelled and its holder releases the claim. Fresh evaluations after reactivation may restart work cancelled specifically for template availability. Explicitly cancelled work is not automatically restarted.

A create request already handed to the external flow may finish after deactivation, expiry or deletion. Such requests retain their receipts and finish independent readback without needing the deleted template. No new create intent is handed out once the worker observes that the template is unavailable. This is not a promise to recall an HTTP request already in flight.

Start dates enable processing during the window; they do not promise an exact-time backfill. Existing records are considered by the regular sweep or explicit replanning. Reactivation does not automatically recreate deleted configuration or rewrite historical files.
