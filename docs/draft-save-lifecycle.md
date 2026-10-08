# Draft save lifecycle

Drafts save automatically, about 1.5 seconds after the last edit; the editor header shows "Saving…" and then "Saved" with the time. One save runs at a time: an edit made while a save is in flight is saved by one more save after it, with the row version the first one returned. A save the server refuses shows "Couldn't save · Retry" and keeps the edits. Conditions that are not valid are not saved until they are fixed; the Folders step is marked until then. Leaving the page, another template, ＋ New or Close save what can be saved first, and ask only about edits that cannot be saved.

A save updates the latest working Draft in place. The first save after a Published revision starts the next numbered Draft; subsequent saves retain its ID and number until publication. Publishing continues to point the template to the immutable published revision. Older historical revisions are retained, not renumbered or deleted.

The editor retains the loaded/saved base revision and concurrency token separately from preview validity. Saving sends that identity even after edits invalidate the preview. Publish saves pending edits first. Edits made while it runs are kept: they are saved as the next Draft with the token publishing gave the revision. Table changes and template deletion clear the base.

The server checks template ownership and expected row version. Existing latest drafts require that exact draft to be loaded; stale saves or attempts to fork another draft are rejected with a reload instruction. Draft children are replaced within the same Dataverse transaction after compare-and-swap of the draft aggregate. Published children and the published pointer are unchanged. Existing Author configuration CRUD privileges cover the replacement.
