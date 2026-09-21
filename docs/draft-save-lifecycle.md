# Draft save lifecycle

Save draft updates the latest working Draft in place. The first save after a Published revision starts the next numbered Draft; subsequent saves retain its ID and number until publication. Publishing continues to point the template to the immutable published revision. Older historical revisions are retained, not renumbered or deleted.

The editor retains the loaded/saved base revision and concurrency token separately from preview validity. Saving sends that identity even after edits invalidate the preview. Publishing reloads the published record to refresh its token before subsequent edits. Table changes and template deletion clear the base.

The server checks template ownership and expected row version. Existing latest drafts require that exact draft to be loaded; stale saves or attempts to fork another draft are rejected with a reload instruction. Draft children are replaced within the same Dataverse transaction after compare-and-swap of the draft aggregate. Published children and the published pointer are unchanged. Existing Author configuration CRUD privileges cover the replacement.
