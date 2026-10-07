'use strict';
(() => {
  const root = document.getElementById('access'),
    $ = (id) => document.getElementById(id),
    xrm = window.parent?.Xrm || window.Xrm;
  if (!root) return;
  const state = {
    sites: [],
    libraries: [],
    teams: new Map(),
    // Teams deleted in Dataverse, by lower-case ID, with their last known name or their ID.
    deleted: new Map(),
    // Team ID to the broader-access warning its group carries.
    warnings: new Map(),
    // The selection whose warnings are shown: { kind: 'apply' | 'provision', key, warnings }.
    consent: null,
    policies: new Map(),
    site: null,
    library: null,
    busy: false,
    polling: false,
    pollPromise: null,
    loaded: false,
    operations: new Map(),
    nextSites: null,
    nextLibraries: null,
    nextActivity: null,
    startPromise: null,
    selectedOperation: null,
    discovery: null,
    progressSignature: null,
    completed: new Map(),
    pollFailures: 0,
    // The site or library command waiting for in-page confirmation: { kind, id, name }.
    confirm: null,
    // What the last re-point changed, such as "old URL → new URL".
    changes: [],
    // Sites and libraries removed in this page. They leave the lists at once, whatever a list
    // read returns, until they are added again.
    removed: new Set(),
    // List IDs, in lower case, of the libraries on the shown discovery page that Documents
    // already has as active rows.
    registered: new Set(),
  };
  const noList = '00000000-0000-0000-0000-000000000000';
  const node = (tag, text) => {
    const n = document.createElement(tag);
    if (text != null) n.textContent = text;
    return n;
  };
  const opt = (value, text) => {
    const o = node('option', text);
    o.value = value;
    return o;
  };
  // Owner teams sync their members. An Entra or Microsoft 365 group team is granted through its
  // group, so it is labelled with its kind. This mirrors the server (TeamPrincipal), with
  // team.teamtype 2 security group, 3 Microsoft 365 group and membershiptype 1 members,
  // 2 owners, 3 guests:
  // - teams SharePoint cannot identify are listed disabled with the reason;
  // - teams whose group reaches more people than the team carry a warning the admin confirms.
  const teamLabel = (t) => {
    if (t.teamtype !== 2 && t.teamtype !== 3) return { text: t.name, reason: null, warning: null };
    const security = t.teamtype === 2,
      kind = security ? 'Entra group' : 'Microsoft 365 group';
    const reason = !t.azureactivedirectoryobjectid
      ? 'no Microsoft Entra group object ID, so SharePoint cannot identify its group'
      : t.membershiptype === 3
        ? 'guests only: SharePoint has no sign-in claim for only the guests of a group'
        : null;
    if (reason) return { text: t.name + ' (' + kind + ') - ' + reason, reason, warning: null };
    const warning =
      t.membershiptype === 1
        ? "The group's guests will also have access to this library."
        : security && t.membershiptype === 2
          ? 'All members of the group will have access to this library, not only its owners.'
          : null;
    const scope =
      t.membershiptype === 1
        ? ', members; guests also get access'
        : t.membershiptype === 2
          ? security
            ? ', owners; all members get access'
            : ', owners'
          : '';
    return { text: t.name + ' (' + kind + scope + ')', reason: null, warning };
  };
  // Warnings for the chosen teams that still need the admin's confirmation, shown in the page.
  // A second select of the same button with the same choice confirms them.
  const warningsFor = (teamIds) => [
    ...new Set(teamIds.map((id) => state.warnings.get(id)).filter(Boolean)),
  ];
  function confirmed(kind, key, warnings) {
    if (!warnings.length) return { go: true, ack: false };
    if (state.consent?.kind === kind && state.consent.key === key) return { go: true, ack: true };
    state.consent = { kind, key, warnings };
    return { go: false };
  }
  const consent = (kind, key, teamIds) => confirmed(kind, key, warningsFor(teamIds));
  // Shown before an admin adds a library that inherits its site's permissions (the server's
  // CatalogAdministration.InheritanceWarning) and before Apply access stops it again.
  const inheritanceWarning =
    'This library inherits permissions from the site. When you approve it, Documents stops the inheritance, keeps a copy of the current site permissions, and then manages team access on it.';
  const reapplyWarning =
    'This library inherits permissions from the site. When you apply access, Documents stops the inheritance, keeps a copy of the current site permissions, and then manages team access on it.';
  const inheritsAgain =
    'This library inherits permissions again. Use Apply access to let Documents stop the inheritance again.';
  // Applied while the queued run could not be replaced yet (PolicyDocument.ApplyPending).
  const pendingNotice =
    'Your change is saved. It is applied right after the access run in progress, as soon as that run is between steps and SharePoint has answered its last write.';
  const stoppedPendingNotice =
    'The access run stopped: Retry or Cancel it, then your change applies.';
  // The last run applied the grants but left team membership unsynced (MembershipIncomplete).
  const incompleteNotice =
    'Needs attention: access is applied, but team membership was not synced. People removed from a team keep access, and people added get none, until this is resolved. See the notices below.';
  // The last access run stopped because the library inherits the site's permissions.
  const inherits = (p) => !!p?.result.Policy?.Inherits;
  // The library's queued access run stopped or waits to retry. A run stopped because the
  // library inherits permissions again is one of them: it needs attention like any other, and
  // Apply access with the acknowledgement is its remedy.
  const stuckRun = (p) => ['Blocked', 'RetryWait'].includes(p?.result.RunStatus);
  // A team deleted in Dataverse keeps its row until access is applied, which removes its access
  // (GetPolicy's Teams marks it); Dataverse can no longer read it by ID.
  const deletedNotice =
    'This team was deleted in Dataverse. Documents removes its access the next time access is applied.';
  // Once a run removed its access (the scheduled refresh, say), only its row is left.
  const deletedRemovedNotice = 'Its access was removed. Apply access to clear it from this list.';
  const deletedTeam = (teamId) => state.deleted.get(String(teamId).toLowerCase());
  const deletedPending = (p) => !!p?.entries.some((e) => deletedTeam(e.TeamId) != null);
  // The library's access state, from its policy and queued run as GetPolicy last read them,
  // never from the catalog flag the library list was read with, which can be older.
  function accessLabel(p) {
    if (!p) return 'Checking access…';
    const status = p.result.Status,
      queued = !!p.result.Policy?.OperationKey;
    if (
      stuckRun(p) ||
      (!queued && status === 'NeedsReview') ||
      (!queued && status === 'Applied' && p.result.Policy?.MembershipIncomplete)
    )
      return 'Needs attention';
    if (queued) return 'Applying access…';
    if (status === 'Applied') return 'Access applied';
    return 'Access setup pending';
  }
  const applyKey = (p) => (state.library?.asx_libraryid || '') + JSON.stringify(p?.entries || []);
  const issue = (text, error = false) => {
    $('ad-message').textContent = text;
    $('ad-message').className = error ? 'ad-issue' : 'ad-status';
  };
  const guid = (v) =>
    /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(v || '');
  // A server time ("/Date(ms)/" or ISO) as "YYYY-MM-DD hh:mm UTC".
  const when = (value) => {
    const ms = /^\/Date\((-?\d+)/.exec(value || '');
    const date = new Date(ms ? Number(ms[1]) : value);
    return isNaN(date) ? String(value) : date.toISOString().slice(0, 16).replace('T', ' ') + ' UTC';
  };
  async function api(name, request) {
    const input = {
      Request: JSON.stringify(request),
      getMetadata: () => ({
        boundParameter: null,
        parameterTypes: { Request: { typeName: 'Edm.String', structuralProperty: 1 } },
        operationType: 0,
        operationName: name,
      }),
    };
    const response = await xrm.WebApi.online.execute(input);
    if (!response.ok) throw new Error(await response.text());
    return JSON.parse((await response.json()).Result);
  }
  const catalog = (request) => api('asx_CatalogAdmin', request),
    security = (request) => api('asx_SecurityAdmin', request);
  async function page(table, options) {
    return xrm.WebApi.retrieveMultipleRecords(table, options, 50);
  }
  function nextOptions(next) {
    const url = new URL(next, xrm.Utility.getGlobalContext().getClientUrl());
    if (url.origin !== new URL(xrm.Utility.getGlobalContext().getClientUrl()).origin)
      throw new Error('Invalid catalog continuation.');
    return url.search;
  }
  // Runs one user action after any active poll finishes and renders its result or error.
  async function action(fn) {
    if (state.busy) return;
    state.busy = true;
    render();
    try {
      if (state.pollPromise) await state.pollPromise.catch(() => {});
      await fn();
    } catch (e) {
      issue(e.message || String(e), true);
    } finally {
      state.busy = false;
      render();
    }
  }
  function policy() {
    return state.library ? state.policies.get(state.library.asx_libraryid) : null;
  }
  function changed(p) {
    return p && JSON.stringify(p.entries) !== p.saved;
  }
  function progressStages(kind) {
    switch (kind) {
      case 'LibrarySetup':
        return ['Queued', 'Set up library', 'Set up access', 'Ready'];
      case 'LibraryValidation':
        return ['Queued', 'Check library', 'Ready'];
      case 'LibraryDiscovery':
        return ['Queued', 'Find libraries', 'Ready'];
      case 'Repoint':
        return ['Queued', 'Read SharePoint', 'Ready'];
      default:
        return ['Queued', 'Check site', 'Finish setup', 'Ready'];
    }
  }

  // Converts worker status to display text and retains the furthest observed setup stage.
  function progressState(operation) {
    const result = operation.result || {};
    const status = result.Status || operation.status || 'Pending';
    const stages = progressStages(operation.kind);
    const done = ['Ready', 'Approved', 'Applied', 'Discovered'].includes(status);
    const stopped =
      ['Blocked', 'Quarantined', 'RecoveryRequired'].includes(status) ||
      (status === 'ExternalUnknown' && !!result.Issue);
    const queued = ['Pending', 'Queued', 'Busy'].includes(status);
    const captured = ['Verified', 'Captured'].includes(status);
    const retrying = status === 'RetryWait';

    let step = 1;
    if (done) {
      step = stages.length - 1;
    } else if (queued || retrying) {
      step = operation.step || 0;
    } else if (
      (operation.kind === 'SiteValidation' && captured) ||
      (operation.kind === 'LibrarySetup' && ['AccessPending', 'Applying access'].includes(status))
    ) {
      step = 2;
    }
    step = Math.max(operation.step || 0, step);
    operation.step = step;

    let label = stages[step];
    if (done) label = 'Ready';
    else if (stopped) label = 'Needs attention';
    else if (retrying) label = 'Waiting to retry';
    else if (queued) label = 'Waiting for a worker';
    else if (status === 'ExternalUnknown') label = 'Waiting for confirmation';
    else if (captured) label = 'Finishing setup';

    let message;
    if (done) {
      // A setup can end Ready with something to do, such as a cancelled first access run.
      message = result.Issue || 'Setup completed.';
    } else if (stopped) {
      message = result.Issue || 'Setup needs review before it can continue.';
    } else if (retrying) {
      message =
        'SharePoint is temporarily unavailable or limiting requests. Retrying automatically.';
    } else if (queued || step === 0) {
      message = 'Your request is queued. Setup will start when a worker is available.';
    } else if (status === 'ExternalUnknown') {
      message = 'The request has been sent. Waiting for SharePoint to confirm the result.';
    } else if (captured) {
      message = 'Checks passed. Saving the result and preparing the site or library.';
    } else if (step === 2) {
      message = 'Applying library access and synchronizing team members.';
    } else {
      message = 'Checking and preparing your selected destination.';
    }
    return { stages, step, done, stopped, label, message, status };
  }
  // Whether a tracked operation belongs to another site or selection. A re-point is matched by
  // its site's ID, since re-pointing a site changes its address.
  function elsewhere(key, o) {
    if (state.selectedOperation && state.selectedOperation !== key) return true;
    if (!state.site) return false;
    if (o.siteId) return o.siteId !== state.site.asx_siteid;
    return !!o.url && state.site.asx_url !== o.url;
  }
  function drawProgress() {
    const area = $('ad-provision-progress');
    area.replaceChildren();
    let loading = false;
    for (const [key, o] of new Map([...state.completed, ...state.operations])) {
      if (elsewhere(key, o)) continue;
      // Found again after a reload and not read yet: its status alone does not say which
      // message and action apply, so nothing is offered until its first read.
      if (o.unread) {
        loading = true;
        continue;
      }
      const p = progressState(o),
        card = node('section'),
        heading = node('div'),
        badge = node('span', p.label),
        bar = node('progress'),
        list = node('ol');
      card.className =
        'ad-progress-card' + (p.stopped ? ' ad-progress-attention' : p.done ? '' : ' is-loading');
      card.setAttribute('style', '--stage-count:' + p.stages.length);
      heading.className = 'ad-row ad-between';
      badge.className = 'ad-progress-badge';
      heading.append(node('h3', o.name), badge);
      bar.max = p.stages.length - 1;
      bar.value = p.step;
      bar.setAttribute('aria-label', o.name + ' setup progress');
      bar.setAttribute(
        'aria-valuetext',
        p.done ? 'Complete' : p.label + '; stage ' + (p.step + 1) + ' of ' + p.stages.length,
      );
      list.className = 'ad-progress-stages';
      p.stages.forEach((name, i) => {
        const item = node('li'),
          mark = node('span', i < p.step || p.done ? '✓' : String(i + 1));
        mark.className = 'ad-step-mark';
        mark.setAttribute('aria-hidden', 'true');
        item.className = i < p.step || p.done ? 'is-complete' : i === p.step ? 'is-current' : '';
        if (i === p.step && !p.done) item.setAttribute('aria-current', 'step');
        item.append(mark, node('span', name));
        list.append(item);
      });
      const message = node('p', p.message);
      message.className = p.stopped ? 'ad-issue' : 'ad-muted';
      card.append(heading, bar, list, message);
      const observed = o.result?.Observation;
      if (p.status === 'Blocked' && o.kind === 'LibraryValidation' && observed?.Inherits) {
        // The card shows the server's refusal, which repeats the warning; this is the consent.
        const add = node('button', 'Stop inheritance and add');
        add.onclick = () =>
          action(async () => {
            await addLibrary(
              observed.SiteId,
              observed.ListId,
              o.name,
              observed.WebUrl || o.url,
              true,
            );
            state.operations.delete(key);
            state.completed.delete(key);
          });
        card.append(add);
      } else if (p.status === 'Blocked' && o.kind === 'Repoint') {
        // Re-point only reads SharePoint; once the cause is fixed it simply reads again. It is
        // never retried in place.
        if (o.command) {
          const again = node('button', 'Re-point again');
          again.onclick = () =>
            action(async () => {
              state.operations.delete(key);
              state.completed.delete(key);
              await runConfirmed({ kind: o.command, id: o.id, name: o.name });
            });
          card.append(again);
        }
      } else if (
        o.kind === 'LibrarySetup' &&
        ['Blocked', 'RecoveryRequired', 'RetryWait'].includes(p.status)
      ) {
        // Through the catalog API, so a Documents Security Administrator needs no Operator
        // role. The server keeps its rules: Retry of a create that may have reached SharePoint
        // is refused with the way out, and Cancel deletes nothing in SharePoint.
        const retry = node('button', 'Retry');
        retry.onclick = () =>
          action(async () => {
            const result = await catalog({ Command: 'RetrySetup', Key: key });
            o.result = result;
            o.status = result.Status;
            state.progressSignature = null;
            issue('Setup queued to run again. It stops again if the cause remains.');
          });
        const cancel = node('button', 'Cancel setup');
        cancel.onclick = () => {
          state.confirm = { kind: 'CancelSetup', id: key, name: o.name };
          render();
        };
        card.append(retry, cancel);
      } else if (p.status === 'Blocked') {
        const retry = node('button', 'Retry after repair');
        retry.onclick = () =>
          action(async () => {
            await api('asx_ManageWork', { Command: 'Retry', Key: o.result?.RecoveryKey || key });
            o.result = { Status: 'Pending' };
            o.status = 'Pending';
            state.progressSignature = null;
            issue('Retry queued.');
          });
        card.append(retry);
      } else if (p.stopped)
        card.append(
          node('p', 'Open Administration to review the original request before retrying.'),
        );
      if (state.completed.has(key) && !state.operations.has(key)) {
        const dismiss = node('button', 'Dismiss');
        dismiss.setAttribute('aria-label', 'Dismiss ' + o.name + ' progress');
        dismiss.onclick = () => {
          state.completed.delete(key);
          drawProgress();
        };
        card.append(dismiss);
      }
      area.append(card);
    }
    if (loading) {
      const wait = node('p', 'Loading…');
      wait.className = 'ad-muted';
      area.append(wait);
    }
    area.hidden = area.children.length === 0;
  }
  function render() {
    $('ad-sites').replaceChildren();
    state.sites.forEach((s) => {
      const b = node('button');
      b.type = 'button';
      b.className = 'ad-site';
      b.setAttribute('aria-pressed', String(s.asx_siteid === state.site?.asx_siteid));
      b.append(
        node('span', s.asx_name),
        node('small', s.asx_approved ? 'Ready' : 'Needs attention'),
      );
      b.onclick = () => action(() => selectSite(s));
      $('ad-sites').append(b);
    });
    for (const [key, o] of state.operations) {
      if (o.kind !== 'SiteValidation') continue;
      const b = node('button', o.name + ' · ' + progressState(o).label);
      b.className = 'ad-site';
      b.type = 'button';
      b.onclick = () => {
        state.selectedOperation = key;
        state.site = null;
        state.library = null;
        state.libraries = [];
        render();
        action(refreshOperations);
      };
      $('ad-sites').append(b);
    }
    $('ad-more-activity').hidden = !state.nextActivity;
    $('ad-more-sites').hidden = !state.nextSites;
    $('ad-more-libraries').hidden = !state.nextLibraries;
    $('ad-site-title').textContent = state.site?.asx_name || 'Select or add a site';
    $('ad-site-health').textContent = state.site
      ? state.site.asx_approved
        ? 'Ready'
        : 'Needs attention'
      : '';
    $('ad-site-error').hidden = !state.site || state.site.asx_approved;
    $('ad-create').disabled = state.busy || !state.site?.asx_approved;
    $('ad-existing').disabled = state.busy || !state.site?.asx_approved;
    $('ad-libraries').replaceChildren();
    state.libraries.forEach((l) => {
      const p = state.policies.get(l.asx_libraryid),
        b = node('button', l.asx_name + (changed(p) ? ' · Unsaved' : ''));
      b.type = 'button';
      b.className = 'ad-library';
      b.setAttribute('aria-pressed', String(l.asx_libraryid === state.library?.asx_libraryid));
      b.onclick = () => action(() => selectLibrary(l));
      $('ad-libraries').append(b);
    });
    $('ad-empty').hidden = !state.site || !!state.library || state.libraries.length > 0;
    $('ad-library-detail').hidden = !state.library;
    const p = policy();
    if (state.library) {
      $('ad-library-title').textContent = state.library.asx_name;
      $('ad-library-status').textContent = state.library.asx_approved
        ? 'Ready for folder templates'
        : 'Needs attention';
      $('ad-library-access').textContent = accessLabel(p);
    }
    $('ad-teams').replaceChildren();
    // The queued access run: one that stopped or waits is shown with its notice and can be
    // retried or cancelled here. Teams stay editable while a run is queued or running: Apply
    // access replaces the run, or, while a flow holds it or SharePoint has not answered its
    // write, the change waits for it (ApplyPending) and is applied right after.
    const run = p?.result.RunStatus,
      stuck = stuckRun(p),
      running = !!p?.result.Policy?.OperationKey && !stuck,
      pending = !!p?.result.Policy?.ApplyPending,
      incomplete = !!p?.result.Policy?.MembershipIncomplete;
    if (p) {
      p.entries.forEach((e) => {
        const tr = node('tr'),
          cell = node('td'),
          select = node('select');
        ['None', 'Read', 'Contribute'].forEach((level) =>
          select.append(opt(level, level === 'None' ? 'Remove access' : level)),
        );
        const gone = deletedTeam(e.TeamId),
          name = gone != null ? 'Deleted team: ' + gone : state.teams.get(e.TeamId) || e.TeamId;
        select.value = e.Access;
        select.disabled = state.busy || gone != null;
        select.setAttribute('aria-label', name + ' access');
        select.onchange = () => {
          e.Access = select.value;
          render();
        };
        cell.append(select);
        const current =
          p.result.Policy?.Applied?.find((a) => a.TeamId === e.TeamId)?.Access || 'None';
        const label = node('td', name);
        if (gone != null) {
          const notice = node('div', current === 'None' ? deletedRemovedNotice : deletedNotice);
          notice.className = 'ad-issue';
          label.append(notice);
        }
        tr.append(
          label,
          cell,
          node(
            'td',
            (current === 'None' ? 'No managed access' : current) + (running ? ' · applying' : ''),
          ),
        );
        $('ad-teams').append(tr);
      });
      if (!p.entries.length) {
        const tr = node('tr'),
          td = node('td', 'No additional teams.');
        td.colSpan = 3;
        tr.append(td);
        $('ad-teams').append(tr);
      }
    }
    // A cancelled run leaves the policy to review: Apply access starts a new run.
    // A library added again after it was removed keeps its policy marked Removed until its next
    // access run. A deleted team's row waits for Apply to remove its access.
    const reapply =
      inherits(p) ||
      ['Missing', 'NeedsReview', 'Removed'].includes(p?.result.Status) ||
      deletedPending(p);
    $('ad-apply').disabled = state.busy || !p || (!changed(p) && !reapply);
    $('ad-add-team').disabled = state.busy || !p;
    // A stopped run comes first: a change waiting behind it applies only after Retry or Cancel.
    // A run that waits to retry carries on by itself, and the waiting change follows it.
    $('ad-change-status').textContent = !p
      ? 'Loading access…'
      : stuck
        ? 'Needs attention: ' +
          (p.result.RunNotice || (inherits(p) ? inheritsAgain : 'the access run stopped.')) +
          (run === 'RetryWait' && p.result.RunNextAttemptUtc
            ? ' Next check: ' + when(p.result.RunNextAttemptUtc) + '.'
            : '') +
          (pending && !changed(p)
            ? run === 'Blocked'
              ? ' ' + stoppedPendingNotice
              : ' ' + pendingNotice
            : '')
        : pending && !changed(p)
          ? pendingNotice
          : running
            ? 'Applying access and syncing members…'
            : inherits(p)
              ? inheritsAgain
              : changed(p)
                ? 'Changes not yet applied.'
                : ['Missing', 'Removed'].includes(p.result.Status)
                  ? 'Apply to confirm this library’s access.'
                  : p.result.Status === 'NeedsReview'
                    ? 'The access run was cancelled. Apply access to run it again.'
                    : p.result.Status === 'Applied' && incomplete
                      ? incompleteNotice
                      : p.result.Status === 'Applied'
                        ? 'Access and team membership confirmed.'
                        : p.result.Status;
    $('ad-change-status').className =
      stuck || (incomplete && p?.result.Status === 'Applied' && !running && !changed(p))
        ? 'ad-issue'
        : 'ad-muted';
    $('ad-run-actions').hidden = !stuck;
    $('ad-run-retry').disabled = state.busy;
    $('ad-run-cancel').disabled = state.busy;
    // What the last access sync skipped or reconciled, such as members SharePoint could not take.
    const notices = (p && p.result.Policy?.Notices) || [];
    $('ad-access-notices').replaceChildren(...notices.map((n) => node('li', n)));
    $('ad-access-notices').hidden = !notices.length;
    // A shown warning belongs to one exact selection; any change asks again.
    if (state.consent?.kind === 'apply' && state.consent.key !== applyKey(p)) state.consent = null;
    const confirmApply = state.consent?.kind === 'apply',
      confirmCreate = state.consent?.kind === 'provision';
    $('ad-apply').textContent = confirmApply ? 'Confirm and apply' : 'Apply access changes';
    $('ad-apply-warning').textContent = confirmApply
      ? state.consent.warnings.join(' ') + ' Select Confirm and apply to continue.'
      : '';
    $('ad-apply-warning').hidden = !confirmApply;
    $('ad-provision').textContent = confirmCreate ? 'Confirm and create' : 'Create library';
    $('ad-provision-warning').textContent = confirmCreate
      ? state.consent.warnings.join(' ') + ' Select Confirm and create to continue.'
      : '';
    $('ad-provision-warning').hidden = !confirmCreate;
    // Site and library commands ask in the page, like Remove in the Tables panel.
    const c = state.confirm;
    $('ad-confirm').hidden = !c;
    $('ad-confirm-text').textContent = c ? confirmText(c) : '';
    $('ad-confirm-go').textContent = c ? confirmLabel(c) : 'Confirm';
    $('ad-confirm-cancel').textContent = c?.kind.startsWith('Cancel') ? 'Keep it' : 'Cancel';
    $('ad-changes').replaceChildren(...state.changes.map((t) => node('li', t)));
    $('ad-changes').hidden = !state.changes.length;
    $('ad-repoint-site').disabled = state.busy || !state.site;
    $('ad-repoint-library').disabled = state.busy || !state.library;
    $('ad-remove-site').disabled = state.busy || !state.site;
    $('ad-remove-library').disabled = state.busy || !state.library;
    [
      'ad-validate',
      'ad-provision',
      'ad-stage-team',
      'ad-add-site',
      'ad-more-sites',
      'ad-more-libraries',
    ].forEach((id) => ($(id).disabled = state.busy));
    drawProgress();
  }
  function confirmText(c) {
    switch (c.kind) {
      case 'RepointSite':
        return (
          'Re-point ' +
          c.name +
          ': Documents reads the site by its ID at the address in its SharePoint site record and updates the site and its libraries to it. Nothing in SharePoint changes.'
        );
      case 'RepointLibrary':
        return (
          'Re-point ' +
          c.name +
          ': Documents reads the library and its entry folder by their IDs and follows a rename or move. Existing record folders keep working. Nothing in SharePoint changes.'
        );
      case 'RemoveLibrary':
        return (
          'Remove ' +
          c.name +
          ' from Documents? Templates can no longer use it, and its unfinished folder and access work is cancelled. Nothing in SharePoint is deleted or changed: the library, its folders, its permissions and the Documents groups stay as they are.'
        );
      case 'RemoveSite':
        return (
          'Remove ' +
          c.name +
          ' from Documents? Nothing in SharePoint is deleted or changed. Remove its libraries first.'
        );
      case 'CancelAccessRun':
        return (
          'Cancel the access run for ' +
          c.name +
          '? Access already set in SharePoint stays as it is; nothing is undone or deleted. Apply access again when you are ready.'
        );
      case 'CancelSetup':
        return (
          'Cancel the setup of ' +
          c.name +
          '? Nothing in SharePoint is deleted. If SharePoint already created the library, add it as an existing library.'
        );
      default:
        return '';
    }
  }
  const confirmLabel = (c) =>
    ({
      RepointSite: 'Re-point',
      RepointLibrary: 'Re-point',
      RemoveLibrary: 'Remove library',
      RemoveSite: 'Remove site',
      CancelAccessRun: 'Cancel access run',
      CancelSetup: 'Cancel setup',
    })[c.kind];
  // Runs a confirmed site or library command.
  async function runConfirmed(c) {
    state.confirm = null;
    state.changes = [];
    if (c.kind === 'CancelAccessRun') {
      const result = await security({
        Command: 'CancelAccessRun',
        LibraryId: c.id,
        OperationKey: c.key,
      });
      const p = state.policies.get(c.id);
      if (p) p.result = result;
      issue(
        'The access run for ' +
          c.name +
          ' was cancelled. Nothing in SharePoint was undone or deleted.',
      );
      return;
    }
    if (c.kind === 'CancelSetup') {
      const result = await catalog({ Command: 'CancelSetup', Key: c.id });
      state.changes = result.Notices || [];
      state.progressSignature = null;
      const o = state.operations.get(c.id);
      if (result.Status === 'Cancelled') {
        state.operations.delete(c.id);
        state.completed.delete(c.id);
      } else if (o) {
        o.result = result;
        o.status = result.Status;
      }
      issue('The setup of ' + c.name + ' was cancelled.');
      return;
    }
    if (c.kind.startsWith('Remove')) {
      // Refused while a Draft or published template uses the library; the error names them.
      const removed = await catalog({ Command: c.kind, CatalogId: c.id });
      state.changes = removed.Notices || [];
      state.removed.add(c.id);
      state.policies.delete(c.id);
      forget(c.id, c.kind === 'RemoveSite' ? state.site?.asx_url : null);
      state.library = null;
      if (c.kind === 'RemoveSite') {
        state.site = null;
        state.libraries = [];
        await loadSites();
      } else await libraries();
      await window.AsxdAdmin?.refreshCatalog();
      issue(c.name + ' was removed from Documents.');
      return;
    }
    const result = await catalog({
      Command: c.kind,
      CatalogId: c.id,
      RequestId: crypto.randomUUID(),
    });
    state.progressSignature = null;
    state.completed.delete(result.Key);
    state.operations.set(result.Key, {
      name: c.name,
      kind: 'Repoint',
      command: c.kind,
      id: c.id,
    });
    issue('Re-pointing ' + c.name + '…');
  }
  async function loadSites(append = false) {
    const term = $('ad-search').value.trim().replace(/'/g, "''");
    const result = await page(
      'asx_site',
      append
        ? nextOptions(state.nextSites)
        : "?$select=asx_siteid,asx_name,asx_approved,asx_url,_asx_nativeid_value&$orderby=asx_name&$filter=statecode eq 0 and contains(asx_name,'" +
            term +
            "')",
    );
    state.sites = (append ? state.sites.concat(result.entities) : result.entities).filter(
      (s) => !state.removed.has(s.asx_siteid),
    );
    state.nextSites = result.nextLink || null;
    render();
  }
  async function libraries(append = false) {
    const result = await page(
      'asx_library',
      append
        ? nextOptions(state.nextLibraries)
        : '?$select=asx_libraryid,asx_name,asx_approved,asx_policyapplied&$orderby=asx_name&$filter=statecode eq 0 and _asx_siteid_value eq ' +
            state.site.asx_siteid,
    );
    state.libraries = (append ? state.libraries.concat(result.entities) : result.entities).filter(
      (l) => !state.removed.has(l.asx_libraryid),
    );
    state.nextLibraries = result.nextLink || null;
    if (state.library)
      state.library =
        state.libraries.find((l) => l.asx_libraryid === state.library.asx_libraryid) ||
        state.library;
    render();
  }
  async function selectSite(s) {
    state.selectedOperation = null;
    state.site = s;
    state.library = null;
    ['ad-library-form', 'ad-team-form', 'ad-existing-form'].forEach((id) => ($(id).hidden = true));
    issue('');
    await libraries();
    if (state.libraries.length) await selectLibrary(state.libraries[0]);
  }
  async function selectLibrary(l) {
    state.library = l;
    $('ad-team-form').hidden = true;
    if (!state.policies.has(l.asx_libraryid)) await loadPolicy(l);
    render();
  }
  async function loadPolicy(l) {
    const old = state.policies.get(l.asx_libraryid);
    if (changed(old)) return;
    const result = await security({ Command: 'GetPolicy', LibraryId: l.asx_libraryid });
    if (changed(state.policies.get(l.asx_libraryid))) return;
    const entries = (result.Policy?.Desired || []).map((e) => ({ ...e }));
    state.policies.set(l.asx_libraryid, { result, entries, saved: JSON.stringify(entries) });
    // The policy names its teams, so a team deleted in Dataverse is never read by ID.
    const named = new Map();
    for (const t of result.Teams || []) {
      const key = String(t.TeamId).toLowerCase();
      if (t.Deleted) state.deleted.set(key, t.Name || t.TeamId);
      else {
        state.deleted.delete(key);
        if (t.Name) named.set(key, t.Name);
      }
    }
    for (const e of entries)
      if (!state.teams.has(e.TeamId) && deletedTeam(e.TeamId) == null) {
        const name = named.get(String(e.TeamId).toLowerCase());
        // Its kind and any warning come with the team list (loadTeams).
        if (name) {
          state.teams.set(e.TeamId, name);
          continue;
        }
        const team = await xrm.WebApi.retrieveRecord(
          'team',
          e.TeamId,
          '?$select=name,teamtype,membershiptype,azureactivedirectoryobjectid',
        );
        const label = teamLabel(team);
        state.teams.set(e.TeamId, label.text);
        if (label.warning) state.warnings.set(e.TeamId, label.warning);
      }
    if (result.Status === 'Applied') {
      l.asx_policyapplied = true;
      await window.AsxdAdmin?.refreshCatalog();
    }
  }
  async function loadTeams() {
    const teams = [];
    // Owner, Entra security group and Microsoft 365 group teams; default and access teams stay out.
    let options =
      '?$select=teamid,name,teamtype,membershiptype,azureactivedirectoryobjectid&$orderby=name' +
      '&$filter=(teamtype eq 0 or teamtype eq 2 or teamtype eq 3) and isdefault eq false';
    do {
      const result = await page('team', options);
      teams.push(...result.entities);
      options = result.nextLink ? nextOptions(result.nextLink) : null;
    } while (options);
    ['ad-team-choice', 'ad-initial-team'].forEach((id) =>
      $(id).replaceChildren(
        opt('', id === 'ad-initial-team' ? 'No additional teams' : 'Select a team'),
      ),
    );
    teams.forEach((t) => {
      const label = teamLabel(t);
      state.teams.set(t.teamid, label.text);
      if (label.warning) state.warnings.set(t.teamid, label.warning);
      ['ad-team-choice', 'ad-initial-team'].forEach((id) => {
        const o = opt(t.teamid, label.text);
        o.disabled = !!label.reason;
        $(id).append(o);
      });
    });
  }
  // The catalog row a card is about, once its first read names it: a re-point's library or
  // site, or the library a setup created.
  function destination(o) {
    const probe = o.result?.Observation;
    if (o.kind === 'Repoint' && guid(probe?.CatalogId))
      return probe.ListId && probe.ListId !== noList
        ? { table: 'asx_library', id: probe.CatalogId }
        : { table: 'asx_site', id: probe.CatalogId };
    if (o.kind === 'LibrarySetup' && guid(o.result?.CatalogId) && o.result.CatalogId !== noList)
      return { table: 'asx_library', id: o.result.CatalogId };
    return null;
  }
  // Whether a site or library is Removed (statecode 1) or deleted: no active row has its ID.
  async function gone(target) {
    if (state.removed.has(target.id)) return true;
    const key = target.table + 'id',
      rows = await page(
        target.table,
        '?$select=' + key + '&$filter=statecode eq 0 and ' + key + ' eq ' + target.id,
      );
    return rows.entities.length === 0;
  }
  // Drops the cards of a site or library just removed: its setup or re-point cannot go on. A
  // removed site also takes the cards of the setups on its address.
  function forget(id, url) {
    for (const cards of [state.operations, state.completed])
      for (const [key, o] of cards)
        if (
          (destination(o)?.id ?? o.id) === id ||
          (url && (o.siteId === id || o.result?.Observation?.SiteId === id || o.url === url))
        )
          cards.delete(key);
    state.progressSignature = null;
  }
  // List IDs of the libraries on a discovery page that Documents already has as active rows,
  // matched by list ID, so a library renamed in SharePoint is still recognized. Removed rows
  // do not match: adding one again reactivates it.
  async function registeredLists(d) {
    const ids = (d.Observation?.Libraries || []).map((l) => l.Id).filter(guid),
      found = new Set();
    let options = ids.length
      ? '?$select=asx_listid&$filter=statecode eq 0 and (' +
        ids.map((v) => "asx_listid eq '" + v + "'").join(' or ') +
        ')'
      : null;
    while (options) {
      const rows = await page('asx_library', options);
      rows.entities.forEach((r) => found.add(String(r.asx_listid).toLowerCase()));
      options = rows.nextLink ? nextOptions(rows.nextLink) : null;
    }
    return found;
  }
  async function discoverActivity(append = false) {
    const rows = await page(
      'asx_operation',
      append
        ? nextOptions(state.nextActivity)
        : "?$select=asx_workkey,asx_workkind,asx_displayname,asx_siteurl,asx_status&$orderby=createdon desc&$filter=(asx_workkind eq 'SiteValidation' or asx_workkind eq 'LibrarySetup' or asx_workkind eq 'LibraryValidation' or asx_workkind eq 'LibraryDiscovery' or asx_workkind eq 'Repoint') and asx_status ne 'Discovered' and asx_status ne 'Applied' and asx_status ne 'Approved' and asx_status ne 'Cancelled' and asx_status ne 'Superseded'",
    );
    // A re-point's command and ID come from its probe when it is first read (see
    // refreshOperations), so a blocked re-point still offers "Re-point again" after a reload.
    // Until that read its card shows neither a message nor an action.
    for (const r of rows.entities)
      state.operations.set(r.asx_workkey, {
        name: r.asx_displayname,
        url: r.asx_workkind === 'Repoint' ? undefined : r.asx_siteurl,
        kind: r.asx_workkind,
        status: r.asx_status,
        unread: true,
      });
    state.nextActivity = rows.nextLink || null;
  }
  const pendingPolicy = () =>
    state.libraries.some((l) => {
      const p = state.policies.get(l.asx_libraryid);
      return p?.result.Policy?.OperationKey && !changed(p);
    });
  // Reads tracked setup operations and refreshes catalogs after confirmed completion.
  async function refreshOperations() {
    const observations = [];
    for (const [key, o] of state.operations) {
      if (elsewhere(key, o)) continue;
      observations.push([key, o, await catalog({ Command: 'Inspect', Key: key })]);
    }
    for (const [key, o, result] of observations) {
      o.result = result;
      o.status = result.Status;
      const probe = result.Observation;
      if (o.kind === 'Repoint' && probe?.CatalogId) {
        o.command = probe.ListId && probe.ListId !== noList ? 'RepointLibrary' : 'RepointSite';
        o.id = probe.CatalogId;
        o.siteId = probe.SiteId;
      }
    }
    // A setup or re-point of a site or library that was removed or deleted cannot go on, so
    // its card is dropped, also when it is found again after a reload, instead of staying as
    // needing attention. Each card's destination is checked once, when it is first known.
    for (const [key, o] of observations) {
      const target = destination(o);
      if (target && !o.checked) {
        o.checked = true;
        if (await gone(target)) {
          state.operations.delete(key);
          state.completed.delete(key);
          state.progressSignature = null;
        }
      }
      o.unread = false;
    }
    const live = observations.filter(([key]) => state.operations.has(key));
    const signature = JSON.stringify([
      state.site?.asx_siteid,
      state.selectedOperation,
      live.map(([key, o]) => [key, progressState(o)]),
    ]);
    if (signature === state.progressSignature && !pendingPolicy()) return;
    for (const [key, o, result] of live) {
      if (['Ready', 'Approved', 'Applied', 'Discovered'].includes(result.Status)) {
        state.completed.set(key, {
          ...o,
          url: o.url || result.Observation?.WebUrl || state.site?.asx_url,
        });
        while (state.completed.size > 4)
          state.completed.delete(state.completed.keys().next().value);
      }
      if (result.Status === 'Discovered') {
        state.operations.delete(key);
        if (state.site?.asx_siteid === result.Observation.SiteId) {
          state.discovery = result;
          state.registered = await registeredLists(result);
          showDiscovery();
        }
        continue;
      }
      if (o.kind === 'Repoint' && result.Status === 'Approved') {
        state.operations.delete(key);
        state.changes = result.Observation?.Changes || [];
        await loadSites();
        const site = state.sites.find((s) => s.asx_siteid === state.site?.asx_siteid);
        if (site) state.site = site;
        if (state.site) await libraries();
        await window.AsxdAdmin?.refreshCatalog();
        issue(o.name + ' re-pointed.');
        continue;
      }
      if (['Ready', 'Approved', 'Applied'].includes(result.Status)) {
        // A removed site or library added again is active again.
        state.removed.delete(result.CatalogId);
        await loadSites();
        if (o.kind === 'SiteValidation' && result.CatalogId) {
          const added = state.sites.find((s) => s.asx_siteid === result.CatalogId);
          if (added) await selectSite(added);
        } else if (state.site) await libraries();
        await window.AsxdAdmin?.refreshCatalog();
        state.operations.delete(key);
        issue(result.Issue ? o.name + ': ' + result.Issue : o.name + ' is ready.');
      }
    }
    for (const l of state.libraries) {
      const p = state.policies.get(l.asx_libraryid);
      if (p?.result.Policy?.OperationKey && !changed(p)) await loadPolicy(l);
    }
    state.progressSignature = signature;
    render();
  }
  async function start() {
    if (state.startPromise) return state.startPromise;
    if (!xrm?.WebApi || state.loaded) return;
    state.startPromise = action(async () => {
      await loadSites();
      if (state.sites.length) await selectSite(state.sites[0]);
      await loadTeams();
      await discoverActivity();
      state.loaded = true;
      // Shows the cards found again as loading, then reads them at once, so each shows its
      // own message and action. A failure here is left to the poll, which reads them again
      // and reports it.
      render();
      await refreshOperations().catch(() => {});
    }).finally(() => {
      state.startPromise = null;
    });
    return state.startPromise;
  }
  window.AsxdSites = {
    open: start,
    selectLibrary: async (id) => {
      await start();
      await action(async () => {
        if (!guid(id)) throw new Error('Select a library.');
        const l = await xrm.WebApi.retrieveRecord(
          'asx_library',
          id,
          '?$select=asx_libraryid,asx_name,_asx_siteid_value,asx_approved,asx_policyapplied,statecode',
        );
        // A link to a removed library does not bring it back into the list.
        if (l.statecode === 1 || state.removed.has(l.asx_libraryid))
          throw new Error('This library was removed from Documents. Add it again to use it.');
        const s = await xrm.WebApi.retrieveRecord(
          'asx_site',
          l._asx_siteid_value,
          '?$select=asx_siteid,asx_name,asx_approved,asx_url,_asx_nativeid_value',
        );
        if (!state.sites.some((v) => v.asx_siteid === s.asx_siteid)) state.sites.push(s);
        await selectSite(s);
        if (!state.libraries.some((v) => v.asx_libraryid === id)) state.libraries.push(l);
        await selectLibrary(l);
      });
    },
  };
  $('ad-more-activity').onclick = () =>
    action(async () => {
      await discoverActivity(true);
      await refreshOperations();
    });
  $('ad-search').onchange = () => action(() => loadSites());
  $('ad-more-sites').onclick = () => action(() => loadSites(true));
  $('ad-more-libraries').onclick = () => action(() => libraries(true));
  $('ad-add-site').onclick = () =>
    action(async () => {
      const result = await page(
        'sharepointsite',
        '?$select=sharepointsiteid,name,absoluteurl&$filter=statecode eq 0&$orderby=name',
      );
      $('ad-native').replaceChildren(opt('', 'Select a registered site'));
      result.entities.forEach((s) =>
        $('ad-native').append(opt(s.sharepointsiteid, s.name || s.absoluteurl)),
      );
      $('ad-site-form').hidden = false;
      $('ad-site-progress').textContent = result.nextLink
        ? 'Showing the first 50 sites. Search by name to narrow the list.'
        : '';
    });
  $('ad-native-search').onchange = () =>
    action(async () => {
      const term = $('ad-native-search').value.replace(/'/g, "''");
      const result = await page(
        'sharepointsite',
        "?$select=sharepointsiteid,name,absoluteurl&$filter=statecode eq 0 and contains(name,'" +
          term +
          "')&$orderby=name",
      );
      $('ad-native').replaceChildren(opt('', 'Select a registered site'));
      result.entities.forEach((s) =>
        $('ad-native').append(opt(s.sharepointsiteid, s.name || s.absoluteurl)),
      );
    });
  $('ad-cancel-site').onclick = () => {
    $('ad-site-form').hidden = true;
  };
  $('ad-validate').onclick = () =>
    action(async () => {
      const id = $('ad-native').value;
      if (!guid(id)) throw new Error('Select a registered site.');
      const name = Array.from($('ad-native').options).find((o) => o.value === id).textContent;
      const result = await catalog({
        Command: 'AddSite',
        NativeSiteId: id,
        Name: name,
        RequestId: crypto.randomUUID(),
      });
      state.progressSignature = null;
      state.completed.delete(result.Key);
      state.operations.set(result.Key, { name, kind: 'SiteValidation' });
      state.selectedOperation = result.Key;
      state.site = null;
      state.library = null;
      state.libraries = [];
      $('ad-site-form').hidden = true;
      issue('Site validation queued.');
    });
  $('ad-recheck').onclick = () =>
    action(async () => {
      if (!state.site?._asx_nativeid_value)
        throw new Error('Select the native site through Add site.');
      const result = await catalog({
        Command: 'AddSite',
        NativeSiteId: state.site._asx_nativeid_value,
        Name: state.site.asx_name,
        RequestId: crypto.randomUUID(),
      });
      state.progressSignature = null;
      state.completed.delete(result.Key);
      state.operations.set(result.Key, {
        name: state.site.asx_name,
        kind: 'SiteValidation',
        url: state.site.asx_url,
      });
    });
  // Adds an existing library; breakInheritance carries the admin's consent for an inheriting one.
  async function addLibrary(siteId, listId, name, url, breakInheritance) {
    const result = await catalog({
      Command: 'AddLibrary',
      SiteId: siteId,
      ListId: listId,
      Name: name,
      RequestId: crypto.randomUUID(),
      ...(breakInheritance ? { BreakInheritance: true } : {}),
    });
    state.consent = null;
    state.progressSignature = null;
    state.completed.delete(result.Key);
    state.operations.set(result.Key, { name, kind: 'LibraryValidation', url });
    $('ad-existing-form').hidden = true;
    issue('Validating ' + name + ' and setting up navigation.');
  }
  function showDiscovery() {
    const d = state.discovery;
    $('ad-existing-form').hidden = false;
    $('ad-existing-choices').replaceChildren();
    const asking = state.consent?.kind === 'inherit' ? state.consent : null;
    $('ad-existing-warning').textContent = asking
      ? asking.warnings.join(' ') + ' Select Confirm and add to continue.'
      : '';
    $('ad-existing-warning').hidden = !asking;
    // A library Documents already has is not offered again, also after a rename in SharePoint:
    // it is matched by list ID. A removed one is offered; adding it again reactivates it.
    const offered = d.Observation.Libraries.filter(
      (l) => !state.registered.has(String(l.Id).toLowerCase()),
    );
    for (const l of offered) {
      const confirming = asking?.key === l.Id,
        b = node('button', (confirming ? 'Confirm and add ' : 'Add ') + l.Title);
      b.type = 'button';
      b.onclick = () =>
        action(async () => {
          // A library that inherits the site's permissions is added only after the admin
          // confirms the warning shown in the page.
          const agreed = confirmed(
            'inherit',
            l.Id,
            l.HasUniqueRoleAssignments === false ? [inheritanceWarning] : [],
          );
          if (!agreed.go) {
            showDiscovery();
            return;
          }
          await addLibrary(d.Observation.SiteId, l.Id, l.Title, d.Observation.WebUrl, agreed.ack);
        });
      $('ad-existing-choices').append(b);
    }
    if (!offered.length)
      $('ad-existing-choices').append(
        node(
          'p',
          d.Observation.Libraries.length
            ? 'Every document library on this page is already in Documents.'
            : 'No document libraries found on this page.',
        ),
      );
    $('ad-existing-more').hidden = !d.Observation.NextLibraries;
  }
  $('ad-existing').onclick = () =>
    action(async () => {
      const result = await catalog({
        Command: 'DiscoverLibraries',
        SiteId: state.site.asx_siteid,
        RequestId: crypto.randomUUID(),
      });
      state.progressSignature = null;
      state.completed.delete(result.Key);
      state.operations.set(result.Key, {
        name: 'Existing libraries',
        kind: 'LibraryDiscovery',
        url: state.site.asx_url,
      });
      issue('Finding document libraries…');
    });
  $('ad-existing-more').onclick = () =>
    action(async () => {
      const d = state.discovery;
      const result = await catalog({
        Command: 'NextLibraries',
        Key: d.Key,
        RowVersion: d.RowVersion,
      });
      state.progressSignature = null;
      state.completed.delete(result.Key);
      state.operations.set(result.Key, {
        name: 'Existing libraries',
        kind: 'LibraryDiscovery',
        url: d.Observation.WebUrl,
      });
      $('ad-existing-form').hidden = true;
      issue('Loading the next library page…');
    });
  $('ad-existing-cancel').onclick = () => {
    $('ad-existing-form').hidden = true;
    if (state.consent?.kind === 'inherit') state.consent = null;
  };
  $('ad-create').onclick = () => {
    $('ad-library-form').hidden = false;
    $('ad-library-name').value = '';
    state.consent = null;
    render();
  };
  $('ad-cancel-library').onclick = () => {
    $('ad-library-form').hidden = true;
    state.consent = null;
    render();
  };
  $('ad-provision').onclick = () =>
    action(async () => {
      const name = $('ad-library-name').value.trim();
      if (!name) throw new Error('Enter a library name.');
      const team = $('ad-initial-team').value,
        access = $('ad-initial-access').value;
      const agreed = consent(
        'provision',
        JSON.stringify([state.site.asx_siteid, name, team, access]),
        team ? [team] : [],
      );
      if (!agreed.go) return;
      const result = await catalog({
        Command: 'CreateLibrary',
        SiteId: state.site.asx_siteid,
        Name: name,
        Entries: team ? [{ TeamId: team, Access: access }] : [],
        RequestId: crypto.randomUUID(),
        ...(agreed.ack ? { AcknowledgeBroaderAccess: true } : {}),
      });
      state.consent = null;
      state.progressSignature = null;
      state.completed.delete(result.Key);
      state.operations.set(result.Key, { name, kind: 'LibrarySetup', url: state.site.asx_url });
      $('ad-library-form').hidden = true;
      issue('Creating ' + name + '. Access and navigation setup follow automatically.');
    });
  $('ad-add-team').onclick = () => {
    $('ad-team-form').hidden = false;
  };
  $('ad-cancel-team').onclick = () => {
    $('ad-team-form').hidden = true;
  };
  $('ad-stage-team').onclick = () => {
    const p = policy(),
      id = $('ad-team-choice').value;
    if (!p || !guid(id)) {
      issue('Select a team.', true);
      return;
    }
    const existing = p.entries.find((e) => e.TeamId === id);
    if (existing) existing.Access = $('ad-team-access').value;
    else p.entries.push({ TeamId: id, Access: $('ad-team-access').value });
    $('ad-team-form').hidden = true;
    render();
  };
  $('ad-apply').onclick = () =>
    action(async () => {
      const p = policy();
      if (!p) throw new Error('Select a library.');
      const broader = warningsFor(
          p.entries.filter((e) => e.Access !== 'None').map((e) => e.TeamId),
        ),
        agreed = confirmed(
          'apply',
          applyKey(p),
          broader.concat(inherits(p) ? [reapplyWarning] : []),
        );
      if (!agreed.go) return;
      const latest = await security({
        Command: 'GetPolicy',
        LibraryId: state.library.asx_libraryid,
      });
      const signature = (entries) =>
        JSON.stringify(
          (entries || [])
            .map((e) => [e.TeamId.toLowerCase(), e.Access])
            .sort((a, b) => a[0].localeCompare(b[0])),
        );
      if (signature(latest.Policy?.Desired) !== signature(JSON.parse(p.saved)))
        throw new Error(
          'Library access changed since you opened it. Reload the page and review the current access before applying.',
        );
      // A queued run is replaced by this apply; one a flow holds, or one whose SharePoint write
      // is unanswered, finishes first and the change is applied right after it.
      const result = await security({
        Command: 'ApplyPolicy',
        LibraryId: state.library.asx_libraryid,
        RowVersion: latest.RowVersion || null,
        Entries: p.entries,
        ...(agreed.ack && broader.length ? { AcknowledgeBroaderAccess: true } : {}),
        ...(agreed.ack && inherits(p) ? { BreakInheritance: true } : {}),
      });
      state.consent = null;
      p.result = result;
      // As saved: deleted teams are left out.
      p.entries = (result.Policy?.Desired || p.entries).map((e) => ({ ...e }));
      p.saved = JSON.stringify(p.entries);
      issue(
        result.Policy?.ApplyPending
          ? pendingNotice
          : 'Access submitted. Team membership syncing is onboarded automatically.',
      );
    });
  $('ad-run-retry').onclick = () =>
    action(async () => {
      const p = policy(),
        key = p?.result.Policy?.OperationKey;
      if (!key) throw new Error('This library has no access run to retry.');
      p.result = await security({
        Command: 'RetryAccessRun',
        LibraryId: state.library.asx_libraryid,
        OperationKey: key,
      });
      issue('Access run queued to run again. It stops again if the cause remains.');
    });
  $('ad-run-cancel').onclick = () => {
    const key = policy()?.result.Policy?.OperationKey;
    if (!key || !state.library) return;
    state.confirm = {
      kind: 'CancelAccessRun',
      id: state.library.asx_libraryid,
      key,
      name: state.library.asx_name,
    };
    render();
  };
  $('ad-repoint-site').onclick = () => {
    if (!state.site) return;
    state.confirm = { kind: 'RepointSite', id: state.site.asx_siteid, name: state.site.asx_name };
    render();
  };
  $('ad-repoint-library').onclick = () => {
    if (!state.library) return;
    state.confirm = {
      kind: 'RepointLibrary',
      id: state.library.asx_libraryid,
      name: state.library.asx_name,
    };
    render();
  };
  $('ad-remove-site').onclick = () => {
    if (!state.site) return;
    state.confirm = { kind: 'RemoveSite', id: state.site.asx_siteid, name: state.site.asx_name };
    render();
  };
  $('ad-remove-library').onclick = () => {
    if (!state.library) return;
    state.confirm = {
      kind: 'RemoveLibrary',
      id: state.library.asx_libraryid,
      name: state.library.asx_name,
    };
    render();
  };
  $('ad-confirm-cancel').onclick = () => {
    state.confirm = null;
    render();
  };
  $('ad-confirm-go').onclick = () =>
    action(async () => {
      if (state.confirm) await runConfirmed(state.confirm);
    });
  $('ad-manage-connection').onclick = () =>
    xrm.Navigation.openUrl('https://make.powerautomate.com/');
  // Refreshes visible setup and policy status, preserving staged local access edits.
  async function poll() {
    try {
      if (
        !root.hidden &&
        !state.busy &&
        xrm?.WebApi &&
        (state.operations.size || pendingPolicy())
      ) {
        state.polling = true;
        try {
          state.pollPromise = refreshOperations();
          await state.pollPromise;
          state.pollFailures = 0;
          $('ad-poll-status').hidden = true;
        } finally {
          state.polling = false;
          state.pollPromise = null;
        }
      }
    } catch (e) {
      state.pollFailures++;
      const message = $('ad-poll-status');
      message.hidden = false;
      message.textContent =
        state.pollFailures < 3
          ? 'Live status is temporarily unavailable. Setup may still be running; checking again…'
          : 'Status updates are unavailable. Check your connection or access; the setup outcome is not yet known.';
    } finally {
      setTimeout(poll, 5000);
    }
  }
  render();
  if (!root.hidden) start();
  setTimeout(poll, 5000);
})();
