'use strict';
// Sites & access: sites, a table of their libraries, each library's team access in a drawer,
// and the setup activity on them. Results report in the line under the page header, or in the
// Add site panel or the drawer's footer; every command asks in the page, next to what asked.
// Text is only ever set with textContent.
(() => {
  const root = document.getElementById('access'),
    $ = (id) => document.getElementById(id),
    ui = window.AsxdUi,
    xrm = window.parent?.Xrm || window.Xrm;
  if (!root) return;
  const state = {
    sites: [],
    sitesLoaded: false,
    libraries: [],
    librariesLoaded: false,
    // Team ID to its name, and to its kind ("Owner team", "Entra group · all members").
    teams: new Map(),
    kinds: new Map(),
    // Teams deleted in Dataverse, by lower-case ID, with their last known name or their ID.
    deleted: new Map(),
    // Team ID to the broader-access warning its group carries.
    warnings: new Map(),
    policies: new Map(),
    // Policy reads in flight, by library ID, and the libraries whose read failed.
    reading: new Map(),
    unreadable: new Set(),
    // Library ID to the number of templates that use it; site ID to its number of libraries.
    usedBy: new Map(),
    siteCounts: null,
    site: null,
    // The library the access drawer shows, or null.
    library: null,
    // What the drawer shows: a library ID, the operation key of a setup that has no library row
    // yet, or null when it is closed.
    drawer: null,
    pill: false,
    busy: false,
    // The feedback area of the action that runs: access (under the page header), access-add
    // (Add site) or access-library (the access drawer).
    area: 'access-library',
    // Where focus goes when the action that runs closes or hides what was focused: element IDs,
    // the first one shown wins.
    focusAfter: null,
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
    // The setup cards were not redrawn because a confirmation was open on one of them.
    progressStale: false,
    completed: new Map(),
    pollFailures: 0,
    // What the last re-point changed, such as "old URL → new URL".
    changes: [],
    // Dynamics' validation status of each site's SharePoint site, by native site ID in lower
    // case: { value, label }, or null when it was read but not answered.
    validation: new Map(),
    // Sites and libraries removed in this page. They leave the lists at once, whatever a list
    // read returns, until they are added again.
    removed: new Set(),
    // List IDs, in lower case, of the libraries on the shown discovery page that Documents
    // already has as active rows.
    registered: new Set(),
    // Add site: the SharePoint sites found, the next page, the search they answer, the one
    // picked ({ id, name }) and the option the arrow keys are on.
    native: null,
    nativeSites: [],
    nativeNext: null,
    nativeTerm: '',
    nativeSearch: 0,
    nativeActive: -1,
  };
  const noList = '00000000-0000-0000-0000-000000000000';
  const INSTALL_DOCS =
    'https://github.com/ascentix-software/Ascentix-Documents/blob/main/docs/customer-installation.md';
  const node = (tag, text, css) => {
    const n = document.createElement(tag);
    if (text != null) n.textContent = text;
    if (css) n.className = css;
    return n;
  };
  const opt = (value, text) => {
    const o = node('option', text);
    o.value = value;
    return o;
  };
  const control = (text, key, onClick, css = 'secondary') => {
    const b = node('button', text, css);
    b.type = 'button';
    if (key) b.dataset.focusKey = key;
    b.onclick = onClick;
    return b;
  };
  const focusKey = (key) => document.querySelector('[data-focus-key="' + key + '"]')?.focus();
  const debounce = (fn, ms) => {
    let t;
    return () => {
      clearTimeout(t);
      t = setTimeout(fn, ms);
    };
  };
  // Owner teams sync their members. An Entra or Microsoft 365 group team is granted through its
  // group, so it is labelled with its kind (kept text #6). This mirrors the server
  // (TeamPrincipal), with team.teamtype 2 security group, 3 Microsoft 365 group and
  // membershiptype 1 members, 2 owners, 3 guests:
  // - teams SharePoint cannot identify are listed disabled with the reason;
  // - teams whose group reaches more people than the team carry a warning the admin confirms.
  const teamLabel = (t) => {
    if (t.teamtype !== 2 && t.teamtype !== 3)
      return {
        text: t.name,
        kind: t.teamtype === 0 ? 'Owner team' : null,
        reason: null,
        warning: null,
      };
    const security = t.teamtype === 2,
      kind = security ? 'Entra group' : 'Microsoft 365 group';
    const reason = !t.azureactivedirectoryobjectid
      ? 'no Microsoft Entra group object ID, so SharePoint cannot identify its group'
      : t.membershiptype === 3
        ? 'guests only: SharePoint has no sign-in claim for only the guests of a group'
        : null;
    if (reason)
      return { text: t.name + ' (' + kind + ') - ' + reason, kind, reason, warning: null };
    const warning =
      t.membershiptype === 1
        ? "The group's guests will also have access to this library."
        : security && t.membershiptype === 2
          ? 'All members of the group will have access to this library, not only its owners.'
          : null;
    const scope =
      t.membershiptype === 1
        ? ' · members + guests'
        : t.membershiptype === 2
          ? security
            ? ' · all members'
            : ' · owners'
          : '';
    return { text: t.name + ' (' + kind + scope + ')', kind: kind + scope, reason: null, warning };
  };
  // Names a team for its row: its name, and its kind for the sub-line.
  function nameTeam(teamId, t) {
    const label = teamLabel(t);
    state.teams.set(teamId, t.name);
    if (label.kind) state.kinds.set(teamId, label.kind);
    if (label.warning) state.warnings.set(teamId, label.warning);
    return label;
  }
  // The broader-access warnings of the chosen teams, which the admin confirms in the page.
  const warningsFor = (teamIds) => [
    ...new Set(teamIds.map((id) => state.warnings.get(id)).filter(Boolean)),
  ];
  // Shown before an admin adds a library that inherits its site's permissions (the server's
  // CatalogAdministration.InheritanceWarning) and before Apply access stops it again.
  const inheritanceWarning =
    'This library inherits permissions from the site. When you approve it, Documents stops the inheritance, keeps a copy of the current site permissions, and then manages team access on it.';
  const reapplyWarning =
    'This library inherits permissions from the site. When you apply access, Documents stops the inheritance, keeps a copy of the current site permissions, and then manages team access on it.';
  const removalNotice =
    'Removed teams lose the access Documents gave them. Access given another way, such as sharing links or site membership, is not changed.';
  // Applied while the queued run could not be replaced yet (PolicyDocument.ApplyPending).
  const pendingNotice = 'Saved · applies after the current run';
  // The last run applied the grants but left team membership unsynced (MembershipIncomplete).
  const incompleteNotice =
    'Needs attention: access is applied, but team membership was not synced. People removed from a team keep access, and people added get none, until this is resolved. See the notices below.';
  // The last access run stopped because the library inherits the site's permissions.
  const inherits = (p) => !!p?.result.Policy?.Inherits;
  // The library's queued access run stopped or waits to retry. A run stopped because the
  // library inherits permissions again is one of them: it needs attention like any other, and
  // Apply access with the acknowledgement is its remedy.
  const stuckRun = (p) => ['Blocked', 'RetryWait'].includes(p?.result.RunStatus);
  // A team deleted in Dataverse keeps its row while it still has access (GetPolicy's Teams marks
  // it); Dataverse can no longer read it by ID. The next access run removes its access.
  const deletedNotice = 'Documents removes its access the next time access is applied.';
  const deletedTeam = (teamId) => state.deleted.get(String(teamId).toLowerCase());
  const appliedAccess = (p, teamId) =>
    p.result.Policy?.Applied?.find((a) => a.TeamId === teamId)?.Access || 'None';
  // A row the admin sees: a team with access, or one removed in this page (Undo until Apply).
  // A team with no access, also a deleted one whose access is already removed, is not listed.
  const listed = (p, e) =>
    !!e.removed ||
    (deletedTeam(e.TeamId) != null ? appliedAccess(p, e.TeamId) !== 'None' : e.Access !== 'None');
  const deletedPending = (p) =>
    !!p?.entries.some((e) => deletedTeam(e.TeamId) != null && listed(p, e));
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
  // Reports in an area's feedback line (fb-<area>); an empty text clears it.
  const issue = (text, error = false, area = state.area) =>
    text ? ui.feedback(area, text, error ? 'error' : 'success') : ui.clearFeedback(area);
  const guid = (v) =>
    /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(v || '');
  // A server time ("/Date(ms)/" or ISO) as "YYYY-MM-DD hh:mm UTC".
  const when = (value) => {
    const ms = /^\/Date\((-?\d+)/.exec(value || '');
    const date = new Date(ms ? Number(ms[1]) : value);
    return isNaN(date) ? String(value) : date.toISOString().slice(0, 16).replace('T', ' ') + ' UTC';
  };
  // Dynamics validates a SharePoint site (sharepointsite.validationstatus 4, Valid) before the
  // document locations made under it get their site collection, which the record's Documents
  // tab needs. Until then the site holds Add existing library and Create library. A status not
  // read holds nothing: the server repairs any location made early.
  const VALID = 4,
    validationReason =
      "This SharePoint site hasn't been validated yet. Validate it in Settings › Document Management Settings, then choose Check again.";
  const nativeKey = (s) => String(s?._asx_nativeid_value || '').toLowerCase();
  const awaitingValidation = (s) => {
    const v = state.validation.get(nativeKey(s));
    return !!v && v.value !== VALID ? v : null;
  };
  function noteValidation(key, row) {
    state.validation.set(
      key,
      row?.validationstatus === undefined
        ? null
        : {
            value: row.validationstatus,
            label: row['validationstatus@OData.Community.Display.V1.FormattedValue'] || '',
          },
    );
  }
  // One query per 25 sites; a site the answer leaves out is noted as not answered.
  async function readValidation(list) {
    const keys = [...new Set(list.map(nativeKey).filter(guid))];
    for (const group of groups(keys)) {
      const rows = await readAll(
        'sharepointsite',
        '?$select=sharepointsiteid,validationstatus&$filter=' + anyOf('sharepointsiteid', group),
      );
      for (const key of group)
        noteValidation(
          key,
          rows.find((r) => String(r.sharepointsiteid).toLowerCase() === key),
        );
    }
  }
  async function checkValidation(s) {
    const key = nativeKey(s);
    if (guid(key))
      noteValidation(
        key,
        await xrm.WebApi.retrieveRecord('sharepointsite', key, '?$select=validationstatus'),
      );
  }
  // A refusal reads as the server's sentence, never as a raw response body.
  const api = (name, request) => ui.api(name, request);
  const catalog = (request) => api('asx_CatalogAdmin', request),
    security = (request) => api('asx_SecurityAdmin', request);
  async function page(table, options, size = 50) {
    return xrm.WebApi.retrieveMultipleRecords(table, options, size);
  }
  function nextOptions(next) {
    const url = new URL(next, xrm.Utility.getGlobalContext().getClientUrl());
    if (url.origin !== new URL(xrm.Utility.getGlobalContext().getClientUrl()).origin)
      throw new Error('Invalid catalog continuation.');
    return url.search;
  }
  // Runs one user action after any active poll finishes and renders its result. An error goes
  // to the feedback line of the area that acted.
  async function action(fn, area = 'access-library') {
    if (state.busy) return;
    state.busy = true;
    state.area = area;
    state.focusAfter = null;
    ui.clearFeedback(area);
    render();
    try {
      if (state.pollPromise) await state.pollPromise.catch(() => {});
      await fn();
    } catch (e) {
      issue(e.message || String(e), true, area);
      state.focusAfter = null;
    } finally {
      state.busy = false;
      render();
      land();
    }
  }
  // An action that succeeded and closed or hid the focused control (a form, the access
  // drawer, the site's actions) moves focus to the next sensible place it named, instead of
  // leaving it on the page body. Focus the admin moved elsewhere meanwhile stays there.
  const shown = (node) => !!node && node.isConnected && !node.closest('[hidden]') && !node.disabled;
  function land() {
    const targets = state.focusAfter;
    state.focusAfter = null;
    const active = document.activeElement;
    if (!targets || (active && active !== document.body && shown(active))) return;
    targets
      .map((id) => $(id))
      .find(shown)
      ?.focus();
  }
  function policy() {
    return state.library ? state.policies.get(state.library.asx_libraryid) : null;
  }
  function changed(p) {
    return p && JSON.stringify(p.entries) !== p.saved;
  }
  // A team's access as last read or applied, or undefined for a team added in this page.
  const savedAccess = (p, teamId) => JSON.parse(p.saved).find((e) => e.TeamId === teamId)?.Access;
  // The changes not applied yet: teams whose access changed, and teams removed or added.
  function changeCount(p) {
    if (!p) return 0;
    return p.entries.filter((e) => e.removed || savedAccess(p, e.TeamId) !== e.Access).length;
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
    // Documents is looking SharePoint up for a library create whose answer was lost.
    const checking = status === 'Reconciling';

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
    else if (checking) label = 'Checking SharePoint…';
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
    } else if (checking) {
      message = 'Checking SharePoint…';
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
  // Records a card's new result after one of its actions; the next draw shows it.
  function settled(o, result) {
    o.result = result;
    o.status = result.Status;
    state.progressSignature = null;
  }
  // Where a card's action reports: the drawer's footer while the drawer is open (a full-width
  // drawer covers the line under the page header), otherwise the line under the page header.
  const cardArea = () => (state.drawer ? 'access-library' : 'access');
  // A confirmation on a card. The cards are not redrawn while it is open, and are once it closes.
  async function cardAsk(invoker, options) {
    try {
      return await ui.confirmInline(invoker, options);
    } finally {
      if (state.progressStale) render();
    }
  }
  // Documents' SharePoint check for a setup whose create answer was lost: its
  // finding, each ambiguous candidate with the checks it failed, and the choices it offers.
  function recoveryChoices(card, actions, key, o, recovery) {
    const resolve = (choice, listId, said) =>
      action(async () => {
        settled(
          o,
          await catalog({
            Command: 'ResolveSetup',
            Key: key,
            Choice: choice,
            ...(listId ? { ListId: listId } : {}),
            RowVersion: o.result.RowVersion,
          }),
        );
        issue(said);
      }, cardArea());
    const named = (b) => {
      b.setAttribute('aria-label', b.textContent + ' for ' + o.name);
      return b;
    };
    const choices = recovery.Choices || [];
    if (choices.includes('UseCandidate')) {
      const list = node('ul', null, 'candidates');
      for (const c of recovery.Candidates || []) {
        const item = node('li');
        const checks = ui.candidateChecks(c);
        item.append(
          node('span', c.Title + ' · ' + c.Url + ' · created '),
          ui.time(c.CreatedUtc),
          node('span', checks ? ' · ' + checks : ''),
        );
        if (c.IsLibrary && c.CatalogEntry !== 'Conflict') {
          const row = node('span', null, 'row');
          row.setAttribute('data-actions', '');
          const use = control('Use this one', 'card:' + key + ':use:' + c.ListId, async () => {
            const ok = await cardAsk(use, {
              text:
                'Use ' +
                c.Title +
                ' at ' +
                c.Url +
                ' for this setup? Documents stops its permission inheritance if needed and manages its team access.',
              confirm: 'Use this library',
              keep: 'Keep looking',
            });
            if (ok)
              await resolve('UseLibrary', c.ListId, 'Using the existing library. Setup continues.');
          });
          use.setAttribute('aria-label', 'Use ' + c.Title + ' at ' + c.Url);
          row.append(use);
          item.append(row);
        }
        list.append(item);
      }
      card.append(list);
    }
    for (const choice of choices) {
      if (choice === 'UseLibrary')
        actions.append(
          named(
            control(
              'Use the library that was created',
              'card:' + key + ':use',
              () =>
                resolve(
                  'UseLibrary',
                  recovery.Candidates[0].ListId,
                  'Using the existing library. Setup continues.',
                ),
              'primary',
            ),
          ),
        );
      else if (choice === 'CreateAgain')
        actions.append(
          named(
            control(
              'Create it again',
              'card:' + key + ':create',
              () => resolve('CreateAgain', null, 'Creating the library again.'),
              'primary',
            ),
          ),
        );
      else if (choice === 'CheckAgain')
        actions.append(
          named(
            control('Check again', 'card:' + key + ':check', () =>
              action(async () => {
                settled(o, await catalog({ Command: 'RecheckSetup', Key: key }));
                issue('Checking SharePoint again.');
              }, cardArea()),
            ),
          ),
        );
      else if (choice === 'Cancel') actions.append(cancelSetupButton(key, o));
    }
  }
  function cancelSetupButton(key, o) {
    const cancel = control(
      'Cancel setup',
      'card:' + key + ':cancel',
      async () => {
        const ok = await cardAsk(cancel, {
          text: confirmText({ kind: 'CancelSetup', name: o.name }),
          confirm: 'Cancel setup',
          keep: 'Keep setup',
          danger: true,
        });
        if (ok)
          await action(
            () => runConfirmed({ kind: 'CancelSetup', id: key, name: o.name }),
            cardArea(),
          );
      },
      'danger',
    );
    cancel.setAttribute('aria-label', 'Cancel setup of ' + o.name);
    return cancel;
  }
  function monitorLink(key, o) {
    const link = control(
      'Open in Monitor',
      'card:' + key + ':monitor',
      () => ui.navigate('monitor', { operation: o.result?.RecoveryKey || key }),
      'link',
    );
    link.setAttribute('aria-label', 'Open ' + o.name + ' in Monitor');
    return link;
  }
  // The cards of this site, finished ones included, with where each one shows: the ID of the
  // library row it belongs to (a library's setup or re-point), the operation key of a setup that
  // has no library row yet (it gets a row of its own), or 'site' above the table.
  function siteCards() {
    const cards = [];
    for (const [key, o] of new Map([...state.completed, ...state.operations])) {
      if (elsewhere(key, o)) continue;
      // Found again after a reload and not read yet: its status alone does not say which
      // message and action apply, so nothing is offered until its first read.
      if (o.unread) {
        cards.push({ key, o, home: 'site' });
        continue;
      }
      const id =
        o.kind === 'Repoint'
          ? o.command === 'RepointLibrary'
            ? o.id
            : null
          : o.kind === 'LibrarySetup'
            ? destination(o)?.id
            : null;
      const home =
        id && state.libraries.some((l) => l.asx_libraryid === id)
          ? id
          : o.kind === 'LibrarySetup'
            ? key
            : 'site';
      cards.push({ key, o, home });
    }
    return cards;
  }
  // Site-level cards show above the table; a library's cards show in the drawer while it shows
  // that library or setup.
  function drawProgress(cards) {
    const site = $('ad-provision-progress'),
      drawer = $('ad-drawer-progress');
    // Never under an open confirmation: cardAsk redraws once it closes.
    state.progressStale = !!(
      site.querySelector('.confirm[role=group]') || drawer.querySelector('.confirm[role=group]')
    );
    if (state.progressStale) return;
    site.replaceChildren();
    drawer.replaceChildren();
    let loading = false;
    for (const { key, o, home } of cards) {
      if (o.unread) loading = true;
      else if (home === 'site') site.append(progressCard(key, o));
      else if (state.drawer && (home === state.drawer || key === state.drawer))
        drawer.append(progressCard(key, o));
    }
    if (loading) site.append(node('p', 'Loading…', 'muted'));
    $('ad-site-cards').hidden = site.children.length === 0;
  }
  // One operation's stage card: its stages, its message, and what can be done about it.
  function progressCard(key, o) {
    const p = progressState(o),
      card = node('section'),
      heading = node('div', null, 'row between'),
      badge = node('span', p.label, 'ad-progress-badge'),
      bar = node('progress'),
      list = node('ol', null, 'ad-progress-stages'),
      actions = node('div', null, 'row');
    card.className =
      'ad-progress-card' + (p.stopped ? ' ad-progress-attention' : p.done ? '' : ' is-loading');
    card.setAttribute('style', '--stage-count:' + p.stages.length);
    card.setAttribute('data-focus-row', '');
    actions.setAttribute('data-actions', '');
    heading.append(node('h3', o.name), badge);
    bar.max = p.stages.length - 1;
    bar.value = p.step;
    bar.setAttribute('aria-label', o.name + ' setup progress');
    bar.setAttribute(
      'aria-valuetext',
      p.done ? 'Complete' : p.label + '; stage ' + (p.step + 1) + ' of ' + p.stages.length,
    );
    p.stages.forEach((name, i) => {
      const item = node(
          'li',
          null,
          i < p.step || p.done ? 'is-complete' : i === p.step ? 'is-current' : '',
        ),
        mark = node('span', i < p.step || p.done ? '✓' : String(i + 1), 'ad-step-mark');
      mark.setAttribute('aria-hidden', 'true');
      if (i === p.step && !p.done) item.setAttribute('aria-current', 'step');
      item.append(mark, node('span', name));
      list.append(item);
    });
    const message = node('p', p.message, p.stopped ? 'ad-issue' : 'muted');
    card.append(heading, bar, list, message);
    const observed = o.result?.Observation;
    const recovery = o.kind === 'LibrarySetup' ? o.result?.Recovery : null;
    if (recovery && (p.stopped || p.status === 'Reconciling')) {
      message.textContent =
        p.status === 'Reconciling' ? 'Checking SharePoint…' : ui.recoverySentence(o.name, recovery);
      recoveryChoices(card, actions, key, o, recovery);
    } else if (p.status === 'Blocked' && o.kind === 'LibraryValidation' && observed?.Inherits) {
      // The card shows the server's refusal, which repeats the warning; this is the consent.
      actions.append(
        control('Stop inheritance and add', 'card:' + key + ':add', () =>
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
          }, cardArea()),
        ),
      );
    } else if (p.status === 'Blocked' && o.kind === 'Repoint') {
      // Re-point only reads SharePoint; once the cause is fixed it simply reads again. It is
      // never retried in place.
      if (o.command)
        actions.append(
          control('Re-point again', 'card:' + key + ':repoint', () =>
            action(async () => {
              state.operations.delete(key);
              state.completed.delete(key);
              await runConfirmed({ kind: o.command, id: o.id, name: o.name });
            }, cardArea()),
          ),
        );
    } else if (
      o.kind === 'LibrarySetup' &&
      ['Blocked', 'RecoveryRequired', 'RetryWait'].includes(p.status)
    ) {
      // Through the catalog API, so a Documents Security Administrator needs no Operator
      // role. The server keeps its rules: Retry of a create that may have reached SharePoint
      // is refused with the way out, and Cancel deletes nothing in SharePoint.
      const retry = control('Retry', 'card:' + key + ':retry', () =>
        action(async () => {
          settled(o, await catalog({ Command: 'RetrySetup', Key: key }));
          issue('Setup queued again.');
        }, cardArea()),
      );
      retry.setAttribute('aria-label', 'Retry setup of ' + o.name);
      actions.append(retry, cancelSetupButton(key, o), monitorLink(key, o));
    } else if (p.status === 'Blocked') {
      const retry = control('Retry', 'card:' + key + ':retry', () =>
        action(async () => {
          await api('asx_ManageWork', { Command: 'Retry', Key: o.result?.RecoveryKey || key });
          settled(o, { Status: 'Pending' });
          issue('Retry queued.');
        }, cardArea()),
      );
      retry.setAttribute('aria-label', 'Retry ' + o.name);
      actions.append(retry, monitorLink(key, o));
    } else if (p.stopped) actions.append(monitorLink(key, o));
    if (state.completed.has(key) && !state.operations.has(key)) {
      const dismiss = control('Dismiss', 'card:' + key + ':dismiss', () => {
        state.completed.delete(key);
        // A setup's drawer closes with its card; focus goes back to the libraries.
        if (state.drawer === key) {
          closeDrawer(false);
          state.focusAfter = ['ad-libraries-heading', 'ad-site-title'];
        }
        render();
        land();
      });
      dismiss.setAttribute('aria-label', 'Dismiss ' + o.name + ' progress');
      actions.append(dismiss);
    }
    if (actions.children.length) card.append(actions);
    return card;
  }
  // A team row: its name and kind, its access, or Removed · Undo until Apply. A row changed in
  // this page is tinted and says the access it had.
  function teamRow(p, e, running) {
    const row = node('div', null, 'team-row'),
      gone = deletedTeam(e.TeamId),
      name = gone != null ? 'Deleted team: ' + gone : state.teams.get(e.TeamId) || e.TeamId,
      key = 'team:' + e.TeamId,
      was = savedAccess(p, e.TeamId),
      label = node('div', null, 'team-label'),
      access = node('div', null, 'team-access'),
      act = node('div', null, 'team-act');
    row.setAttribute('role', 'listitem');
    row.setAttribute('data-focus-row', '');
    if (e.removed || was !== e.Access) row.classList.add('is-changed');
    const sub = [
      gone == null ? state.kinds.get(e.TeamId) : null,
      !e.removed && was && was !== e.Access ? 'was ' + was : null,
      running ? 'applying' : null,
    ].filter(Boolean);
    label.append(node('strong', name), node('span', sub.join(' · '), 'sub'));
    if (gone != null) label.append(node('div', deletedNotice, 'ad-issue'));
    if (e.removed) {
      row.classList.add('is-removed');
      const undo = control(
        'Undo',
        key + ':undo',
        () => {
          delete e.removed;
          render();
          focusKey(key + ':remove');
        },
        'link',
      );
      undo.setAttribute('aria-label', 'Undo removing ' + name);
      undo.disabled = state.busy;
      access.append('Removed · ', undo);
    } else {
      const select = node('select');
      ['Read', 'Contribute'].forEach((level) => select.append(opt(level, level)));
      // A deleted team's saved access may be one the picker does not offer.
      if (!['Read', 'Contribute'].includes(e.Access))
        select.append(opt(e.Access, e.Access === 'None' ? 'No access' : e.Access));
      select.value = e.Access;
      select.disabled = state.busy || gone != null;
      select.setAttribute('aria-label', 'Access for ' + name);
      select.dataset.focusKey = key + ':access';
      select.onchange = () => {
        e.Access = select.value;
        render();
      };
      access.append(select);
      if (gone == null) {
        const remove = control(
          '✕',
          key + ':remove',
          () => {
            e.removed = true;
            render();
            focusKey(key + ':undo');
          },
          'icon',
        );
        remove.setAttribute('aria-label', 'Remove ' + name);
        remove.disabled = state.busy;
        act.append(remove);
      }
    }
    row.append(label, access, act);
    return row;
  }
  // A setup's four stages as a small bar; its Access cell says the same in words.
  // One segment per stage, lit up to and including the current one, so "step 3 of 4" lights 3.
  function miniProgress(s) {
    const bar = node('div', null, 'mini-progress');
    bar.setAttribute('aria-hidden', 'true');
    s.stages.forEach((_, i) =>
      bar.append(node('span', null, i <= s.step || s.done ? 'done' : null)),
    );
    return bar;
  }
  const stageText = (o, s) =>
    o.kind === 'LibrarySetup'
      ? s.label + ' · step ' + (s.step + 1) + ' of ' + s.stages.length
      : s.label;
  const accessTone = (label) =>
    label === 'Access applied' ? 'ok' : label === 'Needs attention' ? 'blocked' : 'pending';
  // A library's row: its name (it opens the drawer), the templates that use it, its teams, and
  // its access: changes not applied, a setup or re-point under way, or the access label.
  function libraryRow(l, card) {
    const id = l.asx_libraryid,
      p = state.policies.get(id),
      tr = node('tr'),
      name = node('td', null, 'library'),
      access = node('td', null, 'access'),
      used = state.usedBy.get(id),
      b = control(l.asx_name, 'library:' + id, () => action(() => openDrawer(l, b)), 'link');
    b.classList.add('library-name');
    tr.setAttribute('data-focus-row', '');
    if (state.drawer === id) tr.classList.add('is-selected');
    name.append(b);
    const s = card && progressState(card.o),
      count = changeCount(p);
    if (s && card.o.kind === 'LibrarySetup' && !s.done) name.append(miniProgress(s));
    if (count > 0) {
      const status = ui.status('attention', ui.plural(count, 'change', 'changes') + ' not applied');
      status.classList.add('is-attention');
      access.append(status);
    } else if (s && !s.done)
      access.append(ui.status(s.stopped ? 'blocked' : 'pending', stageText(card.o, s)));
    else if (!l.asx_approved) access.append(ui.status('blocked', 'Needs attention'));
    else if (p || !state.unreadable.has(id))
      access.append(ui.status(accessTone(accessLabel(p)), accessLabel(p)));
    const teams = p && p.entries.filter((e) => !e.removed && listed(p, e)).length;
    tr.append(
      name,
      node(
        'td',
        used == null ? '' : used ? ui.plural(used, 'template', 'templates') : 'Not used',
        'used-by',
      ),
      node('td', p ? ui.plural(teams, 'team', 'teams') : '', 'teams'),
      access,
    );
    return tr;
  }
  // A setup with no library row yet: its name, its stage, and its card in the drawer.
  function setupRow(key, o) {
    const s = progressState(o),
      tr = node('tr'),
      name = node('td', null, 'library'),
      access = node('td', null, 'access'),
      b = control(o.name, 'setup:' + key, () => openSetup(key, b), 'link');
    b.classList.add('library-name');
    tr.setAttribute('data-focus-row', '');
    if (state.drawer === key) tr.classList.add('is-selected');
    name.append(b, miniProgress(s));
    access.append(
      ui.status(
        s.done ? 'ok' : s.stopped ? 'blocked' : 'pending',
        s.done ? s.label : stageText(o, s),
      ),
    );
    tr.append(name, node('td', 'Not used', 'used-by'), node('td', '', 'teams'), access);
    return tr;
  }
  // Every redraw keeps focus on the same control, by its data-focus-key.
  const render = () => ui.withFocus(draw);
  function draw() {
    const term = $('ad-search').value.trim();
    $('ad-sites').replaceChildren();
    const item = (button) => {
      const row = node('div');
      row.setAttribute('data-focus-row', '');
      row.append(button);
      $('ad-sites').append(row);
    };
    state.sites.forEach((s) => {
      const b = control(
          null,
          'site:' + s.asx_siteid,
          () => action(() => selectSite(s), 'access'),
          'ad-site',
        ),
        title = node('span', null, 'site-name'),
        count = state.siteCounts && (state.siteCounts.get(s.asx_siteid) || 0),
        sub = node('span', null, 'sub');
      if (s.asx_siteid === state.site?.asx_siteid) b.setAttribute('aria-current', 'true');
      title.append(node('strong', s.asx_name));
      if (!s.asx_approved || awaitingValidation(s)) {
        title.append(ui.dot('attention'));
        sub.textContent = 'Needs attention';
        sub.classList.add('attention');
      } else if (state.siteCounts) sub.textContent = ui.plural(count, 'library', 'libraries');
      b.append(title, sub);
      item(b);
    });
    let checking = 0;
    for (const [key, o] of state.operations) {
      if (o.kind !== 'SiteValidation') continue;
      checking++;
      item(
        control(
          o.name + ' · ' + progressState(o).label,
          'site:' + key,
          () => {
            state.selectedOperation = key;
            closeDrawer(false);
            state.site = null;
            state.library = null;
            state.libraries = [];
            render();
            action(refreshOperations, 'access');
          },
          'ad-site',
        ),
      );
    }
    // No sites: "No sites yet" holds ＋ Add; a search that finds none says so.
    const empty = state.sitesLoaded && !state.sites.length && !checking;
    $('ad-sites-empty').hidden = !empty;
    $('ad-sites-empty-text').textContent = term ? 'No sites match' : 'No sites yet';
    const home = empty && !term ? $('ad-sites-empty') : $('ad-sites-header');
    if ($('ad-add-site').parentNode !== home) home.append($('ad-add-site'));
    $('ad-more-activity').hidden = !state.nextActivity;
    $('ad-more-sites').hidden = !state.nextSites;
    $('ad-more-libraries').hidden = !state.nextLibraries;
    $('ad-more-libraries').parentNode.hidden = !state.nextLibraries;
    $('ad-site-title').textContent = state.site?.asx_name || 'Select or add a site';
    $('ad-site-meta').textContent = state.site
      ? [
          (state.site.asx_url || '').replace(/^https:\/\//, ''),
          state.site.asx_approved && !awaitingValidation(state.site) ? 'Ready' : 'Needs attention',
        ]
          .filter(Boolean)
          .join(' · ')
      : '';
    $('ad-site-actions').hidden = !state.site;
    $('ad-site-error').hidden = !state.site || state.site.asx_approved;
    const ready = !state.busy && !!state.site?.asx_approved;
    // A site Dynamics has not validated: its status line, and the library actions held with
    // one reason, kept in that line (ui.disable removes it once they are released).
    const waiting = state.site && awaitingValidation(state.site),
      waitingText = waiting
        ? 'Waiting for site validation' + (waiting.label ? ' (' + waiting.label + ')' : '')
        : '',
      reasonId = 'ad-validation-reason';
    $('ad-site-validation').hidden = !waiting;
    if ($('ad-site-validation-text').textContent !== waitingText)
      $('ad-site-validation-text').textContent = waitingText;
    if (waiting && !$(reasonId)) {
      const note = node('p', null, 'reason');
      note.id = reasonId;
      $('ad-site-validation').append(note);
    }
    for (const id of ['ad-create', 'ad-existing', 'ad-empty-create', 'ad-empty-existing']) {
      $(id).disabled = !ready;
      ui.disable($(id), reasonId, waiting ? validationReason : null);
    }
    ui.disable(
      $('ad-remove-site'),
      'ad-remove-site-reason',
      state.site && state.libraries.length ? 'Remove its libraries first' : null,
    );
    // The libraries table: a row per library, then a row per setup with no library row yet.
    const cards = siteCards(),
      onRow = (id) =>
        cards.find((c) => c.home === id && !progressState(c.o).done) ||
        cards.find((c) => c.home === id),
      setups = state.site ? cards.filter((c) => !c.o.unread && c.home === c.key) : [];
    $('ad-libraries').replaceChildren(
      ...state.libraries.map((l) => libraryRow(l, onRow(l.asx_libraryid))),
      ...setups.map((c) => setupRow(c.key, c.o)),
    );
    $('ad-libraries-card').hidden = !state.site || (!state.libraries.length && !setups.length);
    $('ad-libraries-empty').hidden =
      !state.site || !state.librariesLoaded || state.libraries.length > 0 || setups.length > 0;
    // The access drawer: a library's team access, or only the stage card of a setup.
    const p = policy(),
      setup = !state.library && cards.find((c) => c.key === state.drawer);
    $('ad-library-title').textContent = state.library?.asx_name || setup?.o.name || '';
    $('ad-library-status').textContent = 'Needs attention';
    $('ad-library-status').hidden = !state.library || state.library.asx_approved;
    $('ad-library-menu').hidden = !state.library;
    $('ad-drawer-access').hidden = !state.library;
    // A setup's drawer keeps its footer only to show what its card's actions report.
    $('ad-drawer-footer').hidden = !state.library && !$('fb-access-library').textContent;
    for (const id of ['ad-change-status', 'ad-apply']) $(id).hidden = !state.library;
    $('ad-team-rows').replaceChildren();
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
      const rows = p.entries.filter((e) => listed(p, e));
      rows.forEach((e) => $('ad-team-rows').append(teamRow(p, e, running)));
      if (!rows.length) {
        const none = node('div', 'No additional teams.', 'team-empty muted');
        none.setAttribute('role', 'listitem');
        $('ad-team-rows').append(none);
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
    $('ad-apply').textContent = inherits(p) ? 'Stop inheritance and apply' : 'Apply access';
    $('ad-add-team').disabled = state.busy || !p;
    const count = changeCount(p);
    $('ad-change-status').textContent = count
      ? ui.plural(count, 'change', 'changes') + ' not applied'
      : '';
    $('ad-discard').hidden = !count || !state.library;
    $('ad-discard').disabled = state.busy;
    // A stopped run comes first, as its own notice; a change waiting behind it applies once the
    // run is retried or cancelled. A run that waits to retry carries on by itself. Changes not
    // applied yet are counted in the footer instead.
    $('ad-run-status').textContent = !p
      ? 'Loading access…'
      : stuck
        ? (p.result.RunNotice ||
            (inherits(p)
              ? 'This library inherits permissions again.'
              : 'The access run stopped.')) +
          (run === 'RetryWait' && p.result.RunNextAttemptUtc
            ? ' Next check: ' + when(p.result.RunNextAttemptUtc) + '.'
            : '')
        : pending && !changed(p)
          ? pendingNotice
          : running
            ? 'Applying access and syncing members…'
            : inherits(p)
              ? 'Needs attention'
              : changed(p)
                ? ''
                : ['Missing', 'Removed'].includes(p.result.Status)
                  ? 'Apply to confirm this library’s access.'
                  : p.result.Status === 'NeedsReview'
                    ? 'The access run was cancelled. Apply access to run it again.'
                    : p.result.Status === 'Applied' && incomplete
                      ? incompleteNotice
                      : p.result.Status === 'Applied'
                        ? 'Access and team membership confirmed.'
                        : p.result.Status;
    $('ad-run-status').className =
      stuck ||
      (inherits(p) && !running) ||
      (incomplete && p?.result.Status === 'Applied' && !running && !changed(p))
        ? 'ad-issue'
        : 'muted';
    $('ad-run-actions').hidden = !stuck;
    $('ad-run-retry').disabled = state.busy;
    $('ad-run-cancel').disabled = state.busy;
    // What the last access sync skipped or reconciled, such as members SharePoint could not take.
    const notices = (p && p.result.Policy?.Notices) || [];
    $('ad-access-notices').replaceChildren(...notices.map((n) => node('li', n)));
    $('ad-access-notices').hidden = !notices.length;
    $('ad-changes').replaceChildren(...state.changes.map((t) => node('li', t)));
    $('ad-changes').hidden = !state.changes.length;
    ['ad-validate', 'ad-provision', 'ad-stage-team', 'ad-more-sites', 'ad-more-libraries'].forEach(
      (id) => ($(id).disabled = state.busy),
    );
    drawProgress(cards);
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
        return 'Remove ' + c.name + ' from Documents? Nothing in SharePoint is deleted or changed.';
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
  const keepLabel = (c) =>
    ({
      RepointSite: 'Keep current address',
      RepointLibrary: 'Keep current address',
      RemoveLibrary: 'Keep library',
      RemoveSite: 'Keep site',
      CancelAccessRun: 'Keep access run',
      CancelSetup: 'Keep setup',
    })[c.kind];
  // Asks in the page under what invoked it (the ⋯ button of the site or library, or the button
  // itself), then runs the command; its result or refusal goes to that area's feedback line.
  async function command(c, invoker, area) {
    if (state.busy) return;
    const ok = await ui.confirmInline(invoker, {
      text: confirmText(c),
      confirm: confirmLabel(c),
      keep: keepLabel(c),
      danger: !c.kind.startsWith('Repoint'),
    });
    if (ok) await action(() => runConfirmed(c), area);
  }
  // Runs a confirmed site or library command.
  async function runConfirmed(c) {
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
        // Its drawer closes with it; focus goes back to the libraries.
        if (state.drawer === c.id) {
          closeDrawer(false);
          state.focusAfter = ['ad-libraries-heading', 'ad-site-title'];
          state.area = 'access';
        }
      } else if (o) settled(o, result);
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
      closeDrawer(false);
      state.library = null;
      if (c.kind === 'RemoveSite') {
        state.site = null;
        state.libraries = [];
        await loadSites();
      } else await libraries();
      loadSiteCounts().catch(() => {});
      await window.AsxdAdmin?.refreshCatalog();
      // The drawer is gone with the library, so the site's line reports it, and focus goes to
      // the Libraries heading, or the site's when none is left; for a site, to Sites.
      issue(c.name + ' was removed from Documents.', false, 'access');
      state.focusAfter =
        c.kind === 'RemoveSite' ? ['ad-sites-heading'] : ['ad-libraries-heading', 'ad-site-title'];
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
    await readValidation(result.entities).catch(() => {});
    state.nextSites = result.nextLink || null;
    state.sitesLoaded = true;
    render();
  }
  // Reads the site's libraries. A refresh reads the usage and access of every one shown again;
  // otherwise only the libraries shown for the first time are read.
  async function libraries(append = false, refresh = false) {
    const before = new Set(state.libraries.map((l) => l.asx_libraryid));
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
    state.librariesLoaded = true;
    if (state.library)
      state.library =
        state.libraries.find((l) => l.asx_libraryid === state.library.asx_libraryid) ||
        state.library;
    render();
    preload(
      refresh ? state.libraries : state.libraries.filter((l) => !before.has(l.asx_libraryid)),
    );
  }
  // Reads which templates use these libraries and their access, without holding up the page: a
  // read that fails leaves its cells empty.
  function preload(list) {
    if (!list.length) return;
    loadUsedBy(list).catch(() => {});
    loadPolicies(list).catch(() => {});
  }
  // Every row of a query, page after page.
  async function readAll(table, options, size = 50) {
    const rows = [];
    while (options) {
      const result = await page(table, options, size);
      rows.push(...result.entities);
      options = result.nextLink ? nextOptions(result.nextLink) : null;
    }
    return rows;
  }
  // IDs in groups small enough for one query's filter.
  const groups = (ids) => {
    const out = [];
    for (let i = 0; i < ids.length; i += 25) out.push(ids.slice(i, i + 25));
    return out;
  };
  const anyOf = (column, ids) => ids.map((id) => column + ' eq ' + id).join(' or ');
  // The templates that use each library, by the server's rule for Remove library: a template
  // whose Draft revision, or whose published revision, has a destination on the library. A
  // replaced revision does not count.
  async function loadUsedBy(list) {
    const ids = list.map((l) => l.asx_libraryid),
      destinations = [],
      revisions = new Map(),
      published = new Map();
    for (const group of groups(ids))
      destinations.push(
        ...(await readAll(
          'asx_destination',
          '?$select=_asx_libraryid_value,_asx_revisionid_value&$filter=' +
            anyOf('_asx_libraryid_value', group),
        )),
      );
    const revisionIds = [
      ...new Set(destinations.map((d) => d._asx_revisionid_value).filter(Boolean)),
    ];
    for (const group of groups(revisionIds))
      for (const r of await readAll(
        'asx_revision',
        '?$select=asx_revisionid,asx_status,_asx_templateid_value&$filter=' +
          anyOf('asx_revisionid', group),
      ))
        revisions.set(r.asx_revisionid, r);
    if (revisions.size)
      for (const t of await readAll(
        'asx_template',
        '?$select=asx_templateid,_asx_publishedrevisionid_value',
      ))
        published.set(t.asx_templateid, t._asx_publishedrevisionid_value);
    const users = new Map(ids.map((id) => [id, new Set()]));
    for (const d of destinations) {
      const r = revisions.get(d._asx_revisionid_value),
        template = r?._asx_templateid_value;
      if (
        r &&
        (r.asx_status === 'Draft' || (template && published.get(template) === r.asx_revisionid))
      )
        users.get(d._asx_libraryid_value)?.add(template || r.asx_revisionid);
    }
    for (const [id, used] of users) state.usedBy.set(id, used.size);
    render();
  }
  // Each site's number of libraries, for the sites list: one ID column, 5,000 rows a page.
  async function loadSiteCounts() {
    const counts = new Map();
    for (const r of await readAll(
      'asx_library',
      '?$select=_asx_siteid_value&$filter=statecode eq 0',
      5000,
    ))
      counts.set(r._asx_siteid_value, (counts.get(r._asx_siteid_value) || 0) + 1);
    state.siteCounts = counts;
    render();
  }
  // Reads the access of the libraries shown, four at a time, so a site with many libraries
  // does not send them all at once. These reads name no teams and refresh no catalog: the
  // drawer does that for the library it opens.
  async function loadPolicies(list) {
    const queue = [...list];
    const next = async () => {
      for (let l = queue.shift(); l; l = queue.shift())
        await readPolicy(l).catch(() => state.unreadable.add(l.asx_libraryid));
    };
    await Promise.all([next(), next(), next(), next()]);
    render();
  }
  async function selectSite(s) {
    state.selectedOperation = null;
    closeDrawer(false);
    state.site = s;
    state.library = null;
    state.librariesLoaded = false;
    // A site not in the list read (a link to one of its libraries) has its status read here.
    if (!state.validation.has(nativeKey(s))) await checkValidation(s).catch(() => {});
    ['ad-library-form', 'ad-team-form', 'ad-existing-form'].forEach((id) => ($(id).hidden = true));
    ['access-library', 'access'].forEach((area) => ui.clearFeedback(area));
    await libraries(false, true);
  }
  // The drawer's library: its access read if it is not yet, then its teams named.
  async function selectLibrary(l) {
    state.library = l;
    $('ad-team-form').hidden = true;
    const id = l.asx_libraryid;
    if (!state.policies.has(id) || state.reading.has(id)) await readPolicy(l);
    await completePolicy(l);
    render();
  }
  // Reads a library's access, unless it has changes not applied. The policy names its teams, so
  // a team deleted in Dataverse is never read by ID. One read per library at a time.
  function readPolicy(l) {
    const id = l.asx_libraryid;
    if (state.reading.has(id)) return state.reading.get(id);
    const work = (async () => {
      const old = state.policies.get(id);
      if (changed(old)) return;
      const result = await security({ Command: 'GetPolicy', LibraryId: id });
      if (changed(state.policies.get(id))) return;
      const entries = (result.Policy?.Desired || []).map((e) => ({ ...e }));
      state.policies.set(id, { result, entries, saved: JSON.stringify(entries) });
      state.unreadable.delete(id);
      const named = new Map();
      for (const t of result.Teams || []) {
        const key = String(t.TeamId).toLowerCase();
        if (t.Deleted) state.deleted.set(key, t.Name || t.TeamId);
        else {
          state.deleted.delete(key);
          if (t.Name) named.set(key, t.Name);
        }
      }
      for (const e of entries) {
        const name = named.get(String(e.TeamId).toLowerCase());
        // Its kind and any warning come with the team list (loadTeams).
        if (name && !state.teams.has(e.TeamId)) state.teams.set(e.TeamId, name);
      }
      if (result.Status === 'Applied') l.asx_policyapplied = true;
    })().finally(() => {
      if (state.reading.get(id) === work) state.reading.delete(id);
    });
    state.reading.set(id, work);
    return work;
  }
  // Names the teams the policy did not name, then refreshes the catalog once the access is
  // applied. Once per read.
  async function completePolicy(l) {
    const p = state.policies.get(l.asx_libraryid);
    if (!p || p.complete) return;
    for (const e of p.entries)
      if (!state.teams.has(e.TeamId) && deletedTeam(e.TeamId) == null)
        nameTeam(
          e.TeamId,
          await xrm.WebApi.retrieveRecord(
            'team',
            e.TeamId,
            '?$select=name,teamtype,membershiptype,azureactivedirectoryobjectid',
          ),
        );
    p.complete = true;
    if (p.result.Status === 'Applied') await window.AsxdAdmin?.refreshCatalog();
  }
  // The access drawer, a dialog over the right of the page. It opens on a library's row, on a
  // setup's row, or from a link; closing it returns focus to what opened it, found again by its
  // data-focus-key when the table was redrawn meanwhile.
  let drawerHandle = null;
  function showDrawer(id, invoker) {
    // What the footer said was about the library or setup shown before.
    if (id !== state.drawer) ui.clearFeedback('access-library');
    drawerHandle = ui.sidePanel($('ad-drawer'), invoker, {
      onClose: () => {
        state.drawer = null;
        state.library = null;
        $('ad-team-form').hidden = true;
        render();
        // When the row that opened it is gone (a setup that became a library row), focus goes to
        // the libraries instead of staying in the closed drawer. This runs after the panel gave
        // focus back, so a control that took it keeps it.
        Promise.resolve().then(() => {
          if (!shown(document.activeElement))
            ['ad-libraries-heading', 'ad-site-title']
              .map((id) => $(id))
              .find(shown)
              ?.focus();
        });
      },
    });
    state.drawer = id;
  }
  async function openDrawer(l, invoker) {
    showDrawer(l.asx_libraryid, invoker);
    state.library = l;
    render();
    await selectLibrary(l);
  }
  function openSetup(key, invoker) {
    showDrawer(key, invoker);
    render();
  }
  function closeDrawer(focus = true) {
    drawerHandle?.close(focus);
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
      const label = nameTeam(t.teamid, t);
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
  // Tracks an operation's card. One found again after a reload is unread until its first read.
  function track(key, o) {
    state.progressSignature = null;
    state.completed.delete(key);
    state.operations.set(key, o);
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
      track(r.asx_workkey, {
        name: r.asx_displayname,
        url: r.asx_workkind === 'Repoint' ? undefined : r.asx_siteurl,
        kind: r.asx_workkind,
        status: r.asx_status,
        unread: true,
      });
    state.nextActivity = rows.nextLink || null;
  }
  // Whether a library's access is read again on the next poll: its run is queued or applying
  // and it has no changes not applied. A run that stopped or waits to retry waits for the admin,
  // so it is read again only while its library's drawer is open.
  const polled = (l) => {
    const p = state.policies.get(l.asx_libraryid);
    return (
      !!p?.result.Policy?.OperationKey &&
      !changed(p) &&
      (!stuckRun(p) || l.asx_libraryid === state.drawer)
    );
  };
  const pendingPolicy = () => state.libraries.some(polled);
  // Reads tracked setup operations and refreshes catalogs after confirmed completion. What
  // they report goes to the site's feedback line, which is always shown.
  async function refreshOperations() {
    const observations = [];
    for (const [key, o] of state.operations) {
      if (elsewhere(key, o)) continue;
      observations.push([key, o, await catalog({ Command: 'Inspect', Key: key })]);
    }
    for (const [, o, result] of observations) {
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
        issue(o.name + ' re-pointed.', false, 'access');
        continue;
      }
      if (['Ready', 'Approved', 'Applied'].includes(result.Status)) {
        // A removed site or library added again is active again.
        state.removed.delete(result.CatalogId);
        await loadSites();
        loadSiteCounts().catch(() => {});
        if (o.kind === 'SiteValidation' && result.CatalogId) {
          const added = state.sites.find((s) => s.asx_siteid === result.CatalogId);
          if (added) await selectSite(added);
        } else if (state.site) await libraries();
        await window.AsxdAdmin?.refreshCatalog();
        state.operations.delete(key);
        issue(result.Issue ? o.name + ': ' + result.Issue : o.name + ' is ready.', false, 'access');
      }
    }
    for (const l of state.libraries)
      if (polled(l)) {
        await readPolicy(l);
        // Only the drawer's library names its teams and refreshes the catalog.
        if (l.asx_libraryid === state.drawer) await completePolicy(l);
      }
    state.progressSignature = signature;
    render();
  }
  async function start() {
    if (state.startPromise) return state.startPromise;
    if (!xrm?.WebApi || state.loaded) return;
    state.startPromise = action(async () => {
      // The header's Monitor problem count, read once per page load.
      if (!state.pill) {
        state.pill = true;
        ui.problemPill($('ad-site-problems'));
      }
      await loadSites();
      loadSiteCounts().catch(() => {});
      // A link to a library opens that library's site instead of the first one.
      if (state.sites.length && !ui.deeplink()?.library) await selectSite(state.sites[0]);
      await loadTeams();
      await discoverActivity();
      state.loaded = true;
      // Shows the cards found again as loading, then reads them at once, so each shows its
      // own message and action. A failure here is left to the poll, which reads them again
      // and reports it.
      render();
      await refreshOperations().catch(() => {});
    }, 'access').finally(() => {
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
        await openDrawer(l, $('ad-site-title'));
      }, 'access');
    },
    redraw: render,
    // Tracks a setup or other operation as a card and reads it at once (as a reload finds one).
    trackOperation: async (key, name, kind) => {
      track(key, { name, kind, url: state.site?.asx_url });
      await refreshOperations();
    },
  };
  // The shell starts this tab when it is the one shown, then opens the library a link names.
  ui.onTab('access', async () => {
    await start();
    const link = ui.deeplink();
    if (link?.library) await window.AsxdSites.selectLibrary(link.library);
  });
  ui.menu($('ad-site-menu'), $('ad-site-menu-list'));
  ui.menu($('ad-library-menu'), $('ad-library-menu-list'));
  $('ad-more-activity').onclick = () =>
    action(async () => {
      await discoverActivity(true);
      await refreshOperations();
    }, 'access');
  $('ad-search').oninput = debounce(() => action(() => loadSites(), 'access'), 300);
  $('ad-more-sites').onclick = () => action(() => loadSites(true), 'access');
  $('ad-more-libraries').onclick = () => action(() => libraries(true), 'access');

  // Add site: a combobox that searches the SharePoint sites of Dataverse document management
  // as the admin types (300 ms after the last key), 20 at a time.
  async function searchNative(append = false) {
    const term = $('ad-native-search').value.trim();
    const search = ++state.nativeSearch;
    try {
      const result = await xrm.WebApi.retrieveMultipleRecords(
        'sharepointsite',
        append
          ? nextOptions(state.nativeNext)
          : '?$select=sharepointsiteid,name,absoluteurl&$filter=statecode eq 0' +
              (term ? " and contains(name,'" + term.replace(/'/g, "''") + "')" : '') +
              '&$orderby=name',
        20,
      );
      // An older search answering late does not replace a newer one.
      if (search !== state.nativeSearch) return;
      state.nativeSites = append ? state.nativeSites.concat(result.entities) : result.entities;
      state.nativeNext = result.nextLink || null;
      state.nativeTerm = term;
      drawNative();
    } catch (e) {
      issue(e.message || String(e), true, 'access-add');
    }
  }
  function drawNative() {
    const list = $('ad-native-list'),
      box = $('ad-native-search');
    list.replaceChildren();
    state.nativeActive = -1;
    box.removeAttribute('aria-activedescendant');
    list.hidden = false;
    if (!state.nativeSites.length) {
      // A sentence, not options: the list is no listbox while it holds no sites.
      list.removeAttribute('role');
      box.setAttribute('aria-expanded', 'false');
      const empty = node(
        'li',
        state.nativeTerm
          ? 'No sites match.'
          : 'No SharePoint sites are set up in Dataverse document management yet. ',
        'ad-native-empty',
      );
      if (!state.nativeTerm) {
        const how = node('a', 'How to set one up');
        how.setAttribute('href', INSTALL_DOCS + '#configure-and-enable');
        how.setAttribute('target', '_blank');
        how.setAttribute('rel', 'noopener');
        empty.append(how);
      }
      list.append(empty);
      return;
    }
    list.setAttribute('role', 'listbox');
    box.setAttribute('aria-expanded', 'true');
    state.nativeSites.forEach((s, n) => {
      const option = node('li', s.name || s.absoluteurl);
      option.id = 'ad-native-' + n;
      option.setAttribute('role', 'option');
      option.setAttribute('aria-selected', 'false');
      option.onclick = () => pickNative(n);
      list.append(option);
    });
    if (state.nativeNext) {
      const more = node('li', 'Show more sites', 'ad-native-more');
      more.id = 'ad-native-more';
      more.setAttribute('role', 'option');
      more.setAttribute('aria-selected', 'false');
      more.onclick = () => searchNative(true);
      list.append(more);
    }
  }
  function pickNative(n) {
    const s = state.nativeSites[n];
    state.native = { id: s.sharepointsiteid, name: s.name || s.absoluteurl };
    closeNative();
    $('ad-native-search').value = state.native.name;
    $('ad-native-search').focus();
  }
  function closeNative() {
    $('ad-native-list').hidden = true;
    $('ad-native-search').setAttribute('aria-expanded', 'false');
    $('ad-native-search').removeAttribute('aria-activedescendant');
    state.nativeActive = -1;
  }
  function activateNative(index) {
    const options = $('ad-native-list').querySelectorAll('[role=option]');
    options.forEach((o, i) => o.setAttribute('aria-selected', String(i === index)));
    state.nativeActive = index;
    $('ad-native-search').setAttribute('aria-activedescendant', options[index].id);
    options[index].scrollIntoView?.({ block: 'nearest' });
  }
  const searchSoon = debounce(() => searchNative(), 300);
  $('ad-native-search').oninput = () => {
    // Typing changes the pick: Add and check site needs a site chosen from the list again.
    state.native = null;
    searchSoon();
  };
  // Down moves into the list, Up and Down move in it, Enter picks, Escape clears.
  $('ad-native-search').addEventListener('keydown', (event) => {
    const options = $('ad-native-list').querySelectorAll('[role=option]');
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      if (!options.length || $('ad-native-list').hidden) return;
      event.preventDefault();
      const step = event.key === 'ArrowDown' ? 1 : options.length - 1;
      activateNative(
        state.nativeActive < 0
          ? event.key === 'ArrowDown'
            ? 0
            : options.length - 1
          : (state.nativeActive + step) % options.length,
      );
    } else if (event.key === 'Enter') {
      if (state.nativeActive < 0) return;
      event.preventDefault();
      options[state.nativeActive].onclick();
    } else if (event.key === 'Escape') {
      // With nothing to clear, Escape closes the Add site panel.
      if ($('ad-native-list').hidden && !$('ad-native-search').value) return;
      event.preventDefault();
      $('ad-native-search').value = '';
      state.native = null;
      closeNative();
    }
  });
  // Add site opens as a side panel; Cancel and Escape close it and return focus to ＋ Add.
  let sitePanel = null;
  $('ad-add-site').onclick = () => {
    sitePanel = ui.sidePanel($('ad-site-form'), $('ad-add-site'), { onClose: closeNative });
    state.native = null;
    $('ad-native-search').value = '';
    ui.clearFeedback('access-add');
    return searchNative();
  };
  $('ad-cancel-site').onclick = () => sitePanel?.close();
  $('ad-validate').onclick = () =>
    action(async () => {
      const pick = state.native;
      if (!pick) throw new Error('Choose a SharePoint site from the list.');
      const result = await catalog({
        Command: 'AddSite',
        NativeSiteId: pick.id,
        Name: pick.name,
        RequestId: crypto.randomUUID(),
      });
      track(result.Key, { name: pick.name, kind: 'SiteValidation' });
      state.selectedOperation = result.Key;
      state.site = null;
      closeDrawer(false);
      state.library = null;
      state.libraries = [];
      sitePanel?.close(false);
      issue('Checking ' + pick.name + '.', false, 'access');
      state.focusAfter = ['ad-add-site', 'ad-sites-heading'];
    }, 'access-add');
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
      track(result.Key, {
        name: state.site.asx_name,
        kind: 'SiteValidation',
        url: state.site.asx_url,
      });
    }, 'access');
  // Check again reads the site's validation status again. Once Valid, the line goes and focus
  // moves to the released Add existing library.
  $('ad-check-validation').onclick = () =>
    action(async () => {
      await checkValidation(state.site);
      state.focusAfter = ['ad-existing', 'ad-site-title'];
    }, 'access');
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
    track(result.Key, { name, kind: 'LibraryValidation', url });
    $('ad-existing-form').hidden = true;
    issue('Checking ' + name + ' and setting up its navigation.', false, 'access');
    state.focusAfter = ['ad-existing', 'ad-site-title'];
  }
  // An inheriting library is added only after the admin confirms, in the page, that Documents
  // stops the inheritance.
  async function addExisting(d, l, invoker) {
    const inheriting = l.HasUniqueRoleAssignments === false;
    if (
      inheriting &&
      !(await ui.confirmInline(invoker, {
        text: inheritanceWarning,
        confirm: 'Stop inheritance and add',
        keep: 'Not now',
      }))
    )
      return;
    await action(
      () => addLibrary(d.Observation.SiteId, l.Id, l.Title, d.Observation.WebUrl, inheriting),
      'access',
    );
  }
  function showDiscovery() {
    const d = state.discovery;
    $('ad-existing-form').hidden = false;
    $('ad-existing-choices').replaceChildren();
    // A library Documents already has is not offered again, also after a rename in SharePoint:
    // it is matched by list ID. A removed one is offered; adding it again reactivates it.
    const offered = d.Observation.Libraries.filter(
      (l) => !state.registered.has(String(l.Id).toLowerCase()),
    );
    for (const l of offered) {
      const b = control('Add ' + l.Title, null, () => addExisting(d, l, b));
      $('ad-existing-choices').append(b);
    }
    if (!offered.length)
      $('ad-existing-choices').append(
        node(
          'p',
          d.Observation.Libraries.length
            ? 'All libraries on this site are already added'
            : 'No document libraries on this site',
        ),
      );
    $('ad-existing-more').hidden = !d.Observation.NextLibraries;
  }
  $('ad-existing').onclick = () =>
    !ui.blocked($('ad-existing')) &&
    action(async () => {
      const result = await catalog({
        Command: 'DiscoverLibraries',
        SiteId: state.site.asx_siteid,
        RequestId: crypto.randomUUID(),
      });
      track(result.Key, {
        name: 'Existing libraries',
        kind: 'LibraryDiscovery',
        url: state.site.asx_url,
      });
      issue('Finding document libraries…', false, 'access');
    }, 'access');
  $('ad-existing-more').onclick = () =>
    action(async () => {
      const d = state.discovery;
      const result = await catalog({
        Command: 'NextLibraries',
        Key: d.Key,
        RowVersion: d.RowVersion,
      });
      track(result.Key, {
        name: 'Existing libraries',
        kind: 'LibraryDiscovery',
        url: d.Observation.WebUrl,
      });
      $('ad-existing-form').hidden = true;
      issue('Loading more libraries…', false, 'access');
    }, 'access');
  $('ad-existing-cancel').onclick = () => {
    $('ad-existing-form').hidden = true;
    $('ad-existing').focus();
  };
  $('ad-create').onclick = () => {
    if (ui.blocked($('ad-create'))) return;
    $('ad-library-form').hidden = false;
    $('ad-library-name').value = '';
    ui.clearFeedback('access');
    $('ad-library-name').focus();
  };
  $('ad-cancel-library').onclick = () => {
    $('ad-library-form').hidden = true;
    $('ad-create').focus();
  };
  $('ad-empty-create').onclick = () => $('ad-create').onclick();
  $('ad-empty-existing').onclick = () => $('ad-existing').onclick();
  // Create library asks first when its team's group reaches more people than the team.
  $('ad-provision').onclick = async () => {
    const name = $('ad-library-name').value.trim();
    if (!name) {
      issue('Enter a library name.', true, 'access');
      $('ad-library-name').focus();
      return;
    }
    const team = $('ad-initial-team').value,
      access = $('ad-initial-access').value,
      warnings = warningsFor(team ? [team] : []);
    if (
      warnings.length &&
      !(await ui.confirmInline($('ad-provision'), {
        text: warnings.join(' '),
        confirm: 'Create library',
        keep: 'Not now',
      }))
    )
      return;
    await action(async () => {
      const result = await catalog({
        Command: 'CreateLibrary',
        SiteId: state.site.asx_siteid,
        Name: name,
        Entries: team ? [{ TeamId: team, Access: access }] : [],
        RequestId: crypto.randomUUID(),
        ...(warnings.length ? { AcknowledgeBroaderAccess: true } : {}),
      });
      track(result.Key, { name, kind: 'LibrarySetup', url: state.site.asx_url });
      $('ad-library-form').hidden = true;
      // The form closes, so the site's line reports it and focus returns to ＋ Create library.
      issue('Creating ' + name + '.', false, 'access');
      state.focusAfter = ['ad-create', 'ad-site-title'];
    }, 'access');
  };
  $('ad-add-team').onclick = () => {
    $('ad-team-form').hidden = false;
    $('ad-team-choice').focus();
  };
  $('ad-cancel-team').onclick = () => {
    $('ad-team-form').hidden = true;
    $('ad-add-team').focus();
  };
  $('ad-drawer-close').onclick = () => closeDrawer();
  // Discard puts the teams back as last read or applied.
  function discardChanges(p) {
    p.entries = JSON.parse(p.saved);
    $('ad-team-form').hidden = true;
    render();
    // Apply stays enabled when the library needs it anyway; otherwise focus goes to the heading.
    ($('ad-apply').disabled ? $('ad-library-title') : $('ad-apply')).focus();
  }
  $('ad-discard').onclick = () => {
    const p = policy();
    if (p && !state.busy) discardChanges(p);
  };
  $('ad-stage-team').onclick = () => {
    const p = policy(),
      id = $('ad-team-choice').value;
    if (!p || !guid(id)) {
      issue('Select a team.', true, 'access-library');
      return;
    }
    const existing = p.entries.find((e) => e.TeamId === id);
    if (existing) {
      existing.Access = $('ad-team-access').value;
      delete existing.removed;
    } else p.entries.push({ TeamId: id, Access: $('ad-team-access').value });
    $('ad-team-form').hidden = true;
    render();
  };
  // Apply access asks first when a chosen team's group reaches more people, when it stops the
  // library's inheritance, or when teams are removed; then it sends the access as shown.
  $('ad-apply').onclick = async () => {
    const p = policy();
    if (!p) return;
    const removals = p.entries.some((e) => e.removed),
      broader = warningsFor(
        p.entries.filter((e) => !e.removed && e.Access !== 'None').map((e) => e.TeamId),
      ),
      inheriting = inherits(p),
      warnings = broader.concat(inheriting ? [reapplyWarning] : []);
    if (
      (warnings.length || removals) &&
      !(await ui.confirmInline($('ad-apply'), {
        text: warnings.concat(removals ? [removalNotice] : []).join(' '),
        confirm: 'Apply access',
        keep: 'Not now',
      }))
    )
      return;
    await action(async () => {
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
      // is unanswered, finishes first and the change is applied right after it. A removed team
      // is sent with no access.
      const result = await security({
        Command: 'ApplyPolicy',
        LibraryId: state.library.asx_libraryid,
        RowVersion: latest.RowVersion || null,
        Entries: p.entries.map((e) => ({
          TeamId: e.TeamId,
          Access: e.removed ? 'None' : e.Access,
        })),
        ...(broader.length ? { AcknowledgeBroaderAccess: true } : {}),
        ...(inheriting ? { BreakInheritance: true } : {}),
      });
      p.result = result;
      // As saved: deleted teams are left out.
      p.entries = (result.Policy?.Desired || p.entries).map((e) => ({
        TeamId: e.TeamId,
        Access: e.removed ? 'None' : e.Access,
      }));
      p.saved = JSON.stringify(p.entries);
      issue('Access submitted.');
    }, 'access-library');
  };
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
    }, 'access-library');
  $('ad-run-cancel').onclick = () => {
    const key = policy()?.result.Policy?.OperationKey;
    if (!key || !state.library) return;
    command(
      {
        kind: 'CancelAccessRun',
        id: state.library.asx_libraryid,
        key,
        name: state.library.asx_name,
      },
      $('ad-run-cancel'),
      'access-library',
    );
  };
  // Menu items start their command and return at once: the command waits for the admin's answer.
  $('ad-repoint-site').onclick = () => {
    if (!state.site) return;
    command(
      { kind: 'RepointSite', id: state.site.asx_siteid, name: state.site.asx_name },
      $('ad-site-menu'),
      'access',
    );
  };
  $('ad-remove-site').onclick = () => {
    if (!state.site || ui.blocked($('ad-remove-site'))) return;
    command(
      { kind: 'RemoveSite', id: state.site.asx_siteid, name: state.site.asx_name },
      $('ad-site-menu'),
      'access',
    );
  };
  $('ad-repoint-library').onclick = () => {
    if (!state.library) return;
    command(
      { kind: 'RepointLibrary', id: state.library.asx_libraryid, name: state.library.asx_name },
      $('ad-library-menu'),
      'access-library',
    );
  };
  $('ad-remove-library').onclick = () => {
    if (!state.library) return;
    command(
      { kind: 'RemoveLibrary', id: state.library.asx_libraryid, name: state.library.asx_name },
      $('ad-library-menu'),
      'access-library',
    );
  };
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
  setTimeout(poll, 5000);
})();
