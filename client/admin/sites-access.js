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
  };
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
  // The last access run stopped because the library inherits the site's permissions.
  const inherits = (p) => !!p?.result.Policy?.Inherits;
  const applyKey = (p) => (state.library?.asx_libraryid || '') + JSON.stringify(p?.entries || []);
  const issue = (text, error = false) => {
    $('ad-message').textContent = text;
    $('ad-message').className = error ? 'ad-issue' : 'ad-status';
  };
  const guid = (v) =>
    /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(v || '');
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
      message = 'Setup completed.';
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
  function drawProgress() {
    const area = $('ad-provision-progress');
    area.replaceChildren();
    for (const [key, o] of new Map([...state.completed, ...state.operations])) {
      if (
        (state.selectedOperation && state.selectedOperation !== key) ||
        (state.site && o.url && state.site.asx_url !== o.url)
      )
        continue;
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
      $('ad-library-access').textContent = state.library.asx_policyapplied
        ? 'Access applied'
        : 'Access setup pending';
    }
    $('ad-teams').replaceChildren();
    if (p) {
      p.entries.forEach((e) => {
        const tr = node('tr'),
          cell = node('td'),
          select = node('select');
        ['None', 'Read', 'Contribute'].forEach((level) =>
          select.append(opt(level, level === 'None' ? 'Remove access' : level)),
        );
        select.value = e.Access;
        select.disabled = state.busy || !!p.result.Policy?.OperationKey;
        select.setAttribute('aria-label', (state.teams.get(e.TeamId) || 'Team') + ' access');
        select.onchange = () => {
          e.Access = select.value;
          render();
        };
        cell.append(select);
        const current =
          p.result.Policy?.Applied?.find((a) => a.TeamId === e.TeamId)?.Access || 'None';
        tr.append(
          node('td', state.teams.get(e.TeamId) || e.TeamId),
          cell,
          node(
            'td',
            (current === 'None' ? 'No managed access' : current) +
              (p.result.Policy?.OperationKey ? ' · applying' : ''),
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
    // A run stopped because the library inherits again is replaced by Apply access.
    const running = !!p?.result.Policy?.OperationKey && !inherits(p);
    $('ad-apply').disabled =
      state.busy || !p || running || (!changed(p) && p.result.Status !== 'Missing' && !inherits(p));
    $('ad-add-team').disabled = state.busy || !p || running;
    $('ad-change-status').textContent = !p
      ? 'Loading access…'
      : running
        ? 'Applying access and syncing members…'
        : inherits(p)
          ? inheritsAgain
          : changed(p)
            ? 'Changes not yet applied.'
            : p.result.Status === 'Missing'
              ? 'Apply to confirm this library’s access.'
              : p.result.Status === 'Applied'
                ? 'Access and team membership confirmed.'
                : p.result.Status;
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
  async function loadSites(append = false) {
    const term = $('ad-search').value.trim().replace(/'/g, "''");
    const result = await page(
      'asx_site',
      append
        ? nextOptions(state.nextSites)
        : "?$select=asx_siteid,asx_name,asx_approved,asx_url,_asx_nativeid_value&$orderby=asx_name&$filter=contains(asx_name,'" +
            term +
            "')",
    );
    state.sites = append ? state.sites.concat(result.entities) : result.entities;
    state.nextSites = result.nextLink || null;
    render();
  }
  async function libraries(append = false) {
    const result = await page(
      'asx_library',
      append
        ? nextOptions(state.nextLibraries)
        : '?$select=asx_libraryid,asx_name,asx_approved,asx_policyapplied&$orderby=asx_name&$filter=_asx_siteid_value eq ' +
            state.site.asx_siteid,
    );
    state.libraries = append ? state.libraries.concat(result.entities) : result.entities;
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
    for (const e of entries)
      if (!state.teams.has(e.TeamId)) {
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
  async function discoverActivity(append = false) {
    const rows = await page(
      'asx_operation',
      append
        ? nextOptions(state.nextActivity)
        : "?$select=asx_workkey,asx_workkind,asx_displayname,asx_siteurl,asx_status&$orderby=createdon desc&$filter=(asx_workkind eq 'SiteValidation' or asx_workkind eq 'LibrarySetup' or asx_workkind eq 'LibraryValidation' or asx_workkind eq 'LibraryDiscovery') and asx_status ne 'Discovered' and asx_status ne 'Applied' and asx_status ne 'Approved' and asx_status ne 'Cancelled' and asx_status ne 'Superseded'",
    );
    for (const r of rows.entities)
      state.operations.set(r.asx_workkey, {
        name: r.asx_displayname,
        url: r.asx_siteurl,
        kind: r.asx_workkind,
        status: r.asx_status,
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
      if (
        (state.selectedOperation && state.selectedOperation !== key) ||
        (state.site && o.url && state.site.asx_url !== o.url)
      )
        continue;
      observations.push([key, o, await catalog({ Command: 'Inspect', Key: key })]);
    }
    for (const [key, o, result] of observations) {
      o.result = result;
      o.status = result.Status;
    }
    const signature = JSON.stringify([
      state.site?.asx_siteid,
      state.selectedOperation,
      observations.map(([key, o]) => [key, progressState(o)]),
    ]);
    if (signature === state.progressSignature && !pendingPolicy()) return;
    for (const [key, o, result] of observations) {
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
          showDiscovery();
        }
        continue;
      }
      if (['Ready', 'Approved', 'Applied'].includes(result.Status)) {
        await loadSites();
        if (o.kind === 'SiteValidation' && result.CatalogId) {
          const added = state.sites.find((s) => s.asx_siteid === result.CatalogId);
          if (added) await selectSite(added);
        } else if (state.site) await libraries();
        await window.AsxdAdmin?.refreshCatalog();
        state.operations.delete(key);
        issue(o.name + ' is ready.');
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
          '?$select=asx_libraryid,asx_name,_asx_siteid_value,asx_approved,asx_policyapplied',
        );
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
    for (const l of d.Observation.Libraries) {
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
    if (!d.Observation.Libraries.length)
      $('ad-existing-choices').append(node('p', 'No document libraries found on this page.'));
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
      if (latest.Policy?.OperationKey && !latest.Policy.Inherits)
        throw new Error(
          'An access synchronization is in progress. Wait for it to finish, then apply your changes.',
        );
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
      p.saved = JSON.stringify(p.entries);
      issue('Access submitted. Team membership syncing is onboarded automatically.');
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
