'use strict';
// The admin page shell: which tab a load shows (spec 2.3), tab clicks that reload the page so the
// app's left menu follows, deep links between tabs, the automation chip (2.5), and the helpers
// every tab uses through window.AsxdUi. Text is only ever set with textContent.
(() => {
  const BUILD = 'ui20261006nav1';
  const TABS = ['templates', 'access', 'monitor', 'settings'];
  const LEGACY = {
    runtime: 'settings',
    operations: 'monitor',
    administration: 'monitor',
    author: 'templates',
  };
  const KEYS = { launched: 'asxd.launched', focus: 'asxd.focusTab', link: 'asxd.deeplink' };
  const PARAMS = ['library', 'operation', 'record', 'table', 'run', 'template'];
  const EMPTY = '00000000-0000-0000-0000-000000000000';
  // The server's bounds (Ascentix.Documents.Domain.Bounds, with their reasons); verify-admin.cjs
  // fails when these differ from it.
  const BOUNDS = {
    relatedRecords: 5,
    destinations: 10,
    foldersPerDestination: 100,
    configurationRows: 1216,
    teamEntries: 10,
    previewRecords: 5,
    conditionDepth: 10,
  };
  const ROLES = {
    prvCreateasx_publication: 'Documents Publisher',
    prvCreateasx_operatorcommand: 'Documents Operator',
    prvWriteasx_runtime: 'System Administrator',
    prvCreateasx_policy: 'Documents Security Administrator',
    prvCreateasx_site: 'Documents Security Administrator',
  };
  const xrm = window.parent?.Xrm || window.Xrm;
  const $ = (id) => document.getElementById(id);
  const inits = new Map();
  const links = new Map();
  const runtimeListeners = [];
  const missing = new Set();
  // Open confirmations made by ask(), each with the function that answers it with its keep value.
  const confirmations = new Map();
  const state = { tab: null, link: null, runtime: null, dirty: null, confirms: 0 };

  // sessionStorage can be missing or throw (blocked site data, sandboxed frames). Every access
  // is guarded; a failure reads as "nothing stored", so it never redirects.
  function storage(action) {
    try {
      return action(window.sessionStorage);
    } catch {
      return undefined;
    }
  }
  const read = (key) => storage((s) => s.getItem(key));
  const write = (key, value) => storage((s) => s.setItem(key, value));
  const forget = (key) => storage((s) => s.removeItem(key));

  const known = (name) => {
    const tab = LEGACY[name] || name;
    return TABS.includes(tab) ? tab : null;
  };
  // '#monitor?operation=folderjob:abc' → { tab: 'monitor', operation: 'folderjob:abc' }.
  function fromHash(hash) {
    const [name, query = ''] = String(hash || '')
      .replace(/^#/, '')
      .split('?');
    const tab = known(name);
    if (!tab) return null;
    const params = new URLSearchParams(query);
    const link = { tab };
    for (const key of PARAMS) if (params.get(key)) link[key] = params.get(key);
    return link;
  }
  // '?data=monitor-ui20261006nav1' → 'monitor'.
  const fromData = (search) =>
    known((new URLSearchParams(String(search || '')).get('data') || '').split('-')[0]);
  function storedLink() {
    const text = read(KEYS.link);
    if (text == null) return null;
    forget(KEYS.link);
    try {
      const link = JSON.parse(text);
      const tab = known(link?.tab);
      return tab ? { ...link, tab } : null;
    } catch {
      return null;
    }
  }

  const el = (tag, text, css) => {
    const node = document.createElement(tag);
    if (text != null) node.textContent = text;
    if (css) node.className = css;
    return node;
  };
  function button(text, onClick, style = 'secondary') {
    const b = el('button', text, style);
    b.type = 'button';
    b.onclick = onClick;
    return b;
  }
  const refs = (node, attr) => (node.getAttribute(attr) || '').split(/\s+/).filter(Boolean);
  function addRef(node, attr, id) {
    node.setAttribute(attr, [...new Set([...refs(node, attr), id])].join(' '));
  }
  function dropRef(node, attr, id) {
    const rest = refs(node, attr).filter((r) => r !== id);
    if (rest.length) node.setAttribute(attr, rest.join(' '));
    else node.removeAttribute(attr);
  }
  const blocked = (control) => control.disabled || control.getAttribute('aria-disabled') === 'true';
  // A disabled reason is shown next to its control and linked with aria-describedby; the
  // control stays focusable (aria-disabled). Clearing the reason removes both (spec 4.3 rule 5).
  function disable(control, reasonId, reason) {
    let note = $(reasonId);
    if (reason) {
      if (!note) {
        note = el('p', null, 'reason');
        note.id = reasonId;
        control.after(note);
      }
      note.textContent = reason;
      control.setAttribute('aria-disabled', 'true');
      addRef(control, 'aria-describedby', reasonId);
      return;
    }
    control.removeAttribute('aria-disabled');
    dropRef(control, 'aria-describedby', reasonId);
    note?.remove();
  }
  const help = (id, text) => {
    const node = el('p', text, 'help');
    node.id = id;
    return node;
  };

  function feedback(area, text, kind = 'success', action = null) {
    const line = $('fb-' + area);
    if (!line) return;
    line.className = 'feedback ' + (kind === 'error' ? 'is-error' : 'is-success');
    line.setAttribute('role', kind === 'error' ? 'alert' : 'status');
    line.textContent = text;
    if (action) line.append(' ', button(action.label, action.onClick, 'link'));
  }
  function clearFeedback(area) {
    const line = $('fb-' + area);
    if (!line) return;
    line.textContent = '';
    line.removeAttribute('role');
    line.className = 'feedback';
  }

  // One confirmation per button row. It renders right after the invoker's row (or in host),
  // focuses its consequence text, and resolves with the chosen value; Escape keeps (spec 5.2).
  function ask(invoker, { text, choices, keep, host = null, details = null }) {
    return new Promise((resolve) => {
      const row = host ? null : invoker.closest('[data-actions]') || invoker.parentNode;
      const place = host || row.parentNode;
      // A confirmation already open here is answered with its keep value, so the action waiting
      // on it ends. Only confirmations ask() made are touched; static .confirm markup stays.
      for (const open of [...place.querySelectorAll('.confirm[role=group]')]) {
        const answer = confirmations.get(open);
        if (answer) answer();
        else open.remove();
      }
      const box = el('div', null, 'confirm');
      const message = el('p', text, 'confirm-text');
      message.id = 'confirm-' + ++state.confirms;
      message.tabIndex = -1;
      box.setAttribute('role', 'group');
      box.setAttribute('aria-labelledby', message.id);
      box.append(message);
      if (details) box.append(details);
      const actions = el('div', null, 'row');
      actions.setAttribute('data-actions', '');
      let done = false;
      const finish = (value) => {
        if (done) return;
        done = true;
        confirmations.delete(box);
        box.remove();
        const back = invoker.isConnected
          ? invoker
          : invoker.dataset.focusKey &&
            document.querySelector('[data-focus-key="' + invoker.dataset.focusKey + '"]');
        back?.focus();
        resolve(value);
      };
      for (const choice of choices)
        actions.append(
          button(choice.label, () => finish(choice.value), choice.style || 'secondary'),
        );
      box.append(actions);
      confirmations.set(box, () => finish(keep));
      box.addEventListener('keydown', (event) => {
        if (event.key !== 'Escape') return;
        event.preventDefault();
        finish(keep);
      });
      if (host) host.replaceChildren(box);
      else row.after(box);
      message.focus();
    });
  }
  const confirmInline = (
    invoker,
    { text, confirm, keep, danger = false, details = null, host = null },
  ) =>
    ask(invoker, {
      text,
      details,
      host,
      keep: false,
      choices: [
        { value: true, label: confirm, style: danger ? 'danger' : 'primary' },
        { value: false, label: keep, style: 'secondary' },
      ],
    });

  // Re-renders without losing focus (spec 5.2): the element with the same data-focus-key gets
  // focus back; if it is gone, the row that took its place, else its list or pane heading.
  function withFocus(fn) {
    const active = document.activeElement;
    const key = active?.dataset?.focusKey;
    const scope = active?.closest?.('[data-focus-scope]');
    const scopeId = scope?.id;
    const rows = scope ? [...scope.querySelectorAll('[data-focus-row]')] : [];
    const index = rows.findIndex((r) => r.contains(active));
    fn();
    if (!key) return;
    const again = document.querySelector('[data-focus-key="' + key + '"]');
    if (again) return again.focus();
    const area = scopeId && $(scopeId);
    if (!area) return;
    // Focus that was not in a row (a list action such as Load more) goes to the heading.
    const row = index < 0 ? null : area.querySelectorAll('[data-focus-row]')[index];
    const action = row && [...row.querySelectorAll('button')].find((b) => !b.disabled);
    (action || area.querySelector('[data-focus-heading]'))?.focus();
  }

  const formatter = new Intl.DateTimeFormat(undefined, {
    day: 'numeric',
    month: 'short',
    hour: '2-digit',
    minute: '2-digit',
  });
  function time(value) {
    const ms = /^\/Date\((-?\d+)/.exec(value || '');
    const date = new Date(ms ? Number(ms[1]) : value || NaN);
    const node = el('time', isNaN(date) ? '' : formatter.format(date));
    if (!isNaN(date)) node.setAttribute('datetime', date.toISOString().replace(/\.\d{3}Z$/, 'Z'));
    return node;
  }
  // The only place keys and IDs appear outside Monitor › Advanced (spec 3, "Details"). Copy is a
  // symbol-like button, so its aria-label names what it copies (spec 5.1).
  function details(value, label = 'Details', object = null) {
    const box = el('details', null, 'details');
    const name = label.toLowerCase() + (object ? ' for ' + object : '');
    const copy = button(
      'Copy',
      async () => {
        try {
          await navigator.clipboard.writeText(value);
          copy.textContent = 'Copied';
          copy.setAttribute('aria-label', 'Copied ' + name);
        } catch {
          copy.textContent = 'Copy failed';
          copy.setAttribute('aria-label', 'Copy failed: ' + name);
        }
      },
      'link',
    );
    copy.setAttribute('aria-label', 'Copy ' + name);
    box.append(el('summary', label), el('code', value), copy);
    return box;
  }

  // The clicked button shows a spinner and an "-ing" label; its area is aria-busy and the sibling
  // actions are disabled until the work ends (spec 2.5, F-15). Errors go to the area's feedback.
  async function busy(control, label, area, fn) {
    if (control.dataset.busy) return undefined;
    const region = control.closest('[data-actions]') || control.parentNode;
    const siblings = [...region.querySelectorAll('button')].filter(
      (b) => b !== control && !b.disabled,
    );
    const original = control.textContent;
    control.dataset.busy = '1';
    control.textContent = label;
    control.classList.add('is-busy');
    region.setAttribute('aria-busy', 'true');
    siblings.forEach((b) => (b.disabled = true));
    if (area) clearFeedback(area);
    try {
      return await fn();
    } catch (error) {
      if (!area) throw error;
      feedback(area, error.message || String(error), 'error');
      return undefined;
    } finally {
      delete control.dataset.busy;
      if (control.textContent === label) control.textContent = original;
      control.classList.remove('is-busy');
      region.removeAttribute('aria-busy');
      siblings.forEach((b) => (b.disabled = false));
    }
  }

  async function api(name, request) {
    const response = await xrm.WebApi.online.execute({
      Request: JSON.stringify(request),
      getMetadata: () => ({
        boundParameter: null,
        operationType: 0,
        operationName: name,
        parameterTypes: { Request: { typeName: 'Edm.String', structuralProperty: 1 } },
      }),
    });
    let body = null;
    try {
      body = await response.json();
    } catch {
      body = null;
    }
    if (!response.ok) throw new Error(body?.error?.message || 'The server refused the request.');
    return JSON.parse(body.Result);
  }

  // Privileges are read once; an unknown answer counts as held, because the server decides (6.10).
  async function loadPrivileges() {
    const context = xrm.Utility.getGlobalContext();
    const user = String(context.userSettings?.userId || '').replace(/[{}]/g, '');
    if (!user) return;
    await Promise.all(
      Object.keys(ROLES).map(async (name) => {
        try {
          const response = await fetch(
            context.getClientUrl() +
              '/api/data/v9.2/systemusers(' +
              user +
              ")/Microsoft.Dynamics.CRM.RetrieveUserPrivilegeByPrivilegeName(PrivilegeName='" +
              name +
              "')",
            {
              credentials: 'same-origin',
              headers: { Accept: 'application/json', 'OData-Version': '4.0' },
            },
          );
          if (response.ok && !(await response.json()).RolePrivileges?.length) missing.add(name);
        } catch {
          // Unknown: the control stays available and the server decides.
        }
      }),
    );
  }
  const can = (name) => !missing.has(name);
  const needs = (name) => (can(name) ? null : 'Needs the ' + ROLES[name] + ' role.');

  async function loadRuntime() {
    try {
      state.runtime = await api('asx_RuntimeAdmin', { Command: 'Get' });
    } catch {
      state.runtime = null;
    }
    return state.runtime;
  }
  function setRuntime(result) {
    state.runtime = result;
    renderChip();
    runtimeListeners.forEach((fn) => fn(result));
  }
  async function setAutomation(enabled) {
    const current = state.runtime;
    if (!current) throw new Error('Automation settings are not available to you.');
    // Pause and resume only (spec 6.6): a stale row version is refused, never a profile save.
    const result = await api('asx_RuntimeAdmin', {
      Command: 'SetEnabled',
      Enabled: enabled,
      RowVersion: current.RowVersion,
    });
    setRuntime(result);
    return result;
  }

  // What a runtime result says still needs work in change tracking, or null when nothing does.
  function registrationProblem(result) {
    const registration = result?.Registration;
    if (registration?.Error)
      return 'Turned on, but change tracking needs attention: ' + registration.Error;
    const scopes = (registration?.Readiness || []).filter((r) => r.Status !== 'Ready');
    if (!scopes.length) return null;
    return (
      'Turned on, but change tracking is not ready for ' +
      scopes.map((r) => r.Scope).join(', ') +
      '. Repair it in Settings.'
    );
  }
  // The automation chip (spec 2.5). Its text is a status region and is set only when it changes.
  function renderChip() {
    const runtime = state.runtime;
    const chip = $('automationChip');
    if (!runtime) {
      chip.hidden = true;
      return;
    }
    const unset = !runtime.WorkerId || runtime.WorkerId === EMPTY;
    const attention =
      (runtime.Registration?.Readiness || []).some((r) => r.Status !== 'Ready') ||
      !!runtime.Registration?.Error;
    const view = unset
      ? { text: 'Automation not set up', tone: 'warning', tab: 'settings', label: 'Settings' }
      : !runtime.Enabled
        ? { text: 'Automation paused', tone: 'warning', turnOn: true }
        : attention
          ? {
              text: 'Automation running · needs attention',
              tone: 'warning',
              tab: 'settings',
              label: 'Settings',
            }
          : { text: 'Automation running', tone: 'ok', tab: 'monitor', label: 'Monitor' };
    const text = $('automationChipText');
    if (text.textContent !== view.text) text.textContent = view.text;
    chip.dataset.tone = view.tone;
    const link = $('automationChipLink');
    link.hidden = !view.tab;
    if (view.tab) {
      link.textContent = view.label;
      link.onclick = (event) => {
        event?.preventDefault?.();
        navigate(view.tab);
      };
    }
    const action = $('automationChipAction');
    action.hidden = !view.turnOn;
    disable(
      action,
      'automationChipReason',
      view.turnOn && runtime.CanChange === false
        ? 'Only a System Administrator can turn automation on.'
        : null,
    );
    chip.hidden = false;
  }

  function show(tab, link) {
    state.tab = tab;
    state.link = link;
    for (const name of TABS) {
      const active = name === tab;
      const control = $('tab-' + name);
      control.setAttribute('aria-selected', String(active));
      control.tabIndex = active ? 0 : -1;
      control.classList.toggle('active', active);
      $(name).hidden = !active;
    }
  }
  async function go(tab, link, focus) {
    if (link) write(KEYS.link, JSON.stringify({ ...link, tab }));
    if (focus) write(KEYS.focus, tab);
    await xrm.Navigation.navigateTo({
      pageType: 'webresource',
      webresourceName: 'asx_admin/index.html',
      data: tab + '-' + BUILD,
    });
  }
  // Leaving the tab reloads the page, so unsaved template edits ask first: Save draft,
  // Discard changes or Stay (spec 2.4). Resolves true when the page may go.
  async function leave() {
    const guard = state.dirty?.();
    if (!guard) return true;
    const choice = await ask($('tab-' + state.tab), {
      host: $('tabPrompt'),
      text: 'You have unsaved changes to ' + guard.template + '.',
      keep: 'stay',
      choices: [
        { value: 'save', label: 'Save draft', style: 'primary' },
        { value: 'discard', label: 'Discard changes' },
        { value: 'stay', label: 'Stay' },
      ],
    });
    if (choice === 'stay') return false;
    if (choice === 'discard') {
      guard.discard();
      return true;
    }
    try {
      await guard.save();
      return true;
    } catch (error) {
      feedback('templates', error.message || String(error), 'error');
      return false;
    }
  }
  async function navigate(tab, link = null) {
    if (tab === state.tab && link) return links.get(tab)?.({ ...link, tab });
    if (tab === state.tab) return undefined;
    if (!(await leave())) return undefined;
    return go(tab, link, false);
  }
  async function activate(tab) {
    if (tab === state.tab) return;
    if (await leave()) await go(tab, null, true);
  }
  function wireTabs() {
    const controls = TABS.map((name) => $('tab-' + name));
    controls.forEach((control, index) => {
      control.onclick = () => activate(TABS[index]);
    });
    // Manual activation (spec 2.4): arrows, Home and End move focus only; the native click of
    // Enter and Space activates.
    $('tabs').addEventListener('keydown', (event) => {
      const index = controls.indexOf(document.activeElement);
      const target = {
        ArrowRight: index + 1,
        ArrowLeft: index - 1,
        Home: 0,
        End: controls.length - 1,
      }[event.key];
      if (index < 0 || target === undefined) return;
      event.preventDefault();
      const next = controls[(target + controls.length) % controls.length];
      controls.forEach((c) => (c.tabIndex = c === next ? 0 : -1));
      next.focus();
    });
  }

  async function landing() {
    try {
      const [runtime, tables, libraries, published] = await Promise.all([
        state.runtimePromise,
        xrm.WebApi.retrieveMultipleRecords(
          'asx_runtimetable',
          '?$select=asx_runtimetableid&$top=1',
        ),
        xrm.WebApi.retrieveMultipleRecords(
          'asx_library',
          '?$select=asx_libraryid&$filter=asx_approved eq true&$top=1',
        ),
        xrm.WebApi.retrieveMultipleRecords(
          'asx_template',
          '?$select=asx_templateid&$filter=_asx_publishedrevisionid_value ne null&$top=1',
        ),
      ]);
      // A caller who cannot read the runtime skips the worker condition (spec 2.3).
      const worker = runtime ? !!runtime.WorkerId && runtime.WorkerId !== EMPTY : null;
      const hasLibrary = libraries.entities.length > 0;
      const ready = tables.entities.length > 0 && published.entities.length > 0;
      if (worker !== false && hasLibrary && ready) return 'monitor';
      if (worker === false) return 'settings';
      if (!hasLibrary) return 'access';
      return 'templates';
    } catch {
      return 'templates';
    }
  }

  function offline() {
    const tab =
      fromHash(window.location?.hash)?.tab || fromData(window.location?.search) || 'templates';
    show(tab, null);
    const note = el(
      'div',
      'Open this page from the Ascentix Documents app to connect.',
      'empty-panel',
    );
    note.setAttribute('role', 'alert');
    $(tab).replaceChildren(note);
  }

  async function start() {
    if (!xrm?.WebApi || !xrm?.Utility) return offline();
    const first = read(KEYS.launched) === null;
    write(KEYS.launched, '1');
    const link = storedLink();
    const hashed = link ? null : fromHash(window.location.hash);
    const data = fromData(window.location.search);
    const tab = link?.tab || hashed?.tab || data || 'templates';
    state.runtimePromise = loadRuntime();
    const privileges = loadPrivileges();
    // The app launch opens its first entry, Folder templates; only that load lands elsewhere.
    if (first && !link && !hashed && tab === 'templates') {
      const target = await landing();
      if (target !== 'templates') return go(target, null, false);
    }
    show(tab, link || hashed);
    wireTabs();
    // The tab starts once the privilege checks are in; the chip is drawn, and onRuntime
    // listeners hear the result (null when Get is refused), whenever the runtime Get returns.
    state.runtimePromise.then(setRuntime);
    await privileges;
    $('automationChipAction').onclick = () => {
      const action = $('automationChipAction');
      if (blocked(action)) return undefined;
      return busy(action, 'Turning on…', 'chip', async () => {
        // The chip's status text announces the new state; the feedback line only adds what
        // the resume reported as still needing work.
        const problem = registrationProblem(await setAutomation(true));
        if (problem) feedback('chip', problem, 'error');
      });
    };
    if (read(KEYS.focus) === tab) {
      forget(KEYS.focus);
      $('tab-' + tab).focus();
    }
    return inits.get(tab)?.();
  }

  window.AsxdUi = {
    BUILD,
    BOUNDS,
    TABS,
    activeTab: () => state.tab,
    onTab: (tab, init) => inits.set(tab, init),
    onLink: (tab, handler) => links.set(tab, handler),
    deeplink: () => state.link,
    navigate,
    setDirtyGuard: (fn) => (state.dirty = fn),
    el,
    button,
    feedback,
    clearFeedback,
    ask,
    confirmInline,
    withFocus,
    time,
    details,
    busy,
    disable,
    blocked,
    help,
    api,
    can,
    needs,
    runtime: () => state.runtime,
    // Resolves once the load's runtime Get is back (null when refused), after the chip and the
    // onRuntime listeners have heard it: for a tab that must tell "not back yet" from "refused".
    runtimeReady: () => Promise.resolve(state.runtimePromise).then(() => state.runtime),
    onRuntime: (fn) => runtimeListeners.push(fn),
    setRuntime,
    setAutomation,
  };
  document.addEventListener('DOMContentLoaded', start);
})();
