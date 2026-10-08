'use strict';
// The admin page shell: which page a load shows (the app's left menu opens each page at its own
// address), deep links between pages, the unsaved-changes prompt, and the helpers every page
// uses through window.AsxdUi. Text is only ever set with textContent.
(() => {
  const BUILD = 'ui20261008rc1';
  const TABS = ['templates', 'access', 'monitor', 'settings'];
  const LEGACY = {
    runtime: 'settings',
    operations: 'monitor',
    administration: 'monitor',
    author: 'templates',
  };
  const KEYS = { launched: 'asxd.launched', link: 'asxd.deeplink' };
  const PARAMS = ['library', 'operation', 'record', 'table', 'run', 'template', 'step'];
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
    folderDepth: 10,
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
  // '?data=monitor-ui20261008rc1' → 'monitor'.
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
  // control stays focusable (aria-disabled). Clearing the reason removes both.
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
  // focuses its consequence text, and resolves with the chosen value; Escape keeps.
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

  // A "⋯" menu: Enter, Space or Down opens and focuses the first item (in a menu of choices, the
  // checked one); Up and Down move; Escape closes and returns focus to the trigger. Choosing an
  // item closes it, and so do Tab out of it and a click anywhere else. A busy trigger
  // (AsxdUi.busy) does not open.
  function menu(trigger, list) {
    const items = () =>
      [...list.querySelectorAll('[role=menuitem], [role=menuitemradio]')].filter((i) => !i.hidden);
    const open = () => {
      list.hidden = false;
      trigger.setAttribute('aria-expanded', 'true');
      const all = items();
      (all.find((i) => i.getAttribute('aria-checked') === 'true') || all[0])?.focus();
    };
    const close = (focus = true) => {
      list.hidden = true;
      trigger.setAttribute('aria-expanded', 'false');
      if (focus) trigger.focus();
    };
    trigger.onclick = () => {
      if (trigger.dataset.busy) return;
      if (list.hidden) open();
      else close();
    };
    trigger.addEventListener('keydown', (event) => {
      if (event.key !== 'ArrowDown' || trigger.dataset.busy) return;
      event.preventDefault();
      open();
    });
    list.addEventListener('keydown', (event) => {
      const all = items();
      const index = all.indexOf(document.activeElement);
      if (event.key === 'Escape') {
        event.preventDefault();
        close();
      } else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
        event.preventDefault();
        all[(index + (event.key === 'ArrowDown' ? 1 : all.length - 1)) % all.length].focus();
      }
    });
    // Focus still on an item (a blocked one, or one whose action moved it nowhere) goes back to
    // the button, not to the page.
    list.addEventListener('click', () => close(list.contains(document.activeElement)));
    // Tab out of the button or the list closes it; focus stays where it went. A focus change
    // with no element to go to (another window, a click on a button that does not take focus)
    // leaves it to the click handler below.
    const leaving = (event) => {
      const next = event.relatedTarget;
      if (!list.hidden && next && !trigger.contains(next) && !list.contains(next)) close(false);
    };
    trigger.addEventListener('focusout', leaving);
    list.addEventListener('focusout', leaving);
    // A menu redrawn away (a table row's ⋯) stops listening at the next click.
    const outside = (event) => {
      if (!trigger.isConnected) return document.removeEventListener('click', outside);
      if (!list.hidden && !trigger.contains(event.target) && !list.contains(event.target))
        close(false);
      return undefined;
    };
    document.addEventListener('click', outside);
  }

  // Appearance: Match browser (the default), Light or Dark. A choice sets the root's color-scheme,
  // which every light-dark() token follows; Match browser removes it, so the stylesheet's
  // "light dark" follows the browser again. The choice is this browser's own, kept in
  // localStorage under one key that the four pages share. Storage that is missing or throws reads
  // as Match browser, and a choice still applies to the page that is open.
  const APPEARANCE_KEY = 'asxd.appearance';
  const APPEARANCES = [
    ['browser', 'Match browser'],
    ['light', 'Light'],
    ['dark', 'Dark'],
  ];
  const appearanceItems = [];
  let appearanceChoice = 'browser';
  function local(action) {
    try {
      return action(window.localStorage);
    } catch {
      return undefined;
    }
  }
  const knownAppearance = (value) => (value === 'light' || value === 'dark' ? value : 'browser');
  function applyAppearance(choice) {
    appearanceChoice = knownAppearance(choice);
    document.documentElement.style.colorScheme =
      appearanceChoice === 'browser' ? '' : appearanceChoice;
    for (const item of appearanceItems)
      item.setAttribute('aria-checked', String(item.dataset.appearance === appearanceChoice));
  }
  function setAppearance(choice) {
    applyAppearance(choice);
    local((s) =>
      appearanceChoice === 'browser'
        ? s.removeItem(APPEARANCE_KEY)
        : s.setItem(APPEARANCE_KEY, appearanceChoice),
    );
  }
  // The ◐ button at the right end of a page header and its menu of the three choices.
  function appearanceControl(header, tab) {
    const anchor = el('div', null, 'menu-anchor appearance');
    const trigger = button('◐', null, 'icon');
    const list = el('ul', null, 'menu');
    list.id = 'appearance-' + tab;
    list.setAttribute('role', 'menu');
    list.setAttribute('aria-label', 'Appearance');
    list.hidden = true;
    trigger.setAttribute('aria-label', 'Appearance');
    trigger.setAttribute('aria-haspopup', 'menu');
    trigger.setAttribute('aria-expanded', 'false');
    trigger.setAttribute('aria-controls', list.id);
    for (const [value, label] of APPEARANCES) {
      const item = button(label, () => setAppearance(value), null);
      item.setAttribute('role', 'menuitemradio');
      item.tabIndex = -1;
      item.dataset.appearance = value;
      item.setAttribute('aria-checked', String(value === appearanceChoice));
      appearanceItems.push(item);
      const row = el('li');
      row.setAttribute('role', 'none');
      row.append(item);
      list.append(row);
    }
    anchor.append(trigger, list);
    header.append(anchor);
    menu(trigger, list);
  }
  // Applied as soon as shell.js runs, from the page's head, so the page never paints in the
  // other scheme first.
  applyAppearance(local((s) => s.getItem(APPEARANCE_KEY)));

  // What Documents' SharePoint check found for a library setup whose create answer was lost,
  // as one sentence for Monitor and Sites & access.
  function recoverySentence(name, recovery) {
    const found = recovery.Candidates?.[0];
    switch (recovery.State) {
      case 'Checking':
        return 'Checking SharePoint for ' + name + '…';
      case 'Found': {
        const when = time(found.CreatedUtc).textContent;
        return (
          'SharePoint has a library ' +
          found.Title +
          ' at ' +
          found.Url +
          ', created ' +
          when +
          '. It matches this request.'
        );
      }
      case 'NotFound':
        return 'SharePoint has no library named ' + name + ". The creation didn't happen.";
      default:
        return recovery.Reason;
    }
  }
  // The checks an ambiguous candidate failed: "different name, different address", or "".
  const candidateChecks = (c) =>
    [
      c.TitleMatches ? null : 'different name',
      c.UrlMatches ? null : 'different address',
      c.IsLibrary ? null : 'not a document library',
      c.CreatedAfterRequest ? null : 'there before the request',
      c.CatalogEntry === 'Conflict' ? 'another catalog entry' : null,
    ]
      .filter(Boolean)
      .join(', ');

  // Re-renders without losing focus: the element with the same data-focus-key gets
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
  // A server time ("/Date(ms)/" or ISO) as epoch milliseconds; NaN when unreadable.
  function ms(value) {
    const legacy = /^\/Date\((-?\d+)/.exec(value || '');
    return legacy ? Number(legacy[1]) : new Date(value || NaN).getTime();
  }
  function time(value) {
    const date = new Date(ms(value));
    const node = el('time', isNaN(date) ? '' : formatter.format(date));
    if (!isNaN(date)) node.setAttribute('datetime', date.toISOString().replace(/\.\d{3}Z$/, 'Z'));
    return node;
  }
  // Where a key or ID is shown, it sits in this disclosure with its Copy button. Copy is a
  // symbol-like button, so its aria-label names what it copies.
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
  // actions are disabled until the work ends. Errors go to the area's feedback.
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
    automationCache = null;
    runtimeListeners.forEach((fn) => fn(result));
  }
  async function setAutomation(enabled) {
    const current = state.runtime;
    if (!current) throw new Error('Automation settings are not available to you.');
    // Pause and resume only: a stale row version is refused, never a profile save.
    const result = await api('asx_RuntimeAdmin', {
      Command: 'SetEnabled',
      Enabled: enabled,
      RowVersion: current.RowVersion,
    });
    setRuntime(result);
    return result;
  }

  function show(tab, link) {
    state.tab = tab;
    state.link = link;
    for (const name of TABS) $(name).hidden = name !== tab;
    nameTab(tab);
  }
  // Dynamics names the browser tab after the web resource's address ("asx_admin/index.html?…").
  // The page names it after itself instead, and puts its name back if Dynamics sets the address
  // again while this page is open.
  const PAGE_NAMES = {
    templates: 'Folder templates',
    access: 'Sites & access',
    monitor: 'Monitor',
    settings: 'Settings',
  };
  let tabTitle = null;
  function nameTab(tab) {
    const text = PAGE_NAMES[tab] + ' · Ascentix Documents';
    document.title = text;
    let doc;
    try {
      doc = window.top && window.top !== window ? window.top.document : null;
    } catch {
      doc = null; // Another origin: the tab keeps Dynamics' name.
    }
    if (!doc) return;
    doc.title = text;
    if (tabTitle || !doc.querySelector || typeof MutationObserver !== 'function') return;
    const element = doc.querySelector('title');
    if (!element) return;
    tabTitle = new MutationObserver(() => {
      const name = PAGE_NAMES[state.tab] + ' · Ascentix Documents';
      if (doc.title.startsWith('asx_admin/')) doc.title = name;
    });
    tabTitle.observe(element, { childList: true, characterData: true, subtree: true });
    window.addEventListener?.('pagehide', () => tabTitle.disconnect());
  }
  async function go(tab, link) {
    if (link) write(KEYS.link, JSON.stringify({ ...link, tab }));
    // The same address the left menu uses, so the app highlights this page's menu item.
    await xrm.Navigation.navigateTo({
      pageType: 'webresource',
      webresourceName: 'asx_admin/index.html?data=' + tab + '-' + BUILD,
    });
  }
  // The h1 a page shows: one inside a hidden part of the page (the editor while the overview
  // shows) does not count.
  function pageHeading(tab) {
    const page = $(tab);
    return (
      [...(page?.querySelectorAll('h1') || [])].find((h) => {
        const hidden = h.closest('[hidden]');
        return !hidden || hidden === page;
      }) || null
    );
  }
  // Leaving reloads the page, so unsaved template edits ask first, at the top of the page
  // content: Save draft, Discard changes or Stay. Resolves true when the page may go.
  async function confirmLeave() {
    const guard = state.dirty?.();
    if (!guard) return true;
    const active = document.activeElement;
    const anchor =
      active && active !== document.body ? active : pageHeading(state.tab) || $(state.tab);
    const host = $('leavePrompt');
    const asking = ask(anchor, {
      host,
      text: 'You have unsaved changes to ' + guard.template + '.',
      keep: 'stay',
      choices: [
        { value: 'save', label: 'Save draft', style: 'primary' },
        { value: 'discard', label: 'Discard changes' },
        { value: 'stay', label: 'Stay' },
      ],
    });
    // The prompt sits above the content; bring it into view with its focused message.
    host.querySelector('.confirm-text')?.scrollIntoView?.({ block: 'nearest' });
    const choice = await asking;
    if (choice === 'stay') return false;
    if (choice === 'discard') {
      await guard.discard();
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
  const leave = () => confirmLeave();
  async function navigate(tab, link = null) {
    if (tab === state.tab && link) return links.get(tab)?.({ ...link, tab });
    if (tab === state.tab) return undefined;
    if (!(await leave())) return undefined;
    return go(tab, link);
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
      // A caller who cannot read the runtime skips the worker condition.
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
    for (const tab of TABS) {
      const header = $(tab).querySelector('.page-header');
      if (header) appearanceControl(header, tab);
    }
    if (!xrm?.WebApi || !xrm?.Utility) return offline();
    const first = read(KEYS.launched) === null;
    write(KEYS.launched, '1');
    const link = storedLink();
    const hashed = link ? null : fromHash(window.location.hash);
    const data = fromData(window.location.search);
    const tab = link?.tab || hashed?.tab || data || 'templates';
    state.runtimePromise = loadRuntime();
    const privileges = (state.privileges = loadPrivileges());
    // The app launch opens its first entry, Folder templates; only that load lands elsewhere.
    if (first && !link && !hashed && tab === 'templates') {
      const target = await landing();
      if (target !== 'templates') return go(target, null);
    }
    show(tab, link || hashed);
    // The page starts once the privilege checks are in; onRuntime listeners hear the result
    // (null when Get is refused) whenever the runtime Get returns.
    state.runtimePromise.then(setRuntime);
    await privileges;
    return inits.get(tab)?.();
  }

  // A count with its noun: "1 problem", "1,214 problems".
  const plural = (n, one, many) =>
    Number(n || 0).toLocaleString('en-US') + ' ' + (n === 1 ? one : many);

  // An 8px status dot. It always sits next to text that says the same thing.
  function dot(tone) {
    const node = el('span', null, 'dot');
    node.dataset.tone = tone;
    node.setAttribute('aria-hidden', 'true');
    return node;
  }
  const status = (tone, text) => {
    const node = el('span', null, 'status');
    node.append(dot(tone), text);
    return node;
  };
  // A stored name such as "P-{root.projectnumber}" with each field token as a chip. A token
  // labelOf does not know stays as its literal text. Nothing is parsed as markup.
  function tokens(text, labelOf) {
    const node = el('span', null, 'name');
    const value = String(text ?? '');
    let at = 0;
    for (const match of value.matchAll(/\{([a-z0-9_]+)\.([a-z0-9_]+)\}/gi)) {
      if (match.index > at) node.append(value.slice(at, match.index));
      const label = labelOf(match[1], match[2]);
      node.append(label ? el('span', label, 'token') : match[0]);
      at = match.index + match[0].length;
    }
    if (at < value.length) node.append(value.slice(at));
    return node;
  }
  // A folder's rule as one short sentence: "When Status is Active", "When Status is Active + 2
  // more", "When any of 3…". A group holding only one group reads as that group. clause turns
  // one condition into "Field is Value". An empty group has no sentence ('').
  function ruleSentence(group, clause) {
    const conditions = group.Conditions || [],
      groups = group.Groups || [];
    const n = conditions.length + groups.length;
    if (!n) return '';
    if (conditions.length === 1 && !groups.length) return 'When ' + clause(conditions[0]);
    if (groups.length === 1 && !conditions.length) return ruleSentence(groups[0], clause);
    if (group.All === false) return 'When any of ' + n + '…';
    const first = conditions.length
      ? clause(conditions[0])
      : ruleSentence(groups[0], clause).replace(/^When /, '');
    return 'When ' + first + ' + ' + (n - 1) + ' more';
  }
  // A section card: a head row with the title, a muted summary and an optional link pushed
  // right, then the body.
  function card({ id = null, title, summary = null, action = null, level = 2 }) {
    const box = el('section', null, 'section-card');
    const head = el('div', null, 'card-head');
    const heading = el('h' + level, title);
    if (id) {
      box.id = id;
      heading.id = id + '-title';
      box.setAttribute('aria-labelledby', heading.id);
    }
    head.append(heading);
    if (summary) head.append(el('span', summary, 'muted'));
    if (action) {
      const link = button(action.label, action.onClick, 'link card-action');
      if (action.key) link.dataset.focusKey = action.key;
      head.append(link);
    }
    const body = el('div', null, 'card-body');
    box.append(head, body);
    return { card: box, head, body };
  }

  // A right-side panel or drawer as a dialog named by its heading. Focus moves to the heading;
  // Escape closes it unless a confirmation, menu or popover inside it took that Escape; focus
  // goes back to the invoker, or to the control that replaced it (same data-focus-key). One
  // panel is open at a time.
  let openPanel = null;
  // Below 1000px a panel covers the page, so while it is open the rest of the page is inert and
  // Tab and clicks stay in the panel: what sits beside it, at each level up to <main>. Wider, the
  // page beside it stays usable. The unsaved-changes prompt stays reachable. Returns what puts
  // the page back.
  function inertOutside(panel) {
    const covers = window.matchMedia?.('(max-width: 999px)');
    let made = [];
    const release = () => {
      made.forEach((n) => (n.inert = false));
      made = [];
    };
    const apply = () => {
      release();
      if (covers && !covers.matches) return;
      for (let node = panel; node.parentNode && node.parentNode !== document.body;) {
        for (const sibling of node.parentNode.children)
          if (sibling !== node && sibling.id !== 'leavePrompt' && !sibling.inert) {
            sibling.inert = true;
            made.push(sibling);
          }
        node = node.parentNode;
      }
    };
    apply();
    covers?.addEventListener?.('change', apply);
    return () => {
      covers?.removeEventListener?.('change', apply);
      release();
    };
  }
  function sidePanel(panel, invoker, { onClose = null } = {}) {
    openPanel?.close(false);
    const heading = panel.querySelector('h1, h2, h3');
    if (!heading.id) heading.id = (panel.id || 'panel') + '-title';
    heading.tabIndex = -1;
    panel.setAttribute('role', 'dialog');
    panel.setAttribute('aria-labelledby', heading.id);
    panel.hidden = false;
    const key = invoker?.dataset?.focusKey;
    const onKey = (event) => {
      if (event.key !== 'Escape' || event.defaultPrevented) return;
      if (event.target?.closest?.('.confirm, .menu, .popover')) return;
      event.preventDefault();
      handle.close();
    };
    panel.addEventListener('keydown', onKey);
    const restore = inertOutside(panel);
    const handle = {
      close(focus = true) {
        if (panel.hidden) return;
        // A confirmation open inside it is answered with its keep value, so the action waiting
        // on it ends with the panel and does not come back in its next use.
        for (const open of [...panel.querySelectorAll('.confirm[role=group]')])
          confirmations.get(open)?.();
        panel.hidden = true;
        restore();
        panel.removeEventListener('keydown', onKey);
        if (openPanel === handle) openPanel = null;
        onClose?.();
        if (!focus) return;
        const back = invoker?.isConnected
          ? invoker
          : key && document.querySelector('[data-focus-key="' + key + '"]');
        back?.focus();
      },
    };
    openPanel = handle;
    heading.focus();
    return handle;
  }

  // Whether automation runs. A System Administrator's
  // runtime Get answers it; other roles read the Default runtime row, which every Documents
  // role may read. A row that was read is kept until the runtime changes; a failed or missing
  // read is tried again on the next call.
  let automationCache = null;
  async function automation() {
    const runtime = await Promise.resolve(state.runtimePromise).then(() => state.runtime);
    if (runtime) return { Enabled: !!runtime.Enabled };
    if (!automationCache)
      automationCache = xrm.WebApi.retrieveMultipleRecords(
        'asx_runtime',
        "?$select=asx_enabled&$filter=asx_name eq 'Default'&$top=2",
      ).then((rows) =>
        rows.entities.length === 1 ? { Enabled: !!rows.entities[0].asx_enabled } : null,
      );
    const reading = automationCache;
    try {
      const result = await reading;
      if (!result && automationCache === reading) automationCache = null;
      return result;
    } catch {
      if (automationCache === reading) automationCache = null;
      return null;
    }
  }

  // The problem count other pages show as "Monitor · N problems": the five problem lists of one
  // Summary per page load (a page change reloads), shared by every header that asks. Re-runs are
  // not problems. Null when the caller lacks the Operator role or the read fails.
  const PROBLEMS = [
    'NotCaptured',
    'BlockedRecords',
    'WaitingRecords',
    'BlockedJobs',
    'RetryingJobs',
  ];
  let problemsPromise = null;
  function problems() {
    problemsPromise ??= (async () => {
      await state.privileges;
      if (!can('prvCreateasx_operatorcommand')) return null;
      try {
        const summary = (await api('asx_ManageWork', { Command: 'Summary' })).Summary;
        if (!summary) return null;
        return {
          total: PROBLEMS.reduce((sum, list) => sum + (summary[list] || 0), 0),
          capped: PROBLEMS.some((list) => summary.Capped?.includes(list)),
        };
      } catch {
        return null;
      }
    })();
    return problemsPromise;
  }
  function problemPill(host) {
    const link = button('', () => navigate('monitor'), 'pill problem-pill');
    link.dataset.tone = 'warning';
    link.hidden = true;
    host.replaceChildren(link);
    problems().then((count) => {
      if (!count || count.total === 0) return;
      link.textContent =
        'Monitor · ' +
        (count.capped ? '5,000+ problems' : plural(count.total, 'problem', 'problems'));
      link.hidden = false;
    });
    return link;
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
    menu,
    recoverySentence,
    candidateChecks,
    time,
    ms,
    plural,
    dot,
    status,
    tokens,
    ruleSentence,
    card,
    sidePanel,
    automation,
    problemPill,
    confirmLeave,
    details,
    busy,
    disable,
    blocked,
    api,
    can,
    needs,
    runtime: () => state.runtime,
    // Resolves once the load's runtime Get is back (null when refused), after the onRuntime
    // listeners have heard it: for a page that must tell "not back yet" from "refused".
    runtimeReady: () => Promise.resolve(state.runtimePromise).then(() => state.runtime),
    onRuntime: (fn) => runtimeListeners.push(fn),
    setRuntime,
    setAutomation,
    appearance: () => appearanceChoice,
    setAppearance,
  };
  document.addEventListener('DOMContentLoaded', start);
})();
