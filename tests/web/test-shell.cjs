'use strict';
// Shell contract with a fake DOM and mocked Dataverse: which page a load shows, landing, deep
// links, the unsaved-changes prompt, the shared helpers (side panel, tokens, status, automation,
// problem pill), feedback lines and the inline confirmation. Not a browser test.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createDocument, FakeEvent } = require('./fake-dom.cjs');
const base = path.resolve(__dirname, '../../client/admin');
const html = fs.readFileSync(path.join(base, 'index.html'), 'utf8');
const shell = fs.readFileSync(path.join(base, 'shell.js'), 'utf8');
const BUILD = /const BUILD = '([^']+)'/.exec(shell)[1];
const EMPTY = '00000000-0000-0000-0000-000000000000';

function storage(initial = {}, broken = false) {
  const data = new Map(Object.entries(initial));
  if (broken) {
    const fail = () => {
      throw new Error('SecurityError: storage is disabled');
    };
    return { getItem: fail, setItem: fail, removeItem: fail, data };
  }
  return {
    getItem: (k) => (data.has(k) ? data.get(k) : null),
    setItem: (k, v) => data.set(k, String(v)),
    removeItem: (k) => data.delete(k),
    data,
  };
}

function runtime(overrides = {}) {
  return {
    WorkerId: 'worker-1',
    Enabled: true,
    CanChange: true,
    RowVersion: '7',
    SharePointHosts: ['contoso.sharepoint.com'],
    ProcessRecordUpdates: false,
    Registration: { Readiness: [{ Scope: 'account', Status: 'Ready' }], Error: null },
    ...overrides,
  };
}

// Loads index.html and shell.js the way the page does, fires DOMContentLoaded and returns what
// happened. `setup` sets how many rows the landing queries find; runtime null means Get refused.
async function boot(options = {}) {
  const {
    search = '?data=templates-' + BUILD,
    hash = '',
    session = storage(),
    setup = {},
    profile = runtime(),
    privileges = {},
    register = () => {},
    // Rows a table read returns, by table; other tables return the landing counts.
    rows = {},
    // The asx_ManageWork Summary counts, or null for none.
    summary = null,
    // A promise the runtime Get waits for, to model a slow server.
    getGate = null,
  } = options;
  const document = createDocument(html);
  const navigations = [];
  const calls = [];
  const counts = { asx_runtimetable: 1, asx_library: 1, asx_template: 1, ...setup };
  let current = profile;
  const xrm = {
    // Copies: objects made inside the vm context have its Object prototype, which deepEqual refuses.
    Navigation: {
      navigateTo: async (page) =>
        navigations.push({
          ...page,
          data: new URLSearchParams(page.webresourceName.split('?')[1] || '').get('data'),
        }),
    },
    Utility: {
      getGlobalContext: () => ({
        getClientUrl: () => 'https://example.test',
        userSettings: { userId: '{11111111-1111-1111-1111-111111111111}' },
      }),
    },
    WebApi: {
      retrieveMultipleRecords: async (table, query) => {
        calls.push([table, query]);
        if (rows[table]) return { entities: rows[table] };
        return { entities: Array.from({ length: counts[table] ?? 0 }, (_, i) => ({ i })) };
      },
      online: {
        execute: async (request) => {
          const name = request.getMetadata().operationName;
          const body = JSON.parse(request.Request);
          calls.push([name, body]);
          if (body.Command === 'Get' && getGate) await getGate;
          if (name === 'asx_ManageWork' && body.Command === 'Summary')
            return {
              ok: true,
              json: async () => ({
                Result: JSON.stringify({ Status: 'Summary', Summary: summary || {} }),
              }),
            };
          if (current === null)
            return {
              ok: false,
              json: async () => ({
                error: { message: 'Principal user is missing prvWriteasx_runtime.' },
              }),
            };
          if (body.Command === 'SetEnabled')
            current = { ...current, Enabled: body.Enabled, RowVersion: '8' };
          return { ok: true, json: async () => ({ Result: JSON.stringify(current) }) };
        },
      },
    },
  };
  const fetch = async (url) => {
    const name = /PrivilegeName='([^']+)'/.exec(url)?.[1];
    return {
      ok: true,
      json: async () => ({
        RolePrivileges: privileges[name] === false ? [] : [{ PrivilegeName: name }],
      }),
    };
  };
  const window = { parent: { Xrm: xrm }, location: { search, hash } };
  Object.defineProperty(window, 'sessionStorage', {
    get: () => {
      if (session.throwOnAccess) throw new Error('SecurityError');
      return session;
    },
  });
  const context = vm.createContext({
    window,
    document,
    fetch,
    Intl,
    URLSearchParams,
    navigator: { clipboard: { writeText: async () => {} } },
    console,
    setTimeout,
    clearTimeout,
    setInterval: () => 0,
    clearInterval: () => {},
  });
  vm.runInContext(shell, context);
  register(window.AsxdUi, context);
  await document.fire('DOMContentLoaded');
  return { window, document, navigations, calls, session, xrm, ui: window.AsxdUi };
}
const visible = (d) =>
  ['templates', 'access', 'monitor', 'settings'].filter((id) => !d.getElementById(id).hidden);
const selected = (d) => visible(d)[0];
// The h1s a page shows: a hidden ancestor inside the page (the editor while the overview
// shows) hides its h1; the page section's own hidden does not count.
const shownH1s = (section) =>
  section.querySelectorAll('h1').filter((h) => {
    const hidden = h.closest('[hidden]');
    return !hidden || hidden === section;
  });

(async () => {
  {
    // First launch with setup complete lands on Monitor through the same navigation as a tab click.
    const run = await boot();
    assert.deepEqual(run.navigations, [
      {
        pageType: 'webresource',
        webresourceName: 'asx_admin/index.html?data=monitor-' + BUILD,
        data: 'monitor-' + BUILD,
      },
    ]);
    assert.equal(run.session.data.get('asxd.launched'), '1');
  }
  {
    // First launch with setup incomplete lands on the tab of the first open step.
    const settings = await boot({ profile: runtime({ WorkerId: EMPTY }) });
    assert.equal(settings.navigations[0].data, 'settings-' + BUILD);
    const access = await boot({ setup: { asx_library: 0 } });
    assert.equal(access.navigations[0].data, 'access-' + BUILD);
    const templates = await boot({ setup: { asx_template: 0 } });
    assert.deepEqual(templates.navigations, [], 'No published template stays on Folder templates');
    assert.equal(selected(templates.document), 'templates');
    assert.deepEqual(visible(templates.document), ['templates'], 'Only the active panel is shown');
    // A caller who cannot read the runtime skips the worker condition.
    const refused = await boot({ profile: null });
    assert.equal(refused.navigations[0].data, 'monitor-' + BUILD);
  }
  {
    // Later loads, and loads where storage fails, never redirect.
    const later = await boot({ session: storage({ 'asxd.launched': '1' }) });
    assert.deepEqual(later.navigations, []);
    assert.equal(selected(later.document), 'templates');
    // Review Focus 4: storage that throws counts as "already launched".
    const broken = await boot({ session: storage({}, true) });
    assert.deepEqual(broken.navigations, []);
    assert.equal(selected(broken.document), 'templates');
    const inaccessible = storage();
    inaccessible.throwOnAccess = true;
    const blocked = await boot({ session: inaccessible });
    assert.deepEqual(blocked.navigations, []);
  }
  {
    // Order: stored deep link, then hash, then data, then landing. The stored link is consumed once.
    const link = await boot({
      search: '?data=monitor-' + BUILD,
      hash: '#settings',
      session: storage({ 'asxd.deeplink': JSON.stringify({ tab: 'access', library: 'lib-1' }) }),
    });
    assert.equal(selected(link.document), 'access');
    assert.deepEqual({ ...link.ui.deeplink() }, { tab: 'access', library: 'lib-1' });
    assert.equal(link.session.data.has('asxd.deeplink'), false);
    const hash = await boot({ search: '?data=monitor-' + BUILD, hash: '#settings' });
    assert.equal(selected(hash.document), 'settings');
    const data = await boot({
      search: '?data=monitor-' + BUILD,
      session: storage({ 'asxd.launched': '1' }),
    });
    assert.equal(selected(data.document), 'monitor');
    for (const [legacy, tab] of [
      ['#runtime', 'settings'],
      ['#operations', 'monitor'],
      ['#administration', 'monitor'],
      ['#author', 'templates'],
    ])
      assert.equal(selected((await boot({ hash: legacy })).document), tab, legacy);
    const params = await boot({ hash: '#monitor?operation=folderjob:abc&record=account:r-1' });
    assert.deepEqual(
      { ...params.ui.deeplink() },
      { tab: 'monitor', operation: 'folderjob:abc', record: 'account:r-1' },
    );
    // Check a record's template rides along in the hash (Task 7's Check link carries it).
    const checked = await boot({ hash: '#monitor?record=account:r-1&template=t-1' });
    assert.deepEqual(
      { ...checked.ui.deeplink() },
      { tab: 'monitor', record: 'account:r-1', template: 't-1' },
    );
    // Review Focus 4: a malformed stored link and unknown tab names fall through to the next rule.
    const malformed = await boot({
      search: '?data=settings-' + BUILD,
      session: storage({ 'asxd.launched': '1', 'asxd.deeplink': '{not json' }),
    });
    assert.equal(selected(malformed.document), 'settings');
    const unknown = await boot({
      hash: '#bogus',
      search: '?data=bogus-x',
      session: storage({ 'asxd.launched': '1' }),
    });
    assert.equal(selected(unknown.document), 'templates');
  }
  {
    // No in-page header, tab bar, shell card or chip; each page is a plain section with one h1.
    const run = await boot({ session: storage({ 'asxd.launched': '1' }) });
    const d = run.document;
    for (const id of ['tabs', 'tabPrompt', 'automationChip', 'monitorBadge'])
      assert.equal(d.getElementById(id), null, id + ' is gone');
    assert.equal(d.querySelectorAll('.shell').length, 0);
    assert.equal(d.querySelectorAll('[role=tab]').length, 0);
    for (const page of ['templates', 'access', 'monitor', 'settings']) {
      const section = d.getElementById(page);
      assert.equal(section.getAttribute('role'), null, page);
      assert.equal(shownH1s(section).length, 1, page);
      const name = d.getElementById(section.getAttribute('aria-labelledby'));
      assert.equal(name, shownH1s(section)[0], page + ' is named by the h1 it shows');
      assert.equal(name.getAttribute('tabindex'), '-1', page + ' h1 takes focus');
    }
    assert.deepEqual(visible(d), ['templates']);
  }
  {
    // Pages still open from ?data=, #hash and a stored link (routing unchanged).
    for (const [search, hash, page] of [
      ['?data=monitor-' + BUILD, '', 'monitor'],
      ['', '#settings', 'settings'],
      ['', '#runtime', 'settings'],
    ]) {
      const run = await boot({ search, hash, session: storage({ 'asxd.launched': '1' }) });
      assert.deepEqual(visible(run.document), [page], search + hash);
    }
  }
  {
    // Unsaved template edits ask at the top of the page content before another page opens.
    const choices = [];
    const run = await boot({
      session: storage({ 'asxd.launched': '1' }),
      register: (ui) =>
        ui.setDirtyGuard(() => ({
          template: 'Account onboarding',
          save: async () => choices.push('save'),
          discard: () => choices.push('discard'),
        })),
    });
    const d = run.document;
    const going = run.ui.navigate('settings');
    await d.settle();
    const prompt = d.getElementById('leavePrompt');
    assert.equal(d.querySelector('main').firstElementChild, prompt);
    assert.match(prompt.visibleText, /You have unsaved changes to Account onboarding\./);
    assert.deepEqual(
      prompt.querySelectorAll('button').map((b) => b.textContent),
      ['Save draft', 'Discard changes', 'Stay'],
    );
    prompt
      .querySelectorAll('button')
      .find((b) => b.textContent === 'Stay')
      .click();
    await going;
    assert.deepEqual(run.navigations, []);
    assert.deepEqual(choices, []);
  }
  {
    // The prompt takes focus; Stay returns it to the page's h1, Discard changes and Save draft
    // go, Escape keeps, and a link to another page asks the same question.
    let saved = 0,
      discarded = 0;
    const run = await boot({
      session: storage({ 'asxd.launched': '1' }),
      register: (ui) =>
        ui.setDirtyGuard(() => ({
          template: 'Account onboarding',
          save: async () => saved++,
          discard: () => discarded++,
        })),
    });
    const d = run.document;
    const choose = (label) =>
      d
        .getElementById('leavePrompt')
        .querySelectorAll('button')
        .find((b) => b.textContent === label)
        .click();
    d.track(run.ui.navigate('settings'));
    await d.settle();
    assert.equal(d.activeElement.textContent, 'You have unsaved changes to Account onboarding.');
    choose('Stay');
    await d.settle();
    assert.deepEqual(run.navigations, []);
    assert.equal(d.activeElement.id, 'templates-title', 'Focus returns to the page heading');
    d.track(run.ui.navigate('settings'));
    await d.settle();
    choose('Discard changes');
    await d.settle();
    assert.equal(discarded, 1);
    assert.equal(run.navigations.at(-1).data, 'settings-' + BUILD);
    d.track(run.ui.navigate('monitor'));
    await d.settle();
    choose('Save draft');
    await d.settle();
    assert.equal(saved, 1);
    assert.equal(run.navigations.at(-1).data, 'monitor-' + BUILD);
    // Escape in the prompt keeps the page.
    d.track(run.ui.navigate('access'));
    await d.settle();
    d.activeElement.key('Escape');
    await d.settle();
    assert.equal(run.navigations.length, 2);
    // A link to another page (the library link in the template editor) asks the same question:
    // the reload would lose the edits. Stay keeps the page; Discard changes follows the link.
    d.track(run.ui.navigate('access', { library: 'lib-1' }));
    await d.settle();
    assert.equal(d.activeElement.textContent, 'You have unsaved changes to Account onboarding.');
    choose('Stay');
    await d.settle();
    assert.equal(run.navigations.length, 2);
    assert.equal(run.session.data.has('asxd.deeplink'), false, 'Stay stores no link');
    d.track(run.ui.navigate('access', { library: 'lib-1' }));
    await d.settle();
    choose('Discard changes');
    await d.settle();
    assert.equal(discarded, 2);
    assert.equal(run.navigations.at(-1).data, 'access-' + BUILD);
    assert.equal(
      run.session.data.get('asxd.deeplink'),
      JSON.stringify({ library: 'lib-1', tab: 'access' }),
    );
    // A failed save reports in the Folder templates feedback line and stays.
    const failing = await boot({
      session: storage({ 'asxd.launched': '1' }),
      register: (ui) =>
        ui.setDirtyGuard(() => ({
          template: 'Account onboarding',
          save: async () => {
            throw new Error('Draft changed; reload before publishing.');
          },
          discard: () => {},
        })),
    });
    const answer = failing.ui.confirmLeave();
    await failing.document.settle();
    failing.document
      .getElementById('leavePrompt')
      .querySelectorAll('button')
      .find((b) => b.textContent === 'Save draft')
      .click();
    assert.equal(await answer, false);
    assert.equal(
      failing.document.getElementById('fb-templates').textContent,
      'Draft changed; reload before publishing.',
    );
    // Without unsaved edits there is nothing to ask.
    const clean = await boot({ session: storage({ 'asxd.launched': '1' }) });
    assert.equal(await clean.ui.confirmLeave(), true);
    assert.equal(clean.document.getElementById('leavePrompt').children.length, 0);
  }
  {
    // sidePanel: a dialog named by its heading; focus in; an Escape that starts inside a popover
    // or menu in it does not close it (the guard itself: nothing else handles that Escape);
    // Escape closes it and focus returns to the invoker. The confirmation case is proven in the
    // browser test: here the confirmation removes itself before the event bubbles.
    const run = await boot({ session: storage({ 'asxd.launched': '1' }) });
    const { document: d, ui } = run;
    const host = d.getElementById('monitor');
    const invoker = ui.button('Tools', () => {});
    host.append(invoker);
    const panel = d.createElement('aside');
    panel.hidden = true;
    const heading = ui.el('h2', 'Check a record');
    heading.id = 'test-panel-title';
    const popover = ui.el('div', null, 'popover');
    const field = ui.button('Account Name', () => {});
    popover.append(field);
    const menu = ui.el('ul', null, 'menu');
    const item = ui.button('Copy ID', () => {});
    menu.append(item);
    panel.append(heading, popover, menu);
    host.append(panel);
    const closed = [];
    invoker.focus();
    ui.sidePanel(panel, invoker, { onClose: () => closed.push('closed') });
    assert.equal(panel.hidden, false);
    assert.equal(panel.getAttribute('role'), 'dialog');
    assert.equal(panel.getAttribute('aria-labelledby'), 'test-panel-title');
    assert.equal(d.activeElement, heading);
    field.key('Escape');
    assert.equal(panel.hidden, false, 'An Escape inside a popover belongs to the popover');
    item.key('Escape');
    assert.equal(panel.hidden, false, 'An Escape inside a menu belongs to the menu');
    heading.key('Escape');
    assert.equal(panel.hidden, true);
    assert.equal(d.activeElement, invoker);
    assert.deepEqual(closed, ['closed']);
    // Opening another panel closes the open one without moving focus back; close() on a
    // closed panel does nothing; a redrawn invoker is found again by its focus key.
    const first = ui.sidePanel(panel, invoker);
    const other = d.createElement('aside');
    other.id = 'other-panel';
    other.hidden = true;
    other.append(ui.el('h3', 'Look up an operation'));
    host.append(other);
    const keyed = ui.button('Look up', () => {});
    keyed.dataset.focusKey = 'tools:lookup';
    host.append(keyed);
    const second = ui.sidePanel(other, keyed);
    assert.equal(panel.hidden, true, 'Opening a panel closes the open one');
    assert.equal(other.getAttribute('aria-labelledby'), 'other-panel-title');
    assert.equal(d.activeElement.id, 'other-panel-title');
    first.close();
    assert.equal(d.activeElement.id, 'other-panel-title', 'A closed panel stays closed');
    keyed.remove();
    const again = ui.button('Look up', () => {});
    again.dataset.focusKey = 'tools:lookup';
    host.append(again);
    second.close();
    assert.equal(d.activeElement, again);
  }
  {
    // tokens: field tokens render as chips; an unknown token stays as text; no markup is parsed.
    const run = await boot({ session: storage({ 'asxd.launched': '1' }) });
    const label = (alias, column) =>
      alias === 'root' && column === 'projectnumber' ? 'Project Number' : null;
    const name = run.ui.tokens('P-{root.projectnumber} <b>', label);
    assert.equal(name.className, 'name');
    assert.equal(name.visibleText, 'P-Project Number <b>');
    assert.equal(name.querySelector('.token').textContent, 'Project Number');
    const unknown = run.ui.tokens('{root.gone}', label);
    assert.equal(unknown.visibleText, '{root.gone}');
    assert.equal(unknown.querySelector('.token'), null);
    // status pairs a dot with text; pills carry a tone.
    const ok = run.ui.status('ok', 'Ready');
    assert.equal(ok.querySelector('.dot').dataset.tone, 'ok');
    assert.equal(ok.querySelector('.dot').getAttribute('aria-hidden'), 'true');
    assert.equal(ok.visibleText, 'Ready');
    assert.equal(run.ui.pill('Live v3', 'ok').dataset.tone, 'ok');
    // plural counts with separators; ms reads "/Date(…)/" and ISO, NaN otherwise.
    assert.equal(run.ui.plural(1, 'problem', 'problems'), '1 problem');
    assert.equal(run.ui.plural(1214, 'problem', 'problems'), '1,214 problems');
    assert.equal(run.ui.plural(0, 'problem', 'problems'), '0 problems');
    assert.equal(run.ui.ms('/Date(1700000000000)/'), 1700000000000);
    assert.equal(run.ui.ms('2023-11-14T22:13:20Z'), 1700000000000);
    assert(Number.isNaN(run.ui.ms('not a date')));
    assert(Number.isNaN(run.ui.ms(null)));
  }
  {
    // automation(): the runtime Get when the caller may make it, else the Default runtime row.
    const admin = await boot({ session: storage({ 'asxd.launched': '1' }) });
    assert.deepEqual(
      { ...(await admin.ui.automation()) },
      { Enabled: true, ProcessRecordUpdates: false },
    );
    const operator = await boot({
      profile: null,
      session: storage({ 'asxd.launched': '1' }),
      rows: { asx_runtime: [{ asx_enabled: false, asx_processrecordupdates: true }] },
    });
    assert.deepEqual(
      { ...(await operator.ui.automation()) },
      { Enabled: false, ProcessRecordUpdates: true },
    );
    const read = operator.calls.find(([table]) => table === 'asx_runtime');
    assert.match(read[1], /\$filter=asx_name eq 'Default'/);
    assert.match(read[1], /\$top=2/);
    // A row that was read is kept: the next call reads nothing.
    await operator.ui.automation();
    assert.equal(operator.calls.filter(([table]) => table === 'asx_runtime').length, 1);
    // A read that finds no single Default row gives null and is not kept: the next call reads again.
    const missing = await boot({
      profile: null,
      session: storage({ 'asxd.launched': '1' }),
      rows: { asx_runtime: [] },
    });
    assert.equal(await missing.ui.automation(), null);
    assert.equal(await missing.ui.automation(), null);
    assert.equal(missing.calls.filter(([table]) => table === 'asx_runtime').length, 2);
  }
  {
    // problemPill: one Summary per page load, the five problem lists summed (re-runs are not
    // problems); hidden at 0 and without the Operator role; "1 problem", "{N} problems", and
    // "5,000+ problems" when a list is capped; a click opens Monitor.
    const pillFor = async (summary, privileges = {}) => {
      const run = await boot({ session: storage({ 'asxd.launched': '1' }), summary, privileges });
      const host = run.ui.el('span');
      run.document.getElementById('templates').append(host);
      const pill = run.ui.problemPill(host);
      await run.document.settle();
      return { run, host, pill };
    };
    const summaries = (run) => run.calls.filter(([, body]) => body?.Command === 'Summary').length;
    const zero = await pillFor({
      BlockedRecords: 0,
      WaitingRecords: 0,
      BlockedJobs: 0,
      RetryingJobs: 0,
      NotCaptured: 0,
      TemplateRuns: 2,
    });
    assert.equal(zero.host.firstElementChild, zero.pill);
    assert.equal(zero.pill.hidden, true, 'Hidden at 0');
    const one = await pillFor({ BlockedJobs: 1, TemplateRuns: 3 });
    assert.equal(one.pill.tagName, 'BUTTON');
    assert.ok(one.pill.classList.contains('problem-pill'));
    assert.equal(one.pill.hidden, false);
    assert.equal(one.pill.textContent, 'Monitor · 1 problem');
    const many = await pillFor({
      BlockedRecords: 3,
      WaitingRecords: 6,
      BlockedJobs: 1,
      RetryingJobs: 4,
      NotCaptured: 1200,
    });
    assert.equal(many.pill.textContent, 'Monitor · 1,214 problems');
    const capped = await pillFor({
      BlockedRecords: 5000,
      WaitingRecords: 2,
      Capped: ['BlockedRecords'],
    });
    assert.equal(capped.pill.textContent, 'Monitor · 5,000+ problems');
    const second = many.run.ui.problemPill(many.run.ui.el('span'));
    await many.run.document.settle();
    assert.equal(second.textContent, 'Monitor · 1,214 problems');
    assert.equal(summaries(many.run), 1, 'One Summary per page load');
    const none = await pillFor({ BlockedJobs: 4 }, { prvCreateasx_operatorcommand: false });
    assert.equal(none.pill.hidden, true, 'Hidden without the Operator role');
    assert.equal(summaries(none.run), 0, 'No Summary without the Operator role');
    many.pill.click();
    await many.run.document.settle();
    assert.equal(many.run.navigations.at(-1).data, 'monitor-' + BUILD);
  }
  {
    // Feedback lines replace the banner: success is a status, an error an alert.
    const run = await boot({ session: storage({ 'asxd.launched': '1' }) });
    const d = run.document;
    assert.equal(d.getElementById('status'), null, 'No global banner');
    assert.equal(d.querySelectorAll('footer').length, 0);
    run.ui.feedback('templates', 'Schedule saved.');
    assert.equal(d.getElementById('fb-templates').textContent, 'Schedule saved.');
    assert.equal(d.getElementById('fb-templates').getAttribute('role'), 'status');
    run.ui.feedback('templates', 'Save failed: the draft changed.', 'error');
    assert.equal(d.getElementById('fb-templates').getAttribute('role'), 'alert');
    run.ui.clearFeedback('templates');
    assert.equal(d.getElementById('fb-templates').textContent, '');
    // Details: Copy names what it copies (spec 5.1).
    const box = run.ui.details('folderjob:abc', 'Details', 'Contoso Ltd');
    const copy = box.querySelectorAll('button').find((b) => b.textContent === 'Copy');
    assert.equal(copy.getAttribute('aria-label'), 'Copy details for Contoso Ltd');
  }
  {
    // The shared confirmation renders after the invoker's button row, takes focus on its text,
    // and Escape keeps, returning focus to the invoker.
    const run = await boot({ session: storage({ 'asxd.launched': '1' }) });
    const d = run.document;
    const row = run.ui.el('div', null, 'row');
    row.setAttribute('data-actions', '');
    const invoker = run.ui.button('Delete template', () => {});
    row.append(invoker);
    d.getElementById('templates').append(row);
    invoker.focus();
    const answer = run.ui.confirmInline(invoker, {
      text: 'Delete Account onboarding and all its versions?',
      confirm: 'Delete template',
      keep: 'Keep template',
      danger: true,
    });
    const box = row.nextElementSibling;
    assert(box.classList.contains('confirm'));
    assert.equal(box.getAttribute('role'), 'group');
    assert.equal(d.activeElement.textContent, 'Delete Account onboarding and all its versions?');
    assert.equal(
      box
        .querySelectorAll('button')
        .map((b) => b.textContent)
        .join('|'),
      'Delete template|Keep template',
    );
    d.activeElement.key('Escape');
    assert.equal(await answer, false);
    assert.equal(d.activeElement, invoker);
    assert.equal(box.isConnected, false);
  }
  {
    // No Dataverse connection: the active panel shows one alert and nothing else.
    const document = createDocument(html);
    const window = { location: { search: '', hash: '' } };
    vm.runInNewContext(shell, {
      window,
      document,
      Intl,
      URLSearchParams,
      console,
      navigator: {},
      setTimeout,
      clearTimeout,
      setInterval: () => 0,
      clearInterval: () => {},
    });
    await document.fire('DOMContentLoaded');
    const alert = document.querySelector('[role=alert]');
    assert.equal(alert.textContent, 'Open this page from the Ascentix Documents app to connect.');
  }
  {
    // Fix round 1, item 1: a slow runtime Get does not hold back the page's start. onRuntime
    // listeners hear it when Get returns, and the page sends Get once.
    let release;
    const gate = new Promise((resolve) => (release = resolve));
    let started = 0;
    const heard = [];
    const run = await boot({
      session: storage({ 'asxd.launched': '1' }),
      getGate: gate,
      register: (ui) => {
        ui.onTab('templates', () => started++);
        ui.onRuntime((result) => heard.push(result?.WorkerId ?? null));
      },
    });
    const d = run.document;
    assert.equal(started, 1, 'The page starts while Get is still running');
    assert.equal(run.ui.runtime(), null);
    release();
    await d.settle();
    assert.equal(run.ui.runtime().WorkerId, 'worker-1');
    assert.deepEqual(heard, ['worker-1']);
    assert.equal(run.calls.filter(([, b]) => b?.Command === 'Get').length, 1, 'One Get per load');
    // A refused Get still tells listeners, with null, so they stop waiting.
    const none = [];
    await boot({
      profile: null,
      session: storage({ 'asxd.launched': '1' }),
      register: (ui) => ui.onRuntime((result) => none.push(result)),
    });
    assert.deepEqual(none, [null]);
  }
  {
    // Fix round 1, item 2: withFocus.
    const run = await boot({ session: storage({ 'asxd.launched': '1' }) });
    const d = run.document;
    const { ui } = run;
    const list = ui.el('section');
    list.id = 'jobs';
    list.setAttribute('data-focus-scope', '');
    const heading = ui.el('h3', 'Blocked jobs');
    heading.tabIndex = -1;
    heading.setAttribute('data-focus-heading', '');
    d.getElementById('monitor').append(list);
    const draw = (keys) => {
      const rows = keys.map((key) => {
        const row = ui.el('div', null, 'job');
        row.setAttribute('data-focus-row', '');
        const retry = ui.button('Retry', () => {});
        retry.setAttribute('data-focus-key', 'retry:' + key);
        row.append(retry);
        return row;
      });
      const more = ui.button('Load more', () => {});
      more.setAttribute('data-focus-key', 'more');
      list.replaceChildren(heading, ...rows, ...(keys.length > 2 ? [more] : []));
    };
    draw(['a', 'b', 'c']);
    // The same key after a redraw keeps focus.
    list.querySelector('[data-focus-key="retry:b"]').focus();
    ui.withFocus(() => draw(['a', 'b', 'c']));
    assert.equal(d.activeElement.getAttribute('data-focus-key'), 'retry:b');
    assert.equal(d.activeElement.isConnected, true);
    // Its row gone: the row that took its place.
    ui.withFocus(() => draw(['a', 'c', 'd']));
    assert.equal(d.activeElement.getAttribute('data-focus-key'), 'retry:c');
    // Focus outside any row (Load more) that disappears: the heading, not the first row.
    list.querySelector('[data-focus-key="more"]').focus();
    ui.withFocus(() => draw(['a']));
    // assert(a === b) here: a failing assert.equal would print the whole fake DOM tree.
    assert(
      d.activeElement === heading,
      'Focus moves to the list heading, got ' + d.activeElement.textContent,
    );
    // No focus key: focus is left alone.
    const outside = ui.button('Elsewhere', () => {});
    d.getElementById('monitor').append(outside);
    outside.focus();
    ui.withFocus(() => draw(['a', 'b']));
    assert(d.activeElement === outside, 'Focus stays where it was');

    // time: /Date(ms)/ and ISO values carry a machine-readable datetime; a missing value is empty.
    const legacy = ui.time('/Date(1700000000000)/');
    assert.equal(legacy.tagName, 'TIME');
    assert.equal(legacy.getAttribute('datetime'), '2023-11-14T22:13:20Z');
    assert.notEqual(legacy.textContent, '');
    assert.equal(ui.time('2026-10-06T12:00:00Z').getAttribute('datetime'), '2026-10-06T12:00:00Z');
    const missing = ui.time(null);
    assert.equal(missing.textContent, '');
    assert.equal(missing.hasAttribute('datetime'), false);
    assert.equal(ui.time('not a date').hasAttribute('datetime'), false);

    // help: the only maker of class="help".
    const note = ui.help('saveHelp', 'Saving starts the next draft.');
    assert.equal(note.tagName, 'P');
    assert.equal(note.className, 'help');
    assert.equal(note.id, 'saveHelp');
    assert.equal(note.textContent, 'Saving starts the next draft.');

    // busy: the clicked button shows the -ing label, the row is aria-busy and its other buttons
    // wait; a second press does nothing; everything comes back after the work.
    const row = ui.el('div', null, 'row');
    row.setAttribute('data-actions', '');
    const save = ui.button('Save', () => {});
    const other = ui.button('Publish', () => {});
    row.append(save, other);
    d.getElementById('templates').append(row);
    let finish;
    let calls = 0;
    const work = ui.busy(save, 'Saving…', 'templates', () => {
      calls++;
      return new Promise((resolve) => (finish = resolve));
    });
    assert.equal(save.textContent, 'Saving…');
    assert.equal(save.classList.contains('is-busy'), true);
    assert.equal(row.getAttribute('aria-busy'), 'true');
    assert.equal(other.disabled, true);
    assert.equal(await ui.busy(save, 'Saving…', 'templates', async () => calls++), undefined);
    assert.equal(calls, 1, 'A second press while busy does nothing');
    finish('saved');
    assert.equal(await work, 'saved');
    assert.equal(save.textContent, 'Save');
    assert.equal(save.classList.contains('is-busy'), false);
    assert.equal(row.hasAttribute('aria-busy'), false);
    assert.equal(other.disabled, false);
    // An error goes to the area's feedback line as an alert; without an area it is rethrown.
    assert.equal(
      await ui.busy(save, 'Saving…', 'templates', async () => {
        throw new Error('The draft changed.');
      }),
      undefined,
    );
    assert.equal(d.getElementById('fb-templates').textContent, 'The draft changed.');
    assert.equal(d.getElementById('fb-templates').getAttribute('role'), 'alert');
    await assert.rejects(
      ui.busy(save, 'Saving…', null, async () => {
        throw new Error('Thrown on');
      }),
      /Thrown on/,
    );
    assert.equal(other.disabled, false);

    // api: the server's error message, or a plain refusal when the body is not JSON.
    run.xrm.WebApi.online.execute = async () => ({
      ok: false,
      json: async () => ({ error: { message: 'Select a library.' } }),
    });
    await assert.rejects(ui.api('asx_SitesAccess', {}), /Select a library\./);
    run.xrm.WebApi.online.execute = async () => ({
      ok: false,
      json: async () => {
        throw new SyntaxError('Unexpected token <');
      },
    });
    await assert.rejects(ui.api('asx_SitesAccess', {}), /The server refused the request\./);
    let sent;
    run.xrm.WebApi.online.execute = async (request) => {
      sent = request;
      return { ok: true, json: async () => ({ Result: JSON.stringify({ Status: 'Ready' }) }) };
    };
    assert.equal((await ui.api('asx_SitesAccess', { Command: 'Get' })).Status, 'Ready');
    assert.equal(sent.getMetadata().operationName, 'asx_SitesAccess');
    assert.equal(sent.Request, '{"Command":"Get"}');
  }
  {
    // Fix round 1, item 3: a new confirmation in the same place answers the open one with its
    // keep value, so the action waiting on it ends; static .confirm elements stay.
    const run = await boot({ session: storage({ 'asxd.launched': '1' }) });
    const d = run.document;
    const area = d.getElementById('settings');
    // Static .confirm markup in the same place, which ask() must leave alone.
    const fixed = run.ui.el('div', 'Static confirmation', 'confirm');
    fixed.id = 'staticConfirm';
    area.append(fixed);
    const row = run.ui.el('div', null, 'row');
    row.setAttribute('data-actions', '');
    const first = run.ui.button('Remove', () => {});
    const second = run.ui.button('Remove all', () => {});
    row.append(first, second);
    area.append(row);
    const one = run.ui.confirmInline(first, { text: 'Remove A?', confirm: 'Remove', keep: 'Keep' });
    const two = run.ui.ask(second, {
      text: 'Remove all?',
      keep: 'keep-all',
      choices: [
        { value: 'go', label: 'Remove all' },
        { value: 'keep-all', label: 'Keep all' },
      ],
    });
    // Fails instead of hanging when the replaced confirmation never answers.
    // The timer is not unref'd, so a promise that never settles fails the test, not exit 0.
    const within = (promise) => {
      let timer;
      return Promise.race([
        promise,
        new Promise((_, reject) => {
          timer = setTimeout(
            () => reject(new Error('The replaced confirmation never answered')),
            1000,
          );
        }),
      ]).finally(() => clearTimeout(timer));
    };
    assert.equal(
      await within(one),
      false,
      'The replaced confirmation resolves with its keep value',
    );
    assert.equal(area.querySelectorAll('.confirm[role=group]').length, 1);
    assert.equal(d.activeElement.textContent, 'Remove all?');
    assert(d.getElementById('staticConfirm'), 'A static .confirm element survives');
    const three = run.ui.confirmInline(first, {
      text: 'Remove B?',
      confirm: 'Remove',
      keep: 'Keep',
    });
    assert.equal(await two, 'keep-all');
    assert(d.getElementById('staticConfirm'));
    d.activeElement.key('Escape');
    assert.equal(await three, false);
    assert.equal(area.querySelectorAll('.confirm[role=group]').length, 0);
  }
  {
    // Task 9 fix round 1: a ⋯ menu closes when focus leaves both its button and its list (Tab
    // out of it), and stays open while focus moves between them.
    const run = await boot({ session: storage({ 'asxd.launched': '1' }) });
    const d = run.document;
    const trigger = d.getElementById('template-menu'),
      list = d.getElementById('template-menu-list'),
      item = (id) => d.getElementById(id);
    run.ui.menu(trigger, list);
    trigger.key('ArrowDown');
    assert.equal(list.hidden, false);
    assert.equal(d.activeElement.id, 'menu-history');
    trigger.dispatchEvent(new FakeEvent('focusout', { relatedTarget: item('menu-history') }));
    item('menu-history').dispatchEvent(
      new FakeEvent('focusout', { relatedTarget: item('menu-schedule') }),
    );
    assert.equal(list.hidden, false, 'Moving between items keeps it open');
    item('menu-schedule').dispatchEvent(new FakeEvent('focusout', { relatedTarget: item('save') }));
    assert.equal(list.hidden, true, 'Tab out of the menu closes it');
    assert.equal(trigger.getAttribute('aria-expanded'), 'false');
    // A focus change with no element to go to (another window) leaves it to the click handler.
    trigger.key('ArrowDown');
    item('menu-history').dispatchEvent(new FakeEvent('focusout', { relatedTarget: null }));
    assert.equal(list.hidden, false);
  }
  {
    // A ⋯ menu whose button has left the page stops listening for document clicks at the next
    // click, so menus drawn on every table redraw do not pile up.
    const run = await boot({ session: storage({ 'asxd.launched': '1' }) });
    const d = run.document;
    const clicks = () => (d.listeners.get('click') || []).length;
    const before = clicks();
    const trigger = run.ui.button('⋯', null);
    const list = run.ui.el('ul', null, 'menu');
    list.hidden = true;
    d.getElementById('monitor').append(trigger, list);
    run.ui.menu(trigger, list);
    assert.equal(clicks(), before + 1);
    d.dispatchEvent(new FakeEvent('click', { target: d.body }));
    assert.equal(clicks(), before + 1, 'A connected menu keeps listening');
    trigger.remove();
    list.remove();
    d.dispatchEvent(new FakeEvent('click', { target: d.body }));
    assert.equal(clicks(), before);
  }
  console.log(
    'PASS shell contract: pages without tabs, landing, deep links, unsaved prompt at the top, side panel, tokens, status, plural, ms, automation, problem pill, feedback, confirmation (a replaced one answers keep), offline, a slow Get that does not delay the page, withFocus, time, help, busy, api errors, the ⋯ menu closing when focus leaves it, and a redrawn-away menu letting go of its document listener. Fake DOM; browser QA separate.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
