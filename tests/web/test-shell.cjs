'use strict';
// Shell contract with a fake DOM and mocked Dataverse: which tab a load shows, landing, deep
// links, tab clicks that reload, the unsaved-changes prompt, keyboard focus, the automation chip,
// feedback lines and the inline confirmation. Not a browser test.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createDocument } = require('./fake-dom.cjs');
const base = path.resolve(__dirname, '../../client/admin');
const html = fs.readFileSync(path.join(base, 'index.html'), 'utf8');
const shell = fs.readFileSync(path.join(base, 'shell.js'), 'utf8');
const BUILD = 'ui20261006nav1';
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
  } = options;
  const document = createDocument(html);
  const navigations = [];
  const calls = [];
  const counts = { asx_runtimetable: 1, asx_library: 1, asx_template: 1, ...setup };
  let current = profile;
  const xrm = {
    // Copies: objects made inside the vm context have its Object prototype, which deepEqual refuses.
    Navigation: { navigateTo: async (page) => navigations.push({ ...page }) },
    Utility: {
      getGlobalContext: () => ({
        getClientUrl: () => 'https://example.test',
        userSettings: { userId: '{11111111-1111-1111-1111-111111111111}' },
      }),
    },
    WebApi: {
      retrieveMultipleRecords: async (table, query) => {
        calls.push([table, query]);
        return { entities: Array.from({ length: counts[table] ?? 0 }, (_, i) => ({ i })) };
      },
      online: {
        execute: async (request) => {
          const name = request.getMetadata().operationName;
          const body = JSON.parse(request.Request);
          calls.push([name, body]);
          if (current === null)
            return {
              ok: false,
              json: async () => ({
                error: { message: 'Principal user is missing prvWriteasx_runtime.' },
              }),
            };
          if (body.Command === 'Save')
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
  return { window, document, navigations, calls, session, ui: window.AsxdUi };
}
const selected = (d) =>
  d.querySelectorAll('[role=tab]').find((t) => t.getAttribute('aria-selected') === 'true')?.id;
const visible = (d) =>
  ['templates', 'access', 'monitor', 'settings'].filter((id) => !d.getElementById(id).hidden);

(async () => {
  {
    // First launch with setup complete lands on Monitor through the same navigation as a tab click.
    const run = await boot();
    assert.deepEqual(run.navigations, [
      {
        pageType: 'webresource',
        webresourceName: 'asx_admin/index.html',
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
    assert.equal(selected(templates.document), 'tab-templates');
    assert.deepEqual(visible(templates.document), ['templates'], 'Only the active panel is shown');
    // A caller who cannot read the runtime skips the worker condition.
    const refused = await boot({ profile: null });
    assert.equal(refused.navigations[0].data, 'monitor-' + BUILD);
  }
  {
    // Later loads, and loads where storage fails, never redirect.
    const later = await boot({ session: storage({ 'asxd.launched': '1' }) });
    assert.deepEqual(later.navigations, []);
    assert.equal(selected(later.document), 'tab-templates');
    // Review Focus 4: storage that throws counts as "already launched".
    const broken = await boot({ session: storage({}, true) });
    assert.deepEqual(broken.navigations, []);
    assert.equal(selected(broken.document), 'tab-templates');
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
    assert.equal(selected(link.document), 'tab-access');
    assert.deepEqual({ ...link.ui.deeplink() }, { tab: 'access', library: 'lib-1' });
    assert.equal(link.session.data.has('asxd.deeplink'), false);
    const hash = await boot({ search: '?data=monitor-' + BUILD, hash: '#settings' });
    assert.equal(selected(hash.document), 'tab-settings');
    const data = await boot({
      search: '?data=monitor-' + BUILD,
      session: storage({ 'asxd.launched': '1' }),
    });
    assert.equal(selected(data.document), 'tab-monitor');
    for (const [legacy, tab] of [
      ['#runtime', 'tab-settings'],
      ['#operations', 'tab-monitor'],
      ['#administration', 'tab-monitor'],
      ['#author', 'tab-templates'],
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
    assert.equal(selected(malformed.document), 'tab-settings');
    const unknown = await boot({
      hash: '#bogus',
      search: '?data=bogus-x',
      session: storage({ 'asxd.launched': '1' }),
    });
    assert.equal(selected(unknown.document), 'tab-templates');
  }
  {
    // A tab click reloads to that tab's menu entry and asks for focus on it after the reload.
    const run = await boot({ session: storage({ 'asxd.launched': '1' }) });
    run.document.getElementById('tab-monitor').click();
    await run.document.settle();
    assert.equal(run.navigations.at(-1).data, 'monitor-' + BUILD);
    assert.equal(run.session.data.get('asxd.focusTab'), 'monitor');
    // After the reload the selected tab has focus and the flag is cleared.
    const after = await boot({
      search: '?data=monitor-' + BUILD,
      session: storage({ 'asxd.launched': '1', 'asxd.focusTab': 'monitor' }),
    });
    assert.equal(after.document.activeElement.id, 'tab-monitor');
    assert.equal(after.session.data.has('asxd.focusTab'), false);
  }
  {
    // Unsaved template edits: the prompt takes focus; Stay, Discard changes and Save draft.
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
    d.getElementById('tab-settings').click();
    await d.settle();
    const prompt = d.getElementById('tabPrompt').querySelector('.confirm');
    assert(prompt, 'The prompt renders next to the tab bar');
    assert.equal(d.activeElement.textContent, 'You have unsaved changes to Account onboarding.');
    const choose = (label) =>
      prompt.querySelectorAll('button').find((b) => b.textContent === label);
    choose('Stay').click();
    await d.settle();
    assert.deepEqual(run.navigations, []);
    assert.equal(d.activeElement.id, 'tab-templates');
    d.getElementById('tab-settings').click();
    await d.settle();
    d.getElementById('tabPrompt')
      .querySelectorAll('button')
      .find((b) => b.textContent === 'Discard changes')
      .click();
    await d.settle();
    assert.equal(discarded, 1);
    assert.equal(run.navigations.at(-1).data, 'settings-' + BUILD);
    d.getElementById('tab-monitor').click();
    await d.settle();
    d.getElementById('tabPrompt')
      .querySelectorAll('button')
      .find((b) => b.textContent === 'Save draft')
      .click();
    await d.settle();
    assert.equal(saved, 1);
    assert.equal(run.navigations.at(-1).data, 'monitor-' + BUILD);
    // Escape in the prompt keeps the page.
    d.getElementById('tab-access').click();
    await d.settle();
    d.activeElement.key('Escape');
    await d.settle();
    assert.equal(run.navigations.length, 2);
    // A link to another tab (the library link in the template editor) asks the same question:
    // the reload would lose the edits. Stay keeps the page; Discard changes follows the link.
    d.track(run.ui.navigate('access', { library: 'lib-1' }));
    await d.settle();
    assert.equal(d.activeElement.textContent, 'You have unsaved changes to Account onboarding.');
    d.getElementById('tabPrompt')
      .querySelectorAll('button')
      .find((b) => b.textContent === 'Stay')
      .click();
    await d.settle();
    assert.equal(run.navigations.length, 2);
    assert.equal(run.session.data.has('asxd.deeplink'), false, 'Stay stores no link');
    d.track(run.ui.navigate('access', { library: 'lib-1' }));
    await d.settle();
    d.getElementById('tabPrompt')
      .querySelectorAll('button')
      .find((b) => b.textContent === 'Discard changes')
      .click();
    await d.settle();
    assert.equal(discarded, 2);
    assert.equal(run.navigations.at(-1).data, 'access-' + BUILD);
    assert.equal(
      run.session.data.get('asxd.deeplink'),
      JSON.stringify({ library: 'lib-1', tab: 'access' }),
    );
  }
  {
    // Arrow keys, Home and End move focus only; the selected tab does not change.
    const run = await boot({ session: storage({ 'asxd.launched': '1' }) });
    const d = run.document;
    d.getElementById('tab-templates').focus();
    d.activeElement.key('ArrowRight');
    assert.equal(d.activeElement.id, 'tab-access');
    d.activeElement.key('End');
    assert.equal(d.activeElement.id, 'tab-settings');
    d.activeElement.key('ArrowRight');
    assert.equal(d.activeElement.id, 'tab-templates', 'Arrow keys wrap');
    d.activeElement.key('Home');
    assert.equal(d.activeElement.id, 'tab-templates');
    assert.equal(d.getElementById('tab-templates').getAttribute('tabindex'), '0');
    assert.deepEqual(run.navigations, []);
    assert.equal(selected(d), 'tab-templates');
  }
  {
    // Automation chip: five states (spec 2.5).
    const chip = async (profile) => {
      const run = await boot({ profile, session: storage({ 'asxd.launched': '1' }) });
      const d = run.document;
      return {
        run,
        d,
        text: d.getElementById('automationChipText').textContent,
        hidden: d.getElementById('automationChip').hidden,
      };
    };
    assert.equal((await chip(runtime())).text, 'Automation running');
    assert.equal((await chip(runtime({ Enabled: false }))).text, 'Automation paused');
    assert.equal(
      (
        await chip(
          runtime({
            Registration: { Readiness: [{ Scope: 'account', Status: 'Pending' }], Error: null },
          }),
        )
      ).text,
      'Automation running · needs attention',
    );
    assert.equal((await chip(runtime({ WorkerId: EMPTY }))).text, 'Automation not set up');
    assert.equal(
      (await chip(null)).hidden,
      true,
      'A caller who cannot read the runtime sees no chip',
    );
    // Turn on sends the stored row version and updates the chip.
    const paused = await chip(runtime({ Enabled: false }));
    const turnOn = paused.d.getElementById('automationChipAction');
    assert.equal(turnOn.hidden, false);
    turnOn.click();
    await paused.d.settle();
    const save = paused.run.calls.filter(([name]) => name === 'asx_RuntimeAdmin').at(-1)[1];
    assert.equal(save.Command, 'Save');
    assert.equal(save.Enabled, true);
    assert.equal(save.RowVersion, '7');
    assert.equal(save.WorkerId, 'worker-1');
    assert.equal(paused.d.getElementById('automationChipText').textContent, 'Automation running');
    // Not a System Administrator: Turn on is reachable but disabled, with the reason linked.
    const viewer = await chip(runtime({ Enabled: false, CanChange: false }));
    const off = viewer.d.getElementById('automationChipAction');
    assert.equal(off.getAttribute('aria-disabled'), 'true');
    const reason = viewer.d.getElementById(off.getAttribute('aria-describedby'));
    assert.equal(reason.textContent, 'Only a System Administrator can turn automation on.');
    off.click();
    await viewer.d.settle();
    assert.equal(viewer.run.calls.filter(([, b]) => b?.Command === 'Save').length, 0);
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
  console.log(
    'PASS shell contract: tab order, landing, deep links, menu sync, unsaved prompt, keyboard, chip, feedback, confirmation, offline. Fake DOM; browser QA separate.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
