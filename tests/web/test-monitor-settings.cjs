'use strict';
// Monitor and Settings with a fake DOM and mocked Dataverse: lists load on open, rows read as
// names with actions from the server, re-runs and recovery rows act, Check a record, Advanced and
// Settings. Not a browser test.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createDocument } = require('./fake-dom.cjs');
const base = path.resolve(__dirname, '../../client/admin');
const html = fs.readFileSync(path.join(base, 'index.html'), 'utf8');
const read = (name) => fs.readFileSync(path.join(base, name), 'utf8');
const BUILD = /const BUILD = '([^']+)'/.exec(read('shell.js'))[1];
const GUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;

function row(overrides) {
  return {
    Key: 'folderjob:abc',
    Kind: 'FolderJob',
    KindLabel: 'Folder job',
    Title: 'Contoso Ltd · Account documents',
    Record: {
      Table: 'account',
      TableLabel: 'Account',
      Id: '00000000-0000-0000-0000-0000000000a1',
      Name: 'Contoso Ltd',
    },
    TemplateId: '11111111-2222-3333-4444-555555555555',
    TemplateName: 'Account documents',
    Status: 'Blocked',
    Problem: 'The folder path is too long for SharePoint’s address limit.',
    Fix: 'Shorten the record’s value or the template’s folder names, then re-run the record.',
    Code: 'RequestUrlTooLong',
    More: [],
    SinceUtc: '2026-10-06T17:38:07Z',
    Attempt: 0,
    Actions: ['Retry', 'Cancel', 'OpenRecord', 'Check'],
    ...overrides,
  };
}

// Boots the page on a tab with every script loaded, the way the app opens it.
// `getGate` is a promise the runtime Get waits for, to model a slow server.
async function boot({
  tab = 'monitor',
  lists = {},
  summary = {},
  runtime = {},
  workers,
  // Rows a table read answers with, in place of the defaults below; an Error refuses the read.
  rows = {},
  // Privileges the caller lacks.
  missing = [],
  handle = () => null,
  getGate = null,
  // What Copy ID's clipboard write does with its text (it may throw), and the page's #hash.
  clipboard = () => {},
  hash = '',
  // A link another page stored before it navigated here.
  stored = null,
} = {}) {
  const document = createDocument(html);
  const sent = [];
  const opened = [];
  const ticks = [];
  const profile = {
    WorkerId: 'worker-1',
    Enabled: true,
    CanChange: true,
    RowVersion: '7',
    SharePointHosts: ['contoso.sharepoint.com'],
    ProcessRecordUpdates: false,
    Tables: ['account'],
    Registration: {
      Readiness: [
        { Scope: 'account', Status: 'Ready' },
        { Scope: 'team', Status: 'Ready' },
      ],
      Pending: 0,
      Error: null,
    },
    ...runtime,
  };
  const counts = {
    TemplateRuns: 0,
    NotCaptured: 0,
    BlockedRecords: 0,
    WaitingRecords: 0,
    BlockedJobs: 0,
    RetryingJobs: 0,
    Capped: [],
    CountedUtc: '2026-10-06T17:40:12Z',
    ...summary,
  };
  const xrm = {
    Navigation: {
      // Recorded like an API call: ['navigate', { ...page, data }].
      navigateTo: async (page) =>
        sent.push([
          'navigate',
          {
            ...page,
            data: new URLSearchParams(page.webresourceName.split('?')[1] || '').get('data'),
          },
        ]),
      // A copy: objects made inside the vm context have its Object prototype, which deepEqual refuses.
      openForm: async (options) => opened.push({ ...options }),
    },
    Utility: {
      getGlobalContext: () => ({
        getClientUrl: () => 'https://example.test',
        userSettings: { userId: '{11111111-1111-1111-1111-111111111111}' },
        organizationSettings: { bapEnvironmentId: 'env-1' },
      }),
      lookupObjects: async () => [
        {
          id: '{00000000-0000-0000-0000-0000000000A1}',
          name: 'Contoso Ltd',
          entityType: 'account',
        },
      ],
      getEntityMetadata: async (table) => ({
        DisplayName: table[0].toUpperCase() + table.slice(1),
      }),
    },
    WebApi: {
      retrieveMultipleRecords: async (table) => {
        if (rows[table] instanceof Error) throw rows[table];
        return {
          entities: rows[table]
            ? rows[table]
            : table === 'systemuser'
              ? workers || [{ systemuserid: 'worker-1', fullname: 'Documents worker' }]
              : table === 'asx_runtimetable'
                ? [{ asx_logicalname: 'account' }]
                : table === 'asx_template'
                  ? [
                      {
                        asx_templateid: '11111111-2222-3333-4444-555555555555',
                        asx_name: 'Account documents',
                        asx_table: 'account',
                        _asx_publishedrevisionid_value: 'rev-1',
                      },
                    ]
                  : table === 'connectionreference'
                    ? [
                        { connectionreferencedisplayname: 'Documents HTTP', connectionid: 'c-1' },
                        {
                          connectionreferencedisplayname: 'Documents Dataverse',
                          connectionid: null,
                        },
                      ]
                    : table === 'asx_library'
                      ? [{ asx_libraryid: 'lib-1' }]
                      : [],
        };
      },
      online: {
        execute: async (request) => {
          const api = request.getMetadata().operationName;
          const body = JSON.parse(request.Request);
          sent.push([api, body]);
          if (body.Command === 'Get' && getGate) await getGate;
          const custom = handle(api, body);
          if (custom instanceof Error)
            return { ok: false, json: async () => ({ error: { message: custom.message } }) };
          let result = custom;
          if (!result && api === 'asx_RuntimeAdmin')
            result =
              body.Command === 'SetEnabled'
                ? { ...profile, Enabled: body.Enabled, RowVersion: '8' }
                : profile;
          if (!result && body.Command === 'Summary')
            result = { Status: 'Summary', Summary: counts };
          if (!result && body.Command === 'ListProblems')
            result = { Status: 'Page', Problems: lists[body.List] || [], Next: null };
          if (!result && body.Command === 'InspectRecord')
            result = { Status: 'Planned', Record: { OperationStates: ['Applied'] }, Notices: [] };
          return {
            ok: true,
            json: async () => ({ Result: JSON.stringify(result || { Status: 'Pending' }) }),
          };
        },
      },
    },
  };
  const session = new Map([['asxd.launched', '1']]);
  if (stored) session.set('asxd.deeplink', JSON.stringify(stored));
  const window = {
    parent: { Xrm: xrm },
    location: { search: '?data=' + tab + '-' + BUILD, hash },
    sessionStorage: {
      getItem: (k) => session.get(k) ?? null,
      setItem: (k, v) => session.set(k, v),
      removeItem: (k) => session.delete(k),
    },
  };
  const context = vm.createContext({
    window,
    document,
    Intl,
    URLSearchParams,
    URL,
    console,
    navigator: { clipboard: { writeText: async (text) => clipboard(text) } },
    crypto: require('node:crypto').webcrypto,
    // Privilege checks hold every role; the document-enabled tables are account, contact, lead.
    fetch: async (url) => ({
      ok: true,
      json: async () =>
        String(url).includes('EntityDefinitions')
          ? {
              value: ['account', 'contact', 'lead'].map((t) => ({
                LogicalName: t,
                DisplayName: { UserLocalizedLabel: { Label: t[0].toUpperCase() + t.slice(1) } },
                PrimaryNameAttribute: t === 'contact' ? 'fullname' : 'name',
              })),
            }
          : { RolePrivileges: missing.some((p) => String(url).includes(p)) ? [] : [{}] },
    }),
    setTimeout,
    clearTimeout,
    // The page's 60-second timers are collected; tick() runs them once, as if a minute passed.
    setInterval: (fn) => ticks.push(fn),
    clearInterval: () => {},
  });
  for (const name of ['shell.js', 'operations.js']) vm.runInContext(read(name), context);
  await document.fire('DOMContentLoaded');
  const $ = (id) => document.getElementById(id);
  const buttons = (node) => [...node.querySelectorAll('button')];
  const press = async (node) => {
    node.click();
    await document.settle();
  };
  const tick = async () => {
    for (const fn of ticks) document.track(fn());
    await document.settle();
  };
  return {
    document,
    $,
    sent,
    opened,
    buttons,
    press,
    tick,
    window,
    commands: () => sent.map(([, b]) => b.Command),
  };
}

(async () => {
  // Monitor -------------------------------------------------------------------------------------
  const table = (m) => m.$('problem-table');
  const trs = (m) => [...m.$('problem-rows').querySelectorAll('tr')];
  const named = (m, scope, text) => m.buttons(scope).find((b) => b.textContent === text);
  const chip = (m, start) =>
    [...m.$('monitor-filters').querySelectorAll('.filter-chip')].find((c) =>
      c.textContent.startsWith(start),
    );
  // A row's ⋯ menu items, and its one primary action (fake-dom refuses ">" and ":not(.class)").
  const menuOf = (tr) => [...tr.querySelectorAll('[role=menuitem]')].map((i) => i.textContent);
  const primaryOf = (tr) =>
    [...tr.querySelector('td.actions').querySelectorAll('button')].find(
      (b) => !b.classList.contains('menu-button') && b.getAttribute('role') !== 'menuitem',
    );
  const openTool = async (m, index) => {
    await m.press(m.$('monitor-tools'));
    await m.press(m.$('monitor-tools-list').querySelectorAll('[role=menuitem]')[index]);
  };
  {
    // Filter chips with counts from Summary; zero chips are not buttons; All sums the five lists.
    const m = await boot({
      summary: {
        BlockedRecords: 3,
        WaitingRecords: 6,
        BlockedJobs: 1,
        RetryingJobs: 4,
        NotCaptured: 0,
      },
    });
    const chips = [...m.$('monitor-filters').querySelectorAll('.filter-chip')];
    assert.deepEqual(
      chips.map((c) => c.textContent),
      [
        'All · 14',
        'Blocked records · 3',
        'Waiting for data · 6',
        'Blocked jobs · 1',
        'Retrying · 4',
        'Not captured · 0',
      ],
    );
    assert.equal(chips[0].getAttribute('aria-pressed'), 'true');
    assert.equal(chips[5].tagName, 'SPAN', 'A zero chip is not interactive');
    const capped = await boot({ summary: { BlockedRecords: 5000, Capped: ['BlockedRecords'] } });
    const cappedChips = capped.$('monitor-filters').querySelectorAll('.filter-chip');
    assert.equal(cappedChips[1].textContent, 'Blocked records · 5,000+');
    assert.equal(cappedChips[0].textContent, 'All · 5,000+');
  }
  {
    // One table merging the lists by Since (newest first), with type, item, problem and one primary
    // action plus a ⋯ menu holding the rest. Since mixes ISO and "/Date(…)/" values (ui.ms).
    const m = await boot({
      summary: { BlockedRecords: 1, WaitingRecords: 1, RetryingJobs: 1 },
      lists: {
        BlockedRecords: [
          row({
            Key: 'rec:1',
            SinceUtc: '/Date(' + Date.parse('2026-10-06T10:12:00Z') + ')/',
            Actions: ['Retry', 'Cancel', 'OpenRecord', 'Check'],
          }),
        ],
        WaitingRecords: [
          row({
            Key: 'wait:1',
            Status: 'Waiting',
            SinceUtc: '2026-10-06T10:20:00Z',
            Problem: 'Needs a value in Account Number.',
            Actions: ['Rerun', 'Check', 'OpenRecord'],
          }),
        ],
        RetryingJobs: [
          row({
            Key: 'folderjob:r',
            Kind: 'FolderJob',
            Record: null,
            Title: 'Folder job · Northwind Traders',
            Attempt: 3,
            NextAttemptUtc: '2026-10-06T10:45:00Z',
            SinceUtc: '2026-10-06T10:31:00Z',
            Actions: ['Retry', 'Cancel'],
          }),
        ],
      },
    });
    const rows = trs(m);
    assert.deepEqual(
      rows.map((r) => r.querySelector('.type').textContent),
      ['Retrying', 'Waiting for data', 'Blocked record'],
    );
    assert.deepEqual(
      rows.map((r) => r.querySelector('.type').dataset.tone),
      ['muted', 'warning', 'danger'],
    );
    assert.match(rows[0].querySelector('.sub').textContent, /^Attempt 3 · next /);
    assert.equal(rows[1].querySelector('.sub').textContent, 'Account · Account documents');
    assert.deepEqual(
      rows.map((r) => primaryOf(r)?.textContent),
      ['Retry now', 'Check', 'Retry'],
    );
    assert.deepEqual(menuOf(rows[2]), ['Cancel job', 'Check', 'Copy ID']);
    await m.press(table(m).querySelector('th[aria-sort] button'));
    assert.equal(table(m).querySelector('th[aria-sort]').getAttribute('aria-sort'), 'ascending');
    assert.equal(trs(m)[0].querySelector('.type').textContent, 'Blocked record');
    // A retrying job without a next attempt time shows only its attempt.
    const unscheduled = await boot({
      summary: { RetryingJobs: 1 },
      lists: {
        RetryingJobs: [
          row({
            Key: 'folderjob:u',
            Kind: 'FolderJob',
            Record: null,
            Title: 'Folder job · Contoso',
            Attempt: 2,
            NextAttemptUtc: null,
            SinceUtc: '2026-10-06T10:31:00Z',
            Actions: ['Retry', 'Cancel'],
          }),
        ],
      },
    });
    assert.equal(trs(unscheduled)[0].querySelector('.sub').textContent, 'Attempt 2');
    // Phone widths lay rows out as cards, so the table says what it is in every browser.
    assert.equal(table(m).getAttribute('role'), 'table');
    assert.equal(table(m).querySelector('thead').getAttribute('role'), 'rowgroup');
    assert.equal(table(m).querySelector('thead tr').getAttribute('role'), 'row');
    assert.deepEqual(
      [...new Set([...table(m).querySelectorAll('th')].map((th) => th.getAttribute('role')))],
      ['columnheader'],
    );
    assert.equal(m.$('problem-rows').getAttribute('role'), 'rowgroup');
    assert.deepEqual([...new Set(trs(m).map((tr) => tr.getAttribute('role')))], ['row']);
    assert.deepEqual(
      [
        ...new Set(
          [...m.$('problem-rows').querySelectorAll('td')].map((td) => td.getAttribute('role')),
        ),
      ],
      ['cell'],
    );
  }
  {
    // A filter shows one list; Not captured keeps its checkboxes and Re-run selected (n).
    const m = await boot({
      summary: { NotCaptured: 2, BlockedRecords: 1 },
      lists: {
        NotCaptured: [
          row({ Key: 'cap:1', Kind: 'CaptureJob', Actions: ['Rerun', 'Dismiss'] }),
          row({ Key: 'cap:2', Kind: 'CaptureJob', Actions: ['Rerun', 'Dismiss'] }),
        ],
        BlockedRecords: [row({ Key: 'rec:1' })],
      },
    });
    assert.equal(
      m.$('problem-rows').querySelectorAll('input[type=checkbox]').length,
      0,
      'No boxes under All',
    );
    assert.equal(m.$('rerun-selected-row').hidden, true);
    await m.press(chip(m, 'Not captured'));
    assert.equal(trs(m).length, 2);
    assert.equal(table(m).querySelector('.col-select').getAttribute('role'), 'columnheader');
    const boxes = m.$('problem-rows').querySelectorAll('input[type=checkbox]');
    assert.equal(boxes.length, 2);
    boxes[0].checked = true;
    boxes[0].onchange();
    assert.equal(m.$('rerun-selected-row').hidden, false);
    assert.equal(
      m.$('rerun-selected-row').querySelector('button').textContent,
      'Re-run selected (1)',
    );
  }
  {
    // Re-runs strip: one card per active run with progress and Pause / Cancel re-run.
    const m = await boot({
      summary: { TemplateRuns: 1 },
      lists: {
        TemplateRuns: [
          row({
            Key: 'templaterun:1',
            Kind: 'TemplateRun',
            Title: 'Account onboarding',
            Record: null,
            Run: {
              TemplateName: 'Account onboarding',
              TableLabel: 'Account',
              Version: 3,
              State: 'Running',
              Planned: 812,
              Total: 1284,
              TotalEstimated: true,
              StartedUtc: '2026-10-06T09:58:00Z',
              StartedBy: 'Dana Reyes',
              EstimatedFinishUtc: '2026-10-06T11:30:00Z',
            },
            Actions: ['Pause', 'CancelRun'],
          }),
        ],
      },
    });
    assert.equal(m.$('runs-strip').hidden, false);
    const card = m.$('runs-strip').querySelector('.run-card');
    assert.equal(card.querySelector('h2').textContent, 'Re-run of Account onboarding v3');
    assert.match(card.querySelector('.sub').visibleText, /^Started .+ by Dana Reyes$/);
    assert.match(
      card.querySelector('.progress-text').visibleText,
      /^812 of about 1,284 · ends around /,
    );
    assert.deepEqual(
      [...card.querySelectorAll('button')].map((b) => b.textContent),
      ['Pause', 'Cancel re-run'],
    );
    assert.equal(trs(m).length, 1, 'A re-run is not a problem row');
    const none = await boot();
    assert.equal(none.$('runs-strip').hidden, true);
  }
  {
    // Tools ▾ opens Check a record and Look up an operation as side panels; Escape returns focus.
    const m = await boot();
    assert.equal(m.$('monitor-tools').getAttribute('aria-haspopup'), 'menu');
    assert.deepEqual(
      [...m.$('monitor-tools-list').querySelectorAll('[role=menuitem]')].map((i) => i.textContent),
      ['Check a record…', 'Look up an operation…'],
    );
    await openTool(m, 0);
    assert.equal(m.$('check-panel').hidden, false);
    assert.equal(m.$('check-panel').getAttribute('role'), 'dialog');
    assert.equal(m.$('check-panel').getAttribute('aria-labelledby'), 'check-title');
    assert.equal(m.document.activeElement, m.$('check-title'));
    m.$('check-title').key('Escape');
    assert.equal(m.$('check-panel').hidden, true);
    assert.equal(m.document.activeElement, m.$('monitor-tools'));
    // Close does the same through the panel handle.
    await openTool(m, 1);
    assert.equal(m.$('lookup-panel').hidden, false);
    assert.equal(m.$('lookup-title').textContent, 'Look up an operation');
    await m.press(m.$('lookup-close'));
    assert.equal(m.$('lookup-panel').hidden, true);
    assert.equal(m.document.activeElement, m.$('monitor-tools'));
    // Opening one panel closes the other.
    await openTool(m, 0);
    await openTool(m, 1);
    assert.equal(m.$('check-panel').hidden, true);
    assert.equal(m.$('lookup-panel').hidden, false);
  }
  {
    // The setup checklist still replaces the content while setup is incomplete.
    const m = await boot({ runtime: { WorkerId: '00000000-0000-0000-0000-000000000000' } });
    assert.equal(m.$('setup-checklist').hidden, false);
    for (const id of ['runs-strip', 'monitor-filters', 'rerun-selected-row', 'problem-table-card'])
      assert.equal(m.$(id).hidden, true, id);
  }
  {
    // The 60-second watch does not redraw the table under an open confirmation or menu.
    const m = await boot({
      summary: { BlockedJobs: 1 },
      lists: { BlockedJobs: [row({ Actions: ['Retry', 'Cancel'] })] },
    });
    const tr = trs(m)[0];
    await m.press(tr.querySelector('.menu-button'));
    const item = m.document.activeElement;
    await m.tick();
    assert.equal(item.isConnected, true, 'An open menu survives the tick');
    assert.equal(m.document.activeElement, item);
    await m.press(
      [...tr.querySelectorAll('[role=menuitem]')].find((i) => i.textContent === 'Cancel job'),
    );
    const confirm = table(m).querySelector('.confirm');
    assert.ok(confirm);
    assert.equal(confirm.closest('.confirm-row').getAttribute('role'), 'row');
    assert.equal(confirm.closest('td').getAttribute('role'), 'cell');
    await m.tick();
    assert.equal(table(m).querySelector('.confirm'), confirm, 'Still the same confirmation');
  }
  {
    // A tick's announcement does not replace an error the admin has not read: it follows it.
    const runRow = (State) =>
      row({
        Key: 'templaterun:1',
        Kind: 'TemplateRun',
        Title: 'Account onboarding',
        Record: null,
        Run: {
          TemplateName: 'Account onboarding',
          TableLabel: 'Account',
          Version: 3,
          State,
          Planned: 10,
          Total: 10,
          TotalEstimated: false,
          StartedUtc: '2026-10-06T09:58:00Z',
          StartedBy: 'Dana Reyes',
          EstimatedFinishUtc: null,
        },
        Actions: [],
      });
    const lists = { TemplateRuns: [runRow('Running')], BlockedJobs: [row()] };
    const m = await boot({
      summary: { TemplateRuns: 1, BlockedJobs: 1 },
      lists,
      handle: (api, b) =>
        b.Command === 'Retry' ? new Error('The job changed. Refresh it.') : null,
    });
    await m.press(primaryOf(trs(m)[0]));
    assert.equal(m.$('fb-monitor').textContent, 'The job changed. Refresh it.');
    lists.TemplateRuns = [runRow('Done')];
    await m.tick();
    assert.equal(
      m.$('fb-monitor').textContent,
      'The job changed. Refresh it. Re-run of Account onboarding: Done.',
    );
    assert.equal(m.$('fb-monitor').getAttribute('role'), 'alert', 'The error stays an alert');
    assert.equal(m.$('runs-strip').hidden, true, 'A finished re-run leaves the strip');
  }
  {
    // Opening Monitor loads Summary and every list without a click; the table lists the rows.
    const m = await boot({
      summary: { BlockedJobs: 1, NotCaptured: 0 },
      lists: { BlockedJobs: [row()] },
    });
    assert.deepEqual(
      m.sent
        .filter(([api]) => api === 'asx_ManageWork')
        .map(([, b]) => b.List || b.Command)
        .sort(),
      [
        'BlockedJobs',
        'BlockedRecords',
        'NotCaptured',
        'RetryingJobs',
        'Summary',
        'TemplateRuns',
        'WaitingRecords',
      ],
    );
    assert.equal(trs(m).length, 1);
    assert.equal(m.$('problem-table-card').hasAttribute('aria-busy'), false);
    // Nothing to show: All says nothing needs attention, with no actions.
    const empty = await boot();
    assert.equal(empty.$('problem-rows').visibleText, 'Nothing needs attention.');
    assert.equal(empty.$('problem-rows').querySelectorAll('button').length, 0);
    assert.equal(empty.$('problem-more').hidden, true);
    // A chip shows its list only and keeps focus on itself.
    const both = await boot({
      summary: { BlockedJobs: 1, RetryingJobs: 1 },
      lists: {
        BlockedJobs: [row()],
        RetryingJobs: [
          row({
            Key: 'folderjob:r',
            Attempt: 2,
            NextAttemptUtc: '2026-10-06T17:45:00Z',
            Actions: ['Retry', 'Cancel'],
          }),
        ],
      },
    });
    assert.equal(trs(both).length, 2);
    const retrying = chip(both, 'Retrying');
    retrying.focus();
    await both.press(retrying);
    assert.deepEqual(
      trs(both).map((r) => r.querySelector('.type').textContent),
      ['Retrying'],
    );
    assert.equal(both.document.activeElement.dataset.focusKey, 'filter:RetryingJobs');
    assert.equal(both.document.activeElement.getAttribute('aria-pressed'), 'true');
    assert.equal(chip(both, 'All').getAttribute('aria-pressed'), 'false');
    // Summary counts a blocked job the first page did not bring: the list's own text says so.
    const stale = await boot({ summary: { BlockedJobs: 1 } });
    await stale.press(chip(stale, 'Blocked jobs'));
    assert.equal(stale.$('problem-rows').visibleText, 'No blocked jobs.');
  }
  {
    // When the chosen list's count drops to 0, the chips fall back to All and focus follows.
    let blocked = 1;
    const m = await boot({
      lists: { BlockedJobs: [row()] },
      handle: (api, b) =>
        b.Command === 'Summary'
          ? { Status: 'Summary', Summary: { BlockedJobs: blocked, RetryingJobs: 1, Capped: [] } }
          : null,
    });
    const blockedChip = chip(m, 'Blocked jobs');
    blockedChip.focus();
    await m.press(blockedChip);
    blocked = 0;
    await m.tick();
    assert.equal(chip(m, 'All').getAttribute('aria-pressed'), 'true');
    assert.equal(chip(m, 'Blocked jobs').tagName, 'SPAN');
    assert.equal(m.document.activeElement.dataset.focusKey, 'filter:all');
  }
  {
    // Rows show names, local times and one primary action; the rest and Copy ID are in the ⋯
    // menu. No key or ID shows; Copy ID copies the operation's key.
    const copied = [];
    const m = await boot({
      summary: { BlockedJobs: 1 },
      lists: { BlockedJobs: [row()] },
      clipboard: (text) => copied.push(text),
    });
    const outside = table(m).visibleText;
    assert.match(outside, /Contoso Ltd · Account documents/);
    assert.match(outside, /The folder path is too long/);
    assert.doesNotMatch(outside, /folderjob:abc/);
    assert.doesNotMatch(outside, GUID);
    assert.doesNotMatch(outside, /2026-10-06T/);
    assert.equal(table(m).querySelector('time').getAttribute('datetime'), '2026-10-06T17:38:07Z');
    const [tr] = trs(m);
    assert.equal(tr.querySelector('.type').textContent, 'Blocked job');
    assert.equal(tr.querySelector('.type').dataset.tone, 'danger');
    assert.equal(tr.querySelector('.sub').textContent, 'Folder job');
    assert.equal(primaryOf(tr).textContent, 'Retry');
    assert.deepEqual(menuOf(tr), ['Cancel job', 'Open record', 'Check', 'Copy ID']);
    assert.equal(
      tr.querySelector('.menu-button').getAttribute('aria-label'),
      'More actions for Contoso Ltd · Account documents',
    );
    assert.equal(tr.querySelector('.menu').hidden, true);
    // Copy ID copies the key and says so in the Monitor line.
    const copy = named(m, tr, 'Copy ID');
    assert.equal(copy.getAttribute('aria-label'), 'Copy ID for Contoso Ltd · Account documents');
    await m.press(copy);
    assert.deepEqual(copied, ['folderjob:abc']);
    assert.equal(m.$('fb-monitor').textContent, 'Copied');
    // A job's item is not a link, so Open record is in its menu; it opens the form.
    await m.press(named(m, tr, 'Open record'));
    assert.deepEqual(m.opened.at(-1), {
      entityName: 'account',
      entityId: '00000000-0000-0000-0000-0000000000a1',
      openInNewWindow: true,
    });
    // Retry sends the operation's Retry and leaves the row with its new status.
    const retry = primaryOf(tr);
    await m.press(retry);
    assert.deepEqual(m.sent.at(-1), ['asx_ManageWork', { Command: 'Retry', Key: 'folderjob:abc' }]);
    assert.match(table(m).visibleText, /Queued again/);
    assert.equal(retry.disabled, true);
    assert.equal(named(m, tr, 'Cancel job').disabled, true, 'The row actions disable');
    assert.equal(named(m, tr, 'Open record').disabled, false);
    assert.equal(
      m.$('fb-monitor').textContent,
      'Queued again. It blocks again if the cause remains.',
    );
    // Cancel job from the menu confirms in the page first, with no key in its text; Keep
    // returns focus to the row's ⋯.
    const fresh = await boot({ summary: { BlockedJobs: 1 }, lists: { BlockedJobs: [row()] } });
    const ft = trs(fresh)[0];
    await fresh.press(ft.querySelector('.menu-button'));
    assert.equal(ft.querySelector('.menu').hidden, false);
    assert.equal(fresh.document.activeElement.textContent, 'Cancel job', 'The first item');
    await fresh.press(fresh.document.activeElement);
    assert.equal(ft.querySelector('.menu').hidden, true);
    assert.equal(
      fresh.document.activeElement.textContent,
      'Cancel the folder job for Contoso Ltd · Account documents? Nothing in SharePoint is undone or deleted.',
    );
    await fresh.press(named(fresh, table(fresh), 'Keep job'));
    assert.equal(fresh.commands().includes('Cancel'), false);
    assert.equal(fresh.document.activeElement, ft.querySelector('.menu-button'));
    await fresh.press(named(fresh, ft, 'Cancel job'));
    await fresh.press(
      [...table(fresh).querySelectorAll('.confirm button')].find(
        (b) => b.textContent === 'Cancel job',
      ),
    );
    assert.deepEqual(fresh.sent.at(-1), [
      'asx_ManageWork',
      { Command: 'Cancel', Key: 'folderjob:abc' },
    ]);
    assert.equal(table(fresh).querySelectorAll('.confirm').length, 0, 'The confirmation closes');
    // A clipboard that refuses says so.
    const refused = await boot({
      summary: { BlockedJobs: 1 },
      lists: { BlockedJobs: [row()] },
      clipboard: () => {
        throw new Error('Not allowed.');
      },
    });
    await refused.press(named(refused, trs(refused)[0], 'Copy ID'));
    assert.equal(refused.$('fb-monitor').textContent, 'Copy failed');
    assert.equal(refused.$('fb-monitor').getAttribute('role'), 'alert');
  }
  {
    // A record the caller cannot read says so, with no ID.
    const m = await boot({
      summary: { BlockedRecords: 1 },
      lists: {
        BlockedRecords: [
          row({
            Kind: 'RecordPlan',
            KindLabel: 'Record',
            Key: 'request:1',
            Record: {
              Table: 'account',
              TableLabel: 'Account',
              Id: '00000000-0000-0000-0000-0000000000a1',
              Name: null,
            },
            Actions: ['Retry', 'Check'],
          }),
        ],
      },
    });
    assert.match(table(m).visibleText, /Record not available to you/);
    await m.press(primaryOf(trs(m)[0]));
    assert.deepEqual(m.sent.at(-1)[1], { Command: 'RetryOutbox', Key: 'request:1' });
  }
  {
    // Retrying automatically: the row action is "Retry now" and sends the job's Retry.
    const waiting = row({
      Key: 'folderjob:wait',
      Status: 'RetryWait',
      Problem: 'SharePoint is busy or limiting requests (HTTP 429).',
      Attempt: 3,
      NextAttemptUtc: '2026-10-06T17:45:00Z',
      Actions: ['Retry', 'Cancel'],
    });
    const m = await boot({ summary: { RetryingJobs: 1 }, lists: { RetryingJobs: [waiting] } });
    const [tr] = trs(m);
    assert.deepEqual(menuOf(tr), ['Cancel job', 'Copy ID']);
    const now = primaryOf(tr);
    assert.equal(now.textContent, 'Retry now');
    assert.equal(now.getAttribute('aria-label'), 'Retry now for Contoso Ltd · Account documents');
    await m.press(now);
    assert.deepEqual(m.sent.at(-1), [
      'asx_ManageWork',
      { Command: 'Retry', Key: 'folderjob:wait' },
    ]);
  }
  {
    // Not captured: Re-run selected for checked rows, kept across a redraw; Dismiss removes the
    // row and keeps focus in the table.
    const capture = (id, name) =>
      row({
        Key: id,
        Kind: 'CaptureJob',
        KindLabel: 'Missed change',
        Title: name + ' · Account',
        Record: { Table: 'account', TableLabel: 'Account', Id: id, Name: name },
        Problem: 'Runtime identity is incomplete.',
        Status: 'Failed',
        Actions: ['Rerun', 'Dismiss', 'OpenRecord'],
      });
    const m = await boot({
      summary: { NotCaptured: 2 },
      lists: {
        NotCaptured: [
          capture('aaaaaaaa-0000-0000-0000-000000000001', 'Alpha'),
          capture('aaaaaaaa-0000-0000-0000-000000000002', 'Beta'),
        ],
      },
      handle: (api, b) =>
        b.Command === 'RerunRecord'
          ? { Status: 'Queued', Keys: ['request:x'], Notices: [] }
          : b.Command === 'DismissCaptureJob'
            ? { Status: 'Dismissed' }
            : null,
    });
    await m.press(chip(m, 'Not captured'));
    const boxes = m.$('problem-rows').querySelectorAll('input[type=checkbox]');
    assert.equal(boxes.length, 2);
    // Each checkbox is named by its record cell.
    assert.equal(
      m.document.getElementById(boxes[0].getAttribute('aria-labelledby')).textContent,
      'Alpha',
    );
    assert.equal(table(m).querySelector('thead').visibleText.includes('Select'), true);
    const selected = m.$('rerun-selected-row').querySelector('button');
    assert.equal(selected.textContent, 'Re-run selected (0)');
    boxes[1].checked = true;
    boxes[1].onchange();
    assert.equal(selected.textContent, 'Re-run selected (1)');
    await m.tick();
    const again = m.$('problem-rows').querySelectorAll('input[type=checkbox]');
    assert.equal(again[1].checked, true, 'A redraw keeps the selection');
    assert.equal(selected.textContent, 'Re-run selected (1)');
    await m.press(selected);
    const rerun = m.sent.filter(([, b]) => b.Command === 'RerunRecord');
    assert.equal(rerun.length, 1);
    assert.equal(rerun[0][1].RecordId, 'aaaaaaaa-0000-0000-0000-000000000002');
    assert.equal(rerun[0][1].Table, 'account');
    assert.ok(rerun[0][1].RequestId);
    const first = trs(m)[0];
    assert.equal(primaryOf(first).textContent, 'Re-run');
    assert.deepEqual(menuOf(first), ['Dismiss', 'Copy ID']);
    await m.press(first.querySelector('.menu-button'));
    assert.equal(m.document.activeElement.textContent, 'Dismiss');
    await m.press(m.document.activeElement);
    assert.deepEqual(m.sent.at(-1)[1], {
      Command: 'DismissCaptureJob',
      JobId: 'aaaaaaaa-0000-0000-0000-000000000001',
    });
    assert.equal(trs(m).length, 1);
    assert.equal(m.document.activeElement.textContent, 'Beta', 'Focus moves to the next row');
  }
  {
    // Template re-runs in the strip: progress text and value text, states, Pause, Resume, Retry and
    // Cancel re-run.
    const run = (state, extra = {}) =>
      row({
        Key: 'templaterun:9c4e',
        Kind: 'TemplateRun',
        KindLabel: 'Template re-run',
        Title: 'TEST Account Documents · Account',
        Record: null,
        Status: state === 'Paused' ? 'Paused' : state === 'Blocked' ? 'Blocked' : 'Pending',
        Problem: state === 'Blocked' ? 'The re-run stopped while planning a page of records.' : '',
        Run: {
          TemplateName: 'TEST Account Documents',
          TableLabel: 'Account',
          Version: 2,
          State: state,
          Planned: 12500,
          Queued: 25,
          Total: 40000,
          TotalEstimated: false,
          StartedUtc: '2026-10-06T18:02:00Z',
          StartedBy: 'Alex Rivera',
          EstimatedFinishUtc: '2026-10-07T20:00:00Z',
          ...extra,
        },
        Actions: {
          Running: ['Pause', 'CancelRun'],
          Paused: ['Resume', 'CancelRun'],
          Blocked: ['Retry', 'CancelRun'],
          Done: [],
        }[state],
      });
    const m = await boot({
      summary: { TemplateRuns: 1 },
      lists: { TemplateRuns: [run('Running')] },
      handle: (api, b) =>
        b.Command === 'PauseTemplateRun'
          ? { Status: 'Paused', Key: b.Key, Run: run('Paused').Run }
          : null,
    });
    const strip = m.$('runs-strip');
    const bar = strip.querySelector('progress');
    assert.equal(bar.getAttribute('aria-valuetext'), '12,500 of 40,000, running');
    assert.equal(bar.getAttribute('aria-label'), 'Progress of TEST Account Documents');
    assert.match(strip.visibleText, /12,500 of 40,000 · ends around/);
    assert.match(strip.visibleText, /Re-run of TEST Account Documents v2/);
    assert.match(strip.visibleText, /Alex Rivera/);
    await m.press(named(m, strip, 'Pause'));
    assert.deepEqual(m.sent.at(-1)[1], { Command: 'PauseTemplateRun', Key: 'templaterun:9c4e' });
    assert.match(strip.visibleText, /12,500 of 40,000 · Paused/);
    assert.equal(m.$('fb-monitor').textContent, 'Re-run paused.');
    // Cancel re-run confirms with its text.
    const c = await boot({
      summary: { TemplateRuns: 1 },
      lists: { TemplateRuns: [run('Blocked')] },
    });
    const cs = c.$('runs-strip');
    assert.match(cs.visibleText, /Needs attention/);
    assert.match(cs.visibleText, /The re-run stopped while planning a page of records\./);
    await c.press(named(c, cs, 'Cancel re-run'));
    assert.equal(
      c.document.activeElement.textContent,
      'Cancel the re-run of TEST Account Documents? Records not yet planned are skipped. Folder work already queued for planned records continues, and nothing in SharePoint is undone.',
    );
    await c.press(named(c, cs, 'Keep re-run'));
    await c.press(named(c, cs, 'Retry'));
    assert.deepEqual(c.sent.at(-1)[1], { Command: 'RetryOutbox', Key: 'templaterun:9c4e' });
    assert.match(cs.visibleText, /Queued again/);
    // Estimated totals (the daily row-count snapshot) and finished runs.
    const capped = await boot({
      summary: { TemplateRuns: 1 },
      lists: {
        TemplateRuns: [
          run('Running', { Total: 40000, TotalEstimated: true, EstimatedFinishUtc: null }),
        ],
      },
    });
    assert.match(capped.$('runs-strip').visibleText, /12,500 of about 40,000 · Estimating…/);
    const done = await boot({ lists: { TemplateRuns: [run('Done', { Planned: 40000 })] } });
    assert.equal(done.$('runs-strip').hidden, true, 'A finished re-run is not shown');
    // A link to a run focuses its first action.
    const linked = await boot({
      summary: { TemplateRuns: 1 },
      lists: { TemplateRuns: [run('Running')] },
      hash: '#monitor?run=templaterun:9c4e',
    });
    assert.equal(linked.document.activeElement.textContent, 'Pause');
  }
  {
    // The 60-second re-read while Monitor is open: progress updates silently; a re-run reaching
    // Done, and a setup leaving "Checking SharePoint…", are each announced once.
    const runRow = (State, Planned) =>
      row({
        Key: 'templaterun:9c4e',
        Kind: 'TemplateRun',
        KindLabel: 'Template re-run',
        Title: 'TEST Account Documents · Account',
        Record: null,
        Status: 'Pending',
        Problem: '',
        Run: {
          TemplateName: 'TEST Account Documents',
          TableLabel: 'Account',
          Version: 2,
          State,
          Planned,
          Queued: 0,
          Total: 40000,
          TotalEstimated: false,
          StartedUtc: '2026-10-06T18:02:00Z',
          StartedBy: 'Alex Rivera',
          EstimatedFinishUtc: null,
        },
        Actions: [],
      });
    const setupRow = (recovery, actions) =>
      row({
        Key: 'librarycreate:x',
        RowVersion: 'rv-3',
        Kind: 'LibrarySetup',
        KindLabel: 'Library setup',
        Title: 'Project documents',
        Record: null,
        TemplateId: null,
        Status: recovery.State === 'Checking' ? 'Reconciling' : 'RecoveryRequired',
        Problem: '',
        Recovery: recovery,
        Actions: actions,
      });
    const lists = {
      TemplateRuns: [runRow('Running', 100)],
      BlockedJobs: [
        setupRow({ State: 'Checking', Candidates: [], Choices: ['Cancel'] }, ['Cancel']),
      ],
    };
    const m = await boot({ summary: { TemplateRuns: 1, BlockedJobs: 1 }, lists });
    lists.TemplateRuns = [runRow('Running', 200)];
    await m.tick();
    assert.match(m.$('runs-strip').visibleText, /200 of 40,000/);
    assert.equal(
      m.$('fb-monitor').textContent,
      '',
      'Progress alone and a setup still checking are not announced',
    );
    lists.TemplateRuns = [runRow('Done', 40000)];
    lists.BlockedJobs = [
      setupRow(
        { State: 'NotFound', Candidates: [], Choices: ['CreateAgain', 'CheckAgain', 'Cancel'] },
        ['CreateAgain', 'CheckAgain', 'Cancel'],
      ),
    ];
    await m.tick();
    // Both announcements of the tick share the one Monitor line.
    assert.equal(
      m.$('fb-monitor').textContent,
      'Re-run of TEST Account Documents: Done. ' +
        "SharePoint has no library named Project documents. The creation didn't happen.",
    );
    assert.equal(m.$('fb-monitor').getAttribute('role'), 'status');
    assert.equal(
      primaryOf(trs(m)[0]).textContent,
      'Create it again',
      'The finding brings its choices',
    );
    // Once: an unchanged tick announces nothing again.
    m.$('fb-monitor').textContent = '';
    await m.tick();
    assert.equal(m.$('fb-monitor').textContent, '');
  }
  {
    // A blocked library setup: Retry goes through asx_ManageWork and needs the Operator role;
    // Cancel setup goes through the catalog API and needs Documents Security Administrator.
    const m = await boot({
      summary: { BlockedJobs: 1 },
      missing: ['prvCreateasx_site'],
      lists: {
        BlockedJobs: [
          row({
            Key: 'librarycreate:y',
            Kind: 'LibrarySetup',
            KindLabel: 'Library setup',
            Title: 'Contracts',
            Record: null,
            TemplateId: null,
            Status: 'Blocked',
            Actions: ['Retry', 'Cancel'],
          }),
        ],
      },
    });
    const tr = trs(m)[0];
    assert.equal(primaryOf(tr).textContent, 'Retry');
    assert.equal(primaryOf(tr).hasAttribute('aria-disabled'), false, 'An Operator can retry');
    const cancel = [...tr.querySelectorAll('[role=menuitem]')].find(
      (i) => i.textContent === 'Cancel setup',
    );
    assert.equal(cancel.getAttribute('aria-disabled'), 'true');
  }
  {
    // Recovery rows: each finding shows its sentence and its choices; no input asks for evidence.
    const setup = (recovery, actions) =>
      row({
        Key: 'librarycreate:x',
        RowVersion: 'rv-3',
        Kind: 'LibrarySetup',
        KindLabel: 'Library setup',
        Title: 'Project documents',
        Record: null,
        TemplateId: null,
        Status: recovery.State === 'Checking' ? 'Reconciling' : 'RecoveryRequired',
        Problem: '',
        Recovery: recovery,
        Actions: actions,
      });
    const candidate = {
      ListId: '5e0c0000-0000-0000-0000-000000000001',
      Title: 'Project documents',
      Url: '/sites/x/Project documents',
      CreatedUtc: '2026-10-06T17:41:02Z',
      IsLibrary: true,
      TitleMatches: true,
      UrlMatches: true,
      CreatedAfterRequest: true,
      CatalogEntry: 'None',
    };
    const checking = await boot({
      summary: { BlockedJobs: 1 },
      lists: {
        BlockedJobs: [
          setup({ State: 'Checking', Candidates: [], Choices: ['Cancel'] }, ['Cancel']),
        ],
      },
    });
    assert.match(table(checking).visibleText, /Checking SharePoint for Project documents…/);
    assert.equal(primaryOf(trs(checking)[0]), undefined, 'Cancel is never the primary action');
    assert.deepEqual(menuOf(trs(checking)[0]), ['Cancel setup', 'Copy ID']);
    const found = await boot({
      summary: { BlockedJobs: 1 },
      lists: {
        BlockedJobs: [
          setup(
            {
              State: 'Found',
              Candidates: [candidate],
              Choices: ['UseLibrary', 'CheckAgain', 'Cancel'],
            },
            ['UseLibrary', 'CheckAgain', 'Cancel'],
          ),
        ],
      },
    });
    assert.match(
      table(found).visibleText,
      /SharePoint has a library Project documents at \/sites\/x\/Project documents, created .*\. It matches this request\./,
    );
    const use = primaryOf(trs(found)[0]);
    assert.equal(use.textContent, 'Use the library');
    assert.equal(use.className, 'primary');
    await found.press(use);
    assert.deepEqual(found.sent.at(-1), [
      'asx_CatalogAdmin',
      {
        Command: 'ResolveSetup',
        Key: 'librarycreate:x',
        Choice: 'UseLibrary',
        ListId: candidate.ListId,
        RowVersion: 'rv-3',
      },
    ]);
    assert.equal(found.$('fb-monitor').textContent, 'Using the existing library. Setup continues.');
    const notFound = await boot({
      summary: { BlockedJobs: 1 },
      lists: {
        BlockedJobs: [
          setup(
            { State: 'NotFound', Candidates: [], Choices: ['CreateAgain', 'CheckAgain', 'Cancel'] },
            ['CreateAgain', 'CheckAgain', 'Cancel'],
          ),
        ],
      },
    });
    assert.match(
      table(notFound).visibleText,
      /SharePoint has no library named Project documents\. The creation didn't happen\./,
    );
    assert.deepEqual(menuOf(trs(notFound)[0]), ['Check again', 'Cancel setup', 'Copy ID']);
    await notFound.press(primaryOf(trs(notFound)[0]));
    assert.deepEqual(notFound.sent.at(-1)[1], {
      Command: 'ResolveSetup',
      Key: 'librarycreate:x',
      Choice: 'CreateAgain',
      RowVersion: 'rv-3',
    });
    assert.equal(notFound.$('fb-monitor').textContent, 'Creating the library again.');
    const recheck = await boot({
      summary: { BlockedJobs: 1 },
      lists: {
        BlockedJobs: [
          setup(
            { State: 'NotFound', Candidates: [], Choices: ['CreateAgain', 'CheckAgain', 'Cancel'] },
            ['CreateAgain', 'CheckAgain', 'Cancel'],
          ),
        ],
      },
    });
    await recheck.press(named(recheck, trs(recheck)[0], 'Check again'));
    assert.deepEqual(recheck.sent.at(-1), [
      'asx_CatalogAdmin',
      { Command: 'RecheckSetup', Key: 'librarycreate:x' },
    ]);
    const moved = { ...candidate, Url: '/sites/x/Project documents1', UrlMatches: false };
    const ambiguous = await boot({
      summary: { BlockedJobs: 1 },
      lists: {
        BlockedJobs: [
          setup(
            {
              State: 'Ambiguous',
              Reason: 'A library has this name but a different address.',
              Candidates: [moved],
              Choices: ['UseCandidate', 'CheckAgain', 'Cancel'],
            },
            ['UseCandidate', 'CheckAgain', 'Cancel'],
          ),
        ],
      },
    });
    const a = table(ambiguous);
    assert.match(a.visibleText, /A library has this name but a different address\./);
    assert.match(a.visibleText, /\/sites\/x\/Project documents1/);
    assert.equal(
      ambiguous.buttons(a).some((b) => b.textContent === 'Create it again'),
      false,
    );
    await ambiguous.press(named(ambiguous, a, 'Use this one'));
    assert.equal(
      ambiguous.document.activeElement.textContent,
      'Use Project documents at /sites/x/Project documents1 for this setup? Documents stops its permission inheritance if needed and manages its team access.',
    );
    await ambiguous.press(named(ambiguous, a, 'Use this library'));
    assert.deepEqual(ambiguous.sent.at(-1)[1], {
      Command: 'ResolveSetup',
      Key: 'librarycreate:x',
      Choice: 'UseLibrary',
      ListId: moved.ListId,
      RowVersion: 'rv-3',
    });
    for (const m of [checking, found, notFound, ambiguous])
      assert.equal(
        [...m.document.querySelectorAll('input[type=text], textarea')].filter((n) =>
          /run|token|evidence|response/i.test(n.id),
        ).length,
        0,
      );
  }
  {
    // A list that fails to load says so, with Try again; the others still render.
    let refuse = true;
    const m = await boot({
      summary: { BlockedJobs: 1 },
      lists: { BlockedJobs: [row()] },
      handle: (api, b) =>
        b.List === 'RetryingJobs' && refuse
          ? new Error('Principal user is missing prvReadasx_operation.')
          : null,
    });
    assert.equal(
      m.$('monitor-errors').querySelector('[role=alert]').textContent,
      "Couldn't load Retrying automatically: Principal user is missing prvReadasx_operation.",
    );
    assert.equal(trs(m).length, 1, 'The other lists still show');
    refuse = false;
    await m.press(named(m, m.$('monitor-errors'), 'Try again'));
    assert.equal(m.$('monitor-errors').querySelectorAll('[role=alert]').length, 0);
  }
  {
    // Refresh reloads everything and says what it found.
    const m = await boot({ summary: { TemplateRuns: 1, NotCaptured: 2, BlockedJobs: 1 } });
    const before = m.sent.length;
    await m.press(m.$('monitor-refresh'));
    assert.equal(m.sent.length - before, 7);
    assert.equal(
      m.$('fb-monitor').textContent,
      'Refreshed. 1 re-run in progress, 2 changes not captured, 1 blocked job.',
    );
  }
  {
    // Monitor shows automation as a pill with a link to Settings; the switch lives only in Settings.
    const m = await boot({ runtime: { Enabled: false } });
    assert.equal(m.$('automation-switch-monitor'), null);
    assert.equal(m.$('help-automation-monitor'), null);
    assert.equal(m.$('monitor-automation').textContent, 'Automation paused');
    assert.equal(m.$('monitor-automation').dataset.tone, 'warning');
    assert.equal(m.$('monitor-settings-link').textContent, 'Change in Settings');
    await m.press(m.$('monitor-settings-link'));
    assert.equal(
      m.sent
        .filter(([k]) => k === 'navigate')
        .at(-1)[1]
        .data.split('-')[0],
      'settings',
    );
    const running = await boot();
    assert.equal(running.$('monitor-automation').textContent, 'Automation running');
    assert.equal(running.$('monitor-automation').dataset.tone, 'ok');
  }
  {
    // Row results report in the one Monitor feedback line; there are no per-list lines.
    const m = await boot({ summary: { BlockedJobs: 1 }, lists: { BlockedJobs: [row()] } });
    assert.deepEqual(
      [...m.document.querySelectorAll('.feedback')]
        .map((n) => n.id)
        .filter((id) => id.startsWith('fb-list-')),
      [],
    );
  }
  {
    // Setup incomplete: a checklist replaces the table, each open step linking to its page.
    const m = await boot({
      runtime: { WorkerId: '00000000-0000-0000-0000-000000000000', Enabled: false },
    });
    const checklist = m.$('setup-checklist');
    assert.equal(checklist.hidden, false);
    assert.match(checklist.visibleText, /Choose who runs automation \(Settings\)/);
    assert.equal(m.$('problem-table-card').hidden, true);
  }
  {
    // A paused automation with setup otherwise done is a normal state: the problems stay in view
    // and the header pill says it is paused.
    const m = await boot({
      runtime: { Enabled: false },
      summary: { BlockedJobs: 1 },
      lists: { BlockedJobs: [row({ Actions: ['Retry', 'Cancel'] })] },
    });
    assert.equal(m.$('setup-checklist').hidden, true);
    assert.equal(m.$('problem-table-card').hidden, false);
    assert.equal(m.$('monitor-automation').textContent, 'Automation paused');
  }
  {
    // Check a record: disabled until table, template and record are set; the result is readable.
    const m = await boot();
    const check = m.$('check-run');
    assert.equal(check.getAttribute('aria-disabled'), 'true');
    assert.equal(m.$('check-reason').textContent, 'Choose a record first');
    m.$('check-table').value = 'account';
    await m.$('check-table').onchange();
    await m.press(m.$('check-choose'));
    assert.equal(m.$('check-record-name').textContent, 'Contoso Ltd');
    assert.equal(check.hasAttribute('aria-disabled'), false);
    await m.press(check);
    assert.deepEqual(m.sent.at(-1)[1], {
      Command: 'InspectRecord',
      TemplateId: '11111111-2222-3333-4444-555555555555',
      RecordId: '00000000-0000-0000-0000-0000000000a1',
    });
    assert.equal(m.$('check-result').querySelectorAll('pre').length, 0);
    assert.match(m.$('check-result').visibleText, /Folders are created\./);
    await m.press(m.$('check-rerun'));
    assert.equal(m.sent.at(-1)[1].Command, 'Replan');
    assert.equal(m.$('fb-check').textContent, 'Re-run queued for Contoso Ltd.');
  }
  {
    // Look up an operation: recent operations page on open; Look up validates the ID; there is no
    // recovery panel.
    const m = await boot({
      lists: { RecentOperations: [row({ Status: 'Applied', Actions: [] })] },
      handle: (api, b) =>
        b.Command === 'Inspect'
          ? { Key: b.Key, Status: 'Blocked', Notices: ['RequestUrlTooLong'] }
          : null,
    });
    assert.equal(
      m.sent.some(([, b]) => b.List === 'RecentOperations'),
      false,
      'Recent operations load when the panel opens',
    );
    await openTool(m, 1);
    assert.match(m.$('recent-rows').visibleText, /Contoso Ltd · Account documents/);
    m.$('operation-id').value = 'recordplan:abc';
    await m.press(m.$('operation-lookup'));
    assert.equal(
      m.$('fb-advanced').textContent,
      'Operation IDs start with folderjob:, librarycreate:, catalogprobe: or policywork:',
    );
    m.$('operation-id').value = ' folderjob:abc ';
    await m.press(m.$('operation-lookup'));
    assert.deepEqual(m.sent.at(-1)[1], { Command: 'Inspect', Key: 'folderjob:abc' });
    assert.equal(m.$('operation-result').hidden, false);
    assert.equal(m.document.getElementById('recoveryPanel'), null);
    // A library setup with a SharePoint finding offers its recovery choices here too.
    const finding = {
      State: 'NotFound',
      Candidates: [],
      Choices: ['CreateAgain', 'CheckAgain', 'Cancel'],
    };
    const r = await boot({
      handle: (api, b) =>
        b.Command === 'Inspect'
          ? {
              Key: b.Key,
              Status: 'RecoveryRequired',
              RowVersion: 'rv-3',
              Notices: [],
              Recovery: finding,
            }
          : null,
    });
    await openTool(r, 1);
    r.$('operation-id').value = 'librarycreate:x';
    await r.press(r.$('operation-lookup'));
    assert.match(
      r.$('operation-fields').visibleText,
      /SharePoint has no library named librarycreate:x\. The creation didn't happen\./,
    );
    assert.deepEqual(
      r.buttons(r.$('operation-actions')).map((b) => b.textContent),
      ['Create it again', 'Check again', 'Cancel setup'],
    );
    await r.press(
      r.buttons(r.$('operation-actions')).find((b) => b.textContent === 'Create it again'),
    );
    assert.deepEqual(r.sent.at(-1), [
      'asx_CatalogAdmin',
      {
        Command: 'ResolveSetup',
        Key: 'librarycreate:x',
        Choice: 'CreateAgain',
        RowVersion: 'rv-3',
      },
    ]);
    assert.equal(r.$('fb-advanced').textContent, 'Creating the library again.');
    const moved = {
      ListId: '5e0c0000-0000-0000-0000-000000000009',
      Title: 'Project documents',
      Url: '/sites/x/Project documents1',
      CreatedUtc: '2026-10-06T17:41:02Z',
      IsLibrary: true,
      TitleMatches: true,
      UrlMatches: false,
      CreatedAfterRequest: true,
      CatalogEntry: 'None',
    };
    const amb = await boot({
      handle: (api, b) =>
        b.Command === 'Inspect'
          ? {
              Key: b.Key,
              Status: 'RecoveryRequired',
              RowVersion: 'rv-4',
              Notices: [],
              Recovery: {
                State: 'Ambiguous',
                Reason: 'A library has this name but a different address.',
                Candidates: [moved],
                Choices: ['UseCandidate', 'CheckAgain', 'Cancel'],
              },
            }
          : null,
    });
    await openTool(amb, 1);
    amb.$('operation-id').value = 'librarycreate:y';
    await amb.press(amb.$('operation-lookup'));
    await amb.press(
      amb.buttons(amb.$('operation-actions')).find((b) => b.textContent === 'Use this one'),
    );
    await amb.press(
      amb.buttons(amb.$('operation-actions')).find((b) => b.textContent === 'Use this library'),
    );
    assert.deepEqual(amb.sent.at(-1)[1], {
      Command: 'ResolveSetup',
      Key: 'librarycreate:y',
      Choice: 'UseLibrary',
      ListId: moved.ListId,
      RowVersion: 'rv-4',
    });
  }
  {
    // Links: an operation opens Look up an operation with its result; a record opens Check a
    // record ready to check; a row's Check opens Check a record and Escape returns to that row.
    const op = await boot({
      hash: '#monitor?operation=folderjob:abc',
      handle: (api, b) =>
        b.Command === 'Inspect' ? { Key: b.Key, Status: 'Blocked', Notices: [] } : null,
    });
    assert.equal(op.$('lookup-panel').hidden, false);
    assert.equal(op.$('operation-id').value, 'folderjob:abc');
    assert.equal(op.$('operation-result').hidden, false);
    assert.equal(op.document.activeElement, op.$('operation-result-title'));
    const rec = await boot({
      hash: '#monitor?record=account:00000000-0000-0000-0000-0000000000a1&template=11111111-2222-3333-4444-555555555555',
    });
    assert.equal(rec.$('check-panel').hidden, false);
    assert.equal(rec.$('check-table').value, 'account');
    assert.equal(rec.document.activeElement, rec.$('check-run'));
    rec.$('check-run').key('Escape');
    assert.equal(rec.$('check-panel').hidden, true);
    assert.equal(rec.document.activeElement, rec.$('monitor-tools'));
    const c = await boot({
      summary: { WaitingRecords: 1 },
      lists: {
        WaitingRecords: [
          row({ Key: 'wait:1', Status: 'Waiting', Actions: ['Check', 'OpenRecord'] }),
        ],
      },
    });
    const check = primaryOf(trs(c)[0]);
    assert.equal(check.textContent, 'Check');
    await c.press(check);
    assert.equal(c.$('check-panel').hidden, false);
    assert.equal(c.$('check-record-name').textContent, 'Contoso Ltd');
    assert.equal(c.$('check-template').value, '11111111-2222-3333-4444-555555555555');
    c.$('check-run').key('Escape');
    assert.equal(c.document.activeElement, check);
  }
  {
    // Settings loads at once; hosts are a list; Save sends them; Repair and Repair all; Stop tracking confirms.
    let pending = 2;
    const s = await boot({
      tab: 'settings',
      runtime: {
        Registration: {
          Readiness: [
            { Scope: 'account', Status: 'Missing' },
            { Scope: 'contact', Status: 'WorkerCannotRead' },
            { Scope: 'team', Status: 'Pending' },
          ],
          Pending: 2,
          Error: null,
        },
      },
      handle: (api, b) => {
        if (b.Command === 'Register' && b.Table === '') {
          pending = Math.max(0, pending - 1);
          return {
            WorkerId: 'worker-1',
            Enabled: true,
            CanChange: true,
            RowVersion: '7',
            SharePointHosts: ['contoso.sharepoint.com'],
            Registration: {
              Readiness: [
                { Scope: 'account', Status: pending ? 'Missing' : 'Ready' },
                { Scope: 'contact', Status: 'WorkerCannotRead' },
                { Scope: 'team', Status: 'Ready' },
              ],
              Pending: pending,
              Error: null,
            },
          };
        }
        return null;
      },
    });
    assert.equal(s.commands()[0], 'Get', 'The profile loads when Settings opens');
    assert.equal(s.$('runtimeWorker').value, 'worker-1');
    assert.equal(s.$('hosts-list').querySelectorAll('input').length, 1);
    await s.press(s.$('add-host'));
    const hosts = s.$('hosts-list').querySelectorAll('input');
    assert.equal(s.document.activeElement, hosts[1], 'A new host field takes focus');
    hosts[1].value = 'not a host';
    await s.press(s.$('save-settings'));
    assert.equal(hosts[1].getAttribute('aria-invalid'), 'true');
    assert.equal(
      s.document.getElementById(hosts[1].getAttribute('aria-describedby')).textContent,
      'Enter a host name like contoso.sharepoint.com',
    );
    assert.equal(s.document.activeElement, hosts[1]);
    assert.equal(s.commands().includes('Save'), false, 'Nothing is saved while a field is invalid');
    hosts[1].value = 'Fabrikam.SharePoint.com ';
    await s.press(s.$('save-settings'));
    const save = s.sent.filter(([, b]) => b.Command === 'Save').at(-1)[1];
    assert.deepEqual(save.SharePointHosts, ['contoso.sharepoint.com', 'fabrikam.sharepoint.com']);
    assert.equal(save.Enabled, true);
    assert.equal(s.$('fb-settings').textContent, 'Settings saved.');
    const rows = [...s.$('tables-rows').querySelectorAll('[data-table]')];
    assert.match(rows[0].visibleText, /Account.*Some steps missing.*Repair/);
    assert.match(rows[1].visibleText, /Run-as user can't read this table/);
    assert.ok(
      ![...rows[1].querySelectorAll('button')].some((b) => b.textContent === 'Repair'),
      'WorkerCannotRead has no Repair',
    );
    await s.press(rows[0].querySelector('button'));
    assert.deepEqual(s.sent.at(-1)[1], { Command: 'Register', Table: 'account', RowVersion: '7' });
    await s.press(s.$('repair-all'));
    assert.equal(
      s.sent.filter(([, b]) => b.Command === 'Register' && b.Table === '').length,
      2,
      'Repair all stops when nothing is pending',
    );
    // Once it ends the button shows the current state, not its progress (fix round 1).
    assert.doesNotMatch(s.$('repair-all').textContent, /Repaired/);
    assert.equal(s.$('repair-all').hidden, true);
    await s.press(s.$('stop-tracking'));
    assert.equal(
      s.document.activeElement.textContent,
      'Stop tracking changes for every table? Documents stops capturing record and team changes until you repair change tracking. Use this before uninstalling.',
    );
    await s.press(s.buttons(s.$('danger-zone')).find((b) => b.textContent === 'Keep tracking'));
    assert.equal(s.commands().includes('Unregister'), false);
    assert.match(s.$('connection-list').visibleText, /Documents HTTP.*Connected/);
    assert.match(s.$('connection-list').visibleText, /Documents Dataverse.*Not connected/);
    assert.equal(
      s.$('open-connections').getAttribute('href'),
      'https://make.powerautomate.com/environments/env-1/connections',
    );
  }
  {
    // Repair all with no progress stops and says so.
    const s = await boot({
      tab: 'settings',
      runtime: {
        Registration: {
          Readiness: [
            { Scope: 'account', Status: 'Missing' },
            { Scope: 'contact', Status: 'Missing' },
          ],
          Pending: 2,
          Error: null,
        },
      },
      handle: (api, b) =>
        b.Command === 'Register'
          ? {
              WorkerId: 'worker-1',
              Enabled: true,
              CanChange: true,
              RowVersion: '7',
              SharePointHosts: [],
              Registration: {
                Readiness: [
                  { Scope: 'account', Status: 'Missing' },
                  { Scope: 'contact', Status: 'Missing' },
                ],
                Pending: 2,
                Error: null,
              },
            }
          : null,
    });
    await s.press(s.$('repair-all'));
    assert.equal(s.sent.filter(([, b]) => b.Command === 'Register').length, 1);
    assert.equal(
      s.$('fb-settings').textContent,
      'Repair made no progress. 2 tables still need repair.',
    );
    assert.equal(s.$('repair-all').textContent, 'Repair all (2)');
  }
  {
    // Shared runtime state: pausing from Settings updates the switch and its state text.
    const s = await boot({ tab: 'settings' });
    await s.press(s.$('automation-switch-settings'));
    assert.deepEqual(s.sent.at(-1), [
      'asx_RuntimeAdmin',
      { Command: 'SetEnabled', Enabled: false, RowVersion: '7' },
    ]);
    assert.equal(s.$('automation-switch-settings').getAttribute('aria-checked'), 'false');
    assert.equal(s.$('automation-state-settings').textContent, 'Automation is paused');
    assert.equal(s.$('fb-settings').textContent, '');
    // Turning on gives the Settings form the new row version and keeps its unsaved edits; Save
    // then sends that version and the running state.
    const c = await boot({ tab: 'settings', runtime: { Enabled: false } });
    const host = c.$('hosts-list').querySelector('input');
    host.value = 'fabrikam.sharepoint.com';
    host.oninput();
    await c.press(c.$('automation-switch-settings'));
    assert.deepEqual(c.sent.at(-1)[1], { Command: 'SetEnabled', Enabled: true, RowVersion: '7' });
    assert.equal(c.$('automation-switch-settings').getAttribute('aria-checked'), 'true');
    assert.equal(
      c.$('hosts-list').querySelector('input').value,
      'fabrikam.sharepoint.com',
      'Unsaved edits stay',
    );
    await c.press(c.$('save-settings'));
    const save = c.sent.filter(([, b]) => b.Command === 'Save').at(-1)[1];
    assert.equal(save.RowVersion, '8');
    assert.equal(save.Enabled, true);
    assert.deepEqual(save.SharePointHosts, ['fabrikam.sharepoint.com']);
  }
  {
    // The Settings switch reports a refused change in the Settings line and keeps its state; a
    // non-administrator has no switch, only the state as text.
    const stale = await boot({
      tab: 'settings',
      handle: (api, b) =>
        b.Command === 'SetEnabled'
          ? new Error('Automation settings changed. Reopen the page and try again.')
          : null,
    });
    await stale.press(stale.$('automation-switch-settings'));
    assert.equal(
      stale.$('fb-settings').textContent,
      'Automation settings changed. Reopen the page and try again.',
    );
    assert.equal(stale.$('fb-settings').getAttribute('role'), 'alert');
    assert.equal(stale.$('automation-switch-settings').getAttribute('aria-checked'), 'true');
    const viewer = await boot({ tab: 'settings', runtime: { CanChange: false } });
    const off = viewer.$('automation-switch-settings');
    assert.equal(off.hidden, true);
    assert.equal(viewer.$('automation-state-settings').textContent, 'Automation is running');
    await viewer.press(off);
    assert.equal(viewer.sent.filter(([, b]) => b.Command === 'SetEnabled').length, 0);
  }
  {
    // The 60-second tick reads nothing off Monitor: the problem count elsewhere is read once.
    const s = await boot({ tab: 'settings' });
    const before = s.sent.length;
    // Manage tables lands on the Tables card's heading; a named table on its row.
    const tables = await boot({ tab: 'settings', stored: { tab: 'settings', table: '' } });
    assert.equal(tables.document.activeElement.id, 'tables-title');
    const one = await boot({ tab: 'settings', stored: { tab: 'settings', table: 'account' } });
    assert.ok(
      one
        .$('tables-rows')
        .querySelector('[data-table="account"]')
        .contains(one.document.activeElement),
      'The named table row takes focus',
    );
    await s.tick();
    assert.equal(s.sent.length, before);
  }
  {
    // The shell starts the page before its runtime Get returns (Task 1 fix round): Settings and
    // the Monitor automation pill draw when it arrives, and a refused Get shows the
    // missing-settings alert.
    let open;
    const s = await boot({ tab: 'settings', getGate: new Promise((resolve) => (open = resolve)) });
    assert.equal(s.$('automation-settings').getAttribute('aria-busy'), 'true');
    assert.equal(s.$('settings-missing').hidden, true);
    open();
    await s.document.settle();
    assert.equal(s.$('runtimeWorker').value, 'worker-1');
    assert.equal(s.$('automation-settings').hasAttribute('aria-busy'), false);
    let release;
    const m = await boot({ getGate: new Promise((resolve) => (release = resolve)) });
    assert.equal(m.$('monitor-automation').hidden, true);
    release();
    await m.document.settle();
    assert.equal(m.$('monitor-automation').hidden, false);
    assert.equal(m.$('monitor-automation').textContent, 'Automation running');
    const refused = await boot({
      tab: 'settings',
      handle: (api, b) =>
        b.Command === 'Get' ? new Error('Principal user is missing prvReadasx_runtime.') : null,
    });
    assert.equal(refused.$('settings-missing').hidden, false);
    assert.equal(refused.$('automation-settings').hasAttribute('aria-busy'), false);
  }
  // Settings as cards -------------------------------------------------------------------------
  {
    // Automation card: the switch applies at once through SetEnabled and is not part of Save.
    const s = await boot({ tab: 'settings' });
    assert.equal(s.$('automation-state-settings').textContent, 'Automation is running');
    assert.equal(
      s.$('help-automation-settings').textContent,
      "Turning it off pauses folder and access work. Changes keep queueing and run when it's back on.",
    );
    await s.press(s.$('automation-switch-settings'));
    assert.deepEqual(s.sent.filter(([a]) => a === 'asx_RuntimeAdmin').at(-1)[1], {
      Command: 'SetEnabled',
      Enabled: false,
      RowVersion: '7',
    });
    assert.equal(s.$('automation-state-settings').textContent, 'Automation is paused');
    assert.equal(s.$('settings-footer').hidden, true, 'The switch is not an unsaved change');
  }
  {
    // Automation settings: edits show the sticky footer with a count; Discard resets from the
    // runtime; Save sends the form and hides the footer. "Update folders when records change"
    // is out of 0.1.0.4: no row, no switch, and Save does not send it.
    const s = await boot({ tab: 'settings' });
    assert.equal(s.$('settings-footer').hidden, true);
    assert.equal(s.$('record-updates'), null);
    assert.equal(s.$('help-record-updates'), null);
    assert.doesNotMatch(s.$('automation-settings').textContent, /records change/);
    const host = s.$('hosts-list').querySelector('input');
    host.value = 'fabrikam.sharepoint.com';
    host.oninput();
    assert.equal(s.$('settings-footer').hidden, false);
    assert.equal(s.$('settings-unsaved').textContent, '1 unsaved change');
    assert.ok(host.classList.contains('is-edited'));
    await s.press(s.$('settings-discard'));
    assert.equal(s.$('settings-footer').hidden, true);
    assert.equal(s.$('hosts-list').querySelector('input').value, 'contoso.sharepoint.com');
    const again = s.$('hosts-list').querySelector('input');
    again.value = 'fabrikam.sharepoint.com';
    again.oninput();
    await s.press(s.$('save-settings'));
    const saved = s.sent.filter(([a]) => a === 'asx_RuntimeAdmin').at(-1)[1];
    assert.equal(saved.Command, 'Save');
    assert.deepEqual([...saved.SharePointHosts], ['fabrikam.sharepoint.com']);
    assert.equal('ProcessRecordUpdates' in saved, false);
    assert.equal(s.$('settings-footer').hidden, true);
    assert.equal(s.$('fb-settings').textContent, 'Settings saved.');
  }
  {
    // Tables card: one row per enabled table with its template count and change tracking, team
    // access events last; Repair, the access link, and Remove with the existing confirmation.
    const s = await boot({
      tab: 'settings',
      runtime: {
        Registration: {
          Readiness: [
            { Scope: 'contact', Status: 'Outdated' },
            { Scope: 'account', Status: 'Ready' },
            { Scope: 'lead', Status: 'WorkerCannotRead' },
            { Scope: 'team', Status: 'Ready' },
          ],
          Pending: 1,
          Error: null,
        },
      },
      rows: {
        asx_template: [
          { asx_table: 'account' },
          { asx_table: 'account' },
          { asx_table: 'contact' },
        ],
      },
    });
    const rows = [...s.$('tables-rows').querySelectorAll('[data-table]')];
    assert.deepEqual(
      rows.map((r) => r.dataset.table),
      ['account', 'contact', 'lead', 'team'],
    );
    assert.equal(rows[0].querySelector('.sub').textContent, '2 templates');
    assert.equal(rows[1].querySelector('.sub').textContent, '1 template');
    assert.equal(rows[0].querySelector('.status').visibleText, 'Ready');
    assert.equal(rows[0].querySelector('.dot').dataset.tone, 'ok');
    assert.equal(rows[1].querySelector('.status').visibleText, 'Out of date');
    assert.equal(rows[1].querySelector('.dot').dataset.tone, 'attention');
    assert.ok([...rows[1].querySelectorAll('button')].some((b) => b.textContent === 'Repair'));
    assert.equal(rows[2].querySelector('a').textContent, 'How to grant access');
    assert.equal(rows[3].querySelector('.table-name').textContent, 'Team access events');
    assert.equal(
      rows[3].querySelector('.sub').textContent,
      'Keeps library access in step with team membership',
    );
    assert.ok(![...rows[3].querySelectorAll('button')].some((b) => b.textContent === 'Remove'));
    assert.equal(s.$('repair-all').hidden, true, 'Repair all needs two or more');
    const remove = [...rows[0].querySelectorAll('button')].find((b) => b.textContent === 'Remove');
    await s.press(remove);
    assert.match(
      s.document.activeElement.textContent,
      /^Stop creating folders for .+\? Queued folder work for this table is cancelled\. Templates are kept, and nothing in SharePoint is deleted\.$/,
    );
    await s.press(s.document.querySelector('.confirm').querySelectorAll('button')[0]);
    assert.deepEqual(s.sent.filter(([a]) => a === 'asx_RuntimeAdmin').at(-1)[1], {
      Command: 'RemoveTable',
      Table: 'account',
    });
    assert.match(s.$('fb-settings').textContent, / removed\. Its templates are kept\.$/);
  }
  {
    // ＋ Add table offers the document-enabled tables not enabled yet, and adds one.
    const s = await boot({ tab: 'settings' });
    await s.press(s.$('add-table'));
    const picker = s.$('enableTable');
    assert.deepEqual(
      [...picker.querySelectorAll('option')].map((o) => o.value),
      ['', 'contact', 'lead'],
    );
    picker.value = 'lead';
    await picker.onchange();
    await s.press(s.$('add-chosen-table'));
    assert.deepEqual(s.sent.filter(([a]) => a === 'asx_RuntimeAdmin').at(-1)[1], {
      Command: 'AddTable',
      Table: 'lead',
    });
  }
  {
    // Connections: dot and words; Before uninstalling keeps its confirmation.
    const s = await boot({ tab: 'settings' });
    assert.equal(s.$('open-connections').textContent, 'Open in Power Automate');
    const items = [...s.$('connection-list').querySelectorAll('li')];
    assert.ok(items.every((li) => li.querySelector('.dot')));
    assert.equal(
      s.$('help-stop-tracking').textContent,
      'Stops capturing record and team changes for every table until change tracking is repaired.',
    );
  }
  {
    // Read-only: values as text, no inputs, switches or actions, never the footer.
    const s = await boot({ tab: 'settings', runtime: { CanChange: false } });
    const shown = [...s.$('settings').querySelectorAll('input, select, [role=switch]')].filter(
      (n) => !n.closest('[hidden]') && !n.hidden,
    );
    assert.equal(shown.length, 0);
    assert.match(s.$('automation-settings').visibleText, /contoso\.sharepoint\.com/);
    assert.match(s.$('automation-card').visibleText, /Automation is running/);
    for (const id of ['add-table', 'repair-all', 'stop-tracking', 'add-host'])
      assert.equal(s.$(id).hidden, true, id);
    assert.equal(s.$('tables-rows').querySelectorAll('button').length, 0);
    assert.equal(s.$('settings-footer').hidden, true);
    assert.equal(s.$('settings-meta').textContent, 'Only System Administrators can change these');
  }
  {
    // The header links to Monitor with the problem count (owner decision 4); none at 0.
    const s = await boot({ tab: 'settings', summary: { BlockedRecords: 1, TemplateRuns: 2 } });
    const pill = s.$('settings-problems').querySelector('.problem-pill');
    assert.equal(pill.hidden, false);
    assert.equal(pill.textContent, 'Monitor · 1 problem');
    await s.press(pill);
    assert.equal(
      s.sent
        .filter(([k]) => k === 'navigate')
        .at(-1)[1]
        .data.split('-')[0],
      'monitor',
    );
    const quiet = await boot({ tab: 'settings' });
    assert.equal(quiet.$('settings-problems').querySelector('.problem-pill').hidden, true);
  }
  // Settings fix round 1 ----------------------------------------------------------------------
  {
    // A non-administrator: the runtime Get is refused and the System Administrator privilege is
    // missing. No missing-settings alert; automation reads from the Default runtime row and shows
    // as text; no inputs or switches; a refused connections read hides that card quietly.
    const s = await boot({
      tab: 'settings',
      missing: ['prvWriteasx_runtime'],
      rows: {
        asx_runtime: [{ asx_enabled: true, asx_processrecordupdates: false }],
        connectionreference: new Error('Principal user is missing prvReadconnectionreference.'),
      },
      handle: (api, b) =>
        b.Command === 'Get' ? new Error('Principal user is missing prvWriteasx_runtime.') : null,
    });
    assert.equal(s.$('settings-missing').hidden, true, 'No missing-settings alert');
    assert.equal(s.$('automation-card').hidden, false);
    assert.equal(s.$('automation-state-settings').textContent, 'Automation is running');
    assert.equal(s.$('automation-switch-settings').hidden, true);
    const shown = [...s.$('settings').querySelectorAll('input, select, [role=switch]')].filter(
      (n) => !n.closest('[hidden]') && !n.hidden,
    );
    assert.equal(shown.length, 0);
    assert.equal(s.$('settings-meta').textContent, 'Only System Administrators can change these');
    assert.equal(s.$('settings-footer').hidden, true);
    assert.equal(s.$('connections-card').hidden, true);
    assert.doesNotMatch(s.$('settings').visibleText, /Couldn't read the connections/);
    // No Default row to read: the automation card hides too.
    const none = await boot({
      tab: 'settings',
      missing: ['prvWriteasx_runtime'],
      handle: (api, b) => (b.Command === 'Get' ? new Error('Refused.') : null),
    });
    assert.equal(none.$('automation-card').hidden, true);
    assert.equal(none.$('settings-missing').hidden, true);
  }
  {
    // ＋ Add table: arrowing through the tables only chooses one; Add adds it. Add stays disabled
    // until a table is chosen, and a refused AddTable resets the choice.
    let refuse = true;
    const s = await boot({
      tab: 'settings',
      handle: (api, b) =>
        b.Command === 'AddTable' && refuse ? new Error('The run-as user cannot read Lead.') : null,
    });
    const adds = () => s.sent.filter(([, b]) => b.Command === 'AddTable').length;
    await s.press(s.$('add-table'));
    const picker = s.$('enableTable');
    const add = s.$('add-chosen-table');
    assert.equal(add.textContent.trim(), 'Add');
    assert.equal(add.getAttribute('aria-disabled'), 'true');
    await s.press(add);
    assert.equal(adds(), 0, 'Add does nothing with no table chosen');
    picker.value = 'lead';
    await picker.onchange();
    await s.document.settle();
    assert.equal(adds(), 0, 'Choosing a table does not add it');
    assert.equal(add.hasAttribute('aria-disabled'), false);
    await s.press(add);
    assert.equal(adds(), 1);
    assert.equal(s.$('fb-settings').textContent, 'The run-as user cannot read Lead.');
    assert.equal(picker.value, '', 'A refused AddTable resets the choice');
    assert.equal(add.getAttribute('aria-disabled'), 'true');
    refuse = false;
    picker.value = 'contact';
    await picker.onchange();
    await s.press(add);
    assert.deepEqual(s.sent.filter(([, b]) => b.Command === 'AddTable').at(-1)[1], {
      Command: 'AddTable',
      Table: 'contact',
    });
    assert.equal(s.$('fb-settings').textContent, 'Contact added.');
  }
  {
    // Removing the first of two hosts is one unsaved change.
    const s = await boot({
      tab: 'settings',
      runtime: { SharePointHosts: ['contoso.sharepoint.com', 'fabrikam.sharepoint.com'] },
    });
    await s.press(s.$('hosts-list').querySelectorAll('button')[0]);
    assert.equal(s.$('settings-footer').hidden, false);
    assert.equal(s.$('settings-unsaved').textContent, '1 unsaved change');
  }
  {
    // A runtime change while a Remove table confirmation is open keeps the confirmation; the
    // Tables card redraws once it closes.
    const s = await boot({ tab: 'settings' });
    const remove = [...s.$('tables-rows').querySelectorAll('button')].find(
      (b) => b.textContent === 'Remove',
    );
    await s.press(remove);
    s.window.AsxdUi.setRuntime({
      WorkerId: 'worker-1',
      Enabled: true,
      CanChange: true,
      RowVersion: '8',
      SharePointHosts: ['contoso.sharepoint.com'],
      ProcessRecordUpdates: false,
      Tables: ['account'],
      Registration: {
        Readiness: [
          { Scope: 'account', Status: 'Outdated' },
          { Scope: 'team', Status: 'Ready' },
        ],
        Pending: 1,
        Error: null,
      },
    });
    await s.document.settle();
    const keep = () =>
      [...s.$('tables-rows').querySelectorAll('button')].find(
        (b) => b.textContent === 'Keep table',
      );
    assert.ok(keep(), 'The open confirmation survives');
    assert.match(s.$('tables-rows').visibleText, /Account.*Ready/);
    await s.press(keep());
    assert.match(s.$('tables-rows').visibleText, /Account.*Out of date/);
  }
  {
    // Discard clears a refused save's message in the save bar.
    const s = await boot({
      tab: 'settings',
      handle: (api, b) =>
        b.Command === 'Save' ? new Error('Automation settings changed. Reopen the page.') : null,
    });
    const host = s.$('hosts-list').querySelector('input');
    host.value = 'fabrikam.sharepoint.com';
    host.oninput();
    await s.press(s.$('save-settings'));
    assert.equal(
      s.$('fb-settings-save').textContent,
      'Automation settings changed. Reopen the page.',
    );
    await s.press(s.$('settings-discard'));
    assert.equal(s.$('fb-settings-save').textContent, '');
  }
  // Fix round 1 -------------------------------------------------------------------------------
  {
    // A row acted on and gone from the server's next first page is gone after the tick; a row
    // that "Show N more" appended stays. Show more runs once however often it is pressed.
    const a = row({ Key: 'folderjob:a', Title: 'Alpha · Account documents' });
    const b = row({ Key: 'folderjob:b', Title: 'Beta · Account documents' });
    let first = [a];
    let more = 0;
    const m = await boot({
      summary: { BlockedJobs: 2 },
      handle: (api, q) => {
        if (q.Command !== 'ListProblems' || q.List !== 'BlockedJobs') return null;
        if (!q.Page) return { Status: 'Page', Problems: first, Next: 'p2' };
        more++;
        return { Status: 'Page', Problems: [b], Next: null };
      },
    });
    const showMore = m.$('problem-more');
    assert.equal(showMore.hidden, false);
    assert.equal(showMore.textContent, 'Show 1 more');
    showMore.click();
    showMore.click();
    await m.document.settle();
    assert.equal(more, 1, 'Show more appends once');
    assert.equal(trs(m).length, 2);
    assert.equal(showMore.hidden, true, 'Nothing more to show');
    await m.press(
      m
        .buttons(table(m))
        .find((x) => x.getAttribute('aria-label') === 'Retry for Alpha · Account documents'),
    );
    first = [];
    await m.tick();
    assert.doesNotMatch(table(m).visibleText, /Alpha/, 'The retried row is gone');
    assert.match(table(m).visibleText, /Beta/, 'The appended row stays');
  }
  {
    // Show more pages the merged list: 50 rows at a time, counted from Summary, and "Show 50 more"
    // for a capped list.
    const many = (list, n) =>
      Array.from({ length: n }, (_, i) =>
        row({
          Key: list + ':' + String(i).padStart(3, '0'),
          SinceUtc: new Date(Date.parse('2026-10-06T10:00:00Z') - i * 60000).toISOString(),
        }),
      );
    const m = await boot({
      summary: { BlockedJobs: 40, RetryingJobs: 30 },
      lists: { BlockedJobs: many('job', 40), RetryingJobs: many('retry', 30) },
    });
    assert.equal(trs(m).length, 50);
    assert.equal(m.$('problem-more').textContent, 'Show 20 more');
    await m.press(m.$('problem-more'));
    assert.equal(trs(m).length, 70);
    assert.equal(m.$('problem-more').hidden, true);
    const capped = await boot({
      summary: { BlockedJobs: 5000, Capped: ['BlockedJobs'] },
      handle: (api, q) =>
        q.Command === 'ListProblems' && q.List === 'BlockedJobs'
          ? { Status: 'Page', Problems: many('job', 50), Next: 'p2' }
          : null,
    });
    assert.equal(capped.$('problem-more').textContent, 'Show 50 more');
  }
  {
    // Show more that fails keeps the loaded rows and reports in the Monitor feedback line.
    const m = await boot({
      summary: { BlockedJobs: 1 },
      handle: (api, q) =>
        q.Command === 'ListProblems' && q.List === 'BlockedJobs'
          ? q.Page
            ? new Error('The list changed. Refresh it.')
            : { Status: 'Page', Problems: [row()], Next: 'p2' }
          : null,
    });
    // Summary is behind the server here: the next page is still offered.
    assert.equal(m.$('problem-more').textContent, 'Show 50 more');
    await m.press(m.$('problem-more'));
    assert.equal(trs(m).length, 1);
    assert.equal(m.$('fb-monitor').textContent, 'The list changed. Refresh it.');
    assert.equal(m.$('problem-more').hidden, false, 'Show more stays to try again');
  }
  {
    // The tick does not redraw the table under an open confirmation or focus on a control it
    // cannot restore; it redraws it on a later tick once that is over. Focus on a chip, a row
    // action or the Since sort comes back after the redraw.
    const lists = { BlockedJobs: [row()] };
    const m = await boot({ summary: { BlockedJobs: 1 }, lists });
    const cancel = named(m, trs(m)[0], 'Cancel job');
    await m.press(cancel);
    const question = m.document.activeElement;
    lists.BlockedJobs = [row({ Problem: 'Something else now.' })];
    await m.tick();
    assert.equal(m.document.activeElement, question, 'The open confirmation keeps focus');
    assert.equal(question.isConnected, true, 'The open confirmation survives the tick');
    await m.press(
      [...table(m).querySelectorAll('.confirm button')].find((x) => x.textContent === 'Cancel job'),
    );
    assert.deepEqual(m.sent.at(-1)[1], { Command: 'Cancel', Key: 'folderjob:abc' });
    m.document.body.focus();
    await m.tick();
    assert.match(table(m).visibleText, /Something else now\./, 'Redrawn on the next tick');
    // A control with no focus key (a load error's Try again aside) holds the redraw.
    const unkeyed = m.document.createElement('button');
    trs(m)[0].querySelector('td').append(unkeyed);
    unkeyed.focus();
    lists.BlockedJobs = [row({ Problem: 'Third problem.' })];
    await m.tick();
    assert.equal(unkeyed.isConnected, true, 'Unkeyed focus holds the redraw');
    m.document.body.focus();
    await m.tick();
    assert.match(table(m).visibleText, /Third problem\./);
    // Keyed focus comes back to the same control after the redraw.
    primaryOf(trs(m)[0]).focus();
    await m.tick();
    assert.equal(m.document.activeElement.textContent, 'Retry');
    assert.equal(m.document.activeElement.isConnected, true);
    const blockedChip = chip(m, 'Blocked jobs');
    blockedChip.focus();
    await m.tick();
    assert.equal(m.document.activeElement.textContent, 'Blocked jobs · 1');
    assert.equal(m.document.activeElement.isConnected, true);
  }
  {
    // Not captured: only rows the server can re-run get a checkbox; Re-run selected reports each
    // row it could not queue and still queues the rest.
    const capture = (id, name, record, actions) =>
      row({
        Key: id,
        Kind: 'CaptureJob',
        KindLabel: 'Missed change',
        Title: name,
        Record: record,
        Problem: 'Runtime identity is incomplete.',
        Status: 'Failed',
        Actions: actions,
      });
    const account = (id, name) => ({ Table: 'account', TableLabel: 'Account', Id: id, Name: name });
    const id = (n) => 'aaaaaaaa-0000-0000-0000-00000000000' + n;
    const m = await boot({
      summary: { NotCaptured: 4 },
      lists: {
        NotCaptured: [
          capture(id(1), 'Alpha · Account', account(id(1), 'Alpha'), [
            'Rerun',
            'Dismiss',
            'OpenRecord',
          ]),
          capture(id(2), 'Beta · Account', account(id(2), 'Beta'), [
            'Rerun',
            'Dismiss',
            'OpenRecord',
          ]),
          capture(
            id(3),
            'Sales team · Team',
            { Table: 'team', TableLabel: 'Team', Id: id(3), Name: 'Sales team' },
            ['Dismiss', 'OpenRecord'],
          ),
          capture(id(4), 'Record not available · Missed change', null, ['Dismiss']),
        ],
      },
      handle: (api, b) =>
        b.Command === 'RerunRecord'
          ? b.RecordId === id(2)
            ? new Error('Enable account in Folder templates first.')
            : { Status: 'Queued', Keys: ['request:x'], Notices: [] }
          : null,
    });
    await m.press(chip(m, 'Not captured'));
    const boxes = m.$('problem-rows').querySelectorAll('input[type=checkbox]');
    assert.equal(boxes.length, 2, 'No checkbox on team rows or rows without a record');
    for (const box of boxes) {
      box.checked = true;
      box.onchange();
    }
    await m.press(m.$('rerun-selected-row').querySelector('button'));
    assert.equal(m.sent.filter(([, b]) => b.Command === 'RerunRecord').length, 2);
    assert.equal(
      m.$('fb-monitor').textContent,
      "Queued 1; 1 couldn't be re-run: Beta: Enable account in Folder templates first.",
    );
    assert.equal(m.$('fb-monitor').getAttribute('role'), 'alert');
  }
  {
    // Refresh says "5,000+" for a capped list.
    const m = await boot({ summary: { NotCaptured: 5000, Capped: ['NotCaptured'] } });
    await m.press(m.$('monitor-refresh'));
    assert.equal(m.$('fb-monitor').textContent, 'Refreshed. 5,000+ changes not captured.');
  }
  {
    // A recent operation that cannot be looked up says why in Look up an operation.
    const m = await boot({
      lists: { RecentOperations: [row({ Status: 'Applied', Actions: [] })] },
      handle: (api, b) => (b.Command === 'Inspect' ? new Error('Operation not found.') : null),
    });
    await openTool(m, 1);
    await m.press(m.buttons(m.$('recent-rows'))[0]);
    assert.equal(m.$('fb-advanced').textContent, 'Operation not found.');
  }
  {
    // A per-table Repair that makes the table Ready keeps focus in its row: its Remove.
    const ready = (status, pending) => ({
      WorkerId: 'worker-1',
      Enabled: true,
      CanChange: true,
      RowVersion: '7',
      SharePointHosts: [],
      Registration: {
        Readiness: [{ Scope: 'account', Status: status }],
        Pending: pending,
        Error: null,
      },
    });
    const s = await boot({
      tab: 'settings',
      runtime: { Registration: ready('Missing', 1).Registration },
      handle: (api, b) => (b.Command === 'Register' ? ready('Ready', 0) : null),
    });
    const repair = s.$('tables-rows').querySelector('button');
    repair.focus();
    await s.press(repair);
    assert.equal(s.document.activeElement.getAttribute('aria-label'), 'Remove Account');
  }
  console.log(
    'PASS Monitor and Settings: Monitor as one table (chips counted from Summary with 5,000+ and zero chips, the merged list sorted by Since with one primary action and a ⋯ menu with Copy ID, a filter and its fallback to All, Not captured checkboxes kept across a redraw), the re-runs strip, Tools side panels and links to them, lists on open, row actions, Retry now, 60-second re-read announcements after an unread error, recovery choices (table and Look up an operation), load errors, Refresh, the automation pill and its Settings link, checklist, Check a record, Look up an operation, Settings, Repair all, danger zone, the Settings switch and its refusals, Settings as cards (the automation card, the save bar with its count and Discard, the Tables card with counts, change tracking, Repair, Remove and Add table, Connections, read-only values, the Monitor problem pill; fix round 1: the read-only page for a non-administrator from the Default runtime row, Add table by its Add button and a refused add, removing a host, a runtime change under a Remove confirmation, Discard clears the save bar message), runtime shared with the form, no tick off Monitor, a late or refused Get; fix round 1: the tick keeps appended rows and drops removed ones, defers redraws under confirmations, open menus and unkeyed focus; checkboxes only on re-runnable rows with per-row reasons; Show more once, paged by 50 from Summary, and its errors; capped Refresh text; recent lookup errors; Repair focus and text; one announcement for the switch, one Monitor line for every announcement of a tick; Retry of a library setup on the Operator role and Cancel setup on the Security Administrator role; a Settings link landing on the Tables card or the row of its table. Fake DOM; browser QA separate.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
