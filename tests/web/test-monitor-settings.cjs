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
      Id: '607cba8a-acc1-f111-aaaf-7c1e52067fd0',
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
  handle = () => null,
  getGate = null,
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
      navigateTo: async () => {},
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
          id: '{607CBA8A-ACC1-F111-AAAF-7C1E52067FD0}',
          name: 'Contoso Ltd',
          entityType: 'account',
        },
      ],
      getEntityMetadata: async (table) => ({
        DisplayName: table[0].toUpperCase() + table.slice(1),
      }),
    },
    WebApi: {
      retrieveMultipleRecords: async (table) => ({
        entities:
          table === 'systemuser'
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
                      { connectionreferencedisplayname: 'Documents Dataverse', connectionid: null },
                    ]
                  : table === 'asx_library'
                    ? [{ asx_libraryid: 'lib-1' }]
                    : [],
      }),
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
  const window = {
    parent: { Xrm: xrm },
    location: { search: '?data=' + tab + '-ui20261006nav1', hash: '' },
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
    navigator: { clipboard: { writeText: async () => {} } },
    crypto: require('node:crypto').webcrypto,
    fetch: async () => ({ ok: true, json: async () => ({ RolePrivileges: [{}] }) }),
    setTimeout,
    clearTimeout,
    // The page's 60-second timers are collected; tick() runs them once, as if a minute passed.
    setInterval: (fn) => ticks.push(fn),
    clearInterval: () => {},
  });
  for (const name of ['shell.js', 'operations.js']) vm.runInContext(read(name), context);
  await document.fire('DOMContentLoaded');
  const $ = (id) => document.getElementById(id);
  const buttons = (node) => node.querySelectorAll('button');
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
    commands: () => sent.map(([, b]) => b.Command),
  };
}

(async () => {
  {
    // Opening Monitor loads Summary and every list without a click; counts show in headings and tiles.
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
    assert.equal(m.$('h-BlockedJobs').textContent, 'Blocked jobs · 1');
    assert.equal(m.$('h-NotCaptured').textContent, 'Changes not captured · 0');
    assert.equal(m.$('list-NotCaptured').visibleText.includes('No missed changes.'), true);
    assert.equal(
      m.$('list-NotCaptured').querySelectorAll('button').length,
      0,
      'An empty list offers no actions',
    );
    const tile = m
      .$('monitor-tiles')
      .querySelectorAll('button')
      .find((b) => b.textContent.startsWith('Blocked jobs'));
    assert.equal(tile.textContent, 'Blocked jobs 1');
    await m.press(tile);
    assert.equal(m.document.activeElement.id, 'h-BlockedJobs');
    // The badge counts the five problem lists.
    assert.equal(m.$('monitorBadge').textContent, '1');
    assert.equal(m.$('tab-monitor').getAttribute('aria-label'), 'Monitor, 1 problem');
  }
  {
    // Rows show names, local times and server actions; keys only inside Details.
    const m = await boot({ summary: { BlockedJobs: 1 }, lists: { BlockedJobs: [row()] } });
    const list = m.$('list-BlockedJobs');
    const outside = list.visibleText;
    assert.match(outside, /Contoso Ltd · Account documents/);
    assert.match(outside, /The folder path is too long/);
    assert.doesNotMatch(outside, /folderjob:abc/);
    assert.doesNotMatch(outside, GUID);
    assert.doesNotMatch(outside, /2026-10-06T/);
    assert.equal(list.querySelector('time').getAttribute('datetime'), '2026-10-06T17:38:07Z');
    const details = list.querySelector('details');
    details.open = true;
    assert.match(details.visibleText, /folderjob:abc/);
    assert.deepEqual(
      m
        .buttons(list.querySelector('tbody'))
        .map((b) => b.textContent)
        .filter((t) => t !== 'Copy'),
      ['Retry', 'Cancel job', 'Open record', 'Check'],
    );
    // Open record opens the form in a new window.
    await m.press(m.buttons(list).find((b) => b.textContent === 'Open record'));
    assert.deepEqual(m.opened.at(-1), {
      entityName: 'account',
      entityId: '607cba8a-acc1-f111-aaaf-7c1e52067fd0',
      openInNewWindow: true,
    });
    // Retry sends the operation's Retry and leaves the row with its new status.
    const retry = m.buttons(list).find((b) => b.textContent === 'Retry');
    await m.press(retry);
    assert.deepEqual(m.sent.at(-1), ['asx_ManageWork', { Command: 'Retry', Key: 'folderjob:abc' }]);
    assert.match(list.visibleText, /Queued again/);
    assert.equal(retry.disabled, true);
    assert.equal(
      m.buttons(list).find((b) => b.textContent === 'Cancel job').disabled,
      true,
      'The row actions disable',
    );
    assert.equal(m.buttons(list).find((b) => b.textContent === 'Open record').disabled, false);
    assert.equal(
      m.$('fb-list-BlockedJobs').textContent,
      'Queued again. It blocks again if the cause remains.',
    );
    // Cancel job confirms in the page first, with no key in its text.
    const fresh = await boot({ summary: { BlockedJobs: 1 }, lists: { BlockedJobs: [row()] } });
    const cancel = fresh
      .buttons(fresh.$('list-BlockedJobs'))
      .find((b) => b.textContent === 'Cancel job');
    await fresh.press(cancel);
    assert.equal(
      fresh.document.activeElement.textContent,
      'Cancel the folder job for Contoso Ltd · Account documents? Nothing in SharePoint is undone or deleted.',
    );
    const keep = fresh
      .buttons(fresh.$('list-BlockedJobs'))
      .find((b) => b.textContent === 'Keep job');
    await fresh.press(keep);
    assert.equal(fresh.commands().includes('Cancel'), false);
    await fresh.press(cancel);
    await fresh.press(
      fresh
        .buttons(fresh.$('list-BlockedJobs'))
        .find((b) => b.textContent === 'Cancel job' && b !== cancel),
    );
    assert.deepEqual(fresh.sent.at(-1), [
      'asx_ManageWork',
      { Command: 'Cancel', Key: 'folderjob:abc' },
    ]);
    assert.equal(
      fresh.$('list-BlockedJobs').querySelectorAll('.confirm').length,
      0,
      'The confirmation closes',
    );
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
              Id: '607cba8a-acc1-f111-aaaf-7c1e52067fd0',
              Name: null,
            },
            Actions: ['Retry', 'Check'],
          }),
        ],
      },
    });
    assert.match(m.$('list-BlockedRecords').visibleText, /Record not available to you/);
    await m.press(m.buttons(m.$('list-BlockedRecords')).find((b) => b.textContent === 'Retry'));
    assert.deepEqual(m.sent.at(-1)[1], { Command: 'RetryOutbox', Key: 'request:1' });
  }
  {
    // Retrying automatically: the row action is "Retry now" (spec 3.3) and sends the job's Retry.
    const waiting = row({
      Key: 'folderjob:wait',
      Status: 'RetryWait',
      Problem: 'SharePoint is busy or limiting requests (HTTP 429).',
      Attempt: 3,
      NextAttemptUtc: '2026-10-06T17:45:00Z',
      Actions: ['Retry', 'Cancel'],
    });
    const m = await boot({ summary: { RetryingJobs: 1 }, lists: { RetryingJobs: [waiting] } });
    const list = m.$('list-RetryingJobs');
    assert.deepEqual(
      m
        .buttons(list.querySelector('tbody'))
        .map((b) => b.textContent)
        .filter((t) => t !== 'Copy'),
      ['Retry now', 'Cancel job'],
    );
    const now = m.buttons(list).find((b) => b.textContent === 'Retry now');
    assert.equal(now.getAttribute('aria-label'), 'Retry now for Contoso Ltd · Account documents');
    await m.press(now);
    assert.deepEqual(m.sent.at(-1), [
      'asx_ManageWork',
      { Command: 'Retry', Key: 'folderjob:wait' },
    ]);
    // Copy names what it copies (spec 5.1).
    assert.equal(
      m
        .buttons(list)
        .find((b) => b.textContent === 'Copy')
        .getAttribute('aria-label'),
      'Copy details for Contoso Ltd · Account documents',
    );
  }
  {
    // Missed changes: Re-run selected for checked rows; Dismiss removes the row and keeps focus in the list.
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
    const list = m.$('list-NotCaptured');
    const boxes = list.querySelectorAll('input[type=checkbox]');
    assert.equal(boxes.length, 2);
    // Each checkbox is named by its record cell.
    assert.equal(
      m.document.getElementById(boxes[0].getAttribute('aria-labelledby')).textContent,
      'Alpha',
    );
    const selected = m.buttons(list).find((b) => b.textContent.startsWith('Re-run selected'));
    assert.equal(selected.textContent, 'Re-run selected (0)');
    boxes[1].checked = true;
    boxes[1].onchange();
    assert.equal(selected.textContent, 'Re-run selected (1)');
    await m.press(selected);
    const rerun = m.sent.filter(([, b]) => b.Command === 'RerunRecord');
    assert.equal(rerun.length, 1);
    assert.equal(rerun[0][1].RecordId, 'aaaaaaaa-0000-0000-0000-000000000002');
    assert.equal(rerun[0][1].Table, 'account');
    assert.ok(rerun[0][1].RequestId);
    const dismiss = m.buttons(list).find((b) => b.textContent === 'Dismiss');
    dismiss.focus();
    await m.press(dismiss);
    assert.deepEqual(m.sent.at(-1)[1], {
      Command: 'DismissCaptureJob',
      JobId: 'aaaaaaaa-0000-0000-0000-000000000001',
    });
    assert.equal(list.querySelectorAll('tbody tr').length, 1);
    assert.equal(m.document.activeElement.textContent, 'Beta', 'Focus moves to the next row');
  }
  {
    // Template re-runs: progress text and value text, states, Pause, Resume, Retry and Cancel re-run.
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
          StartedBy: 'Matt LaCasse',
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
    const list = m.$('list-TemplateRuns');
    const bar = list.querySelector('progress');
    assert.equal(bar.getAttribute('aria-valuetext'), '12,500 of 40,000, running');
    assert.match(list.visibleText, /12,500 of 40,000/);
    assert.match(list.visibleText, /TEST Account Documents · Account · v2/);
    assert.match(list.visibleText, /Ends around/);
    assert.match(list.visibleText, /Matt LaCasse/);
    await m.press(m.buttons(list).find((b) => b.textContent === 'Pause'));
    assert.deepEqual(m.sent.at(-1)[1], { Command: 'PauseTemplateRun', Key: 'templaterun:9c4e' });
    assert.match(list.visibleText, /Paused/);
    // Cancel re-run confirms with the spec text.
    const c = await boot({
      summary: { TemplateRuns: 1 },
      lists: { TemplateRuns: [run('Blocked')] },
    });
    await c.press(
      c.buttons(c.$('list-TemplateRuns')).find((b) => b.textContent === 'Cancel re-run'),
    );
    assert.equal(
      c.document.activeElement.textContent,
      'Cancel the re-run of TEST Account Documents? Records not yet planned are skipped. Folder work already queued for planned records continues, and nothing in SharePoint is undone.',
    );
    await c.press(c.buttons(c.$('list-TemplateRuns')).find((b) => b.textContent === 'Keep re-run'));
    await c.press(c.buttons(c.$('list-TemplateRuns')).find((b) => b.textContent === 'Retry'));
    assert.deepEqual(c.sent.at(-1)[1], { Command: 'RetryOutbox', Key: 'templaterun:9c4e' });
    assert.match(c.$('list-TemplateRuns').visibleText, /Needs attention|Queued again/);
    // Estimated totals (the daily row-count snapshot) and finished runs.
    const capped = await boot({
      summary: { TemplateRuns: 1 },
      lists: {
        TemplateRuns: [
          run('Running', { Total: 40000, TotalEstimated: true, EstimatedFinishUtc: null }),
        ],
      },
    });
    assert.match(capped.$('list-TemplateRuns').visibleText, /12,500 of about 40,000/);
    assert.match(capped.$('list-TemplateRuns').visibleText, /Estimating/);
    const done = await boot({ lists: { TemplateRuns: [run('Done', { Planned: 40000 })] } });
    assert.match(done.$('list-TemplateRuns').visibleText, /Done/);
    assert.equal(
      done
        .buttons(done.$('list-TemplateRuns').querySelector('tbody'))
        .filter((b) => b.textContent !== 'Copy').length,
      0,
    );
  }
  {
    // The 60-second re-read while Monitor is open (spec 5.3): progress updates silently; a re-run
    // reaching Done, and a setup leaving "Checking SharePoint…", are each announced once.
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
          StartedBy: 'Matt LaCasse',
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
    assert.match(m.$('list-TemplateRuns').visibleText, /200 of 40,000/);
    assert.equal(m.$('fb-monitor').textContent, '', 'Progress alone is not announced');
    assert.equal(m.$('fb-list-BlockedJobs').textContent, '', 'Still checking: nothing to announce');
    lists.TemplateRuns = [runRow('Done', 40000)];
    lists.BlockedJobs = [
      setupRow(
        { State: 'NotFound', Candidates: [], Choices: ['CreateAgain', 'CheckAgain', 'Cancel'] },
        ['CreateAgain', 'CheckAgain', 'Cancel'],
      ),
    ];
    await m.tick();
    assert.equal(m.$('fb-monitor').textContent, 'Re-run of TEST Account Documents: Done.');
    assert.equal(m.$('fb-monitor').getAttribute('role'), 'status');
    assert.equal(
      m.$('fb-list-BlockedJobs').textContent,
      "SharePoint has no library named Project documents. The creation didn't happen.",
    );
    assert.ok(
      m.buttons(m.$('list-BlockedJobs')).find((b) => b.textContent === 'Create it again'),
      'The finding brings its choices',
    );
    // Once: an unchanged tick announces nothing again.
    m.$('fb-monitor').textContent = '';
    m.$('fb-list-BlockedJobs').textContent = '';
    await m.tick();
    assert.equal(m.$('fb-monitor').textContent, '');
    assert.equal(m.$('fb-list-BlockedJobs').textContent, '');
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
    assert.match(
      checking.$('list-BlockedJobs').visibleText,
      /Checking SharePoint for Project documents…/,
    );
    assert.deepEqual(
      checking
        .buttons(checking.$('list-BlockedJobs').querySelector('tbody'))
        .map((b) => b.textContent)
        .filter((t) => t !== 'Copy'),
      ['Cancel setup'],
    );
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
    const list = found.$('list-BlockedJobs');
    assert.match(
      list.visibleText,
      /SharePoint has a library Project documents at \/sites\/x\/Project documents, created .*\. It matches this request\./,
    );
    await found.press(
      found.buttons(list).find((b) => b.textContent === 'Use the library that was created'),
    );
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
    assert.equal(
      found.$('fb-list-BlockedJobs').textContent,
      'Using the existing library. Setup continues.',
    );
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
      notFound.$('list-BlockedJobs').visibleText,
      /SharePoint has no library named Project documents\. The creation didn't happen\./,
    );
    await notFound.press(
      notFound
        .buttons(notFound.$('list-BlockedJobs'))
        .find((b) => b.textContent === 'Create it again'),
    );
    assert.deepEqual(notFound.sent.at(-1)[1], {
      Command: 'ResolveSetup',
      Key: 'librarycreate:x',
      Choice: 'CreateAgain',
      RowVersion: 'rv-3',
    });
    assert.equal(notFound.$('fb-list-BlockedJobs').textContent, 'Creating the library again.');
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
    await recheck.press(
      recheck.buttons(recheck.$('list-BlockedJobs')).find((b) => b.textContent === 'Check again'),
    );
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
    const a = ambiguous.$('list-BlockedJobs');
    assert.match(a.visibleText, /A library has this name but a different address\./);
    assert.match(a.visibleText, /\/sites\/x\/Project documents1/);
    assert.equal(
      ambiguous.buttons(a).some((b) => b.textContent === 'Create it again'),
      false,
    );
    await ambiguous.press(ambiguous.buttons(a).find((b) => b.textContent === 'Use this one'));
    assert.equal(
      ambiguous.document.activeElement.textContent,
      'Use Project documents at /sites/x/Project documents1 for this setup? Documents stops its permission inheritance if needed and manages its team access.',
    );
    await ambiguous.press(ambiguous.buttons(a).find((b) => b.textContent === 'Use this library'));
    assert.deepEqual(ambiguous.sent.at(-1)[1], {
      Command: 'ResolveSetup',
      Key: 'librarycreate:x',
      Choice: 'UseLibrary',
      ListId: moved.ListId,
      RowVersion: 'rv-3',
    });
    for (const m of [checking, found, notFound, ambiguous])
      assert.equal(
        m.document
          .querySelectorAll('input[type=text], textarea')
          .filter((n) => /run|token|evidence|response/i.test(n.id)).length,
        0,
      );
  }
  {
    // A list that fails to load says so, with Try again; the others still render.
    const m = await boot({
      handle: (api, b) =>
        b.List === 'RetryingJobs'
          ? new Error('Principal user is missing prvReadasx_operation.')
          : null,
    });
    assert.equal(
      m.$('list-RetryingJobs').querySelector('[role=alert]').textContent,
      "Couldn't load Retrying automatically: Principal user is missing prvReadasx_operation.",
    );
    assert.ok(m.buttons(m.$('list-RetryingJobs')).find((b) => b.textContent === 'Try again'));
    assert.match(m.$('list-BlockedJobs').visibleText, /No blocked jobs\./);
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
    // Automation switch: SetEnabled with the row version; a non-administrator sees the reason.
    const m = await boot();
    const toggle = m.$('automation-switch-monitor');
    assert.equal(toggle.getAttribute('aria-checked'), 'true');
    assert.equal(toggle.getAttribute('aria-describedby'), 'help-automation-monitor');
    await m.press(toggle);
    assert.deepEqual(m.sent.at(-1), [
      'asx_RuntimeAdmin',
      { Command: 'SetEnabled', Enabled: false, RowVersion: '7' },
    ]);
    assert.equal(
      m.$('fb-automation-monitor').textContent,
      'Automation paused. Changes keep queueing.',
    );
    assert.equal(m.$('automationChipText').textContent, 'Automation paused');
    const viewer = await boot({ runtime: { CanChange: false } });
    const off = viewer.$('automation-switch-monitor');
    assert.equal(off.getAttribute('aria-disabled'), 'true');
    assert.match(off.getAttribute('aria-describedby'), /help-automation-monitor/);
    assert.equal(
      viewer.document.getElementById('automation-reason-monitor').textContent,
      'Only a System Administrator can pause or resume automation.',
    );
  }
  {
    // Setup incomplete: a checklist replaces the counts, each open step linking to its tab.
    const m = await boot({
      runtime: { WorkerId: '00000000-0000-0000-0000-000000000000', Enabled: false },
    });
    const checklist = m.$('setup-checklist');
    assert.equal(checklist.hidden, false);
    assert.match(checklist.visibleText, /Choose who runs automation \(Settings\)/);
    assert.equal(m.$('monitor-tiles').hidden, true);
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
      RecordId: '607cba8a-acc1-f111-aaaf-7c1e52067fd0',
    });
    assert.equal(m.$('check-result').querySelectorAll('pre').length, 0);
    assert.match(m.$('check-result').visibleText, /Folders are created\./);
    await m.press(m.$('check-rerun'));
    assert.equal(m.sent.at(-1)[1].Command, 'Replan');
    assert.equal(m.$('fb-check').textContent, 'Re-run queued for Contoso Ltd.');
  }
  {
    // Advanced: recent operations page on open; Look up validates the ID; there is no recovery panel.
    const m = await boot({
      lists: { RecentOperations: [row({ Status: 'Applied', Actions: [] })] },
      handle: (api, b) =>
        b.Command === 'Inspect'
          ? { Key: b.Key, Status: 'Blocked', Notices: ['RequestUrlTooLong'] }
          : null,
    });
    const advanced = m.$('advanced');
    advanced.open = true;
    await advanced.ontoggle();
    await m.document.settle();
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
    // A library setup with a SharePoint finding offers its recovery choices here too (spec 3.3).
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
    r.$('advanced').open = true;
    await r.$('advanced').ontoggle();
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
    amb.$('advanced').open = true;
    await amb.$('advanced').ontoggle();
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
    const rows = s.$('tracking-rows').querySelectorAll('tr');
    assert.match(rows[0].visibleText, /Account.*Some steps missing.*Repair/);
    assert.match(rows[1].visibleText, /Run-as user can't read this table/);
    assert.equal(rows[1].querySelectorAll('button').length, 0, 'WorkerCannotRead has no Repair');
    await s.press(rows[0].querySelector('button'));
    assert.deepEqual(s.sent.at(-1)[1], { Command: 'Register', Table: 'account', RowVersion: '7' });
    await s.press(s.$('repair-all'));
    assert.equal(
      s.sent.filter(([, b]) => b.Command === 'Register' && b.Table === '').length,
      2,
      'Repair all stops when nothing is pending',
    );
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
    // A non-administrator sees read-only settings.
    const viewer = await boot({ tab: 'settings', runtime: { CanChange: false } });
    assert.equal(viewer.$('settings-readonly').hidden, false);
    assert.equal(viewer.$('save-settings').hidden, true);
    assert.equal(viewer.$('runtimeWorker').disabled, true);
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
      s.$('fb-tracking').textContent,
      'Repair made no progress. 2 tables still need repair.',
    );
  }
  {
    // Shared runtime state (controller ruling 2): pausing from Settings updates the chip.
    const s = await boot({ tab: 'settings' });
    await s.press(s.$('automation-switch-settings'));
    assert.deepEqual(s.sent.at(-1), [
      'asx_RuntimeAdmin',
      { Command: 'SetEnabled', Enabled: false, RowVersion: '7' },
    ]);
    assert.equal(s.$('automationChipText').textContent, 'Automation paused');
    assert.equal(s.$('automation-switch-settings').getAttribute('aria-checked'), 'false');
    assert.equal(s.$('automation-state-settings').textContent, 'Paused');
    assert.equal(
      s.$('fb-automation-settings').textContent,
      'Automation paused. Changes keep queueing.',
    );
    // Turning on from the chip gives the Settings form the new row version and keeps its
    // unsaved edits; Save then sends that version and the running state.
    const c = await boot({ tab: 'settings', runtime: { Enabled: false } });
    const host = c.$('hosts-list').querySelector('input');
    host.value = 'fabrikam.sharepoint.com';
    host.oninput();
    await c.press(c.$('automationChipAction'));
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
    // The shell starts the tab before its runtime Get returns (Task 1 fix round): Settings and
    // the Monitor switch draw when it arrives, and a refused Get shows the missing-settings alert.
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
    assert.equal(m.$('automation-switch-monitor').hidden, true);
    release();
    await m.document.settle();
    assert.equal(m.$('automation-switch-monitor').hidden, false);
    assert.equal(m.$('automation-switch-monitor').getAttribute('aria-checked'), 'true');
    const refused = await boot({
      tab: 'settings',
      handle: (api, b) =>
        b.Command === 'Get' ? new Error('Principal user is missing prvReadasx_runtime.') : null,
    });
    assert.equal(refused.$('settings-missing').hidden, false);
    assert.equal(refused.$('automation-settings').hasAttribute('aria-busy'), false);
  }
  console.log(
    'PASS Monitor and Settings: lists on open, names and Details, row actions, Retry now, re-runs, 60-second re-read announcements, recovery choices (lists and Advanced), load errors, Refresh, switch, checklist, Check a record, Advanced, Settings, Repair all, danger zone, runtime shared with the chip, a late or refused Get. Fake DOM; browser QA separate.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
