'use strict';
// Folder templates with a fake DOM and mocked Dataverse: the templates list and the overview,
// its ⋯ menu, Schedule and Versions panels, Delete, the editor's header, steps, autosave and
// change tracking, step 1 Destinations, Publish, step 2 Folders (the tree, the folder panel, ＋ Field,
// the condition builder and test records), step 3 Review and publish, focus, and
// Re-run for existing records.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createDocument, FakeEvent } = require('./fake-dom.cjs');
const base = path.resolve(__dirname, '../../client/admin');
const html = fs.readFileSync(path.join(base, 'index.html'), 'utf8');
const read = (name) => fs.readFileSync(path.join(base, name), 'utf8');
const BUILD = /const BUILD = '([^']+)'/.exec(read('shell.js'))[1];
const GUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;
const TEMPLATE = '11111111-0000-0000-0000-000000000001';
const OTHER = '33333333-0000-0000-0000-000000000003';
const CONTACT = '22222222-0000-0000-0000-000000000002';
const RECORD = '00000000-0000-0000-0000-0000000000a1';

const attribute = (name, type, label) => ({
  LogicalName: name,
  AttributeType: type,
  DisplayName: { UserLocalizedLabel: { Label: label } },
});
const metadata = {
  account: [
    attribute('name', 'String', 'Account Name'),
    attribute('accountnumber', 'String', 'Account Number'),
    attribute('revenue', 'Money', 'Annual Revenue'),
    attribute('closedate', 'DateTime', 'Close Date'),
    attribute('reviewedon', 'DateTime', 'Reviewed On'),
    attribute('statecode', 'State', 'Status'),
    attribute('industrycode', 'Picklist', 'Industry'),
    attribute('primarycontactid', 'Lookup', 'Primary Contact'),
    attribute('old_code', 'String', '(Deprecated) Old Code'),
  ],
  contact: [attribute('fullname', 'String', 'Full Name')],
};

function draft(overrides = {}) {
  return {
    TemplateId: TEMPLATE,
    Name: 'Account onboarding',
    Table: 'account',
    Sources: [
      { Alias: 'root', Table: 'account', Lookup: null, Columns: [{ Name: 'name', Kind: 'Text' }] },
    ],
    Destinations: [
      {
        Key: 'general',
        Name: 'Business documents',
        LibraryId: 'lib-a',
        Folders: [
          { Key: 'root', Parent: null, Name: '{root.name}', Condition: null },
          { Key: 'general_docs', Parent: 'root', Name: 'General', Condition: null },
        ],
      },
    ],
    ...overrides,
  };
}

async function boot({
  pickMany = false,
  enabled = ['account'],
  templates = [
    {
      asx_templateid: TEMPLATE,
      asx_name: 'Account onboarding',
      asx_table: 'account',
      _asx_publishedrevisionid_value: 'rev-1',
      asx_disabled: false,
    },
  ],
  loaded = {
    RevisionId: 'rev-1',
    RowVersion: '3',
    Status: 'Published',
    Version: 1,
    Draft: draft(),
  },
  runtime = {},
  privileges = {},
  handle = () => null,
  libraries,
  // Fix round 1 hooks: a template or revision read's answer (or Error), saved revisions' numbers,
  // more lookups on account, and a metadata request's delay or failure.
  reads = () => null,
  versions = { 'rev-2': 2 },
  extraLookups = [],
  delay = async () => {},
  failFetch = () => false,
  // The asx_ManageWork Summary counts, and a table read's rows (a non-null answer wins).
  summary = {},
  rows = () => null,
  // The page address's #hash: a deep link.
  hash = '',
  // A single-select lookup's answer (test records), instead of the default record.
  pick = null,
  // A record update's refusal (an Error), or null to accept it.
  update = () => null,
} = {}) {
  const document = createDocument(html);
  // Timers of 100 ms or more (the autosave and preview pauses) wait for flush(); shorter ones run
  // as usual. Each waiting timer has its own ID, so clearing an old one never clears a new one.
  const timers = new Map();
  let timerId = 0;
  const sent = [];
  const deleted = [];
  const looked = [];
  const profile = {
    WorkerId: 'worker-1',
    Enabled: true,
    CanChange: true,
    RowVersion: '7',
    SharePointHosts: [],
    Registration: {
      Readiness: enabled
        .map((t) => ({ Scope: t, Status: 'Ready' }))
        .concat([{ Scope: 'team', Status: 'Ready' }]),
      Pending: 0,
      Error: null,
    },
    ...runtime,
  };
  const xrm = {
    Navigation: {
      navigateTo: async (page) =>
        sent.push([
          'navigate',
          {
            ...page,
            data: new URLSearchParams(page.webresourceName.split('?')[1] || '').get('data'),
          },
        ]),
      openForm: async () => {},
      openConfirmDialog: async () => assert.fail('No platform dialogs'),
    },
    Utility: {
      getGlobalContext: () => ({
        getClientUrl: () => 'https://example.test',
        userSettings: { userId: '{11111111-1111-1111-1111-111111111111}' },
      }),
      lookupObjects: async (options) => {
        looked.push(options);
        if (pick && !options.allowMultiSelect) return pick(options);
        if (options.allowMultiSelect && pickMany)
          return Array.from({ length: window.AsxdUi.BOUNDS.previewRecords + 1 }, (_, i) => ({
            id: '{' + String(i).padStart(8, '0') + '-0000-0000-0000-000000000000}',
            name: 'Account ' + i,
            entityType: 'account',
          }));
        return options.entityTypes[0] === 'contact'
          ? [{ id: '{' + CONTACT.toUpperCase() + '}', name: 'Jane Smith', entityType: 'contact' }]
          : [{ id: '{' + RECORD.toUpperCase() + '}', name: 'Contoso Ltd', entityType: 'account' }];
      },
    },
    WebApi: {
      retrieveRecord: async (table, id) => {
        const custom = reads(table, id);
        if (custom instanceof Error) throw custom;
        if (custom) return custom;
        if (table === 'asx_revision') return { asx_revisionid: id, asx_version: versions[id] };
        return {
          asx_templateid: id,
          asx_name: 'Account onboarding',
          asx_table: 'account',
          asx_disabled: false,
          _asx_publishedrevisionid_value: 'rev-1',
          ...templates.find((t) => t.asx_templateid === id),
        };
      },
      retrieveMultipleRecords: async (table, options) => ({
        entities:
          rows(table, options) ||
          (table === 'asx_runtimetable'
            ? enabled.map((asx_logicalname) => ({ asx_logicalname }))
            : table === 'asx_template'
              ? templates
              : table === 'asx_revision'
                ? [
                    {
                      asx_revisionid: 'rev-1',
                      asx_version: 1,
                      asx_status: 'Published',
                      _asx_templateid_value: TEMPLATE,
                      modifiedon: '2026-10-03T09:00:00Z',
                      '_modifiedby_value@OData.Community.Display.V1.FormattedValue': 'Dana Reyes',
                    },
                  ]
                : table === 'asx_library'
                  ? libraries || [
                      { asx_libraryid: 'lib-a', asx_name: 'General', _asx_siteid_value: 'site-a' },
                      {
                        asx_libraryid: 'lib-b',
                        asx_name: 'Sensitive',
                        _asx_siteid_value: 'site-b',
                      },
                    ]
                  : table === 'asx_site'
                    ? [
                        { asx_siteid: 'site-a', asx_name: 'Delivery' },
                        { asx_siteid: 'site-b', asx_name: 'Commercial' },
                      ]
                    : []),
      }),
      updateRecord: async (table, id, data) => {
        sent.push(['update', table, data]);
        const refused = update(table, id, data);
        if (refused instanceof Error) throw refused;
      },
      deleteRecord: async (table, id) => deleted.push([table, id]),
      online: {
        execute: async (request) => {
          const api = request.getMetadata().operationName;
          const body = request.Request
            ? JSON.parse(request.Request)
            : { RevisionId: request.RevisionId, RowVersion: request.RowVersion };
          sent.push([api, body]);
          const custom = await handle(api, body);
          if (custom instanceof Error)
            return { ok: false, json: async () => ({ error: { message: custom.message } }) };
          const result =
            custom ||
            (api === 'asx_RuntimeAdmin'
              ? profile
              : api === 'asx_LoadDraft'
                ? loaded
                : api === 'asx_CreateDraft'
                  ? { TemplateId: TEMPLATE, RevisionId: 'rev-2', RowVersion: '4', Status: 'Draft' }
                  : api === 'asx_PreviewTemplate'
                    ? {
                        Folders: [
                          {
                            Section: 'general',
                            Node: 'root',
                            Name: 'Contoso Ltd',
                            RelativePath: 'Contoso Ltd',
                          },
                          {
                            Section: 'general',
                            Node: 'general_docs',
                            Name: 'General',
                            RelativePath: 'Contoso Ltd/General',
                          },
                        ],
                        Notices: [],
                      }
                    : api === 'asx_PublishTemplate'
                      ? { Status: 'Published', Notices: [] }
                      : body.Command === 'Summary'
                        ? { Status: 'Summary', Summary: summary }
                        : body.Command === 'CountRecords'
                          ? { Status: 'Counted', Run: { Total: 1240, TotalEstimated: false } }
                          : body.Command === 'ListProblems'
                            ? { Status: 'Page', Problems: [], Next: null }
                            : body.Command === 'StartTemplateRun'
                              ? {
                                  Status: 'Pending',
                                  Key: 'templaterun:1',
                                  Run: {
                                    State: 'Running',
                                    Planned: 0,
                                    Total: 1240,
                                    TotalEstimated: false,
                                  },
                                }
                              : { Status: 'Pending' });
          return { ok: true, json: async () => ({ Result: JSON.stringify(result) }) };
        },
      },
    },
  };
  const fetch = async (url) => {
    if (url.includes('RetrieveUserPrivilegeByPrivilegeName')) {
      const name = /PrivilegeName='([^']+)'/.exec(url)[1];
      return {
        ok: true,
        json: async () => ({ RolePrivileges: privileges[name] === false ? [] : [{}] }),
      };
    }
    const table = /LogicalName='(\w+)'/.exec(url)?.[1];
    await delay(table, url);
    if (failFetch(url, table))
      return { ok: false, json: async () => ({ error: { message: 'Too many requests.' } }) };
    const all = metadata[table] || [];
    const value = url.includes('DateTimeAttributeMetadata')
      ? [
          { LogicalName: 'closedate', DateTimeBehavior: { Value: 'DateOnly' } },
          { LogicalName: 'reviewedon', DateTimeBehavior: { Value: 'UserLocal' } },
        ]
      : url.includes('CRM.PicklistAttributeMetadata')
        ? [
            {
              LogicalName: 'industrycode',
              OptionSet: {
                Options: [
                  { Value: 1, Label: { UserLocalizedLabel: { Label: 'Government' } } },
                  { Value: 2, Label: { UserLocalizedLabel: { Label: 'Retail' } } },
                ],
              },
            },
          ]
        : url.includes('StateAttributeMetadata')
          ? [
              {
                LogicalName: 'statecode',
                OptionSet: {
                  Options: [
                    { Value: 0, Label: { UserLocalizedLabel: { Label: 'Active' } } },
                    { Value: 1, Label: { UserLocalizedLabel: { Label: 'Inactive' } } },
                  ],
                },
              },
            ]
          : url.includes('LookupAttributeMetadata')
            ? table === 'account'
              ? [
                  {
                    LogicalName: 'primarycontactid',
                    Targets: ['contact'],
                    DisplayName: { UserLocalizedLabel: { Label: 'Primary Contact' } },
                  },
                  ...extraLookups,
                ]
              : []
            : url.includes('AttributeMetadata')
              ? []
              : url.includes('/Attributes')
                ? all
                : [
                    {
                      LogicalName: 'account',
                      EntitySetName: 'accounts',
                      PrimaryIdAttribute: 'accountid',
                      PrimaryNameAttribute: 'name',
                      DisplayName: { UserLocalizedLabel: { Label: 'Account' } },
                      DisplayCollectionName: { UserLocalizedLabel: { Label: 'Accounts' } },
                    },
                    {
                      LogicalName: 'contact',
                      EntitySetName: 'contacts',
                      PrimaryIdAttribute: 'contactid',
                      PrimaryNameAttribute: 'fullname',
                      DisplayName: { UserLocalizedLabel: { Label: 'Contact' } },
                      DisplayCollectionName: { UserLocalizedLabel: { Label: 'Contacts' } },
                    },
                  ];
    return { ok: true, json: async () => ({ value }) };
  };
  const session = new Map([['asxd.launched', '1']]);
  const window = {
    parent: { Xrm: xrm },
    location: { search: '?data=templates-' + BUILD, hash },
    sessionStorage: {
      getItem: (k) => session.get(k) ?? null,
      setItem: (k, v) => session.set(k, v),
      removeItem: (k) => session.delete(k),
    },
  };
  const context = vm.createContext({
    window,
    document,
    fetch,
    Intl,
    URL,
    URLSearchParams,
    console,
    navigator: { clipboard: { writeText: async () => {} } },
    crypto: require('node:crypto').webcrypto,
    structuredClone,
    setTimeout: (fn, ms) => {
      if (ms < 100) return setTimeout(fn, ms);
      timers.set(++timerId, fn);
      return timerId;
    },
    clearTimeout: (id) => {
      if (timers.has(id)) timers.delete(id);
      else if (typeof id !== 'number') clearTimeout(id);
    },
    setInterval: () => 0,
    clearInterval: () => {},
  });
  for (const name of ['shell.js', 'admin.js']) vm.runInContext(read(name), context);
  await document.fire('DOMContentLoaded');
  document.defaultBounds = window.AsxdUi.BOUNDS;
  const $ = (id) => document.getElementById(id);
  const press = async (node) => {
    node.click();
    await document.settle();
  };
  // A button by its text and, when given, its aria-label.
  const find = (root, text, label = null) =>
    [...root.querySelectorAll('button')].find(
      (b) => b.textContent === text && (label === null || b.getAttribute('aria-label') === label),
    );
  const labelled = (root, name) =>
    [...root.querySelectorAll('input, select, textarea')].find(
      (n) =>
        n.getAttribute('aria-label') === name ||
        n.parentNode?._text === name ||
        n.parentNode?.firstElementChild?.textContent === name,
    );
  const change = async (node, value) => {
    node.value = value;
    await node.onchange?.();
    await node.oninput?.();
    await document.settle();
  };
  // A templates list row by its name (the row's whole text is its name and its state).
  const row = (name) =>
    [...$('template-groups').querySelectorAll('button')].find(
      (b) => b.classList.contains('list-row') && b.querySelector('.row-name')?.textContent === name,
    );
  const overview = async (name = 'Account onboarding') => press(row(name));
  // A step 2 tree row by its folder's shown name (chips read as their labels), and selecting it.
  const node = (name) =>
    [...$('folder-tree').querySelectorAll('.tree-row')].find(
      (r) =>
        !r.classList.contains('is-removed') && r.querySelector('.node .name')?.visibleText === name,
    );
  const select = async (name) => press(node(name).querySelector('.node'));
  const open = async (name = 'Account onboarding') => {
    await overview(name);
    await press($('overview-edit'));
  };
  return {
    document,
    $,
    sent,
    deleted,
    looked,
    press,
    find,
    labelled,
    change,
    row,
    overview,
    open,
    window,
    last: (api) => sent.filter(([a]) => a === api).at(-1)?.[1],
    // Runs the queued timers, and the ones they queue, until none is left.
    flush: async () => {
      while (timers.size) {
        const due = [...timers.values()];
        timers.clear();
        for (const fn of due) fn();
        await document.settle();
      }
    },
    step: (n) => press($('step-tab-' + n)),
    // admin.js state, for checks the page does not show.
    state: () => window.AsxdAdmin.state(),
    node,
    select,
    // Renames a folder in step 2.
    rename: async (folder, text) => {
      await select(folder);
      await change($('folder-name'), text);
    },
    // Create this folder: 0 Always, 1 Only when….
    mode: (n) => press($('create-mode').querySelectorAll('[role=radio]')[n]),
    // An edit autosave cannot save: a folder renamed "Dropped" under a condition whose number
    // the server cannot read.
    unsaveable: async () => {
      await select('General');
      await change($('folder-name'), 'Dropped');
      await press($('create-mode').querySelectorAll('[role=radio]')[1]);
      await change(labelled($('conditions'), 'Field, condition 1'), 'root.revenue');
      await change(labelled($('conditions'), 'Value, condition 1'), '1e5');
    },
  };
}

(async () => {
  {
    // Empty states: no tables yet in the list, and no template selected in the overview.
    const t = await boot({ enabled: [], templates: [] });
    assert.match(t.$('template-groups').visibleText, /No tables yet/);
    assert.equal(t.$('template-editor').hidden, true, 'The editor waits for a template');
    assert.equal(t.$('overview-title').textContent, 'Folder templates');
    assert.match(t.$('overview-cards').visibleText, /No template selected/);
    assert.equal(t.$('new-template').hidden, true, 'No ＋ New without a table');
  }
  {
    // Tables are managed in Settings: the list has no Add table, Remove or Enable, and links there.
    const t = await boot({ enabled: [], templates: [] });
    assert.equal(t.$('addTable') === null, true, 'No Add table in the list');
    assert.equal(t.$('tablePicker') === null, true, 'No table picker in the list');
    await t.press(t.$('manage-tables'));
    assert.equal(
      t.sent
        .filter(([k]) => k === 'navigate')
        .at(-1)[1]
        .data.split('-')[0],
      'settings',
    );
  }
  {
    // Edit template opens the editor: Publish explains why it is off; no IDs in the page.
    const t = await boot();
    await t.open();
    assert.equal(t.$('step-1').hidden, false);
    assert.equal(t.$('editor-title').textContent, 'Account onboarding');
    assert.doesNotMatch(t.$('templates').visibleText, GUID);
    assert.doesNotMatch(t.$('templates').visibleText, /rev-1/);
    assert.equal(
      t.node('Account Name').querySelector('.token').textContent,
      'Account Name',
      'Tokens read as field labels',
    );
    assert.equal(t.$('publish').getAttribute('aria-disabled'), 'true');
    assert.equal(t.$('publish-reason').textContent, 'Already published');
  }
  {
    // The list groups templates by table with their state; search filters by name.
    const t = await boot({
      enabled: ['account', 'contact'],
      templates: [
        {
          asx_templateid: TEMPLATE,
          asx_name: 'Account onboarding',
          asx_table: 'account',
          _asx_publishedrevisionid_value: 'rev-1',
          asx_disabled: false,
        },
        {
          asx_templateid: OTHER,
          asx_name: 'Key accounts',
          asx_table: 'account',
          _asx_publishedrevisionid_value: null,
          asx_disabled: false,
        },
        {
          asx_templateid: 'tpl-c',
          asx_name: 'Contact files',
          asx_table: 'contact',
          _asx_publishedrevisionid_value: 'rev-c',
          asx_disabled: true,
        },
      ],
      rows: (table) =>
        table === 'asx_revision'
          ? [
              {
                asx_revisionid: 'rev-1',
                asx_version: 1,
                asx_status: 'Published',
                _asx_templateid_value: TEMPLATE,
              },
              {
                asx_revisionid: 'rev-k',
                asx_version: 1,
                asx_status: 'Draft',
                _asx_templateid_value: OTHER,
              },
              {
                asx_revisionid: 'rev-c',
                asx_version: 2,
                asx_status: 'Published',
                _asx_templateid_value: 'tpl-c',
              },
            ]
          : null,
    });
    const groups = [...t.$('template-groups').querySelectorAll('.list-group')];
    assert.deepEqual(
      groups.map((g) => g.querySelector('.group-label').textContent),
      ['Account', 'Contact'],
    );
    const states = [...t.$('template-groups').querySelectorAll('.list-row')].map((r) => [
      r.querySelector('.row-name').textContent,
      r.querySelector('.row-state').textContent,
    ]);
    assert.deepEqual(states, [
      ['Account onboarding', 'Live v1'],
      ['Key accounts', 'Draft v1'],
      ['Contact files', 'Off'],
    ]);
    assert.deepEqual(
      [...t.$('template-groups').querySelectorAll('.row-state')].map((s) =>
        s.getAttribute('data-tone'),
      ),
      ['ok', 'warning', 'muted'],
    );
    t.$('template-search').value = 'key';
    t.$('template-search').oninput();
    assert.deepEqual(
      [...t.$('template-groups').querySelectorAll('.list-row')].map(
        (r) => r.querySelector('.row-name').textContent,
      ),
      ['Key accounts'],
    );
  }
  {
    // Overview: title, pill and meta (no record count); cards with chips and rule sentences.
    const conditional = draft();
    conditional.Sources[0].Columns.push({ Name: 'statecode', Kind: 'Choice' });
    conditional.Destinations[0].Folders.push({
      Key: 'tenders',
      Parent: 'root',
      Name: 'Public tenders',
      Condition: {
        All: true,
        Groups: [],
        Conditions: [{ Source: 'root', Column: 'statecode', Operator: 'Equal', Literal: '0' }],
      },
    });
    const t = await boot({
      loaded: {
        RevisionId: 'rev-1',
        RowVersion: '3',
        Status: 'Published',
        Version: 1,
        Draft: conditional,
      },
    });
    await t.overview();
    assert.equal(t.$('overview-title').textContent, 'Account onboarding');
    assert.equal(t.$('overview-pill').textContent, 'Live v1');
    assert.match(t.$('overview-meta').textContent, /^Account table · published .+ by Dana Reyes$/);
    assert.doesNotMatch(t.$('overview-meta').textContent, /records/);
    assert.equal(t.row('Account onboarding').getAttribute('aria-current'), 'true');
    const cards = [...t.$('overview-cards').querySelectorAll('.section-card')];
    assert.deepEqual(
      cards.map((c) => c.querySelector('h2').textContent),
      ['Destinations', 'Folders', 'Schedule and runs', 'Versions'],
    );
    // ui.card: a head with the h2, the summary and the action, then a body.
    for (const card of cards) {
      assert.equal(card.tagName, 'SECTION');
      assert.ok(card.querySelector('.card-head h2'));
      assert.ok(card.querySelector('.card-body'));
    }
    assert.equal(cards[3].querySelector('.card-head button').textContent, 'All versions');
    assert.equal(cards[0].querySelector('.card-head .muted').textContent, '1 library');
    const dest = cards[0].querySelector('.dest-row');
    assert.equal(dest.querySelector('strong').textContent, 'Business documents');
    assert.equal(dest.querySelector('.muted').textContent, 'Delivery › General');
    assert.equal(dest.querySelector('.teams').textContent, '0 teams');
    assert.equal(
      cards[1].querySelector('.card-head .muted').textContent,
      '3 folders · 1 conditional',
    );
    const folderRows = cards[1].querySelectorAll('.folder-row');
    assert.equal(folderRows[0].querySelector('.token').textContent, 'Account Name');
    assert.equal(folderRows[0].querySelector('.rule').textContent, 'Always');
    assert.equal(folderRows[2].querySelector('.rule').textContent, '◆ When Status is Active');
    assert.ok(folderRows[2].querySelector('.rule').classList.contains('conditional'));
    assert.equal(cards[2].querySelector('.card-head .muted').textContent, 'On · no end date');
    assert.equal(t.$('overview-edit').textContent, 'Edit template');
    assert.doesNotMatch(t.$('templates').visibleText, GUID);
    // A destination's folders collapse to a count; the first destination starts open.
    const group = cards[1].querySelector('.group-row');
    assert.equal(group.getAttribute('aria-expanded'), 'true');
    group.focus();
    await t.press(group);
    const again = t.$('overview-cards').querySelectorAll('.section-card')[1];
    assert.equal(again.querySelectorAll('.folder-row').length, 0);
    assert.equal(again.querySelector('.group-row').getAttribute('aria-expanded'), 'false');
    assert.match(again.querySelector('.group-row').textContent, /3 folders$/);
    assert.equal(t.document.activeElement, again.querySelector('.group-row'));
  }
  {
    // Versions: Live, Replaced and Draft from the status and the published pointer.
    const t = await boot({
      templates: [
        {
          asx_templateid: TEMPLATE,
          asx_name: 'Account onboarding',
          asx_table: 'account',
          _asx_publishedrevisionid_value: 'rev-3',
          asx_disabled: false,
        },
      ],
      rows: (table) =>
        table === 'asx_revision'
          ? [
              {
                asx_revisionid: 'rev-4',
                asx_version: 4,
                asx_status: 'Draft',
                _asx_templateid_value: TEMPLATE,
              },
              {
                asx_revisionid: 'rev-3',
                asx_version: 3,
                asx_status: 'Published',
                _asx_templateid_value: TEMPLATE,
              },
              {
                asx_revisionid: 'rev-2',
                asx_version: 2,
                asx_status: 'Published',
                _asx_templateid_value: TEMPLATE,
              },
            ]
          : null,
    });
    await t.overview();
    assert.equal(t.$('overview-pill').textContent, 'Live v3 · Draft v4');
    assert.equal(t.$('overview-edit').textContent, 'Continue Draft v4');
    const versions = [...t.$('overview-cards').querySelectorAll('.version-row')];
    assert.deepEqual(
      versions.map((v) => [
        v.querySelector('.v').textContent,
        v.querySelector('.state').textContent,
      ]),
      [
        ['v4', 'Draft'],
        ['v3', 'Live'],
        ['v2', 'Replaced'],
      ],
    );
  }
  {
    // Schedule opens as a side panel from its card; Delete is in ⋯; Manage tables goes to Settings.
    const t = await boot();
    await t.overview();
    const schedule = t.$('overview-cards').querySelectorAll('.section-card')[2];
    await t.press(schedule.querySelector('.card-head button'));
    assert.equal(t.$('schedule-panel').hidden, false);
    assert.equal(t.$('schedule-panel').getAttribute('role'), 'dialog');
    assert.equal(t.document.activeElement, t.$('schedule-title'));
    t.$('schedule-title').key('Escape');
    assert.equal(t.$('schedule-panel').hidden, true);
    await t.press(t.$('overview-menu'));
    assert.deepEqual(
      [...t.$('overview-menu-list').querySelectorAll('[role=menuitem]')].map((i) =>
        i.textContent.trim(),
      ),
      ['Re-run for existing records…', 'Delete template'],
    );
    await t.press(t.$('manage-tables'));
    assert.equal(
      t.sent
        .filter(([k]) => k === 'navigate')
        .at(-1)[1]
        .data.split('-')[0],
      'settings',
    );
  }
  {
    // Edit template opens the editor; Close returns to the overview with focus on Edit.
    const t = await boot();
    await t.overview();
    await t.press(t.$('overview-edit'));
    assert.equal(t.$('template-overview').hidden, true);
    assert.equal(t.$('template-editor').hidden, false);
    assert.equal(t.document.activeElement, t.$('editor-title'));
    await t.press(t.$('editor-close'));
    assert.equal(t.$('template-overview').hidden, false);
    assert.equal(t.$('template-editor').hidden, true);
    assert.equal(t.document.activeElement, t.$('overview-edit'));
    // Close with edits that cannot be saved asks first, with the page's unsaved-changes prompt;
    // Discard changes closes, and Edit template opens the saved version again.
    await t.press(t.$('overview-edit'));
    await t.unsaveable();
    await t.press(t.$('editor-close'));
    assert.match(
      t.$('leavePrompt').visibleText,
      /You have unsaved changes to Account onboarding\./,
    );
    assert.equal(t.$('template-editor').hidden, false, 'Nothing closes before the answer');
    await t.press(t.find(t.$('leavePrompt'), 'Discard changes'));
    assert.equal(t.$('template-overview').hidden, false);
    assert.equal(t.last('asx_CreateDraft'), undefined, 'Discard saves nothing');
    await t.press(t.$('overview-edit'));
    assert.doesNotMatch(t.$('folder-tree').visibleText, /Dropped/);
    assert.equal(t.$('editor-pill').textContent, 'Draft v2');
    await t.flush();
    assert.equal(t.last('asx_CreateDraft'), undefined, 'The discarded edit is never saved');
  }
  {
    // Without the Operator role the Last re-run row is left out; without the Security
    // Administrator role so are team counts, and GetPolicy is never called.
    const t = await boot({
      privileges: { prvCreateasx_operatorcommand: false, prvCreateasx_policy: false },
      summary: { BlockedJobs: 3 },
    });
    await t.overview();
    assert.doesNotMatch(t.$('overview-cards').visibleText, /Last re-run/);
    assert.doesNotMatch(t.$('overview-cards').visibleText, /\d teams?\b/);
    assert.equal(t.sent.filter(([, b]) => b?.Command === 'GetPolicy').length, 0);
    assert.equal(t.sent.filter(([, b]) => b?.Command === 'ListProblems').length, 0);
    assert.equal(t.$('overview-problems').querySelector('.problem-pill').hidden, true);
    assert.equal(t.sent.filter(([, b]) => b?.Command === 'Summary').length, 0);
  }
  {
    // With both roles: the last re-run of this template from TemplateRuns, and team counts from
    // one GetPolicy per library per page load.
    const t = await boot({
      handle: (api, b) =>
        b.Command === 'ListProblems'
          ? {
              Status: 'Page',
              Problems: [
                {
                  Key: 'templaterun:2',
                  Kind: 'TemplateRun',
                  Run: { TemplateId: OTHER, State: 'Done', Planned: 5, Total: 5 },
                },
                {
                  Key: 'templaterun:1',
                  Kind: 'TemplateRun',
                  Run: {
                    TemplateId: TEMPLATE.toUpperCase(),
                    State: 'Done',
                    Planned: 1284,
                    Total: 1284,
                    TotalEstimated: false,
                    StartedUtc: '2026-10-01T10:00:00Z',
                  },
                },
              ],
              Next: null,
            }
          : b.Command === 'GetPolicy'
            ? {
                Status: 'Applied',
                Policy: {
                  Desired: [
                    { TeamId: 'team-1', Access: 'Contribute' },
                    { TeamId: 'team-2', Access: 'Read' },
                    { TeamId: 'team-3', Access: 'None' },
                  ],
                  Applied: [],
                },
                Teams: [],
              }
            : null,
    });
    await t.overview();
    const run = t.$('overview-cards').querySelector('.run-row');
    assert.match(run.textContent, /^Last re-run.+ · 1,284 of 1,284 records$/);
    assert.equal(t.$('overview-cards').querySelector('.teams').textContent, '2 teams');
    await t.overview();
    assert.equal(
      t.sent.filter(([, b]) => b?.Command === 'GetPolicy').length,
      1,
      'The policy is read once per page load',
    );
  }
  {
    // The overview header links to Monitor with the problem count (owner decision 4); an
    // overview redraw keeps the pill and makes no second Summary call.
    const t = await boot({ summary: { BlockedJobs: 2, WaitingRecords: 1, TemplateRuns: 4 } });
    await t.overview();
    const pill = t.$('overview-problems').querySelector('.problem-pill');
    assert.equal(pill.hidden, false);
    assert.equal(pill.textContent, 'Monitor · 3 problems');
    await t.overview();
    assert.equal(t.$('overview-problems').querySelector('.problem-pill'), pill);
    assert.equal(t.sent.filter(([, b]) => b?.Command === 'Summary').length, 1);
    await t.press(pill);
    assert.equal(
      t.sent
        .filter(([k]) => k === 'navigate')
        .at(-1)[1]
        .data.split('-')[0],
      'monitor',
    );
  }
  {
    // View opens that version read-only, not the latest revision over it.
    const t = await boot({
      handle: (api, body) =>
        api === 'asx_LoadDraft'
          ? {
              RevisionId: body.RevisionId,
              RowVersion: '3',
              Status: body.RevisionId === 'rev-4' ? 'Draft' : 'Published',
              Version: Number(body.RevisionId.split('-')[1]),
              Draft: draft(),
            }
          : null,
      templates: [
        {
          asx_templateid: TEMPLATE,
          asx_name: 'Account onboarding',
          asx_table: 'account',
          _asx_publishedrevisionid_value: 'rev-3',
          asx_disabled: false,
        },
      ],
      rows: (table) =>
        table === 'asx_revision'
          ? [
              {
                asx_revisionid: 'rev-4',
                asx_version: 4,
                asx_status: 'Draft',
                _asx_templateid_value: TEMPLATE,
              },
              {
                asx_revisionid: 'rev-3',
                asx_version: 3,
                asx_status: 'Published',
                _asx_templateid_value: TEMPLATE,
              },
              {
                asx_revisionid: 'rev-2',
                asx_version: 2,
                asx_status: 'Published',
                _asx_templateid_value: TEMPLATE,
              },
            ]
          : null,
    });
    await t.overview();
    await t.press(t.find(t.$('overview-cards'), 'View', 'View v2'));
    const loads = t.sent.filter(([a]) => a === 'asx_LoadDraft').map(([, b]) => b.RevisionId);
    assert.equal(loads.at(-1), 'rev-2', 'The viewed version is loaded last');
    assert.equal(t.$('template-editor').hidden, false);
    // Read-only: the pill names the version and its state; nothing can be edited or published.
    assert.equal(t.$('editor-pill').textContent, 'v2 · Replaced');
    assert.equal(t.$('editor-pill').getAttribute('data-tone'), 'muted');
    assert.equal(t.$('editor-note').hidden, true);
    assert.equal(t.$('publish').hidden, true);
    assert.equal(t.$('add-destination').hidden, true);
    assert.ok(t.labelled(t.$('step-1'), 'Name').disabled);
    assert.ok(t.$('folder-name').disabled);
    assert.equal(t.node('General').querySelector('.node').disabled, false, 'Folders can be opened');
    assert.equal(t.$('create-mode').querySelectorAll('[role=radio]')[1].disabled, true);
  }
  {
    // A related record's field reads through its lookup, as a chip; a lookup condition names the
    // record; one of several conditions reads "any of".
    const d = draft();
    d.Sources[0].Columns.push({ Name: 'primarycontactid', Kind: 'Lookup' });
    d.Sources.push({
      Alias: 'lookup_1',
      Table: 'contact',
      Lookup: 'primarycontactid',
      Columns: [{ Name: 'fullname', Kind: 'Text' }],
    });
    const folders = d.Destinations[0].Folders;
    folders[1].Name = 'Contact {lookup_1.fullname}';
    folders[1].Condition = {
      All: true,
      Groups: [],
      Conditions: [
        {
          Source: 'root',
          Column: 'primarycontactid',
          Operator: 'Equal',
          LiteralKind: 'Lookup',
          Literal: CONTACT,
          LiteralLabel: 'Jane Smith',
          LiteralTable: 'contact',
        },
      ],
    };
    folders.push({
      Key: 'either',
      Parent: 'root',
      Name: 'Either',
      Condition: {
        All: false,
        Groups: [],
        Conditions: [
          { Source: 'root', Column: 'name', Operator: 'IsNull' },
          { Source: 'lookup_1', Column: 'fullname', Operator: 'IsNotNull' },
        ],
      },
    });
    const t = await boot({
      loaded: { RevisionId: 'rev-1', RowVersion: '3', Status: 'Published', Version: 1, Draft: d },
    });
    await t.overview();
    const rows = [...t.$('overview-cards').querySelectorAll('.folder-row')];
    assert.equal(rows[1].querySelector('.token').textContent, 'Primary Contact › Full Name');
    assert.equal(
      rows[1].querySelector('.rule').textContent,
      '◆ When Primary Contact is Jane Smith',
    );
    assert.equal(rows[2].querySelector('.rule').textContent, '◆ When any of 2…');
    assert.doesNotMatch(t.$('overview-cards').visibleText, GUID);
  }
  {
    // A link to a template opens on its overview.
    const t = await boot({
      hash: '#templates?template=' + OTHER,
      templates: [
        { asx_templateid: TEMPLATE, asx_name: 'Account onboarding', asx_table: 'account' },
        { asx_templateid: OTHER, asx_name: 'Contract documents', asx_table: 'account' },
      ],
    });
    assert.equal(t.$('overview-title').textContent, 'Contract documents');
    assert.equal(t.row('Contract documents').getAttribute('aria-current'), 'true');
    // Without one, the first template listed.
    const first = await boot();
    assert.equal(first.$('overview-title').textContent, 'Account onboarding');
  }
  {
    // Fix round 1: a row chosen with the keyboard keeps focus once the list redraws.
    const t = await boot({
      templates: [
        { asx_templateid: TEMPLATE, asx_name: 'Account onboarding', asx_table: 'account' },
        { asx_templateid: OTHER, asx_name: 'Contract documents', asx_table: 'account' },
      ],
    });
    const before = t.row('Contract documents');
    before.focus();
    await t.press(before);
    assert.equal(t.$('overview-title').textContent, 'Contract documents');
    assert.ok(
      t.document.activeElement === t.row('Contract documents'),
      'Focus is on the chosen row',
    );
    assert.ok(t.document.activeElement.isConnected, 'Focus is on a row in the page');
  }
  {
    // Fix round 1: a re-run still under way reads "Re-run in progress", an ended one "Last re-run".
    const run = (State) => ({
      Status: 'Page',
      Problems: [
        {
          Key: 'templaterun:1',
          Kind: 'TemplateRun',
          Run: {
            TemplateId: TEMPLATE,
            State,
            Planned: 500,
            Total: 1240,
            StartedUtc: '2026-10-01T10:00:00Z',
          },
        },
      ],
      Next: null,
    });
    for (const [State, label] of [
      ['Running', 'Re-run in progress'],
      ['Paused', 'Re-run in progress'],
      ['Done', 'Last re-run'],
    ]) {
      const t = await boot({
        handle: (api, b) => (b.Command === 'ListProblems' ? run(State) : null),
      });
      await t.overview();
      const row = t.$('overview-cards').querySelector('.run-row');
      assert.equal(row.querySelector('span').textContent, label, State);
      assert.match(row.textContent, / · 500 of 1,240 records$/);
    }
  }
  {
    // Fix round 1: the ⋯ menu's separator hides with Re-run for existing records.
    const draftOnly = await boot({
      templates: [
        {
          asx_templateid: TEMPLATE,
          asx_name: 'Account onboarding',
          asx_table: 'account',
          _asx_publishedrevisionid_value: null,
        },
      ],
    });
    await draftOnly.overview();
    const separator = (t) => t.$('overview-menu-list').querySelector('[role=separator]');
    assert.equal(draftOnly.$('menu-rerun').hidden, true);
    assert.equal(separator(draftOnly).hidden, true);
    const live = await boot();
    await live.overview();
    assert.equal(live.$('menu-rerun').hidden, false);
    assert.equal(separator(live).hidden, false);
  }
  {
    // Fix round 1: arrow keys move through the table picker without opening a template; leaving
    // it forgets that, so a later pick opens one; Escape hides it and returns to ＋ New.
    const t = await boot({ enabled: ['account', 'contact'] });
    const picker = t.$('new-template-table');
    await t.press(t.$('new-template'));
    picker.key('ArrowDown');
    await t.change(picker, 'account');
    assert.equal(t.$('template-editor').hidden, true, 'An arrow key opens nothing');
    assert.equal(picker.hidden, false);
    picker.key('Escape');
    assert.equal(picker.hidden, true);
    assert.ok(t.document.activeElement === t.$('new-template'), 'Focus returns to ＋ New');
    await t.press(t.$('new-template'));
    picker.key('ArrowDown');
    picker.dispatchEvent(new FakeEvent('focusout'));
    await t.change(picker, 'contact');
    assert.equal(t.$('template-editor').hidden, false, 'A pick after leaving the picker opens');
    assert.equal(t.$('templateName').hidden, false);
    assert.equal(t.$('templateName').value, '', 'A new template starts without a name');
    assert.equal(t.$('editor-table').textContent, 'Contact ›');
    // Enter opens the table the picker shows.
    const e = await boot({ enabled: ['account', 'contact'] });
    await e.press(e.$('new-template'));
    e.$('new-template-table').key('ArrowDown');
    e.$('new-template-table').value = 'contact';
    e.$('new-template-table').key('Enter');
    await e.document.settle();
    assert.equal(e.$('template-editor').hidden, false);
  }
  {
    // Fix round 1: ＋ New with edits that cannot be saved asks first; Keep editing keeps them.
    const t = await boot();
    await t.open();
    await t.unsaveable();
    await t.press(t.$('new-template'));
    assert.equal(
      t.document.activeElement.textContent,
      'Open a new template? Your unsaved changes to Account onboarding are discarded.',
    );
    await t.press(t.find(t.$('templates-list'), 'Keep editing'));
    assert.equal(t.$('editor-pill').textContent, 'Draft v2');
    assert.equal(t.$('editor-name').textContent, 'Account onboarding');
    await t.press(t.$('new-template'));
    await t.press(t.find(t.$('templates-list'), 'Open a new template'));
    assert.equal(t.$('templateName').value, '');
    assert.equal(t.$('editor-pill').textContent, 'Draft v1');
    assert.equal(t.$('editor-note').hidden, true, 'A new template has no live version');
    await t.flush();
    assert.equal(t.last('asx_CreateDraft'), undefined, 'Nothing was saved');
  }
  {
    // ＋ New: with more than one table it asks which; the choice opens a new template's editor.
    const t = await boot({ enabled: ['account', 'contact'] });
    await t.press(t.$('new-template'));
    const table = t.$('new-template-table');
    assert.equal(table.hidden, false);
    assert.deepEqual(
      [...table.querySelectorAll('option')].map((o) => o.textContent),
      ['Choose a table', 'Account', 'Contact'],
    );
    await t.change(table, 'contact');
    assert.equal(table.hidden, true);
    assert.equal(t.$('template-editor').hidden, false);
    assert.equal(t.$('templateName').value, '', 'A new template starts without a name');
    assert.equal(t.$('templateName').hidden, false, 'A new template is named in the heading');
  }
  {
    // An edit enables Publish, which saves it first.
    const t = await boot();
    await t.open();
    await t.rename('General', 'General documents');
    assert.equal(t.$('editor-pill').textContent, 'Draft v2');
    assert.equal(t.$('publish').hasAttribute('aria-disabled'), false);
    // Leaving the page with unsaved edits asks first (the shell owns the prompt).
    t.document.track(t.window.AsxdUi.navigate('monitor'));
    await t.document.settle();
    assert.match(
      t.$('leavePrompt').visibleText,
      /You have unsaved changes to Account onboarding\./,
    );
    // The autosave, then Publish on step 3, without another confirmation.
    await t.press(t.find(t.$('leavePrompt'), 'Stay'));
    await t.flush();
    assert.equal(t.last('asx_CreateDraft').Destinations[0].Folders[1].Name, 'General documents');
    assert.equal(t.$('editor-pill').textContent, 'Draft v2');
    await t.step(3);
    await t.press(t.$('publish'));
    assert.deepEqual(t.last('asx_PublishTemplate'), { RevisionId: 'rev-2', RowVersion: '4' });
    assert.match(t.$('fb-templates').textContent, /^Published v2\./);
    // Publishing ends the editing session on the overview.
    assert.equal(t.$('template-overview').hidden, false);
    assert.equal(t.state().root, null, 'The next Edit template loads the published version');
  }
  {
    // Publish with an edit not saved yet saves it first, then publishes what was saved.
    const t = await boot();
    await t.open();
    await t.rename('General', 'Unsaved');
    await t.step(3);
    await t.press(t.$('publish'));
    assert.equal(t.last('asx_CreateDraft').Destinations[0].Folders[1].Name, 'Unsaved');
    assert.deepEqual(t.last('asx_PublishTemplate'), { RevisionId: 'rev-2', RowVersion: '4' });
    assert.deepEqual(
      t.sent.map(([a]) => a).filter((a) => /CreateDraft|PublishTemplate/.test(a)),
      ['asx_CreateDraft', 'asx_PublishTemplate'],
    );
  }
  {
    // The links to Sites & access save the edits first and go without asking; they ask only
    // when the edits cannot be saved, and Save draft then says why and stays. A page change
    // with Discard changes goes without saving.
    const t = await boot();
    await t.open();
    await t.rename('General', 'Renamed');
    await t.press(t.find(t.$('step-1'), 'Change in Sites & access'));
    assert.equal(t.$('leavePrompt').textContent, '', 'Saved, so nothing to ask');
    assert.equal(t.last('asx_CreateDraft').Destinations[0].Folders[1].Name, 'Renamed');
    assert.equal(t.sent.filter(([k]) => k === 'navigate').at(-1)[1].data, 'access-' + BUILD);
    // The Library list's setup option saves first too.
    const l = await boot();
    await l.open();
    await l.rename('General', 'Set up');
    await l.change(l.labelled(l.$('step-1'), 'Library'), '__setup');
    assert.equal(l.$('leavePrompt').textContent, '');
    assert.equal(l.last('asx_CreateDraft').Destinations[0].Folders[1].Name, 'Set up');
    assert.equal(l.sent.filter(([k]) => k === 'navigate').at(-1)[1].data, 'access-' + BUILD);
    // An edit that cannot be saved asks first.
    const u = await boot();
    await u.open();
    await u.unsaveable();
    await u.press(u.find(u.$('step-1'), 'Change in Sites & access'));
    assert.match(
      u.$('leavePrompt').visibleText,
      /You have unsaved changes to Account onboarding\./,
    );
    assert.equal(u.sent.filter(([k]) => k === 'navigate').length, 0, 'Nothing leaves yet');
    assert.equal(u.last('asx_CreateDraft'), undefined);
    await u.press(u.find(u.$('leavePrompt'), 'Discard changes'));
    assert.equal(u.sent.filter(([k]) => k === 'navigate').at(-1)[1].data, 'access-' + BUILD);
    const d = await boot();
    await d.open();
    await d.rename('General', 'Dropped');
    d.document.track(d.window.AsxdUi.navigate('settings'));
    await d.document.settle();
    await d.press(d.find(d.$('leavePrompt'), 'Discard changes'));
    assert.equal(d.last('asx_CreateDraft'), undefined, 'Discard saves nothing');
    assert.equal(d.sent.filter(([k]) => k === 'navigate').at(-1)[1].data, 'settings-' + BUILD);
  }
  {
    // Paused automation shows a warning on step 3; no Publisher role disables Publish.
    const paused = await boot({
      runtime: { Enabled: false },
      loaded: { RevisionId: 'rev-2', RowVersion: '4', Status: 'Draft', Version: 2, Draft: draft() },
    });
    await paused.open();
    await paused.step(3);
    assert.equal(
      paused.$('publish-paused').textContent,
      "Automation is paused, so no folders are created until it's on.",
    );
    const role = await boot({
      privileges: { prvCreateasx_publication: false },
      loaded: { RevisionId: 'rev-2', RowVersion: '4', Status: 'Draft', Version: 2, Draft: draft() },
    });
    await role.open();
    assert.equal(role.$('publish-reason').textContent, 'Needs the Documents Publisher role.');
  }
  {
    // The overview's ⋯ menu names its template; Delete template confirms in the page.
    const t = await boot();
    await t.overview();
    const menu = t.$('overview-menu');
    assert.equal(menu.getAttribute('aria-haspopup'), 'menu');
    assert.equal(
      menu.getAttribute('aria-label'),
      'More actions for Account onboarding',
      'The ⋯ button names its template',
    );
    menu.key('ArrowDown');
    await t.document.settle();
    assert.equal(t.$('overview-menu-list').hidden, false);
    assert.equal(t.document.activeElement.id, 'menu-rerun');
    t.document.activeElement.key('ArrowDown');
    assert.equal(t.document.activeElement.id, 'menu-delete');
    t.document.activeElement.key('Escape');
    assert.equal(t.$('overview-menu-list').hidden, true);
    assert.equal(t.document.activeElement, menu);
    await t.press(menu);
    await t.press(t.$('menu-delete'));
    assert.equal(
      t.document.activeElement.textContent,
      'Delete Account onboarding and all its versions? No new folder work starts for it. Folders, documents and access in SharePoint stay as they are. Work already sent to SharePoint may still finish.',
    );
    await t.press(t.find(t.$('overview-header').querySelector('.confirm'), 'Delete template'));
    assert.deepEqual(t.deleted, [['asx_template', TEMPLATE]]);
    // Schedule: Save schedule writes the template row, says so, and the card follows it.
    let saved = false;
    const s = await boot({
      reads: (table, id) =>
        table === 'asx_template' && saved
          ? {
              asx_templateid: id,
              asx_name: 'Account onboarding',
              asx_table: 'account',
              asx_disabled: true,
              _asx_publishedrevisionid_value: 'rev-1',
            }
          : null,
    });
    await s.overview();
    const edit = s.$('overview-cards').querySelectorAll('.section-card')[2].querySelector('button');
    assert.equal(edit.getAttribute('aria-label'), 'Edit schedule');
    await s.press(edit);
    assert.equal(s.$('schedule-panel').hidden, false);
    assert.equal(s.$('schedule-on').getAttribute('aria-checked'), 'true');
    await s.press(s.$('schedule-on'));
    saved = true;
    await s.press(s.$('saveAvailability'));
    const updates = s.sent.filter(([k]) => k === 'update');
    assert.equal(updates.length, 1);
    assert.equal(updates[0][2].asx_disabled, true);
    assert.equal(s.$('fb-schedule').textContent, 'Schedule saved.');
    assert.equal(
      s.$('overview-cards').querySelectorAll('.section-card')[2].querySelector('.card-head .muted')
        .textContent,
      'Off',
    );
    assert.equal(s.$('overview-pill').textContent, 'Off');
    assert.equal(s.row('Account onboarding').querySelector('.row-state').textContent, 'Off');
    // The panel's Close returns focus to the card's Edit.
    await s.press(s.$('schedule-panel').querySelector('.panel-close'));
    assert.equal(s.$('schedule-panel').hidden, true);
    assert.equal(s.document.activeElement.getAttribute('aria-label'), 'Edit schedule');
    // All versions lists every version in a side panel; View opens it in the editor.
    const h = await boot();
    await h.overview();
    await h.press(h.find(h.$('overview-cards'), 'All versions'));
    assert.equal(h.$('history-panel').hidden, false);
    assert.equal(h.$('history-panel').getAttribute('role'), 'dialog');
    assert.equal(h.document.activeElement, h.$('history-title'));
    assert.match(h.$('history-list').visibleText, /^v1Live/);
    const loads = h.sent.filter(([k]) => k === 'asx_LoadDraft').length;
    await h.press(h.find(h.$('history-list'), 'View', 'View v1'));
    const after = h.sent.filter(([k]) => k === 'asx_LoadDraft');
    assert.ok(after.length > loads);
    assert.equal(after.at(-1)[1].RevisionId, 'rev-1');
    assert.equal(h.$('history-panel').hidden, true);
    assert.equal(h.$('template-editor').hidden, false);
  }
  {
    // After Delete template, focus goes to the next template in the list, else the search box,
    // and the result stays visible.
    const list = [
      { asx_templateid: TEMPLATE, asx_name: 'Account onboarding', asx_table: 'account' },
      { asx_templateid: OTHER, asx_name: 'Contract documents', asx_table: 'account' },
    ];
    const t = await boot({ templates: list });
    await t.overview();
    await t.press(t.$('overview-menu'));
    await t.press(t.$('menu-delete'));
    list.splice(0, 1);
    await t.press(t.find(t.$('overview-header').querySelector('.confirm'), 'Delete template'));
    assert.equal(
      t.document.activeElement.querySelector('.row-name').textContent,
      'Contract documents',
    );
    assert.match(
      t.$('fb-templates').visibleText,
      /^Account onboarding deleted\. Nothing in SharePoint changed\.$/,
    );
    assert.match(t.$('overview-cards').visibleText, /No template selected/);
    const only = [
      { asx_templateid: TEMPLATE, asx_name: 'Account onboarding', asx_table: 'account' },
    ];
    const last = await boot({ templates: only });
    await last.overview();
    await last.press(last.$('overview-menu'));
    await last.press(last.$('menu-delete'));
    only.splice(0);
    await last.press(
      last.find(last.$('overview-header').querySelector('.confirm'), 'Delete template'),
    );
    assert.equal(last.document.activeElement.id, 'template-search');
  }
  {
    // Another template with edits that cannot be saved asks before it discards them.
    const t = await boot({
      templates: [
        { asx_templateid: TEMPLATE, asx_name: 'Account onboarding', asx_table: 'account' },
        { asx_templateid: OTHER, asx_name: 'Contract documents', asx_table: 'account' },
      ],
    });
    await t.open();
    await t.unsaveable();
    const loads = t.sent.filter(([k]) => k === 'asx_LoadDraft').length;
    await t.press(t.row('Contract documents'));
    assert.equal(
      t.document.activeElement.textContent,
      'Open Contract documents? Your unsaved changes to Account onboarding are discarded.',
    );
    await t.press(t.find(t.$('template-groups'), 'Keep editing'));
    assert.equal(t.sent.filter(([k]) => k === 'asx_LoadDraft').length, loads);
    assert.equal(t.$('editor-pill').textContent, 'Draft v2');
    assert.equal(t.$('template-editor').hidden, false);
    assert.equal(t.document.activeElement, t.row('Contract documents'));
    // Discarding opens the other template's overview.
    await t.press(t.row('Contract documents'));
    await t.press(t.find(t.$('template-groups'), 'Open Contract documents'));
    assert.equal(t.$('template-editor').hidden, true);
    assert.equal(t.$('overview-title').textContent, 'Contract documents');
  }
  {
    // Folders and the folder panel: selection keeps focus, readable copy, the path, Remove folder
    // in the ⋯ menu; the destination's Remove in step 1 keeps its confirmation.
    const t = await boot();
    await t.open();
    await t.step(2);
    const node = t.node('General').querySelector('.node');
    node.focus();
    await t.press(node);
    const again = t.node('General').querySelector('.node');
    assert.equal(again.getAttribute('aria-current'), 'true');
    assert.ok(t.document.activeElement === again, 'Focus stays on the selected node after render');
    assert.equal(t.$('folder-path').textContent, 'Business documents › Account Name ›');
    assert.equal(t.$('folder-title').textContent, 'General');
    assert.equal(t.$('folder-menu').getAttribute('aria-label'), 'More actions for General');
    await t.select('Account Name');
    assert.equal(t.$('folder-path').textContent, 'Business documents ›');
    assert.equal(t.$('folder-menu').hidden, true, 'The top folder is removed with its destination');
    assert.equal(t.$('folder-shows-as').visibleText, 'Shows as: Account Name');
    assert.ok(t.$('help-folder-access') === null, 'Step 1 says who can open the folders');
    assert.equal(
      t.$('help-include-root').textContent,
      "If these conditions don't match, Documents skips this whole destination for the record. Folders it already created stay.",
    );
    assert.equal(t.$('create-mode').getAttribute('aria-describedby'), 'help-include-root');
    // ＋ Subfolder adds under the selected folder, selects it and focuses its name.
    await t.press(t.$('add-subfolder'));
    assert.ok(t.document.activeElement === t.$('folder-name'), 'The new folder name takes focus');
    assert.equal(t.node('New folder').querySelector('.node').getAttribute('aria-current'), 'true');
    assert.equal(t.state().sections[0].Folders.at(-1).Parent, 'root');
    // A folder with folders under it is removed after them.
    await t.select('General');
    await t.press(t.$('add-subfolder'));
    await t.select('General');
    await t.press(t.$('folder-menu'));
    const removeFolder = [...t.$('folder-menu-list').querySelectorAll('[role=menuitem]')].find(
      (i) => i.textContent === 'Remove folder',
    );
    assert.equal(removeFolder.getAttribute('aria-disabled'), 'true');
    assert.equal(
      t.document.getElementById(removeFolder.getAttribute('aria-describedby')).textContent,
      'Remove its folders first',
    );
    await t.press(removeFolder);
    assert.ok(t.node('General'), 'Not removed');
    // Removing a folder selects its parent and focuses its node.
    await t.select('New folder');
    await t.press(t.$('folder-menu'));
    await t.press(
      [...t.$('folder-menu-list').querySelectorAll('[role=menuitem]')].find(
        (i) => i.textContent === 'Remove folder',
      ),
    );
    assert.equal(
      t.document.activeElement.getAttribute('data-focus-key'),
      'node:general:general_docs',
    );
    await t.press(t.find(t.$('step-1'), 'Remove', 'Remove Business documents'));
    assert.equal(
      t.document.activeElement.textContent,
      'Remove Business documents and its 3 folders from this draft? Nothing changes in SharePoint until you publish.',
    );
  }
  {
    // ＋ Field lists this record's fields, then one group per related record: the primary name
    // first, deprecated fields last; a related field adds its record under an alias.
    const t = await boot();
    await t.open();
    await t.step(2);
    await t.select('General');
    await t.press(t.$('add-field'));
    const groups = [...t.$('field-options').querySelectorAll('[role=group]')];
    assert.deepEqual(
      groups.map((g) => g.getAttribute('aria-label')),
      ['This record', 'Primary Contact → Contact'],
    );
    const own = [...groups[0].querySelectorAll('[role=option]')].map(
      (o) => o.querySelector('.label').textContent,
    );
    assert.equal(own[0], 'Account Name');
    assert.equal(own.at(-1), '(Deprecated) Old Code');
    assert.deepEqual(
      [...groups[0].querySelectorAll('[role=option]')].map(
        (o) => o.querySelector('.kind').textContent,
      ),
      ['Text', 'Text', 'Number', 'Date', 'Choice', 'Choice', 'Text'],
      'The name kinds only, with their type',
    );
    await t.press(
      [...t.$('field-options').querySelectorAll('[role=option]')].find(
        (o) => o.dataset.value === 'lookup:primarycontactid:contact:fullname',
      ),
    );
    assert.match(t.$('folder-name').value, /\{lookup_1\.fullname\}$/);
    assert.equal(
      [...t.$('folder-tree').querySelectorAll('.token')].at(-1).textContent,
      'Primary Contact › Full Name',
    );
  }
  {
    // Condition builder: names on every control, typed values, the lookup picker, validation.
    const t = await boot();
    await t.open();
    await t.select('General');
    const editor = t.$('conditions');
    await t.mode(1);
    const field = [...editor.querySelectorAll('select')].find(
      (s) => s.getAttribute('aria-label') === 'Field, condition 1',
    );
    assert.ok(t.document.activeElement === field, 'A new condition focuses its Field');
    const group = editor.querySelector('.condition-group');
    assert.equal(group.getAttribute('role'), 'group');
    assert.equal(group.getAttribute('aria-label'), 'Conditions for General');
    for (const control of editor.querySelectorAll('select, input'))
      assert.ok(control.getAttribute('aria-label'), 'Unnamed control in the condition builder');
    assert.equal(
      [...t.labelled(editor, 'Match, Conditions for General').querySelectorAll('option')]
        .map((o) => o.textContent)
        .join('|'),
      'All|Any',
    );
    const kinds = {
      'root.revenue': ['number', 'decimal'],
      'root.closedate': ['date', null],
      'root.reviewedon': ['datetime-local', null],
    };
    for (const [value, [type, mode]] of Object.entries(kinds)) {
      await t.change(field, value);
      const input = [...editor.querySelectorAll('input')].find(
        (i) => i.getAttribute('aria-label') === 'Value, condition 1',
      );
      assert.equal(input.type, type, value);
      assert.equal(input.getAttribute('inputmode'), mode);
    }
    await t.change(field, 'root.statecode');
    const choice = [...editor.querySelectorAll('select')].find(
      (s) => s.getAttribute('aria-label') === 'Value, condition 1',
    );
    assert.deepEqual(
      [...choice.querySelectorAll('option')].map((o) => o.textContent),
      ['Another field…', 'Choose a value', 'Active', 'Inactive'],
    );
    await t.change(field, 'root.primarycontactid');
    await t.press(t.find(editor, 'Choose record…'));
    assert.deepEqual([...t.looked.at(-1).entityTypes], ['contact']);
    assert.match(editor.visibleText, /Jane Smith/);
    // The chosen record is a removable chip (spec 3.1): ✕ is named, clears the value, and focus
    // returns to Choose record….
    const clear = [...editor.querySelectorAll('button')].find(
      (b) => b.getAttribute('aria-label') === 'Remove Jane Smith, condition 1',
    );
    assert.ok(clear, 'The lookup chip has a named remove button');
    await t.press(clear);
    assert.doesNotMatch(editor.visibleText, /Jane Smith/);
    assert.equal(t.document.activeElement.getAttribute('aria-label'), 'Choose record, condition 1');
    await t.press(t.find(editor, 'Choose record…'));
    assert.match(editor.visibleText, /Jane Smith/);
    await t.change(t.labelled(editor, 'Operator, condition 1'), 'IsNull');
    assert.equal(
      [...editor.querySelectorAll('input, select')].filter(
        (n) => n.getAttribute('aria-label') === 'Value, condition 1',
      ).length,
      0,
      'is empty has no value control',
    );
    await t.change(t.labelled(editor, 'Operator, condition 1'), 'Equal');
    await t.press(t.find(editor, 'Choose record…'));
    // Removing a condition moves focus to the next condition's Field.
    await t.press(
      [...editor.querySelectorAll('button')].find(
        (b) => b.getAttribute('aria-label') === 'Add condition to conditions for General',
      ),
    );
    assert.equal(t.document.activeElement.getAttribute('aria-label'), 'Field, condition 2');
    await t.press(
      [...editor.querySelectorAll('button')].find(
        (b) => b.getAttribute('aria-label') === 'Remove condition 2',
      ),
    );
    assert.equal(
      t.document.activeElement.getAttribute('aria-label'),
      'Add condition to conditions for General',
    );
    // An empty group is refused at Save, at the group, with focus on it.
    await t.press(
      [...editor.querySelectorAll('button')].find(
        (b) => b.getAttribute('aria-label') === 'Add group to conditions for General',
      ),
    );
    await t.flush();
    assert.equal(t.sent.filter(([k]) => k === 'asx_CreateDraft').length, 0);
    const alert = editor.querySelector('[role=alert]');
    assert.equal(alert.textContent, 'Add a condition or remove this group');
    assert.equal(
      t.document.activeElement.getAttribute('aria-label'),
      'Match, Group 2 conditions',
      'A new group focuses its All/Any',
    );
    await t.press(
      [...editor.querySelectorAll('button')].find(
        (b) => b.getAttribute('aria-label') === 'Remove group 2',
      ),
    );
    await t.flush();
    const saved = t.last('asx_CreateDraft').Destinations[0].Folders[1].Condition.Conditions[0];
    assert.deepEqual(
      [saved.Column, saved.LiteralKind, saved.Literal],
      ['primarycontactid', 'Lookup', CONTACT],
    );
  }
  {
    // Add group stops at the condition depth bound, with the traced reason (ruling 4).
    const d = draft();
    const t = await boot({
      loaded: { RevisionId: 'rev-1', RowVersion: '3', Status: 'Published', Version: 1, Draft: d },
    });
    const depth = t.document.defaultBounds.conditionDepth;
    const nest = (n) => ({
      All: true,
      Conditions: [{ Source: 'root', Column: 'name', Operator: 'IsNotNull' }],
      Groups: n > 1 ? [nest(n - 1)] : [],
    });
    d.Destinations[0].Folders[1].Condition = nest(depth);
    await t.open();
    await t.select('General');
    const editor = t.$('conditions');
    assert.equal(editor.querySelectorAll('fieldset').length, 0, 'No fieldset boxes');
    assert.equal(editor.querySelectorAll('.condition-group.nested').length, depth - 1);
    const addGroup = (n) =>
      [...editor.querySelectorAll('button')].find(
        (b) => b.getAttribute('aria-label') === 'Add group to group ' + n + ' conditions',
      );
    const deepest = addGroup(depth);
    assert.equal(deepest.getAttribute('aria-disabled'), 'true');
    assert.equal(
      t.document.getElementById(deepest.getAttribute('aria-describedby')).textContent,
      'Condition groups can be nested up to ' +
        depth +
        " levels deep, because deeper templates can't be sent to Dataverse.",
    );
    assert.equal(addGroup(depth - 1).hasAttribute('aria-disabled'), false);
    const groups = editor.querySelectorAll('.condition-group').length;
    await t.press(deepest);
    assert.equal(
      editor.querySelectorAll('.condition-group').length,
      groups,
      'No group past the bound',
    );
  }
  {
    // A loaded lookup condition shows the record's name, never its ID.
    const condition = {
      All: true,
      Groups: [],
      Conditions: [
        {
          Source: 'root',
          Column: 'primarycontactid',
          Operator: 'Equal',
          LiteralKind: 'Lookup',
          Literal: CONTACT,
          LiteralLabel: 'Jane Smith',
          LiteralTable: 'contact',
        },
      ],
    };
    const d = draft();
    d.Sources[0].Columns.push({ Name: 'primarycontactid', Kind: 'Lookup' });
    d.Destinations[0].Folders[1].Condition = condition;
    const t = await boot({
      loaded: { RevisionId: 'rev-1', RowVersion: '3', Status: 'Published', Version: 1, Draft: d },
    });
    await t.open();
    await t.step(2);
    await t.select('General');
    assert.match(t.$('conditions').visibleText, /Jane Smith/);
    assert.doesNotMatch(t.$('step-2').visibleText, GUID);
    const rule = t.node('General').querySelector('.rule');
    assert.equal(rule.textContent, '◆ When Primary Contact is Jane Smith', 'The rule in words');
    assert.equal(
      t.node('General').querySelector('.node').getAttribute('aria-describedby'),
      rule.id,
      'The node is described by its rule',
    );
  }
  {
    // Result for: an edit previews the draft again after the pause, and the result follows it.
    const t = await boot({
      handle: (api, body) =>
        api === 'asx_PreviewTemplate'
          ? {
              Folders: [
                {
                  Section: 'general',
                  Node: 'root',
                  Name: 'Contoso Ltd',
                  RelativePath: 'Contoso Ltd',
                },
                ...JSON.parse(JSON.stringify(body.Draft.Destinations[0].Folders))
                  .filter((f) => f.Parent === 'root')
                  .map((f) => ({
                    Section: 'general',
                    Node: f.Key,
                    Name: f.Name,
                    RelativePath: 'Contoso Ltd/' + f.Name,
                  })),
              ],
              Notices: ['Name adjusted for SharePoint.'],
            }
          : null,
    });
    await t.open();
    await t.step(3);
    await t.press(t.$('result-choose'));
    assert.equal(t.$('previewTrees').hasAttribute('aria-live'), false);
    assert.equal(
      t.$('previewTrees').querySelector('.callout').textContent,
      'Name adjusted for SharePoint.',
    );
    await t.rename('General', 'Edited');
    await t.flush();
    const request = t.last('asx_PreviewTemplate');
    assert.equal(request.RevisionId, '');
    assert.equal(request.Draft.Destinations[0].Folders[1].Name, 'Edited');
    await t.step(3);
    assert.match(t.$('previewTrees').visibleText, /Edited/);
    // A failed preview says so in place of the tree.
    const f = await boot({
      handle: (api) => (api === 'asx_PreviewTemplate' ? new Error('Record not found.') : null),
    });
    await f.open();
    await f.step(3);
    await f.press(f.$('result-choose'));
    assert.equal(f.$('previewTrees').textContent, 'Preview failed: Record not found.');
  }
  {
    // Re-run for existing records: the count labels the button, the confirmation, a background start.
    const t = await boot();
    await t.overview();
    await t.press(t.$('overview-menu'));
    await t.press(t.$('menu-rerun'));
    assert.equal(t.$('rerun-all').textContent, 'Re-run all 1,240 records');
    await t.press(t.$('rerun-all'));
    assert.equal(
      t.document.activeElement.textContent,
      'Re-run v1 for all 1,240 Account records? Documents works through them in the background, after other work, so this can take a while. You can close this page and follow it in Monitor.',
    );
    await t.press(t.find(t.$('rerun-panel'), 'Re-run all records'));
    const start = t.last('asx_ManageWork');
    assert.equal(start.Command, 'StartTemplateRun');
    assert.equal(start.TemplateId, TEMPLATE);
    assert.match(start.RequestId, GUID);
    assert.match(t.$('fb-rerun').textContent, /^Re-run started\./);
    await t.press(t.find(t.$('fb-rerun'), 'Follow it in Monitor →'));
    const nav = t.sent.filter(([k]) => k === 'navigate').at(-1)[1];
    assert.equal(nav.data, 'monitor-' + BUILD);
    // CountRecords reads Dataverse's daily snapshot, so its total is usually "about N" (ruling 5).
    const about = await boot({
      handle: (api, b) =>
        b.Command === 'CountRecords'
          ? { Status: 'Counted', Run: { Total: 1240, TotalEstimated: true } }
          : null,
    });
    await about.overview();
    await about.press(about.$('overview-menu'));
    await about.press(about.$('menu-rerun'));
    assert.equal(about.$('rerun-all').textContent, 'Re-run all about 1,240 records');
    await about.press(about.$('rerun-all'));
    assert.match(
      about.document.activeElement.textContent,
      /^Re-run v1 for all about 1,240 Account records\? /,
    );
    // An active run replaces the button with its progress line.
    const busy = await boot({
      handle: (api, b) =>
        b.Command === 'ListProblems'
          ? {
              Status: 'Page',
              Problems: [
                {
                  Key: 'templaterun:1',
                  Kind: 'TemplateRun',
                  TemplateId: TEMPLATE,
                  Run: {
                    TemplateId: TEMPLATE,
                    State: 'Running',
                    Planned: 500,
                    Total: 1240,
                    TotalEstimated: false,
                  },
                },
              ],
              Next: null,
            }
          : null,
    });
    await busy.overview();
    await busy.press(busy.$('overview-menu'));
    await busy.press(busy.$('menu-rerun'));
    assert.equal(busy.$('rerun-all').hidden, true);
    assert.match(
      busy.$('rerun-progress').visibleText,
      /Re-run in progress: 500 of 1,240 · Open in Monitor/,
    );
    // Preview re-run: more records than the traced bound is refused at the button.
    const many = await boot({ pickMany: true });
    await many.overview();
    await many.press(many.$('overview-menu'));
    await many.press(many.$('menu-rerun'));
    await many.press(many.$('rerun-preview'));
    const bound = many.document.defaultBounds.previewRecords;
    assert.equal(
      many.$('rerun-reason').textContent,
      'Preview covers up to ' +
        bound +
        " records at a time so it finishes within Dataverse's 2-minute limit. To re-run every record, use Re-run all.",
    );
    assert.equal(
      many.sent.some(([, b]) => b?.Command === 'PreviewBatch'),
      false,
    );
  }
  {
    // Fix 1: the next editing session after Publish saves with the row version publishing gave.
    let published = false;
    const t = await boot({
      loaded: { RevisionId: 'rev-2', RowVersion: '4', Status: 'Draft', Version: 2, Draft: draft() },
      handle: (api) => {
        if (api === 'asx_PublishTemplate') {
          published = true;
          return { Status: 'Published', Notices: [] };
        }
        if (api === 'asx_LoadDraft' && published)
          return {
            RevisionId: 'rev-2',
            RowVersion: '5',
            Status: 'Published',
            Version: 2,
            Draft: draft(),
          };
        return null;
      },
    });
    await t.open();
    await t.step(3);
    await t.press(t.$('publish'));
    assert.equal(t.$('template-overview').hidden, false);
    await t.press(t.$('overview-edit'));
    assert.equal(t.$('editor-pill').textContent, 'Draft v3');
    assert.equal(t.$('editor-note').textContent, 'v2 stays live until you publish');
    await t.rename('General', 'After publish');
    await t.flush();
    const after = t.last('asx_CreateDraft');
    assert.deepEqual([after.RevisionId, after.RowVersion], ['rev-2', '5']);
  }
  {
    // Fix 2: related tables load after the first render, at most 4 tables at a time; a group
    // still loading says so.
    const related = Array.from({ length: 10 }, (_, i) => 'rel' + i);
    let release;
    const gate = new Promise((resolve) => (release = resolve));
    const active = new Map();
    let peak = 0;
    const t = await boot({
      extraLookups: related.map((table) => ({
        LogicalName: table + 'id',
        Targets: [table],
        DisplayName: { UserLocalizedLabel: { Label: 'Link ' + table } },
      })),
      delay: async (table) => {
        if (!related.includes(table)) return;
        active.set(table, (active.get(table) || 0) + 1);
        peak = Math.max(peak, active.size);
        await gate;
        active.set(table, active.get(table) - 1);
        if (!active.get(table)) active.delete(table);
      },
    });
    await t.open();
    assert.ok(t.node('Account Name'), 'The template renders before its related tables load');
    assert.equal(peak, 4, 'Four related tables load at a time');
    await t.step(2);
    await t.select('General');
    await t.press(t.$('add-field'));
    const group = () =>
      [...t.$('field-options').querySelectorAll('[role=group]')].find(
        (g) => g.getAttribute('aria-label') === 'Link rel9 → rel9',
      );
    assert.equal(group().getAttribute('aria-disabled'), 'true');
    assert.equal(group().querySelectorAll('[role=option]')[0].textContent, 'Loading fields…');
    release();
    await t.document.settle();
    assert.equal(peak, 4, 'Never more than four');
    // The open list follows the tables (these have no fields a name can use, so no group).
    assert.doesNotMatch(t.$('field-options').textContent, /Loading fields…/);
    assert.ok(group() === undefined, 'No group for a table without name fields');
    assert.ok(t.document.activeElement === t.$('field-search'), 'Typing in the search goes on');
    // A table that fails twice (one retry) shows as unavailable, not as missing.
    let attempts = 0;
    const failed = await boot({
      failFetch: (url, table) =>
        table === 'contact' && url.includes('/Attributes?') && ++attempts > 0,
    });
    await failed.open();
    await failed.select('General');
    await failed.press(failed.$('add-field'));
    const contact = [...failed.$('field-options').querySelectorAll('[role=group]')].find(
      (g) => g.getAttribute('aria-label') === 'Primary Contact → Contact',
    );
    assert.equal(contact.getAttribute('aria-disabled'), 'true');
    assert.equal(
      contact.querySelectorAll('[role=option]')[0].textContent,
      "Couldn't load these fields",
    );
    assert.equal(attempts, 2, 'One retry');
  }
  {
    // Task 9 ruling 7: a related table that arrives while the admin types in the editor leaves
    // the pickers stale; they redraw once focus leaves the field.
    let release;
    const gate = new Promise((resolve) => (release = resolve));
    const t = await boot({
      extraLookups: [
        {
          LogicalName: 'rel0id',
          Targets: ['rel0'],
          DisplayName: { UserLocalizedLabel: { Label: 'Link rel0' } },
        },
      ],
      delay: async (table) => {
        if (table === 'rel0') await gate;
      },
    });
    await t.open();
    await t.select('General');
    await t.mode(1);
    const group = () =>
      [...t.labelled(t.$('conditions'), 'Field, condition 1').querySelectorAll('optgroup')].find(
        (g) => g.getAttribute('label') === 'Link rel0 → rel0',
      );
    const name = t.$('folder-name');
    name.focus();
    release();
    await t.document.settle();
    assert.equal(group().disabled, true, 'No redraw while the admin types');
    t.document.activeElement = t.document.body;
    name.dispatchEvent(new FakeEvent('focusout'));
    await new Promise((resolve) => setTimeout(resolve, 20));
    await t.document.settle();
    assert.equal(group().disabled, false, 'The pickers redraw once focus leaves the field');
  }
  {
    // Task 9 ruling 7: switching templates stops the old template's background reads from
    // taking more tables.
    const related = Array.from({ length: 10 }, (_, i) => 'rel' + i);
    let release;
    const gate = new Promise((resolve) => (release = resolve));
    const requested = new Set();
    const t = await boot({
      enabled: ['account', 'contact'],
      extraLookups: related.map((table) => ({
        LogicalName: table + 'id',
        Targets: [table],
        DisplayName: { UserLocalizedLabel: { Label: 'Link ' + table } },
      })),
      delay: async (table) => {
        if (!related.includes(table)) return;
        requested.add(table);
        await gate;
      },
    });
    await t.open();
    assert.equal(requested.size, 4);
    await t.press(t.$('new-template'));
    await t.change(t.$('new-template-table'), 'contact');
    release();
    await t.document.settle();
    assert.equal(requested.size, 4, 'The old template takes no more tables after a switch');
  }
  {
    // Fix 3: a save reads the new version's number. v1 opened while v3 is the latest saves as v4.
    const t = await boot({ versions: { 'rev-2': 4 } });
    await t.open();
    await t.rename('General', 'From v1');
    await t.flush();
    assert.equal(t.$('editor-pill').textContent, 'Draft v4');
    await t.step(3);
    assert.equal(t.$('publishing-title').textContent, 'Publishing v4');
    assert.equal(t.$('publish').textContent, 'Publish v4 and re-run');
  }
  {
    // Fix 4: a save whose follow-up reads fail still records the template, so the next save
    // updates it instead of creating a second one.
    let failing = false;
    const t = await boot({
      templates: [],
      versions: { 'rev-2': 1 },
      reads: (table) => (failing && table === 'asx_template' ? new Error('Network down.') : null),
    });
    await t.press(t.$('new-template'));
    await t.change(t.$('templateName'), 'Client onboarding');
    await t.press(t.$('add-destination'));
    failing = true;
    await t.flush();
    assert.match(
      t.$('fb-editor').textContent,
      /^Draft saved, but the page could not refresh: Network down\./,
    );
    assert.match(t.$('save-status').textContent, /^Saved /, 'The draft itself was saved');
    assert.equal(t.$('editor-pill').textContent, 'Draft v1');
    failing = false;
    await t.change(t.$('folder-name'), 'Second save');
    await t.flush();
    const second = t.last('asx_CreateDraft');
    assert.deepEqual(
      [second.TemplateId, second.RevisionId, second.RowVersion],
      [TEMPLATE, 'rev-2', '4'],
    );
  }
  {
    // Fix 6: an unavailable condition field is not shown by its internal name.
    const d = draft();
    d.Destinations[0].Folders[1].Condition = {
      All: true,
      Groups: [],
      Conditions: [{ Source: 'root', Column: 'retired_code', Operator: 'IsNotNull' }],
    };
    const t = await boot({
      loaded: { RevisionId: 'rev-1', RowVersion: '3', Status: 'Published', Version: 1, Draft: d },
    });
    await t.open();
    await t.select('General');
    assert.match(t.$('conditions').visibleText, /Unavailable field\. Select a replacement\./);
    assert.doesNotMatch(t.$('step-2').visibleText, /retired_code/);
  }
  {
    // Fix 7: a number is sent as typed; one the server cannot read is refused at the control.
    const t = await boot();
    await t.open();
    await t.select('General');
    const editor = t.$('conditions');
    await t.mode(1);
    await t.change(t.labelled(editor, 'Field, condition 1'), 'root.revenue');
    await t.change(t.labelled(editor, 'Value, condition 1'), '1e5');
    await t.flush();
    assert.equal(t.sent.filter(([k]) => k === 'asx_CreateDraft').length, 0);
    assert.equal(editor.querySelector('[role=alert]').textContent, 'Enter a number');
    await t.change(t.labelled(editor, 'Value, condition 1'), ' 12345678901234567890.5 ');
    await t.flush();
    const saved = t.last('asx_CreateDraft').Destinations[0].Folders[1].Condition.Conditions[0];
    assert.equal(saved.Literal, '12345678901234567890.5');
  }
  {
    // Fix 8: without the Operator role every re-run action is disabled with the reason.
    const t = await boot({ privileges: { prvCreateasx_operatorcommand: false } });
    await t.overview();
    await t.press(t.$('overview-menu'));
    await t.press(t.$('menu-rerun'));
    for (const id of ['rerun-all', 'rerun-preview', 'rerun-these']) {
      const control = t.$(id);
      assert.equal(control.getAttribute('aria-disabled'), 'true', id);
      const reasons = control
        .getAttribute('aria-describedby')
        .split(' ')
        .map((r) => t.document.getElementById(r).textContent);
      assert.ok(reasons.includes('Needs the Documents Operator role.'), id);
    }
    const picks = t.looked.length;
    await t.press(t.$('rerun-preview'));
    assert.equal(t.looked.length, picks, 'A blocked preview opens no picker');
  }
  {
    // Header and stepper: table, name, Draft pill, the live note; three tabs, freely clickable,
    // arrows move focus only.
    const t = await boot();
    await t.open();
    assert.equal(t.$('editor-table').textContent, 'Account ›');
    assert.equal(t.$('editor-title').textContent, 'Account onboarding');
    assert.equal(t.$('editor-pill').textContent, 'Draft v2');
    assert.equal(t.$('editor-note').textContent, 'v1 stays live until you publish');
    const tabs = [...t.$('editor-steps').querySelectorAll('[role=tab]')];
    assert.deepEqual(
      tabs.map((b) => b.querySelector('.step-label').textContent),
      ['Destinations', 'Folders', 'Review and publish'],
    );
    assert.equal(tabs[0].getAttribute('aria-selected'), 'true');
    assert.equal(t.$('step-back').hidden, true);
    assert.equal(t.$('step-next').textContent, 'Next: Folders');
    tabs[0].focus();
    tabs[0].key('ArrowRight');
    assert.equal(t.document.activeElement, tabs[1]);
    assert.equal(tabs[1].getAttribute('aria-selected'), 'false', 'Arrows move focus only');
    await t.press(tabs[2]);
    assert.equal(t.$('step-3').hidden, false);
    assert.equal(t.$('step-1').hidden, true);
    await t.press(t.$('step-back'));
    assert.equal(t.$('step-2').hidden, false);
    assert.equal(t.$('step-next').textContent, 'Next: Review');
  }
  {
    // Autosave: 1.5 s after the last edit; one save in flight; an edit made meanwhile is saved
    // next, with the revision and row version the first save returned (Review Focus 1).
    let release;
    const gate = new Promise((resolve) => (release = resolve));
    let saves = 0;
    const t = await boot({
      handle: async (api) => {
        if (api !== 'asx_CreateDraft') return null;
        saves++;
        if (saves === 1) await gate;
        return {
          TemplateId: TEMPLATE,
          RevisionId: 'rev-2',
          RowVersion: String(3 + saves),
          Status: 'Draft',
        };
      },
    });
    await t.open();
    await t.step(2);
    await t.rename('General', 'One');
    const drafts = () => t.sent.filter(([a]) => a === 'asx_CreateDraft').map(([, b]) => b);
    assert.equal(drafts().length, 0, 'Nothing saves before the pause');
    await t.flush();
    assert.equal(t.$('save-status').textContent, 'Saving…');
    await t.change(t.$('folder-name'), 'Two');
    await t.flush();
    assert.equal(drafts().length, 1, 'One save in flight at a time');
    release();
    await t.document.settle();
    await t.flush();
    assert.equal(drafts().length, 2);
    assert.equal(drafts()[1].RevisionId, 'rev-2');
    assert.equal(drafts()[1].RowVersion, '4');
    assert.equal(drafts()[1].Destinations[0].Folders[1].Name, 'Two');
    assert.match(t.$('save-status').textContent, /^Saved \d{1,2}:\d{2}/);
    assert.equal(t.$('editor-pill').textContent, 'Draft v2');
  }
  {
    // A new template is created once, even when an edit lands while its first save runs.
    let release;
    const gate = new Promise((resolve) => (release = resolve));
    let saves = 0;
    const t = await boot({
      templates: [],
      reads: (table, id) =>
        table === 'asx_template' && id === OTHER
          ? { asx_templateid: OTHER, asx_name: 'Client onboarding', asx_table: 'account' }
          : null,
      handle: async (api) => {
        if (api !== 'asx_CreateDraft') return null;
        saves++;
        if (saves === 1) await gate;
        return {
          TemplateId: OTHER,
          RevisionId: 'rev-new',
          RowVersion: String(saves),
          Status: 'Draft',
        };
      },
    });
    await t.press(t.$('new-template'));
    // A new template saves once it has a name and a destination.
    assert.equal(t.$('templateName').hidden, false);
    await t.change(t.$('templateName'), 'Client onboarding');
    await t.press(t.$('add-destination'));
    const drafts = () => t.sent.filter(([a]) => a === 'asx_CreateDraft').map(([, b]) => b);
    await t.flush(); // save 1 starts and waits at the gate
    assert.equal(drafts().length, 1);
    assert.equal(drafts()[0].Name, 'Client onboarding');
    await t.change(t.labelled(t.$('step-1'), 'Name'), 'Client files'); // an edit during the flight
    await t.flush();
    assert.equal(drafts().length, 1, 'No second save while the first runs');
    release();
    await t.document.settle();
    await t.flush();
    assert.deepEqual(
      drafts().map((d) => d.TemplateId),
      [null, OTHER],
      'One template, then its follow-up save',
    );
    assert.equal(drafts()[1].Destinations[0].Name, 'Client files');
    assert.equal(t.$('templateName').hidden, true, 'The saved name shows as the heading');
    assert.equal(t.$('editor-name').textContent, 'Client onboarding');
  }
  {
    // A failed save says so with Retry; Retry saves again.
    let fail = true;
    const t = await boot({
      handle: (api) =>
        api === 'asx_CreateDraft' && fail ? new Error('The server refused the request.') : null,
    });
    await t.open();
    await t.change(t.labelled(t.$('step-1'), 'Name'), 'Client files');
    await t.flush();
    assert.equal(t.$('save-status').visibleText, "Couldn't save · Retry");
    assert.ok(t.$('save-status').classList.contains('is-error'));
    assert.equal(t.$('fb-editor').textContent, 'The server refused the request.');
    const retry = t.find(t.$('save-status'), 'Retry');
    assert.equal(retry.getAttribute('aria-label'), 'Retry saving Account onboarding');
    fail = false;
    await t.press(retry);
    assert.match(t.$('save-status').textContent, /^Saved /);
    assert.equal(t.$('fb-editor').textContent, '');
  }
  {
    // Invalid conditions block the autosave and mark the Folders tab, and never move focus.
    const t = await boot();
    await t.open();
    await t.step(2);
    await t.select('General');
    await t.mode(1);
    await t.change(t.labelled(t.$('conditions'), 'Field, condition 1'), 'root.revenue');
    const value = t.labelled(t.$('conditions'), 'Value, condition 1');
    value.focus();
    await t.change(value, '1e5');
    await t.flush();
    assert.equal(t.sent.filter(([a]) => a === 'asx_CreateDraft').length, 0);
    assert.ok(t.$('step-tab-2').classList.contains('has-error'));
    assert.equal(t.$('step-tab-2').getAttribute('aria-label'), 'Folders, has errors');
    assert.equal(t.$('step-tab-2').querySelector('.step-dot').hidden, true, 'No dot with an error');
    assert.equal(t.document.activeElement, value, 'Autosave never takes focus');
    // The same problems are not announced again at the next pause.
    const alert = t.$('conditions').querySelector('[role=alert]');
    assert.equal(alert.textContent, 'Enter a number');
    await t.change(t.labelled(t.$('step-1'), 'Name'), 'Client files');
    await t.flush();
    // assert.ok: a failed equal on two elements would print the whole page.
    assert.ok(t.$('conditions').querySelector('[role=alert]') === alert, 'Not re-inserted');
    assert.equal(t.$('conditions').querySelectorAll('[role=alert]').length, 1);
    // Fixed, the next autosave saves and the error state clears.
    await t.change(value, '100000');
    await t.flush();
    assert.equal(t.sent.filter(([a]) => a === 'asx_CreateDraft').length, 1);
    assert.equal(t.$('step-tab-2').classList.contains('has-error'), false);
  }
  {
    // The dirty guard covers an edit not saved yet: Save draft flushes the autosave first.
    const t = await boot();
    await t.open();
    await t.change(t.labelled(t.$('step-1'), 'Name'), 'Client files');
    const going = t.window.AsxdUi.navigate('monitor');
    await t.document.settle();
    await t.press(t.find(t.$('leavePrompt'), 'Save draft'));
    await going;
    assert.equal(t.last('asx_CreateDraft').Destinations[0].Name, 'Client files');
    assert.equal(
      t.sent
        .filter(([k]) => k === 'navigate')
        .at(-1)[1]
        .data.split('-')[0],
      'monitor',
    );
  }
  {
    // The dirty guard also covers a save still in flight: Save draft waits for it.
    let release;
    const gate = new Promise((resolve) => (release = resolve));
    const t = await boot({
      handle: async (api) => (api === 'asx_CreateDraft' ? gate.then(() => null) : null),
    });
    await t.open();
    await t.change(t.labelled(t.$('step-1'), 'Name'), 'Client files');
    await t.flush();
    assert.equal(t.$('save-status').textContent, 'Saving…');
    const going = t.window.AsxdUi.navigate('monitor');
    await t.document.settle();
    assert.match(
      t.$('leavePrompt').visibleText,
      /You have unsaved changes to Account onboarding\./,
    );
    await t.press(t.find(t.$('leavePrompt'), 'Save draft'));
    assert.equal(t.sent.filter(([k]) => k === 'navigate').length, 0, 'Nothing leaves mid-save');
    release();
    await going;
    assert.equal(t.sent.filter(([a]) => a === 'asx_CreateDraft').length, 1);
    assert.equal(t.sent.filter(([k]) => k === 'navigate').length, 1);
  }
  {
    // changesSince: alias renumbering and date formats are not changes; added, removed and
    // condition changes are listed in order and drive marks and step counts (Review Focus 2).
    const t = await boot();
    const { changesSince, nextKey } = t.window.AsxdAdmin;
    const sources = (alias) => [
      { Alias: 'root', Table: 'account', Lookup: null, columns: [] },
      { Alias: alias, Table: 'contact', Lookup: 'primarycontactid', columns: [] },
    ];
    const f = (Key, Parent, Name, Condition = null) => ({ Key, Parent, Name, Condition });
    const when = (field, Literal) => ({
      All: true,
      Groups: [],
      Conditions: [{ field, Operator: 'Equal', Literal, right: null }],
    });
    const dest = (Folders) => [
      { Key: 'general', Name: 'Client files', LibraryId: 'lib-a', Folders },
    ];
    const published = {
      sources: sources('lookup_1'),
      sections: dest([
        f('root', null, '{root.name}'),
        f('folder_1', 'root', 'Contact {lookup_1.fullname}'),
        f('folder_2', 'root', 'Reviewed', when('root.reviewedon', '2026-10-01T00:00:00Z')),
        f('folder_3', 'root', 'Invoices'),
        f('folder_4', 'root', 'Contracts'),
      ]),
    };
    const current = {
      sources: sources('lookup_2'),
      sections: dest([
        f('root', null, '{root.name}'),
        f('folder_1', 'root', 'Contact {lookup_2.fullname}'),
        f('folder_2', 'root', 'Reviewed', when('root.reviewedon', '2026-10-01T00:00:00.000Z')),
        f('folder_4', 'root', 'Contracts', when('root.statecode', '0')),
        f('folder_5', 'root', 'Projects'),
      ]),
    };
    const changes = changesSince(published, current);
    // Array.from: the page's arrays come from another realm, which deepEqual compares too.
    assert.deepEqual(
      Array.from(changes.folders, (c) => [c.kind, c.key]),
      [
        ['added', 'folder_5'],
        ['removed', 'folder_3'],
        ['condition', 'folder_4'],
      ],
    );
    assert.deepEqual({ ...changes.byStep }, { 1: 0, 2: 3 });
    assert.equal(changes.marks.get('general/folder_5'), 'new');
    assert.equal(changes.marks.get('general/folder_4'), 'edited');
    assert.equal(
      changes.marks.get('general/folder_1'),
      undefined,
      'Alias renumbering is not a change',
    );
    assert.equal(changes.marks.get('general/folder_2'), undefined, 'Same instant, other format');
    assert.deepEqual(
      Array.from(changes.removed, (r) => r.folder.Key),
      ['folder_3'],
    );
    // A new folder never takes a key the published revision used.
    assert.equal(
      nextKey(
        'folder_',
        new Set(['root', 'folder_1', 'folder_2', 'folder_3', 'folder_4', 'folder_5']),
      ),
      'folder_6',
    );
    assert.equal(
      changesSince(null, current).folders.length,
      0,
      'No published revision: no changes',
    );
    // Destinations: added, removed, renamed and a new library, keyed by Key; a folder of an
    // added destination is marked new without its own entry.
    const two = {
      sources: sources('lookup_1'),
      sections: [
        {
          Key: 'general',
          Name: 'Client files',
          LibraryId: 'lib-a',
          Folders: [f('root', null, 'A')],
        },
        { Key: 'legal', Name: 'Legal', LibraryId: 'lib-b', Folders: [f('root', null, 'B')] },
      ],
    };
    const moved = changesSince(two, {
      sources: sources('lookup_1'),
      sections: [
        { Key: 'general', Name: 'Clients', LibraryId: 'lib-c', Folders: [f('root', null, 'A')] },
        { Key: 'destination_1', Name: 'New', LibraryId: 'lib-a', Folders: [f('root', null, 'C')] },
      ],
    });
    assert.deepEqual(
      Array.from(moved.destinations, (c) => [c.kind, c.key, c.name, c.before, c.after]),
      [
        ['added', 'destination_1', 'New', undefined, undefined],
        ['removed', 'legal', 'Legal', undefined, undefined],
        ['renamed', 'general', 'Clients', 'Client files', 'Clients'],
        ['library', 'general', 'Clients', 'lib-a', 'lib-c'],
      ],
    );
    assert.equal(moved.folders.length, 0);
    assert.equal(moved.marks.get('destination_1/root'), 'new');
    assert.deepEqual({ ...moved.byStep }, { 1: 4, 2: 0 });
  }
  {
    // Change tracking in the page: the Folders tab gets a dot and the footer a count.
    const t = await boot();
    await t.open();
    assert.equal(t.$('step-tab-2').querySelector('.step-dot').hidden, true);
    await t.step(2);
    await t.rename('General', 'General documents');
    assert.equal(t.$('step-tab-2').querySelector('.step-dot').hidden, false);
    assert.equal(t.$('step-tab-2').getAttribute('aria-label'), 'Folders, has changes');
    assert.equal(t.$('step-count').textContent, '1 change in this step');
    // A new folder never reuses a key of the published revision.
    const published = await boot({
      loaded: {
        RevisionId: 'rev-1',
        RowVersion: '3',
        Status: 'Published',
        Version: 1,
        Draft: draft({
          Destinations: [
            {
              Key: 'general',
              Name: 'Business documents',
              LibraryId: 'lib-a',
              Folders: [
                { Key: 'root', Parent: null, Name: '{root.name}', Condition: null },
                { Key: 'folder_1', Parent: 'root', Name: 'Invoices', Condition: null },
              ],
            },
          ],
        }),
      },
    });
    await published.open();
    await published.step(2);
    await published.select('Invoices');
    await published.press(published.$('folder-menu'));
    await published.press(
      [...published.$('folder-menu-list').querySelectorAll('[role=menuitem]')].find(
        (i) => i.textContent === 'Remove folder',
      ),
    );
    await published.select('Account Name');
    await published.press(published.$('add-subfolder'));
    await published.flush();
    assert.deepEqual(
      published.last('asx_CreateDraft').Destinations[0].Folders.map((x) => x.Key),
      ['root', 'folder_2'],
    );
  }
  {
    // Step 1: the expanded destination with Name, Site and Library, the team panel and the
    // library link; the Library list ends with Sites & access.
    const t = await boot({
      handle: (api, body) =>
        body?.Command === 'GetPolicy'
          ? {
              Status: 'Applied',
              Policy: { Desired: [{ TeamId: 'team-1', Access: 'Contribute' }], Applied: [] },
              Teams: [{ TeamId: 'team-1', Name: 'Account managers' }],
            }
          : null,
    });
    await t.open();
    assert.equal(
      t.$('help-destinations').textContent,
      'Each destination is a SharePoint library. Every team with access to the library can open the folders created in it; to restrict a folder, give it its own library.',
    );
    const card = t.$('step-1').querySelector('.destination-card');
    assert.equal(card.querySelector('.eyebrow').textContent, 'Destination 1');
    assert.equal(t.labelled(card, 'Name').value, 'Business documents');
    const library = t.labelled(card, 'Library');
    assert.equal(
      [...library.querySelectorAll('option')].at(-1).textContent,
      'Set up a library in Sites & access…',
    );
    const who = card.querySelector('.who-can-open');
    assert.equal(who.querySelector('h3').textContent, 'Who can open these folders');
    assert.equal(who.querySelector('.team-line').textContent, 'Account managers · Contribute');
    await t.press(t.find(who, 'Change in Sites & access'));
    assert.equal(
      t.sent
        .filter(([k]) => k === 'navigate')
        .at(-1)[1]
        .data.split('-')[0],
      'access',
    );
    // Choosing "Set up a library…" keeps the chosen library and goes to Sites & access.
    const s = await boot();
    await s.open();
    const picker = s.labelled(s.$('step-1'), 'Library');
    await s.change(picker, '__setup');
    assert.equal(picker.value, 'lib-a');
    assert.equal(
      s.sent
        .filter(([k]) => k === 'navigate')
        .at(-1)[1]
        .data.split('-')[0],
      'access',
    );
    await s.flush();
    assert.equal(s.last('asx_CreateDraft'), undefined, 'Not an edit');
  }
  {
    // Add destination: the new one expands and the other collapses to one row; at the bound the
    // button is disabled with its reason; without the Security Administrator role the team panel
    // says what is needed.
    const t = await boot({ privileges: { prvCreateasx_policy: false } });
    await t.open();
    assert.equal(t.$('destination-limit-note').textContent, 'Up to 10 per template');
    await t.press(t.$('add-destination'));
    assert.equal(t.$('step-1').querySelectorAll('.destination-card').length, 1);
    const row = t.$('step-1').querySelector('.destination-row');
    assert.equal(row.querySelector('.eyebrow').textContent, 'Destination 1');
    assert.equal(row.querySelector('strong').textContent, 'Business documents');
    assert.equal(row.querySelector('.where').textContent, 'Delivery › General');
    assert.equal(
      t.$('step-1').querySelector('.who-can-open .reason').textContent,
      'Needs the Documents Security Administrator role.',
    );
    assert.equal(t.sent.filter(([, b]) => b?.Command === 'GetPolicy').length, 0);
    assert.equal(t.$('step-tab-1').getAttribute('aria-label'), 'Destinations, has changes');
    // Edit expands the row and collapses the other.
    await t.press(t.find(row, 'Edit', 'Edit Business documents'));
    assert.equal(
      t.labelled(t.$('step-1').querySelector('.destination-card'), 'Name').value,
      'Business documents',
    );
    // At the bound the button says why it is off.
    for (let i = 2; i < t.window.AsxdUi.BOUNDS.destinations; i++)
      await t.press(t.$('add-destination'));
    assert.equal(t.$('add-destination').getAttribute('aria-disabled'), 'true');
    assert.equal(
      t.$('destination-reason').textContent,
      'A template can have up to 10 destinations, because each record plans all of them in one step that Dataverse stops after 2 minutes.',
    );
    // With a policy read, a collapsed row names its teams.
    const p = await boot({
      handle: (api, body) =>
        body?.Command === 'GetPolicy'
          ? {
              Status: 'Applied',
              Policy: {
                Desired: [
                  { TeamId: 'team-1', Access: 'Contribute' },
                  { TeamId: 'team-2', Access: 'None' },
                ],
                Applied: [],
              },
              Teams: [
                { TeamId: 'team-1', Name: 'Legal' },
                { TeamId: 'team-2', Name: 'Finance' },
              ],
            }
          : null,
    });
    await p.open();
    await p.press(p.$('add-destination'));
    assert.equal(
      p.$('step-1').querySelector('.destination-row .teams').textContent,
      '1 team · Legal (Contribute)',
    );
  }
  {
    // A link to a step opens that template's editor at the step.
    const t = await boot({ hash: '#templates?template=' + TEMPLATE + '&step=2' });
    assert.equal(t.$('template-editor').hidden, false);
    assert.equal(t.$('step-2').hidden, false);
    assert.equal(t.$('step-tab-2').getAttribute('aria-selected'), 'true');
  }
  {
    // Fix round 1: a new template starts without a name and is not saved while its name is
    // typed; the pause counts once the name box is left.
    const t = await boot({ templates: [] });
    await t.press(t.$('new-template'));
    assert.equal(t.$('templateName').value, '');
    await t.press(t.$('add-destination'));
    await t.flush();
    assert.equal(t.last('asx_CreateDraft'), undefined, 'No name, no template');
    const name = t.$('templateName');
    name.focus();
    await t.change(name, 'Client on');
    await t.flush();
    assert.equal(t.last('asx_CreateDraft'), undefined, 'Not while the name is typed');
    await t.change(name, 'Client onboarding');
    t.document.activeElement = t.document.body;
    name.dispatchEvent(new FakeEvent('blur'));
    await t.flush();
    assert.equal(t.sent.filter(([a]) => a === 'asx_CreateDraft').length, 1);
    assert.equal(t.last('asx_CreateDraft').Name, 'Client onboarding');
    // Save draft in the unsaved-changes prompt for a nameless new template stays, and the name
    // box takes focus.
    const n = await boot({ templates: [] });
    await n.press(n.$('new-template'));
    await n.press(n.$('add-destination'));
    const going = n.window.AsxdUi.navigate('monitor');
    await n.document.settle();
    await n.press(n.find(n.$('leavePrompt'), 'Save draft'));
    assert.equal(await going, undefined);
    assert.equal(n.sent.filter(([k]) => k === 'navigate').length, 0, 'The page stays');
    assert.equal(
      n.$('fb-templates').textContent,
      'Enter a template name of 200 characters or fewer.',
    );
    assert.equal(n.document.activeElement, n.$('templateName'));
  }
  {
    // Fix round 1: picking another template or Close soon after an edit saves it first, without
    // asking.
    const t = await boot({
      templates: [
        { asx_templateid: TEMPLATE, asx_name: 'Account onboarding', asx_table: 'account' },
        { asx_templateid: OTHER, asx_name: 'Contract documents', asx_table: 'account' },
      ],
    });
    await t.open();
    await t.change(t.labelled(t.$('step-1'), 'Name'), 'Client files');
    await t.press(t.row('Contract documents'));
    assert.equal(t.last('asx_CreateDraft').Destinations[0].Name, 'Client files');
    assert.equal(t.$('template-groups').querySelector('.confirm'), null, 'No question');
    assert.equal(t.$('template-editor').hidden, true);
    assert.equal(t.$('overview-title').textContent, 'Contract documents');
    const c = await boot();
    await c.open();
    await c.change(c.labelled(c.$('step-1'), 'Name'), 'Client files');
    await c.press(c.$('editor-close'));
    assert.equal(c.last('asx_CreateDraft').Destinations[0].Name, 'Client files');
    assert.equal(c.$('leavePrompt').textContent, '');
    assert.equal(c.$('template-overview').hidden, false);
    // ＋ New saves the edit too, and the new template starts clean.
    const n = await boot({ enabled: ['account'] });
    await n.open();
    await n.change(n.labelled(n.$('step-1'), 'Name'), 'Client files');
    await n.press(n.$('new-template'));
    assert.equal(n.last('asx_CreateDraft').Destinations[0].Name, 'Client files');
    assert.equal(n.$('templateName').hidden, false);
    assert.equal(n.$('editor-pill').textContent, 'Draft v1');
  }
  {
    // Fix round 1: a save still in flight when the admin discards and starts a new template never
    // reaches the new one: its first save creates a template of its own.
    let release;
    const gate = new Promise((resolve) => (release = resolve));
    let saves = 0;
    const t = await boot({
      handle: async (api) => {
        if (api !== 'asx_CreateDraft') return null;
        saves++;
        if (saves === 1) await gate;
        return null;
      },
    });
    await t.open();
    await t.change(t.labelled(t.$('step-1'), 'Name'), 'Client files');
    await t.flush();
    assert.equal(t.$('save-status').textContent, 'Saving…');
    const leaving = t.window.AsxdUi.navigate('monitor');
    await t.document.settle();
    await t.press(t.find(t.$('leavePrompt'), 'Discard changes'));
    await leaving;
    await t.press(t.$('new-template'));
    release();
    await t.document.settle();
    await t.flush();
    assert.equal(t.$('templateName').hidden, false, 'The new template is open');
    assert.equal(t.$('editor-pill').textContent, 'Draft v1');
    await t.change(t.$('templateName'), 'Client onboarding');
    await t.press(t.$('add-destination'));
    await t.flush();
    const last = t.last('asx_CreateDraft');
    assert.deepEqual(
      [last.TemplateId, last.RevisionId, last.RowVersion, last.Name],
      [null, null, null, 'Client onboarding'],
    );
  }
  {
    // Fix round 1: an edit made while Publish runs is kept, and saved next as the next version
    // with the row version read back after publishing.
    let release;
    const gate = new Promise((resolve) => (release = resolve));
    let published = false;
    const t = await boot({
      loaded: { RevisionId: 'rev-2', RowVersion: '4', Status: 'Draft', Version: 2, Draft: draft() },
      versions: { 'rev-3': 3 },
      handle: async (api) => {
        if (api === 'asx_PublishTemplate') {
          await gate;
          published = true;
          return { Status: 'Published', Notices: [] };
        }
        if (api === 'asx_LoadDraft' && published)
          return {
            RevisionId: 'rev-2',
            RowVersion: '5',
            Status: 'Published',
            Version: 2,
            Draft: draft(),
          };
        if (api === 'asx_CreateDraft')
          return { TemplateId: TEMPLATE, RevisionId: 'rev-3', RowVersion: '1', Status: 'Draft' };
        return null;
      },
    });
    await t.open();
    await t.step(3);
    await t.press(t.$('publish'));
    await t.step(2);
    await t.rename('General', 'During');
    await t.flush();
    assert.equal(t.sent.filter(([a]) => a === 'asx_CreateDraft').length, 0, 'Waits for Publish');
    release();
    await t.document.settle();
    // Saved as the next version before the editor closes, so nothing is lost.
    const saved = t.last('asx_CreateDraft');
    assert.deepEqual([saved.RevisionId, saved.RowVersion], ['rev-2', '5']);
    assert.equal(saved.Destinations[0].Folders[1].Name, 'During');
    assert.equal(t.$('template-overview').hidden, false);
    assert.match(t.$('fb-templates').textContent, /^Published v2\./);
    // When those edits cannot be saved, the editor stays open with them.
    let free;
    const hold = new Promise((resolve) => (free = resolve));
    let done = false;
    const k = await boot({
      loaded: { RevisionId: 'rev-2', RowVersion: '4', Status: 'Draft', Version: 2, Draft: draft() },
      handle: async (api) => {
        if (api === 'asx_PublishTemplate') {
          await hold;
          done = true;
          return { Status: 'Published', Notices: [] };
        }
        if (api === 'asx_CreateDraft' && done) return new Error('The server is busy.');
        return null;
      },
    });
    await k.open();
    await k.step(3);
    await k.press(k.$('publish'));
    await k.rename('General', 'Kept');
    free();
    await k.document.settle();
    assert.equal(k.$('template-editor').hidden, false);
    assert.equal(k.$('folder-name').value, 'Kept');
    assert.equal(k.$('fb-editor').textContent, 'The server is busy.');
    assert.match(k.$('fb-templates').textContent, /^Published v2\./);
    // Close while Publish runs: its answer leaves the closed editor alone.
    let let_go;
    const wait = new Promise((resolve) => (let_go = resolve));
    const c = await boot({
      loaded: { RevisionId: 'rev-2', RowVersion: '4', Status: 'Draft', Version: 2, Draft: draft() },
      handle: async (api) =>
        api === 'asx_PublishTemplate'
          ? wait.then(() => ({ Status: 'Published', Notices: [] }))
          : null,
    });
    await c.open();
    await c.step(3);
    await c.press(c.$('publish'));
    await c.press(c.$('editor-close'));
    assert.equal(c.$('template-overview').hidden, false);
    let_go();
    await c.document.settle();
    assert.equal(c.$('template-editor').hidden, true);
    assert.equal(
      c.$('fb-templates').classList.contains('is-error'),
      false,
      c.$('fb-templates').textContent,
    );
  }
  {
    // Fix round 1: the published version is read again once when its read fails; if it still
    // fails, the footer says why, and new keys cannot collide with the unknown published ones.
    let failing = false;
    let tries = 0;
    const t = await boot({
      templates: [
        {
          asx_templateid: TEMPLATE,
          asx_name: 'Account onboarding',
          asx_table: 'account',
          _asx_publishedrevisionid_value: 'rev-1',
        },
      ],
      rows: (table) =>
        table === 'asx_revision'
          ? [
              {
                asx_revisionid: 'rev-2',
                asx_version: 2,
                asx_status: 'Draft',
                _asx_templateid_value: TEMPLATE,
              },
              {
                asx_revisionid: 'rev-1',
                asx_version: 1,
                asx_status: 'Published',
                _asx_templateid_value: TEMPLATE,
              },
            ]
          : null,
      handle: (api, body) => {
        if (api !== 'asx_LoadDraft') return null;
        if (body.RevisionId === 'rev-2')
          return {
            RevisionId: 'rev-2',
            RowVersion: '4',
            Status: 'Draft',
            Version: 2,
            Draft: draft(),
          };
        if (!failing) return null;
        tries++;
        return new Error('The published version could not be read.');
      },
    });
    await t.overview();
    failing = true;
    await t.press(t.$('overview-edit'));
    assert.equal(tries, 2, 'One retry');
    assert.equal(t.$('fb-editor').textContent, 'The published version could not be read.');
    await t.select('Account Name');
    await t.press(t.$('add-subfolder'));
    await t.press(t.$('add-destination'));
    await t.flush();
    const sent = t.last('asx_CreateDraft').Destinations;
    const folder = sent[0].Folders.at(-1).Key;
    assert.match(folder, /^folder_[a-z0-9]+$/);
    assert.doesNotMatch(
      folder,
      /^folder_\d{1,3}$/,
      'Not a small number a published folder may have',
    );
    assert.doesNotMatch(sent[1].Key, /^destination_\d{1,3}$/);
  }
  {
    // Fix round 1: a failed policy read says so in the footer, not as an empty team list.
    const t = await boot({
      handle: (api, body) =>
        body?.Command === 'GetPolicy' ? new Error('The policy could not be read.') : null,
    });
    await t.open();
    assert.equal(t.$('step-1').querySelectorAll('.team-line').length, 0);
    assert.equal(t.$('fb-editor').textContent, 'The policy could not be read.');
  }
  {
    // Fix round 1: a destination name changed only by spaces around it is not a change.
    const t = await boot();
    const { changesSince } = t.window.AsxdAdmin;
    const sections = (Name) => [{ Key: 'general', Name, LibraryId: 'lib-a', Folders: [] }];
    const changes = changesSince(
      { sources: [], sections: sections('Client files') },
      { sources: [], sections: sections(' Client files ') },
    );
    assert.equal(changes.destinations.length, 0);
  }
  {
    // The tree: destination pills, rule sentences, New and edited marks, Removed with Undo,
    // and no drag hint.
    const published = draft();
    published.Destinations[0].Folders.push({
      Key: 'invoices',
      Parent: 'root',
      Name: 'Invoices',
      Condition: null,
    });
    const t = await boot({
      handle: (api, body) =>
        api === 'asx_LoadDraft'
          ? body.RevisionId === 'rev-1'
            ? {
                RevisionId: 'rev-1',
                RowVersion: '3',
                Status: 'Published',
                Version: 1,
                Draft: published,
              }
            : null
          : null,
    });
    await t.open();
    await t.step(2);
    assert.deepEqual(
      [...t.$('dest-pills').querySelectorAll('button')].map((b) => b.textContent),
      ['Business documents'],
    );
    assert.equal(t.node('Account Name').querySelector('.token').textContent, 'Account Name');
    assert.equal(t.node('General').querySelector('.rule').textContent, 'Always');
    await t.select('Invoices');
    await t.press(t.$('folder-menu'));
    await t.press(
      [...t.$('folder-menu-list').querySelectorAll('[role=menuitem]')].find(
        (i) => i.textContent === 'Remove folder',
      ),
    );
    const removed = t.$('folder-tree').querySelector('.tree-row.is-removed');
    assert.equal(removed.querySelector('.tag').textContent, 'Removed');
    assert.equal(t.find(removed, 'Undo').dataset.focusKey, 'undo:general:invoices');
    await t.select('General');
    await t.press(t.$('add-subfolder'));
    assert.equal(t.node('New folder').querySelector('.tag').textContent, 'New');
    await t.press(t.find(t.$('folder-tree').querySelector('.tree-row.is-removed'), 'Undo'));
    assert.ok(t.node('Invoices'), 'Undo puts the folder back');
    assert.doesNotMatch(t.$('step-2').visibleText, /Drag folders/);
  }
  {
    // ＋ Folder adds a sibling; with the top folder selected it adds under it (one top folder).
    const t = await boot();
    await t.open();
    await t.step(2);
    await t.select('Account Name');
    await t.press(t.$('add-folder'));
    const added = t.state().sections[0].Folders.at(-1);
    assert.equal(added.Parent, 'root');
    assert.equal(added.Key, 'folder_1');
    // With another folder selected it adds beside it.
    await t.select('General');
    await t.press(t.$('add-folder'));
    assert.equal(t.state().sections[0].Folders.at(-1).Parent, 'root');
    assert.equal(t.state().sections[0].Folders.at(-1).Key, 'folder_2');
    // At the folder bound both are off, with one reason.
    const d = draft();
    const bound = t.window.AsxdUi.BOUNDS.foldersPerDestination;
    for (let i = 2; i < bound; i++)
      d.Destinations[0].Folders.push({
        Key: 'f' + i,
        Parent: 'root',
        Name: 'F' + i,
        Condition: null,
      });
    const full = await boot({
      loaded: { RevisionId: 'rev-1', RowVersion: '3', Status: 'Published', Version: 1, Draft: d },
    });
    await full.open();
    await full.step(2);
    assert.equal(full.state().sections[0].Folders.length, bound);
    for (const id of ['add-folder', 'add-subfolder']) {
      assert.equal(full.$(id).getAttribute('aria-disabled'), 'true', id);
      assert.equal(full.$(id).getAttribute('aria-describedby'), 'folder-reason', id);
    }
    assert.equal(
      full.$('folder-reason').textContent,
      'A destination can have up to ' +
        bound +
        " folders, because a record's folders for one destination are kept in one Dataverse row.",
    );
    await full.press(full.$('add-folder'));
    assert.equal(full.state().sections[0].Folders.length, bound, 'Nothing past the bound');
  }
  {
    // ＋ Field: a searchable popover grouped like the field picker, with the type on the right;
    // choosing a field inserts its token at the caret and replaces a selection; Escape closes only
    // the popover and returns focus to Name.
    const t = await boot();
    await t.open();
    await t.step(2);
    await t.select('General');
    const name = t.$('folder-name');
    name.value = 'P- files';
    name.oninput();
    name.selectionStart = name.selectionEnd = 2;
    await t.press(t.$('add-field'));
    assert.equal(t.$('field-popover').hidden, false);
    assert.equal(t.$('add-field').getAttribute('aria-expanded'), 'true');
    assert.ok(t.document.activeElement === t.$('field-search'));
    const groups = [...t.$('field-options').querySelectorAll('[role=group]')].map((g) =>
      g.getAttribute('aria-label'),
    );
    assert.equal(groups[0], 'This record');
    t.$('field-search').value = 'number';
    t.$('field-search').oninput();
    const options = [...t.$('field-options').querySelectorAll('[role=option]')];
    assert.deepEqual(
      options.map((o) => o.querySelector('.label').textContent),
      ['Account Number'],
    );
    assert.equal(options[0].querySelector('.kind').textContent, 'Text');
    await t.press(options[0]);
    assert.equal(name.value, 'P-{root.accountnumber} files');
    assert.equal(t.$('field-popover').hidden, true);
    assert.ok(t.document.activeElement === name);
    assert.equal(name.selectionStart, 'P-{root.accountnumber}'.length);
    assert.equal(t.$('folder-shows-as').visibleText, 'Shows as: P-Account Number files');
    assert.equal(
      t.node('P-Account Number files').querySelector('.token').textContent,
      'Account Number',
    );
    name.selectionStart = 0;
    name.selectionEnd = 2;
    await t.press(t.$('add-field'));
    t.$('field-search').key('Escape');
    assert.equal(t.$('field-popover').hidden, true);
    assert.ok(t.document.activeElement === name);
    await t.press(t.$('add-field'));
    t.$('field-search').value = 'account name';
    t.$('field-search').oninput();
    t.$('field-search').key('ArrowDown');
    assert.equal(
      t.$('field-search').getAttribute('aria-activedescendant'),
      t.$('field-options').querySelector('[role=option]').id,
    );
    t.$('field-search').key('Enter');
    assert.equal(name.value, '{root.name}{root.accountnumber} files', 'A selection is replaced');
  }
  {
    // Create this folder: Always | Only when…; Only when… seeds one condition; the conditions
    // read as a sentence with short operator labels; Another field… switches the value control.
    const t = await boot();
    await t.open();
    await t.step(2);
    await t.select('General');
    const modes = [...t.$('create-mode').querySelectorAll('[role=radio]')];
    assert.deepEqual(
      modes.map((m) => m.textContent),
      ['Always', 'Only when…'],
    );
    await t.press(modes[1]);
    assert.equal(modes[1].getAttribute('aria-checked'), 'true');
    const conditions = t.$('conditions');
    assert.match(conditions.querySelector('.sentence').visibleText, /of these are true$/);
    assert.equal(conditions.querySelectorAll('fieldset').length, 0, 'No fieldset boxes');
    const operator = t.labelled(conditions, 'Operator, condition 1');
    assert.deepEqual(
      [...operator.querySelectorAll('option')].map((o) => o.textContent),
      ['is', 'is not', 'is empty', 'has a value', 'contains', "doesn't contain"],
    );
    await t.change(t.labelled(conditions, 'Field, condition 1'), 'root.statecode');
    const value = t.labelled(conditions, 'Value, condition 1');
    assert.equal(value.querySelectorAll('option')[0].textContent, 'Another field…');
    await t.change(value, 'value:0');
    assert.equal(t.node('General').querySelector('.rule').textContent, '◆ When Status is Active');
    await t.change(t.labelled(conditions, 'Value, condition 1'), '__field');
    const other = t.labelled(t.$('conditions'), 'Other field, condition 1');
    assert.equal(other.querySelectorAll('option')[0].textContent, 'A value');
    await t.press(t.find(t.$('conditions'), '＋ Condition'));
    assert.match(t.node('General').querySelector('.rule').textContent, /^◆ When .+ \+ 1 more$/);
    await t.press(modes[0]);
    assert.equal(t.node('General').querySelector('.rule').textContent, 'Always');
  }
  {
    // A typed value offers Another field… in its ▾ menu; A value switches back. Arrow keys
    // switch Create this folder and keep focus on it.
    const t = await boot();
    await t.open();
    await t.step(2);
    await t.select('General');
    const always = t.$('create-mode').querySelectorAll('[role=radio]')[0];
    always.focus();
    always.key('ArrowRight');
    await t.document.settle();
    const when = t.$('create-mode').querySelectorAll('[role=radio]')[1];
    assert.equal(when.getAttribute('aria-checked'), 'true');
    assert.equal(when.tabIndex, 0);
    assert.ok(t.document.activeElement === when, 'Focus stays in the radio group');
    assert.ok(t.state().sections[0].Folders[1].Condition, 'Only when… seeds a condition');
    const more = [...t.$('conditions').querySelectorAll('button')].find(
      (b) => b.getAttribute('aria-label') === 'Value options, condition 1',
    );
    assert.equal(more.getAttribute('aria-haspopup'), 'menu');
    await t.press(more);
    await t.press(
      [...t.$('conditions').querySelectorAll('[role=menuitem]')].find(
        (i) => i.textContent === 'Another field…',
      ),
    );
    const other = t.labelled(t.$('conditions'), 'Other field, condition 1');
    assert.ok(t.document.activeElement === other);
    assert.equal(
      t.state().sections[0].Folders[1].Condition.Conditions[0].right,
      'root.accountnumber',
    );
    await t.change(other, '__value');
    assert.equal(t.state().sections[0].Folders[1].Condition.Conditions[0].right, null);
    assert.equal(t.labelled(t.$('conditions'), 'Value, condition 1').tagName, 'INPUT');
    // Always and back by the arrow keys: the conditions come back, never lost to a key press.
    await t.change(t.labelled(t.$('conditions'), 'Value, condition 1'), 'Contoso');
    await t.press(t.find(t.$('conditions'), '＋ Condition'));
    await t.flush();
    when.key('ArrowLeft');
    await t.document.settle();
    assert.equal(t.state().sections[0].Folders[1].Condition, null);
    assert.equal(always.getAttribute('aria-checked'), 'true');
    always.key('ArrowRight');
    await t.document.settle();
    const back = t.state().sections[0].Folders[1].Condition;
    assert.deepEqual(
      Array.from(back.Conditions, (c) => [c.field, c.Literal]),
      [
        ['root.name', 'Contoso'],
        ['root.name', ''],
      ],
      'The conditions return as they were',
    );
    await t.flush();
    const saved = t.last('asx_CreateDraft').Destinations[0].Folders[1].Condition;
    assert.deepEqual(
      saved.Conditions.map((c) => c.Literal),
      ['Contoso', ''],
      'The next save carries them',
    );
    // A click on Always and on Only when… brings them back too.
    await t.mode(0);
    await t.mode(1);
    assert.equal(t.state().sections[0].Folders[1].Condition.Conditions[0].Literal, 'Contoso');
  }
  {
    // ＋ Subfolder stops at the folder depth the server takes, with the reason; ＋ Folder beside
    // the deepest folder still works.
    const d = draft();
    const depth = 10;
    d.Destinations[0].Folders = [
      { Key: 'root', Parent: null, Name: '{root.name}', Condition: null },
      ...Array.from({ length: depth - 1 }, (_, i) => ({
        Key: 'l' + (i + 1),
        Parent: i ? 'l' + i : 'root',
        Name: 'Level ' + (i + 2),
        Condition: null,
      })),
    ];
    const t = await boot({
      loaded: { RevisionId: 'rev-1', RowVersion: '3', Status: 'Published', Version: 1, Draft: d },
    });
    assert.equal(t.window.AsxdUi.BOUNDS.folderDepth, depth);
    await t.open();
    await t.step(2);
    await t.select('Level ' + depth);
    const sub = t.$('add-subfolder');
    assert.equal(sub.getAttribute('aria-disabled'), 'true');
    assert.equal(
      t.document.getElementById(sub.getAttribute('aria-describedby')).textContent,
      'Folders can be nested up to ' + depth + ' levels deep.',
    );
    assert.equal(t.$('add-folder').hasAttribute('aria-disabled'), false);
    await t.press(sub);
    assert.equal(t.state().sections[0].Folders.length, depth, 'Nothing past the depth');
    await t.select('Level ' + (depth - 1));
    assert.equal(t.$('add-subfolder').hasAttribute('aria-disabled'), false);
    await t.press(t.$('add-subfolder'));
    assert.equal(t.state().sections[0].Folders.length, depth + 1);
    assert.equal(
      t.$('add-subfolder').getAttribute('aria-disabled'),
      'true',
      'The new one is as deep',
    );
  }
  {
    // Test records: previews of the draft, debounced; Created or Skipped for the selected folder,
    // with the failing value; a late answer never replaces a newer one.
    const answers = [];
    const t = await boot({
      handle: (api, body) => {
        if (api !== 'asx_PreviewTemplate') return null;
        let resolve;
        const answer = new Promise((r) => (resolve = r));
        answers.push({ body, resolve });
        return answer;
      },
      reads: (table) =>
        table === 'account'
          ? { statecode: 1, 'statecode@OData.Community.Display.V1.FormattedValue': 'Inactive' }
          : null,
    });
    await t.open();
    await t.step(2);
    await t.select('General');
    await t.press(t.$('add-test-record'));
    assert.equal(answers.length, 1, 'A new test record previews at once');
    assert.equal(answers[0].body.RevisionId, '');
    assert.equal(answers[0].body.RecordId, RECORD);
    assert.ok(answers[0].body.Draft, 'The draft payload, not the saved revision');
    const plan = (folders) => ({
      Folders: folders.map((Node) => ({
        Section: 'general',
        Node,
        Name: Node,
        RelativePath: Node,
      })),
      Notices: [],
    });
    answers[0].resolve(plan(['root', 'general_docs']));
    await t.document.settle();
    // Rows are redrawn on every answer: read the row afresh each time, never a captured node.
    const record = () => t.$('test-records').querySelector('.test-record');
    assert.equal(record().querySelector('strong').textContent, 'Contoso Ltd');
    assert.equal(record().querySelector('.state').textContent, 'Created');
    await t.press(t.$('create-mode').querySelectorAll('[role=radio]')[1]);
    await t.change(t.labelled(t.$('conditions'), 'Field, condition 1'), 'root.statecode');
    await t.change(t.labelled(t.$('conditions'), 'Value, condition 1'), 'value:0');
    assert.equal(answers.length, 1, 'Edits wait for the pause');
    await t.flush();
    assert.equal(answers.length, 2, 'One preview per record after the pause');
    await t.change(t.labelled(t.$('conditions'), 'Value, condition 1'), 'value:1');
    await t.flush();
    assert.equal(record().getAttribute('aria-busy'), 'true');
    answers[2].resolve(plan(['root']));
    await t.document.settle();
    assert.equal(record().querySelector('.state').textContent, 'Skipped');
    answers[1].resolve(plan(['root', 'general_docs']));
    await t.document.settle();
    assert.equal(
      record().querySelector('.state').textContent,
      'Skipped',
      'The older answer is ignored',
    );
    assert.equal(record().querySelector('.why').textContent, 'Inactive');
    assert.equal(record().hasAttribute('aria-busy'), false);
    // The top folder was created: no failing value there.
    await t.select('Account Name');
    assert.equal(record().querySelector('.state').textContent, 'Created');
    assert.ok(record().querySelector('.why') === null, 'No value shown');
    // A folder under a skipped folder is skipped because of it: no value of its own.
    await t.select('General');
    await t.press(t.$('add-subfolder'));
    await t.mode(1);
    await t.change(t.labelled(t.$('conditions'), 'Field, condition 1'), 'root.statecode');
    assert.equal(record().querySelector('.state').textContent, 'Skipped');
    assert.ok(record().querySelector('.why') === null, 'No value shown');
  }
  {
    // Test records: at most the preview bound, with its reason; a record is listed once; ✕
    // removes it; a draft that cannot be previewed marks the rows out of date without a call; a
    // failed preview says why; another template starts with none.
    const ids = ['a1', 'a2', 'a3', 'a4', 'a5'].map((n) => '00000000-0000-0000-0000-0000000000' + n);
    let next = 0;
    let fail = false;
    const t = await boot({
      enabled: ['account', 'contact'],
      pick: () => [
        {
          id: '{' + ids[next].toUpperCase() + '}',
          name: 'Account ' + (next + 1),
          entityType: 'account',
        },
      ],
      handle: (api) =>
        api === 'asx_PreviewTemplate' && fail ? new Error('The record could not be read.') : null,
    });
    await t.open();
    await t.step(2);
    await t.select('General');
    const rows = () => [...t.$('test-records').querySelectorAll('.test-record')];
    await t.press(t.$('add-test-record'));
    await t.press(t.$('add-test-record'));
    assert.equal(rows().length, 1, 'Listed once');
    for (next = 1; next < 5; next++) await t.press(t.$('add-test-record'));
    assert.equal(rows().length, 5);
    assert.deepEqual(
      rows().map((r) => r.querySelector('.state').textContent),
      ['Created', 'Created', 'Created', 'Created', 'Created'],
    );
    const add = t.$('add-test-record');
    assert.equal(add.getAttribute('aria-disabled'), 'true');
    assert.equal(
      t.document.getElementById(add.getAttribute('aria-describedby')).textContent,
      "Preview covers up to 5 records at a time so it finishes within Dataverse's 2-minute limit.",
    );
    const looked = t.looked.length;
    await t.press(add);
    assert.equal(t.looked.length, looked, 'No picker at the bound');
    await t.press(t.find(rows()[4], '✕', 'Remove Account 5 from test records'));
    assert.equal(rows().length, 4);
    assert.equal(add.hasAttribute('aria-disabled'), false);
    const previews = () => t.sent.filter(([a]) => a === 'asx_PreviewTemplate').length;
    const before = previews();
    await t.mode(1);
    await t.change(t.labelled(t.$('conditions'), 'Field, condition 1'), 'root.revenue');
    await t.change(t.labelled(t.$('conditions'), 'Value, condition 1'), '1e5');
    await t.flush();
    assert.equal(previews(), before, 'Nothing is sent that the server would refuse');
    assert.equal(rows()[0].querySelector('.state').textContent, 'Out of date');
    fail = true;
    await t.change(t.labelled(t.$('conditions'), 'Value, condition 1'), '5');
    await t.flush();
    assert.equal(previews(), before + 4, 'One preview per record');
    assert.equal(
      rows()[0].querySelector('.state').textContent,
      'Preview failed: The record could not be read.',
    );
    await t.press(t.$('new-template'));
    await t.change(t.$('new-template-table'), 'contact');
    assert.equal(t.state().testRecords.length, 0);
  }
  {
    // Undo puts a removed folder back with its fields under the draft's own aliases, after its
    // removed parent; a removed folder's removed child shows under it; edited folders have a dot.
    const published = draft({
      Sources: [
        {
          Alias: 'root',
          Table: 'account',
          Lookup: null,
          Columns: [{ Name: 'name', Kind: 'Text' }],
        },
        {
          Alias: 'lookup_2',
          Table: 'contact',
          Lookup: 'primarycontactid',
          Columns: [{ Name: 'fullname', Kind: 'Text' }],
        },
      ],
    });
    published.Destinations[0].Folders.push(
      { Key: 'contacts', Parent: 'root', Name: '{lookup_2.fullname}', Condition: null },
      { Key: 'archive', Parent: 'contacts', Name: 'Archive', Condition: null },
    );
    const t = await boot({
      rows: (table) =>
        table === 'asx_revision'
          ? [
              {
                asx_revisionid: 'rev-2',
                asx_version: 2,
                asx_status: 'Draft',
                _asx_templateid_value: TEMPLATE,
              },
              {
                asx_revisionid: 'rev-1',
                asx_version: 1,
                asx_status: 'Published',
                _asx_templateid_value: TEMPLATE,
              },
            ]
          : null,
      handle: (api, body) =>
        api !== 'asx_LoadDraft'
          ? null
          : body.RevisionId === 'rev-1'
            ? {
                RevisionId: 'rev-1',
                RowVersion: '3',
                Status: 'Published',
                Version: 1,
                Draft: published,
              }
            : { RevisionId: 'rev-2', RowVersion: '4', Status: 'Draft', Version: 2, Draft: draft() },
    });
    await t.open();
    await t.step(2);
    const removed = [...t.$('folder-tree').querySelectorAll('.tree-row.is-removed')];
    assert.deepEqual(
      removed.map((r) => r.querySelector('.name').textContent),
      ['Primary Contact › Full Name', 'Archive'],
    );
    assert.ok(
      parseInt(removed[1].style.paddingLeft) > parseInt(removed[0].style.paddingLeft),
      'Under its removed parent',
    );
    assert.equal(t.find(removed[1], 'Undo').getAttribute('aria-label'), 'Undo removing Archive');
    await t.press(t.find(removed[1], 'Undo'));
    assert.deepEqual(
      Array.from(t.state().sections[0].Folders.slice(-2), (f) => [f.Key, f.Parent, f.Name]),
      [
        ['contacts', 'root', '{lookup_1.fullname}'],
        ['archive', 'contacts', 'Archive'],
      ],
    );
    assert.equal(t.$('folder-tree').querySelectorAll('.tree-row.is-removed').length, 0);
    assert.equal(t.document.activeElement.getAttribute('data-focus-key'), 'node:general:archive');
    await t.flush();
    assert.deepEqual(
      t.last('asx_CreateDraft').Sources.map((s) => s.Alias),
      ['root', 'lookup_1'],
    );
    await t.rename('General', 'General files');
    assert.equal(t.node('General files').querySelector('.edit-dot').textContent, ' (edited)');
  }
  {
    // Undo stops at the folder bound, counting the removed parent it puts back first.
    const bound = 100;
    const published = draft();
    published.Destinations[0].Folders.push(
      { Key: 'contacts', Parent: 'root', Name: 'Contacts', Condition: null },
      { Key: 'archive', Parent: 'contacts', Name: 'Archive', Condition: null },
    );
    const current = draft();
    for (let i = 2; i < bound - 1; i++)
      current.Destinations[0].Folders.push({
        Key: 'f' + i,
        Parent: 'root',
        Name: 'F' + i,
        Condition: null,
      });
    const t = await boot({
      rows: (table) =>
        table === 'asx_revision'
          ? [
              {
                asx_revisionid: 'rev-2',
                asx_version: 2,
                asx_status: 'Draft',
                _asx_templateid_value: TEMPLATE,
              },
              {
                asx_revisionid: 'rev-1',
                asx_version: 1,
                asx_status: 'Published',
                _asx_templateid_value: TEMPLATE,
              },
            ]
          : null,
      handle: (api, body) =>
        api !== 'asx_LoadDraft'
          ? null
          : body.RevisionId === 'rev-1'
            ? {
                RevisionId: 'rev-1',
                RowVersion: '3',
                Status: 'Published',
                Version: 1,
                Draft: published,
              }
            : {
                RevisionId: 'rev-2',
                RowVersion: '4',
                Status: 'Draft',
                Version: 2,
                Draft: current,
              },
    });
    assert.equal(t.window.AsxdUi.BOUNDS.foldersPerDestination, bound);
    await t.open();
    await t.step(2);
    assert.equal(t.state().sections[0].Folders.length, bound - 1);
    const undo = (name) =>
      [...t.$('folder-tree').querySelectorAll('button')].find(
        (b) => b.getAttribute('aria-label') === 'Undo removing ' + name,
      );
    // Archive needs Contacts back too: two folders where one fits.
    assert.equal(undo('Archive').getAttribute('aria-disabled'), 'true');
    assert.equal(
      t.document.getElementById(undo('Archive').getAttribute('aria-describedby')).textContent,
      'A destination can have up to ' +
        bound +
        " folders, because a record's folders for one destination are kept in one Dataverse row.",
    );
    await t.press(undo('Archive'));
    assert.equal(t.state().sections[0].Folders.length, bound - 1, 'Nothing past the bound');
    assert.equal(undo('Contacts').hasAttribute('aria-disabled'), false);
    await t.press(undo('Contacts'));
    assert.equal(t.state().sections[0].Folders.length, bound);
    assert.equal(undo('Archive').getAttribute('aria-disabled'), 'true');
    assert.equal(undo('Archive').getAttribute('aria-describedby'), 'folder-reason');
    assert.equal(t.$('add-folder').getAttribute('aria-describedby'), 'folder-reason');
  }
  {
    // Review: the change list against the published version, each line linking to its step.
    const t = await boot();
    await t.open();
    await t.step(2);
    await t.select('General');
    await t.press(t.$('add-subfolder'));
    t.$('folder-name').value = 'P-{root.accountnumber}';
    t.$('folder-name').oninput();
    await t.flush();
    await t.step(3);
    assert.equal(t.$('review-title').textContent, 'Changes since v1');
    const line = t.$('change-list').querySelector('.change-row');
    assert.equal(line.querySelector('.mark').textContent, '＋');
    assert.equal(line.querySelector('.mark').getAttribute('aria-hidden'), 'true');
    assert.equal(
      line.querySelector('.text').visibleText,
      'Folder General › P-Account Number added to Business documents',
    );
    assert.equal(line.querySelector('.text strong .token').textContent, 'Account Number');
    const link = t.find(line, 'Folders');
    assert.equal(link.getAttribute('aria-label'), 'Show P-Account Number in Folders');
    await t.press(link);
    assert.equal(t.$('step-2').hidden, false);
    assert.equal(t.document.activeElement.dataset.focusKey, 'node:general:folder_1');
    // A removed folder links to its Undo; a destination change links to its card.
    const r = await boot();
    await r.open();
    await r.step(2);
    await r.select('General');
    await r.press(r.find(r.$('folder-menu-list'), 'Remove folder'));
    await r.change(r.labelled(r.$('step-1'), 'Name'), 'Client files');
    await r.step(3);
    const rows = [...r.$('change-list').querySelectorAll('.change-row')];
    assert.deepEqual(
      rows.map((row) => row.querySelector('.mark').textContent),
      ['◆', '−'],
      'Destinations first, then folders',
    );
    assert.equal(
      rows[1].querySelector('.text').visibleText,
      'Folder General removed from Client files. Existing General folders stay in SharePoint.',
    );
    await r.press(r.find(rows[1], 'Folders'));
    assert.equal(r.document.activeElement.dataset.focusKey, 'undo:general:general_docs');
    await r.step(3);
    await r.press(r.find(r.$('change-list').querySelectorAll('.change-row')[0], 'Destinations'));
    assert.equal(r.$('step-1').hidden, false);
    assert.equal(r.document.activeElement.dataset.focusKey, 'dest:general:name');
  }
  {
    // Publishing: the three consequences, the re-run box with generated text, Starts, the
    // paused warning, and Publish v2 and re-run, which returns to the overview.
    const t = await boot({ runtime: { Enabled: false, ProcessRecordUpdates: true } });
    await t.open();
    await t.step(2);
    await t.select('General');
    await t.press(t.$('add-subfolder'));
    t.$('folder-name').value = 'Projects';
    t.$('folder-name').oninput();
    await t.flush();
    await t.step(3);
    assert.equal(t.$('publishing-title').textContent, 'Publishing v2');
    assert.deepEqual(
      [...t.$('consequences').querySelectorAll('li')].map((li) => li.visibleText),
      [
        'New Account records get v2 folders from now on.',
        'Changed records are updated, because “Update folders when records change” is on.',
        'About 1,240 existing Account records keep their v1 folders until they are re-run. Folders are never removed.',
      ],
    );
    assert.deepEqual(
      [...t.$('consequences').querySelectorAll('li')].map(
        (li) => li.querySelector('.dot').dataset.tone,
      ),
      ['ok', 'ok', 'pending'],
    );
    assert.equal(
      t.$('consequences').querySelectorAll('li')[0].querySelector('strong').textContent,
      'New Account records',
    );
    assert.equal(t.$('publish-rerun').checked, true);
    assert.equal(
      t.$('publish-rerun-desc').textContent,
      'Adds Projects folders to existing accounts. Progress shows in Monitor.',
    );
    assert.equal(t.$('publish-paused').hidden, false);
    assert.equal(
      t.$('publish-paused').textContent,
      "Automation is paused, so no folders are created until it's on.",
    );
    assert.equal(t.$('publish').hidden, false);
    assert.equal(t.$('step-next').hidden, true);
    assert.equal(t.$('publish').textContent, 'Publish v2 and re-run');
    t.$('publish-rerun').checked = false;
    t.$('publish-rerun').onchange();
    assert.equal(t.$('publish').textContent, 'Publish v2');
    t.$('publish-rerun').checked = true;
    t.$('publish-rerun').onchange();
    await t.press(t.$('publish'));
    assert.equal(
      t.document.querySelectorAll('.confirm').length,
      0,
      'This page is the confirmation',
    );
    const order = t.sent
      .map(([a, b]) => b?.Command || a)
      .filter((x) => ['asx_PublishTemplate', 'StartTemplateRun'].includes(x));
    assert.deepEqual(order, ['asx_PublishTemplate', 'StartTemplateRun']);
    assert.deepEqual(t.last('asx_PublishTemplate'), { RevisionId: 'rev-2', RowVersion: '4' });
    assert.equal(t.$('template-overview').hidden, false);
    assert.equal(t.$('template-editor').hidden, true);
    assert.equal(t.$('fb-templates').visibleText.startsWith('Published v2. Re-run started.'), true);
    assert.equal(t.document.activeElement, t.$('overview-title'));
    await t.press(t.find(t.$('fb-templates'), 'Follow it in Monitor →'));
    assert.equal(t.sent.filter(([k]) => k === 'navigate').at(-1)[1].data, 'monitor-' + BUILD);
    // Changed records are not updated when record updates are off; the line goes when automation
    // cannot be read.
    const off = await boot({ runtime: { ProcessRecordUpdates: false } });
    await off.open();
    await off.change(off.labelled(off.$('step-1'), 'Name'), 'Client files');
    await off.step(3);
    assert.equal(
      off.$('consequences').querySelectorAll('li')[1].visibleText,
      'Changed records are not updated, because “Update folders when records change” is off.',
    );
    assert.equal(off.$('publish-paused').hidden, true);
  }
  {
    // Starts Later writes the start date after publishing and turns the re-run off with its reason.
    const t = await boot();
    await t.open();
    await t.change(t.labelled(t.$('step-1'), 'Name'), 'Client files');
    await t.flush();
    await t.step(3);
    // No folder added: the re-run text says it applies the changes.
    assert.equal(
      t.$('publish-rerun-desc').textContent,
      'Applies these changes to existing accounts. Progress shows in Monitor.',
    );
    assert.equal(t.$('publish-starts').value, 'now');
    assert.equal(t.$('publish-start-label').hidden, true);
    await t.change(t.$('publish-starts'), 'later');
    assert.equal(t.$('publish-start-label').hidden, false);
    assert.equal(t.$('publish-start').hidden, false);
    assert.equal(t.$('publish-rerun').checked, false);
    assert.equal(t.$('publish-rerun').getAttribute('aria-disabled'), 'true');
    assert.equal(
      t.document.getElementById(
        t.$('publish-rerun').getAttribute('aria-describedby').split(' ').at(-1),
      ).textContent,
      'Re-run after the template starts, from Re-run for existing records.',
    );
    assert.equal(t.$('publish').textContent, 'Publish v2');
    // A click on the blocked box does not tick it.
    const click = new FakeEvent('click');
    t.$('publish-rerun').dispatchEvent(click);
    assert.equal(click.defaultPrevented, true);
    t.$('publish-start').value = '2026-12-01T09:00';
    t.$('publish-start').onchange();
    assert.match(t.$('publish-start-label').textContent, /^Start/);
    await t.press(t.$('publish'));
    const update = t.sent.filter(([k]) => k === 'update').at(-1);
    assert.equal(update[1], 'asx_template');
    assert.deepEqual(Object.keys(update[2]), ['asx_startsutc']);
    assert.equal(update[2].asx_startsutc, new Date('2026-12-01T09:00').toISOString());
    assert.equal(t.sent.filter(([, b]) => b?.Command === 'StartTemplateRun').length, 0);
    assert.match(t.$('fb-templates').visibleText, /^Published v2\.$/);
    // A stored start ahead opens on Later with its time; Now clears it, before the re-run.
    const later = '2027-01-04T08:30:00.000Z';
    const s = await boot({
      templates: [
        {
          asx_templateid: TEMPLATE,
          asx_name: 'Account onboarding',
          asx_table: 'account',
          _asx_publishedrevisionid_value: 'rev-1',
          asx_disabled: false,
          asx_startsutc: later,
        },
      ],
    });
    await s.open();
    await s.change(s.labelled(s.$('step-1'), 'Name'), 'Client files');
    await s.step(3);
    assert.equal(s.$('publish-starts').value, 'later');
    assert.equal(new Date(s.$('publish-start').value).toISOString(), later);
    assert.equal(s.$('publish-rerun').checked, false);
    await s.change(s.$('publish-starts'), 'now');
    assert.equal(s.$('publish-rerun').checked, true, 'The box comes back as it was');
    assert.equal(s.$('publish-rerun').hasAttribute('aria-disabled'), false);
    await s.press(s.$('publish'));
    const steps = s.sent
      .map(([a, b, c]) => (a === 'update' ? 'update:' + c.asx_startsutc : b?.Command || a))
      .filter((x) => /^(asx_PublishTemplate|StartTemplateRun|update:)/.test(x));
    assert.deepEqual(steps, ['asx_PublishTemplate', 'update:null', 'StartTemplateRun']);
    // Later without a date publishes nothing and says what is missing.
    const u = await boot();
    await u.open();
    await u.change(u.labelled(u.$('step-1'), 'Name'), 'Client files');
    await u.step(3);
    await u.change(u.$('publish-starts'), 'later');
    await u.press(u.$('publish'));
    assert.equal(u.last('asx_PublishTemplate'), undefined);
    assert.equal(u.$('fb-editor').textContent, 'Pick a date');
    assert.equal(u.$('step-3').hidden, false);
  }
  {
    // A refused re-run after a successful publish lands on the overview with the error and the
    // existing action.
    const t = await boot({
      handle: (api, body) =>
        body?.Command === 'StartTemplateRun'
          ? new Error('Account is not enabled for Documents.')
          : null,
    });
    await t.open();
    await t.change(t.labelled(t.$('step-1'), 'Name'), 'Client files');
    await t.flush();
    await t.step(3);
    await t.press(t.$('publish'));
    assert.equal(t.$('template-overview').hidden, false);
    assert.equal(t.$('fb-templates').getAttribute('role'), 'alert');
    assert.match(
      t.$('fb-templates').visibleText,
      /^Published v2\. The re-run didn't start: Account is not enabled for Documents\./,
    );
    assert.ok(t.find(t.$('fb-templates'), 'Re-run existing records…'));
    assert.equal(t.$('step-3').hidden, true, 'Not left on the publish step to publish again');
    await t.press(t.find(t.$('fb-templates'), 'Re-run existing records…'));
    assert.equal(t.$('rerun-panel').hidden, false);
    // A refused publish stays on the step with the server's reason, and starts nothing.
    const p = await boot({
      handle: (api) =>
        api === 'asx_PublishTemplate'
          ? new Error('Draft changed; reload before publishing.')
          : null,
    });
    await p.open();
    await p.change(p.labelled(p.$('step-1'), 'Name'), 'Client files');
    await p.step(3);
    await p.press(p.$('publish'));
    assert.equal(p.$('step-3').hidden, false);
    assert.equal(p.$('fb-editor').textContent, 'Draft changed; reload before publishing.');
    assert.equal(p.sent.filter(([, b]) => b?.Command === 'StartTemplateRun').length, 0);
    assert.equal(p.sent.filter(([k]) => k === 'update').length, 0);
  }
  {
    // Without the Operator role: no count in the third line and the re-run box is disabled.
    const t = await boot({ privileges: { prvCreateasx_operatorcommand: false } });
    await t.open();
    await t.change(t.labelled(t.$('step-1'), 'Name'), 'Client files');
    await t.flush();
    await t.step(3);
    assert.equal(
      t.$('consequences').querySelectorAll('li')[2].visibleText,
      'Existing Account records keep their v1 folders until they are re-run. Folders are never removed.',
    );
    assert.equal(t.$('publish-rerun').checked, false);
    assert.equal(t.$('publish-rerun').getAttribute('aria-disabled'), 'true');
    assert.equal(t.$('publish-rerun-reason').textContent, 'Needs the Documents Operator role.');
    assert.equal(t.$('publish').textContent, 'Publish v2');
    assert.equal(t.sent.filter(([, b]) => b?.Command === 'CountRecords').length, 0);
    // A refused count reads as no count.
    const n = await boot({
      handle: (api, body) =>
        body?.Command === 'CountRecords'
          ? new Error('Account is not enabled for Documents.')
          : null,
    });
    await n.open();
    await n.change(n.labelled(n.$('step-1'), 'Name'), 'Client files');
    await n.step(3);
    assert.match(
      n.$('consequences').querySelectorAll('li')[2].visibleText,
      /^Existing Account records keep their v1 folders/,
    );
    assert.equal(n.$('publish-rerun').checked, true);
  }
  {
    // Narrow screens: the templates list is also offered as a select above the overview.
    const t = await boot({
      templates: [
        {
          asx_templateid: TEMPLATE,
          asx_name: 'Account onboarding',
          asx_table: 'account',
          _asx_publishedrevisionid_value: 'rev-1',
          asx_disabled: false,
        },
        {
          asx_templateid: OTHER,
          asx_name: 'Contracts',
          asx_table: 'account',
          _asx_publishedrevisionid_value: null,
          asx_disabled: false,
        },
      ],
    });
    const picker = t.$('template-picker');
    assert.equal(picker.closest('.narrow-only').childNodes[0].textContent, 'Template');
    assert.deepEqual(
      [...picker.querySelectorAll('option')].map((o) => o.textContent),
      ['Account onboarding', 'Contracts'],
    );
    assert.equal(picker.value, TEMPLATE);
    await t.change(picker, OTHER);
    assert.equal(t.$('overview-title').textContent, 'Contracts');
    assert.equal(picker.value, OTHER);
    await t.change(picker, TEMPLATE);
    assert.equal(t.$('overview-title').textContent, 'Account onboarding');
  }
  {
    // changeText for every kind the change list shows, read against the open draft
    // (destination "Business documents", key "general").
    const t = await boot();
    await t.open();
    const { changeText, state } = t.window.AsxdAdmin;
    const library = state().sections[0].LibraryId;
    const text = (entry) => changeText(entry).visibleText;
    const folder = { destination: 'general', key: 'general_docs', name: 'General' };
    const cases = [
      [{ ...folder, kind: 'added' }, 'Folder General added to Business documents'],
      [
        { ...folder, kind: 'removed' },
        'Folder General removed from Business documents. Existing General folders stay in SharePoint.',
      ],
      [
        { ...folder, kind: 'renamed', before: 'Docs', after: 'General' },
        'Folder Docs renamed to General in Business documents',
      ],
      [
        { ...folder, kind: 'moved', before: 'other', after: 'root' },
        'Folder General moved under Account Name in Business documents',
      ],
      [{ ...folder, kind: 'condition' }, 'General now always created'],
      [
        { kind: 'added', key: 'general', name: 'Business documents' },
        'Destination Business documents added',
      ],
      [
        { kind: 'removed', key: 'general', name: 'Business documents' },
        'Destination Business documents removed. Its folders stay in SharePoint.',
      ],
      [
        {
          kind: 'renamed',
          key: 'general',
          name: 'Business documents',
          before: 'Client files',
          after: 'Business documents',
        },
        'Destination Client files renamed to Business documents',
      ],
      [
        {
          kind: 'library',
          key: 'general',
          name: 'Business documents',
          before: 'lib-old',
          after: library,
        },
        'Destination Business documents now uses Delivery › General',
      ],
    ];
    for (const [entry, expected] of cases) assert.equal(text(entry), expected, entry.kind);
    // With a condition on the draft folder, the condition line reads the rule sentence.
    state().sections[0].Folders.find((f) => f.Key === 'general_docs').Condition = {
      All: true,
      Groups: [],
      Conditions: [{ field: 'root.statecode', Operator: 'Equal', Literal: '0', right: null }],
    };
    assert.equal(
      text({ ...folder, kind: 'condition' }),
      'General now created only when Status is Active',
    );
    assert.ok(
      changeText({ ...folder, kind: 'renamed', before: 'Docs', after: 'General' }).querySelectorAll(
        'strong',
      ).length >= 2,
    );
  }
  {
    // First publish (no published revision): no change list, and the third consequence says
    // existing records get folders when re-run.
    const t = await boot({
      templates: [
        {
          asx_templateid: TEMPLATE,
          asx_name: 'Account onboarding',
          asx_table: 'account',
          _asx_publishedrevisionid_value: null,
          asx_disabled: false,
        },
      ],
      loaded: { RevisionId: 'rev-1', RowVersion: '3', Status: 'Draft', Version: 1, Draft: draft() },
    });
    await t.open();
    await t.step(3);
    assert.equal(t.$('review-title').hidden, true, 'No change list heading');
    assert.equal(t.$('change-list').hidden, true, 'No change list');
    assert.equal(t.$('publishing-title').textContent, 'Publishing v1');
    assert.equal(
      t.$('consequences').querySelectorAll('li')[2].visibleText,
      'About 1,240 existing Account records get folders when they are re-run.',
    );
    assert.equal(t.$('publish').textContent, 'Publish v1 and re-run');
  }
  {
    // A published version that could not be read is not a first publish: the server's message
    // takes the change list's place, Publish stays available, and existing records keep v1.
    let failing = false;
    const t = await boot({
      rows: (table) =>
        table === 'asx_revision'
          ? [
              {
                asx_revisionid: 'rev-2',
                asx_version: 2,
                asx_status: 'Draft',
                _asx_templateid_value: TEMPLATE,
              },
              {
                asx_revisionid: 'rev-1',
                asx_version: 1,
                asx_status: 'Published',
                _asx_templateid_value: TEMPLATE,
              },
            ]
          : null,
      handle: (api, body) => {
        if (api !== 'asx_LoadDraft') return null;
        if (body.RevisionId === 'rev-2')
          return {
            RevisionId: 'rev-2',
            RowVersion: '4',
            Status: 'Draft',
            Version: 2,
            Draft: draft(),
          };
        return failing ? new Error('The published version could not be read.') : null;
      },
    });
    await t.overview();
    failing = true;
    await t.press(t.$('overview-edit'));
    await t.step(3);
    assert.equal(t.$('review-title').hidden, false);
    assert.equal(t.$('review-title').textContent, 'Changes since v1');
    assert.equal(t.$('change-list').textContent, 'The published version could not be read.');
    assert.match(
      t.$('consequences').querySelectorAll('li')[2].visibleText,
      /keep their v1 folders until they are re-run/,
    );
    assert.equal(t.$('publish').hasAttribute('aria-disabled'), false);
    assert.equal(t.$('publish').textContent, 'Publish v2 and re-run');
  }
  {
    // Result for: with no test record a Choose record… button adds one; its preview of the
    // draft shows one card per destination, the rows indented by depth.
    const t = await boot();
    await t.open();
    await t.step(3);
    assert.equal(t.$('result-record').hidden, true);
    assert.equal(t.$('result-choose').textContent.trim(), 'Choose record…');
    await t.press(t.$('result-choose'));
    assert.equal(t.$('result-record').hidden, false);
    assert.equal(t.$('result-choose').hidden, true);
    assert.deepEqual(
      [...t.$('result-record').querySelectorAll('option')].map((o) => o.textContent),
      ['Contoso Ltd'],
    );
    const request = t.last('asx_PreviewTemplate');
    assert.equal(request.RevisionId, '');
    assert.equal(request.RecordId, RECORD);
    assert.equal(t.$('preview-status').textContent, 'Preview updated: 1 destination, 2 folders.');
    const card = t.$('previewTrees').querySelector('.preview-card');
    assert.equal(card.querySelector('.eyebrow').textContent, 'Business documents · Delivery');
    const nodes = [...card.querySelectorAll('.preview-node')];
    assert.deepEqual(
      nodes.map((n) => [n.visibleText, n.style.paddingLeft]),
      [
        ['Contoso Ltd', '0px'],
        ['General', '18px'],
      ],
    );
    // The test record is the same one step 2 shows.
    await t.step(2);
    assert.equal(t.$('test-record-list').querySelectorAll('.test-record').length, 1);
  }
  {
    // A start date the server refuses after the version was published: the overview says so
    // with the server's reason, the re-run does not start, and Re-run existing records… is
    // offered.
    const later = '2027-01-04T08:30:00.000Z';
    const t = await boot({
      templates: [
        {
          asx_templateid: TEMPLATE,
          asx_name: 'Account onboarding',
          asx_table: 'account',
          _asx_publishedrevisionid_value: 'rev-1',
          asx_disabled: false,
          asx_startsutc: later,
        },
      ],
      update: () => new Error('The template could not be updated.'),
    });
    await t.open();
    await t.change(t.labelled(t.$('step-1'), 'Name'), 'Client files');
    await t.step(3);
    await t.change(t.$('publish-starts'), 'now');
    assert.equal(t.$('publish-rerun').checked, true);
    await t.press(t.$('publish'));
    assert.equal(t.last('asx_PublishTemplate')?.RevisionId, 'rev-2');
    assert.equal(t.$('template-overview').hidden, false);
    assert.equal(t.$('fb-templates').getAttribute('role'), 'alert');
    assert.match(
      t.$('fb-templates').visibleText,
      /^Published v2\. The template could not be updated\./,
    );
    assert.equal(t.sent.filter(([, b]) => b?.Command === 'StartTemplateRun').length, 0);
    assert.ok(t.find(t.$('fb-templates'), 'Re-run existing records…'));
  }
  {
    // An install with a table and no templates yet: ＋ New and the Template select stay, so a
    // narrow screen can start the first template.
    const t = await boot({ templates: [] });
    assert.equal(t.$('new-template').hidden, false);
    assert.equal(t.$('template-picker').closest('label').hidden, false);
    await t.press(t.$('new-template'));
    assert.equal(t.$('template-editor').hidden, false);
  }
  console.log(
    'PASS Folder templates: empty states, the templates list with states and search, the overview (pill, meta, cards, chips, rule sentences, versions, last re-run, team counts, problem pill, roles), Edit template and Close, View read-only, ＋ New, the editor header and its Draft pill, Publish and its reasons, unsaved-changes prompts, the ⋯ menu, Delete and focus after it, Schedule and All versions side panels, Manage tables, focus after a keyboard pick, Re-run in progress or Last re-run, the ⋯ separator, the ＋ New table picker and its unsaved-changes prompt, folders and focus, Insert field, condition builder and its depth bound, lookup labels, preview of edits, Re-run all with exact and estimated totals; fix round 1: Save after Publish, related tables loaded four at a time after the first render with a retry and unavailable groups, saved version numbers, no second template after a failed reload, unavailable fields unnamed, numbers as typed, the Operator reason on every re-run action; Task 9: stale pickers redraw once focus leaves the field, and a template switch stops the old preload; Task 6: the stepper and its keys, single-flight autosave with the saved row version, one new template under an edit in flight, a failed save and Retry, invalid conditions blocking the autosave, the unsaved-changes prompt over an unsaved or in-flight save, changesSince and nextKey, change dots and counts, step 1 cards, rows, team panel, library setup option, Add destination and its bound, a link to a step; Task 6 fix round 1: a new template unnamed until named and not saved mid-name, saving before switching or closing, a late save answer kept out of a new template, an edit during Publish kept and saved next, a published read retried and its failure, a failed policy read, trimmed destination names, problems not announced twice, and Save draft refusing a nameless template; Task 7: the folder tree with destination pills, rules, New, edited and Removed with Undo (aliases mapped, a removed parent first), ＋ Folder and ＋ Subfolder with their bound, the folder panel and its ⋯ menu, ＋ Field inserting at the caret, Create this folder, the condition sentence with Another field…, debounced test-record previews with Created or Skipped and the failing value, their bound, Out of date and failures, and links to Sites & access that save first; Task 7 fix round 1: conditions kept across Always and back, Undo at the folder bound counting removed parents, and ＋ Subfolder at the folder depth; Task 8: Review and publish (the change list with every kind in words and links to its step, Result for with Choose record…, the consequences with the record count, record updates and the first-publish wording, a published version that could not be read, the re-run box and its role and Later reasons, Starts with a stored start, the paused warning, Publish and re-run ending on the overview, a refused publish, a refused re-run after a publish, edits made during Publish saved or kept) and the narrow-screen Template select. Fake DOM; browser QA separate.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
