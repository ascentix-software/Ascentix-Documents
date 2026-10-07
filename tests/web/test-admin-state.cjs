'use strict';
// Folder templates with a fake DOM and mocked Dataverse: the templates list and the overview,
// its ⋯ menu, Schedule and Versions panels, Delete, the editor's template bar and version chip,
// Publish, folders, Insert field, the condition builder, preview of unsaved edits, focus, and
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
} = {}) {
  const document = createDocument(html);
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
      updateRecord: async (table, id, data) => sent.push(['update', table, data]),
      deleteRecord: async (table, id) => deleted.push([table, id]),
      online: {
        execute: async (request) => {
          const api = request.getMetadata().operationName;
          const body = request.Request
            ? JSON.parse(request.Request)
            : { RevisionId: request.RevisionId, RowVersion: request.RowVersion };
          sent.push([api, body]);
          const custom = handle(api, body);
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
                  },
                  {
                    LogicalName: 'contact',
                    EntitySetName: 'contacts',
                    PrimaryIdAttribute: 'contactid',
                    PrimaryNameAttribute: 'fullname',
                    DisplayName: { UserLocalizedLabel: { Label: 'Contact' } },
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
    setTimeout,
    clearTimeout,
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
    root
      .querySelectorAll('button')
      .find(
        (b) => b.textContent === text && (label === null || b.getAttribute('aria-label') === label),
      );
  const labelled = (root, name) =>
    root
      .querySelectorAll('input, select, textarea')
      .find(
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
    $('template-groups')
      .querySelectorAll('button')
      .find(
        (b) =>
          b.classList.contains('list-row') && b.querySelector('.row-name')?.textContent === name,
      );
  const overview = async (name = 'Account onboarding') => press(row(name));
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
    assert.equal(t.$('template-bar').hidden, false);
    assert.equal(t.$('editor-title').textContent, 'Account onboarding');
    assert.doesNotMatch(t.$('templates').visibleText, GUID);
    assert.doesNotMatch(t.$('templates').visibleText, /rev-1/);
    assert.match(
      t.$('destinations').visibleText,
      /\[Account Name\]/,
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
    const groups = t.$('template-groups').querySelectorAll('.list-group');
    assert.deepEqual(
      groups.map((g) => g.querySelector('.group-label').textContent),
      ['Account', 'Contact'],
    );
    const states = t
      .$('template-groups')
      .querySelectorAll('.list-row')
      .map((r) => [
        r.querySelector('.row-name').textContent,
        r.querySelector('.row-state').textContent,
      ]);
    assert.deepEqual(states, [
      ['Account onboarding', 'Live v1'],
      ['Key accounts', 'Draft v1'],
      ['Contact files', 'Off'],
    ]);
    assert.deepEqual(
      t
        .$('template-groups')
        .querySelectorAll('.row-state')
        .map((s) => s.getAttribute('data-tone')),
      ['ok', 'warning', 'muted'],
    );
    t.$('template-search').value = 'key';
    t.$('template-search').oninput();
    assert.deepEqual(
      t
        .$('template-groups')
        .querySelectorAll('.list-row')
        .map((r) => r.querySelector('.row-name').textContent),
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
    const cards = t.$('overview-cards').querySelectorAll('.section-card');
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
    const versions = t.$('overview-cards').querySelectorAll('.version-row');
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
      t
        .$('overview-menu-list')
        .querySelectorAll('[role=menuitem]')
        .map((i) => i.textContent.trim()),
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
    // Close with unsaved edits asks first, with the page's unsaved-changes prompt; Discard
    // changes closes, and Edit template opens the saved version again.
    await t.press(t.$('overview-edit'));
    await t.press(t.find(t.$('destinations'), 'General'));
    await t.change(t.labelled(t.$('folderEditor'), 'Folder name'), 'Dropped');
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
    assert.doesNotMatch(t.$('destinations').visibleText, /Dropped/);
    assert.equal(t.$('version-chip').textContent, 'Published v1');
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
    // Read-only: Save draft and Publish say why they are off.
    assert.equal(t.$('save').getAttribute('aria-disabled'), 'true');
    assert.equal(t.$('publish-reason').textContent, 'Viewing an earlier version.');
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
    const rows = t.$('overview-cards').querySelectorAll('.folder-row');
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
    assert.equal(t.$('editor-title').textContent, 'New template');
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
    // Fix round 1: ＋ New with unsaved edits in the editor asks first; Keep editing keeps them.
    const t = await boot();
    await t.open();
    await t.press(t.find(t.$('destinations'), 'General'));
    await t.change(t.labelled(t.$('folderEditor'), 'Folder name'), 'Kept');
    await t.press(t.$('new-template'));
    assert.equal(
      t.document.activeElement.textContent,
      'Open a new template? Your unsaved changes to Account onboarding are discarded.',
    );
    await t.press(t.find(t.$('templates-list'), 'Keep editing'));
    assert.equal(t.$('version-chip').textContent, 'Draft v2 · unsaved changes');
    assert.equal(t.$('templateName').value, 'Account onboarding');
    await t.press(t.$('new-template'));
    await t.press(t.find(t.$('templates-list'), 'Open a new template'));
    assert.equal(t.$('templateName').value, 'New template');
    assert.equal(t.$('version-chip').textContent, 'Draft v1');
    assert.equal(t.last('asx_CreateDraft'), undefined, 'Nothing was saved');
  }
  {
    // ＋ New: with more than one table it asks which; the choice opens a new template's editor.
    const t = await boot({ enabled: ['account', 'contact'] });
    await t.press(t.$('new-template'));
    const table = t.$('new-template-table');
    assert.equal(table.hidden, false);
    assert.deepEqual(
      table.querySelectorAll('option').map((o) => o.textContent),
      ['Choose a table', 'Account', 'Contact'],
    );
    await t.change(table, 'contact');
    assert.equal(table.hidden, true);
    assert.equal(t.$('template-editor').hidden, false);
    assert.equal(t.$('templateName').value, 'New template');
    assert.equal(t.$('editor-title').textContent, 'New template');
  }
  {
    // An edit marks Save draft, explains Publish, and dims the preview.
    const t = await boot();
    await t.open();
    await t.press(t.find(t.$('destinations'), 'General'));
    await t.change(t.labelled(t.$('folderEditor'), 'Folder name'), 'General documents');
    assert.equal(t.$('version-chip').textContent, 'Draft v2 · unsaved changes');
    assert.equal(t.$('save').getAttribute('aria-label'), 'Save draft, unsaved changes');
    assert.equal(t.$('publish-reason').textContent, 'Save your changes first');
    // Leaving the page with unsaved edits asks first (the shell owns the prompt).
    t.document.track(t.window.AsxdUi.navigate('monitor'));
    await t.document.settle();
    assert.match(
      t.$('leavePrompt').visibleText,
      /You have unsaved changes to Account onboarding\./,
    );
    // Save draft, then Publish with its confirmation (kept text #3).
    await t.press(t.find(t.$('leavePrompt'), 'Stay'));
    await t.press(t.$('save'));
    assert.equal(t.last('asx_CreateDraft').Destinations[0].Folders[1].Name, 'General documents');
    assert.equal(t.$('version-chip').textContent, 'Draft v2');
    assert.equal(t.$('publish').hasAttribute('aria-disabled'), false);
    await t.press(t.$('publish'));
    assert.equal(t.document.activeElement.textContent, 'Publish v2?');
    assert.equal(
      t.$('help-publish').textContent,
      'Documents uses v2 for records created from now on, and for changed records when record updates are on. Existing records keep their folders until you re-run them.',
    );
    await t.press(t.find(t.$('template-bar').querySelector('.confirm'), 'Publish v2'));
    assert.deepEqual(t.last('asx_PublishTemplate'), { RevisionId: 'rev-2', RowVersion: '4' });
    assert.match(t.$('fb-templates').textContent, /^Published v2\./);
    assert.ok(t.find(t.$('fb-templates'), 'Re-run existing records…'));
  }
  {
    // Unsaved edits ask before the library link leaves the page too: Save draft saves and then
    // goes; a page change with Discard changes goes without saving.
    const t = await boot();
    await t.open();
    await t.press(t.find(t.$('destinations'), 'General'));
    await t.change(t.labelled(t.$('folderEditor'), 'Folder name'), 'Renamed');
    await t.press(t.find(t.$('folderEditor'), 'View this library’s team access →'));
    assert.match(
      t.$('leavePrompt').visibleText,
      /You have unsaved changes to Account onboarding\./,
    );
    assert.equal(t.sent.filter(([k]) => k === 'navigate').length, 0, 'Nothing leaves yet');
    await t.press(t.find(t.$('leavePrompt'), 'Save draft'));
    assert.equal(t.last('asx_CreateDraft').Destinations[0].Folders[1].Name, 'Renamed');
    assert.equal(t.sent.filter(([k]) => k === 'navigate').at(-1)[1].data, 'access-' + BUILD);
    const d = await boot();
    await d.open();
    await d.press(d.find(d.$('destinations'), 'General'));
    await d.change(d.labelled(d.$('folderEditor'), 'Folder name'), 'Dropped');
    d.document.track(d.window.AsxdUi.navigate('settings'));
    await d.document.settle();
    await d.press(d.find(d.$('leavePrompt'), 'Discard changes'));
    assert.equal(d.last('asx_CreateDraft'), undefined, 'Discard saves nothing');
    assert.equal(d.sent.filter(([k]) => k === 'navigate').at(-1)[1].data, 'settings-' + BUILD);
  }
  {
    // Paused automation adds a sentence to the Publish confirmation; no Publisher role disables it.
    const paused = await boot({
      runtime: { Enabled: false },
      loaded: { RevisionId: 'rev-2', RowVersion: '4', Status: 'Draft', Version: 2, Draft: draft() },
    });
    await paused.open();
    await paused.press(paused.$('publish'));
    assert.match(
      paused.$('help-publish').textContent,
      /Automation is paused, so no folders are created until it's on\.$/,
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
    // Another template with unsaved edits asks before it discards them.
    const t = await boot({
      templates: [
        { asx_templateid: TEMPLATE, asx_name: 'Account onboarding', asx_table: 'account' },
        { asx_templateid: OTHER, asx_name: 'Contract documents', asx_table: 'account' },
      ],
    });
    await t.open();
    await t.press(t.find(t.$('destinations'), 'General'));
    await t.change(t.labelled(t.$('folderEditor'), 'Folder name'), 'Kept');
    const loads = t.sent.filter(([k]) => k === 'asx_LoadDraft').length;
    await t.press(t.row('Contract documents'));
    assert.equal(
      t.document.activeElement.textContent,
      'Open Contract documents? Your unsaved changes to Account onboarding are discarded.',
    );
    await t.press(t.find(t.$('template-groups'), 'Keep editing'));
    assert.equal(t.sent.filter(([k]) => k === 'asx_LoadDraft').length, loads);
    assert.equal(t.$('version-chip').textContent, 'Draft v2 · unsaved changes');
    assert.equal(t.$('template-editor').hidden, false);
    assert.equal(t.document.activeElement, t.row('Contract documents'));
    // Discarding opens the other template's overview.
    await t.press(t.row('Contract documents'));
    await t.press(t.find(t.$('template-groups'), 'Open Contract documents'));
    assert.equal(t.$('template-editor').hidden, true);
    assert.equal(t.$('overview-title').textContent, 'Contract documents');
  }
  {
    // Folders and folder settings: selection keeps focus, readable copy, Remove folder and Remove destination.
    const t = await boot();
    await t.open();
    const node = t.find(t.$('destinations'), 'General');
    node.focus();
    await t.press(node);
    const again = t.find(t.$('destinations'), 'General');
    assert.equal(again.getAttribute('aria-current'), 'true');
    assert.equal(t.document.activeElement, again, 'Focus stays on the selected node after render');
    await t.press(t.find(t.$('destinations'), '[Account Name]'));
    const editor = t.$('folderEditor');
    const removeFolder = t.find(editor, 'Remove folder');
    assert.equal(removeFolder.getAttribute('aria-disabled'), 'true');
    assert.equal(
      t.document.getElementById(removeFolder.getAttribute('aria-describedby')).textContent,
      'Remove its folders first',
    );
    assert.match(editor.visibleText, /Shows as: \[Account Name\]/);
    assert.equal(
      t.$('help-folder-access').textContent,
      'Everyone with access to General can open this folder. To restrict a folder, use a separate library.',
    );
    assert.equal(
      t.$('help-include-root').textContent,
      "If these conditions don't match, Documents skips this whole destination for the record. Folders it already created stay.",
    );
    await t.press(t.find(t.$('destinations'), '＋ Add folder inside [Account Name]'));
    assert.match(
      t.document.activeElement.getAttribute('data-focus-key'),
      /^node:general:/,
      'A new folder takes focus',
    );
    await t.press(t.find(editor, 'Remove destination'));
    assert.equal(
      t.document.activeElement.textContent,
      'Remove Business documents and its 3 folders from this draft? Nothing changes in SharePoint until you publish.',
    );
  }
  {
    // Insert field: one picker with optgroups, the primary name first, deprecated fields last.
    const t = await boot();
    await t.open();
    await t.press(t.find(t.$('destinations'), 'General'));
    const picker = t.labelled(t.$('folderEditor'), 'Insert field');
    const groups = picker.querySelectorAll('optgroup').map((g) => g.getAttribute('label'));
    assert.deepEqual(groups, ['This record', 'Primary Contact → Contact']);
    assert.equal(picker.value, 'root.name');
    const own = picker
      .querySelectorAll('optgroup')[0]
      .querySelectorAll('option')
      .map((o) => o.textContent);
    assert.equal(own.at(-1), '(Deprecated) Old Code');
    await t.change(picker, 'lookup:primarycontactid:contact:fullname');
    await t.press(t.find(t.$('folderEditor'), 'Insert field'));
    assert.match(t.labelled(t.$('folderEditor'), 'Folder name').value, /\{lookup_1\.fullname\}$/);
    assert.match(t.$('destinations').visibleText, /\[Primary Contact › Full Name\]/);
  }
  {
    // Condition builder: names on every control, typed values, the lookup picker, validation.
    const t = await boot();
    await t.open();
    await t.press(t.find(t.$('destinations'), 'General'));
    const editor = t.$('folderEditor');
    await t.change(t.labelled(editor, 'When should this folder appear?'), 'conditional');
    const field = editor
      .querySelectorAll('select')
      .find((s) => s.getAttribute('aria-label') === 'Field, condition 1');
    assert.equal(t.document.activeElement, field, 'A new condition focuses its Field');
    const fieldset = editor.querySelector('fieldset');
    assert.equal(fieldset.querySelector('legend').textContent, 'Conditions for General');
    for (const control of editor.querySelectorAll('select, input'))
      assert.ok(
        control.getAttribute('aria-label') || control.parentNode.tagName === 'LABEL',
        'Unnamed control in the condition builder',
      );
    assert.equal(
      t
        .labelled(editor, 'Match')
        .querySelectorAll('option')
        .map((o) => o.textContent)
        .join('|'),
      'all conditions|any condition',
    );
    const kinds = {
      'root.revenue': ['number', 'decimal'],
      'root.closedate': ['date', null],
      'root.reviewedon': ['datetime-local', null],
    };
    for (const [value, [type, mode]] of Object.entries(kinds)) {
      await t.change(field, value);
      const input = editor
        .querySelectorAll('input')
        .find((i) => i.getAttribute('aria-label') === 'Value, condition 1');
      assert.equal(input.type, type, value);
      assert.equal(input.getAttribute('inputmode'), mode);
    }
    await t.change(field, 'root.statecode');
    const choice = editor
      .querySelectorAll('select')
      .find((s) => s.getAttribute('aria-label') === 'Value, condition 1');
    assert.deepEqual(
      choice
        .querySelectorAll('option')
        .map((o) => o.textContent)
        .slice(1),
      ['Active', 'Inactive'],
    );
    await t.change(field, 'root.primarycontactid');
    await t.press(t.find(editor, 'Choose record…'));
    assert.deepEqual([...t.looked.at(-1).entityTypes], ['contact']);
    assert.match(editor.visibleText, /Jane Smith/);
    // The chosen record is a removable chip (spec 3.1): ✕ is named, clears the value, and focus
    // returns to Choose record….
    const clear = editor
      .querySelectorAll('button')
      .find((b) => b.getAttribute('aria-label') === 'Remove Jane Smith, condition 1');
    assert.ok(clear, 'The lookup chip has a named remove button');
    await t.press(clear);
    assert.doesNotMatch(editor.visibleText, /Jane Smith/);
    assert.equal(t.document.activeElement.getAttribute('aria-label'), 'Choose record, condition 1');
    await t.press(t.find(editor, 'Choose record…'));
    assert.match(editor.visibleText, /Jane Smith/);
    await t.change(t.labelled(editor, 'Operator, condition 1'), 'IsNull');
    assert.equal(
      editor
        .querySelectorAll('input, select')
        .filter((n) => n.getAttribute('aria-label') === 'Value, condition 1').length,
      0,
      'is empty has no value control',
    );
    await t.change(t.labelled(editor, 'Operator, condition 1'), 'Equal');
    await t.press(t.find(editor, 'Choose record…'));
    // Removing a condition moves focus to the next condition's Field.
    await t.press(
      editor
        .querySelectorAll('button')
        .find((b) => b.getAttribute('aria-label') === 'Add condition to conditions for General'),
    );
    assert.equal(t.document.activeElement.getAttribute('aria-label'), 'Field, condition 2');
    await t.press(
      editor
        .querySelectorAll('button')
        .find((b) => b.getAttribute('aria-label') === 'Remove condition 2'),
    );
    assert.equal(
      t.document.activeElement.getAttribute('aria-label'),
      'Add condition to conditions for General',
    );
    // An empty group is refused at Save, at the group, with focus on it.
    await t.press(
      editor
        .querySelectorAll('button')
        .find((b) => b.getAttribute('aria-label') === 'Add group to conditions for General'),
    );
    await t.press(t.$('save'));
    assert.equal(t.sent.filter(([k]) => k === 'asx_CreateDraft').length, 0);
    const alert = editor.querySelector('[role=alert]');
    assert.equal(alert.textContent, 'Add a condition or remove this group');
    assert.equal(t.document.activeElement.tagName, 'LEGEND');
    await t.press(
      editor
        .querySelectorAll('button')
        .find((b) => b.getAttribute('aria-label') === 'Remove group 2'),
    );
    await t.press(t.$('save'));
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
    await t.press(t.find(t.$('destinations'), 'General'));
    const editor = t.$('folderEditor');
    const addGroup = (n) =>
      editor
        .querySelectorAll('button')
        .find((b) => b.getAttribute('aria-label') === 'Add group to group ' + n + ' conditions');
    const deepest = addGroup(depth);
    assert.equal(deepest.getAttribute('aria-disabled'), 'true');
    assert.equal(
      t.document.getElementById(deepest.getAttribute('aria-describedby')).textContent,
      'Condition groups can be nested up to ' +
        depth +
        " levels deep, because deeper templates can't be sent to Dataverse.",
    );
    assert.equal(addGroup(depth - 1).hasAttribute('aria-disabled'), false);
    const groups = editor.querySelectorAll('fieldset').length;
    await t.press(deepest);
    assert.equal(editor.querySelectorAll('fieldset').length, groups, 'No group past the bound');
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
    await t.press(t.find(t.$('destinations'), 'General'));
    assert.match(t.$('folderEditor').visibleText, /Jane Smith/);
    assert.doesNotMatch(t.$('folderEditor').visibleText, GUID);
    assert.match(
      t.$('destinations').visibleText,
      /\(conditional\)/,
      'The conditional marker has text, not only colour',
    );
  }
  {
    // Preview: picking a record runs it (saved revision); an edit makes it out of date; Refresh previews the edit.
    const t = await boot();
    await t.open();
    await t.press(t.$('chooseRecord'));
    assert.equal(t.$('preview-record-name').textContent, 'Contoso Ltd');
    assert.deepEqual(t.last('asx_PreviewTemplate'), { RevisionId: 'rev-1', RecordId: RECORD });
    assert.equal(t.$('preview-status').textContent, 'Preview updated: 1 destination, 2 folders.');
    assert.equal(t.$('previewTrees').hasAttribute('aria-live'), false);
    await t.press(t.find(t.$('destinations'), 'General'));
    await t.change(t.labelled(t.$('folderEditor'), 'Folder name'), 'Edited');
    assert.equal(t.$('preview-stale').hidden, false);
    await t.press(t.$('refreshPreview'));
    const request = t.last('asx_PreviewTemplate');
    assert.equal(request.RevisionId, '');
    assert.equal(request.Draft.Destinations[0].Folders[1].Name, 'Edited');
    assert.equal(t.$('preview-stale').hidden, true);
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
    // Fix 1: Publish reloads the revision, so the next Save sends its new row version.
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
    await t.press(t.$('publish'));
    await t.press(t.find(t.$('template-bar').querySelector('.confirm'), 'Publish v2'));
    assert.equal(t.$('version-chip').textContent, 'Published v2');
    await t.press(t.find(t.$('destinations'), 'General'));
    await t.change(t.labelled(t.$('folderEditor'), 'Folder name'), 'After publish');
    await t.press(t.$('save'));
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
    assert.match(
      t.$('destinations').visibleText,
      /\[Account Name\]/,
      'The template renders before its related tables load',
    );
    assert.equal(peak, 4, 'Four related tables load at a time');
    await t.press(t.find(t.$('destinations'), 'General'));
    const group = () =>
      t
        .labelled(t.$('folderEditor'), 'Insert field')
        .querySelectorAll('optgroup')
        .find((g) => g.getAttribute('label') === 'Link rel9 → rel9');
    assert.equal(group().disabled, true);
    assert.equal(group().querySelectorAll('option')[0].textContent, 'Loading fields…');
    release();
    await t.document.settle();
    assert.equal(peak, 4, 'Never more than four');
    assert.equal(group().disabled, false, 'The picker redraws as tables arrive');
    assert.equal(
      group()
        .querySelectorAll('option')
        .some((o) => o.textContent === 'Loading fields…'),
      false,
    );
    // A table that fails twice (one retry) shows as unavailable, not as missing.
    let attempts = 0;
    const failed = await boot({
      failFetch: (url, table) =>
        table === 'contact' && url.includes('/Attributes?') && ++attempts > 0,
    });
    await failed.open();
    await failed.press(failed.find(failed.$('destinations'), 'General'));
    const contact = failed
      .labelled(failed.$('folderEditor'), 'Insert field')
      .querySelectorAll('optgroup')
      .find((g) => g.getAttribute('label') === 'Primary Contact → Contact');
    assert.equal(contact.disabled, true);
    assert.equal(contact.querySelectorAll('option')[0].textContent, "Couldn't load these fields");
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
    await t.press(t.find(t.$('destinations'), 'General'));
    const group = () =>
      t
        .labelled(t.$('folderEditor'), 'Insert field')
        .querySelectorAll('optgroup')
        .find((g) => g.getAttribute('label') === 'Link rel0 → rel0');
    const name = t.labelled(t.$('folderEditor'), 'Folder name');
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
    await t.press(t.find(t.$('destinations'), 'General'));
    await t.change(t.labelled(t.$('folderEditor'), 'Folder name'), 'From v1');
    await t.press(t.$('save'));
    assert.equal(t.$('version-chip').textContent, 'Draft v4');
    await t.press(t.$('publish'));
    assert.equal(t.document.activeElement.textContent, 'Publish v4?');
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
    await t.press(t.$('addDestination'));
    failing = true;
    await t.press(t.$('save'));
    assert.match(
      t.$('fb-templates').textContent,
      /^Draft saved, but the page could not refresh: Network down\./,
    );
    assert.equal(t.$('version-chip').textContent, 'Draft v1');
    failing = false;
    await t.change(t.labelled(t.$('folderEditor'), 'Folder name'), 'Second save');
    await t.press(t.$('save'));
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
    await t.press(t.find(t.$('destinations'), 'General'));
    assert.match(t.$('folderEditor').visibleText, /Unavailable field\. Select a replacement\./);
    assert.doesNotMatch(t.$('folderEditor').visibleText, /retired_code/);
  }
  {
    // Fix 7: a number is sent as typed; one the server cannot read is refused at the control.
    const t = await boot();
    await t.open();
    await t.press(t.find(t.$('destinations'), 'General'));
    const editor = t.$('folderEditor');
    await t.change(t.labelled(editor, 'When should this folder appear?'), 'conditional');
    await t.change(t.labelled(editor, 'Field, condition 1'), 'root.revenue');
    await t.change(t.labelled(editor, 'Value, condition 1'), '1e5');
    await t.press(t.$('save'));
    assert.equal(t.sent.filter(([k]) => k === 'asx_CreateDraft').length, 0);
    assert.equal(editor.querySelector('[role=alert]').textContent, 'Enter a number');
    await t.change(t.labelled(editor, 'Value, condition 1'), ' 12345678901234567890.5 ');
    await t.press(t.$('save'));
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
  console.log(
    'PASS Folder templates: empty states, the templates list with states and search, the overview (pill, meta, cards, chips, rule sentences, versions, last re-run, team counts, problem pill, roles), Edit template and Close, View read-only, ＋ New, version chip, Publish and its reasons, unsaved-changes prompts, the ⋯ menu, Delete and focus after it, Schedule and All versions side panels, Manage tables, focus after a keyboard pick, Re-run in progress or Last re-run, the ⋯ separator, the ＋ New table picker and its unsaved-changes prompt, folders and focus, Insert field, condition builder and its depth bound, lookup labels, preview of edits, Re-run all with exact and estimated totals; fix round 1: Save after Publish, related tables loaded four at a time after the first render with a retry and unavailable groups, saved version numbers, no second template after a failed reload, unavailable fields unnamed, numbers as typed, the Operator reason on every re-run action; Task 9: stale pickers redraw once focus leaves the field, and a template switch stops the old preload. Fake DOM; browser QA separate.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
