'use strict';
// Folder templates with a fake DOM and mocked Dataverse: template bar and version chip, Publish,
// the ⋯ menu, Schedule, Version history, Delete, the Tables rail, folders, Insert field, the
// condition builder, preview of unsaved edits, focus, and Re-run for existing records.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createDocument } = require('./fake-dom.cjs');
const base = path.resolve(__dirname, '../../client/admin');
const html = fs.readFileSync(path.join(base, 'index.html'), 'utf8');
const read = (name) => fs.readFileSync(path.join(base, name), 'utf8');
const GUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;
const TEMPLATE = '11111111-0000-0000-0000-000000000001';
const OTHER = '33333333-0000-0000-0000-000000000003';
const CONTACT = '22222222-0000-0000-0000-000000000002';
const RECORD = '607cba8a-acc1-f111-aaaf-7c1e52067fd0';

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
  templates = [{ asx_templateid: TEMPLATE, asx_name: 'Account onboarding', asx_table: 'account' }],
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
      navigateTo: async (page) => sent.push(['navigate', page]),
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
      retrieveRecord: async (table, id) => ({
        asx_templateid: id,
        asx_name: 'Account onboarding',
        asx_table: 'account',
        asx_disabled: false,
      }),
      retrieveMultipleRecords: async (table) => ({
        entities:
          table === 'asx_runtimetable'
            ? enabled.map((asx_logicalname) => ({ asx_logicalname }))
            : table === 'asx_template'
              ? templates
              : table === 'asx_revision'
                ? [{ asx_revisionid: 'rev-1', asx_version: 1, asx_status: 'Published' }]
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
                    : [],
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
    location: { search: '?data=templates-ui20261006nav1', hash: '' },
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
  const find = (root, text) => root.querySelectorAll('button').find((b) => b.textContent === text);
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
  const open = async () => press(find($('templateTree'), 'Account onboarding'));
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
    open,
    last: (api) => sent.filter(([a]) => a === api).at(-1)?.[1],
  };
}

(async () => {
  {
    // Empty states: no tables yet, and no template selected.
    const t = await boot({ enabled: [], templates: [] });
    const rail = t.$('templateTree');
    assert.match(rail.visibleText, /No tables yet/);
    assert.ok(t.find(rail, '＋ Add table'), 'Add table sits inside the empty state');
    assert.equal(t.$('template-bar').hidden, true, 'The template bar waits for a table');
    assert.match(t.$('no-template').visibleText, /No template selected/);
  }
  {
    // Opening a template: the bar and the version chip; no IDs in the panel.
    const t = await boot();
    await t.open();
    assert.equal(t.$('template-bar').hidden, false);
    assert.equal(t.$('version-chip').textContent, 'Published v1');
    assert.equal(t.$('version-chip').getAttribute('aria-describedby'), 'version-chip-desc');
    assert.equal(
      t.$('version-chip-desc').textContent,
      'Published v1 keeps running while you edit. Saving creates Draft v2.',
    );
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
    // An edit marks Save draft, explains Publish, and dims the preview.
    const t = await boot();
    await t.open();
    await t.press(t.find(t.$('destinations'), 'General'));
    await t.change(t.labelled(t.$('folderEditor'), 'Folder name'), 'General documents');
    assert.equal(t.$('version-chip').textContent, 'Draft v2 · unsaved changes');
    assert.equal(t.$('save').getAttribute('aria-label'), 'Save draft, unsaved changes');
    assert.equal(t.$('publish-reason').textContent, 'Save your changes first');
    // Leaving the tab with unsaved edits asks first (the shell owns the prompt).
    t.$('tab-monitor').click();
    await t.document.settle();
    assert.match(t.$('tabPrompt').visibleText, /You have unsaved changes to Account onboarding\./);
    // Save draft, then Publish with its confirmation (kept text #3).
    await t.press(t.find(t.$('tabPrompt'), 'Stay'));
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
    // Unsaved edits ask before the library link leaves the tab too (ruling 3): Save draft saves
    // and then goes; a tab change with Discard changes goes without saving.
    const t = await boot();
    await t.open();
    await t.press(t.find(t.$('destinations'), 'General'));
    await t.change(t.labelled(t.$('folderEditor'), 'Folder name'), 'Renamed');
    await t.press(t.find(t.$('folderEditor'), 'View this library’s team access →'));
    assert.match(t.$('tabPrompt').visibleText, /You have unsaved changes to Account onboarding\./);
    assert.equal(t.sent.filter(([k]) => k === 'navigate').length, 0, 'Nothing leaves yet');
    await t.press(t.find(t.$('tabPrompt'), 'Save draft'));
    assert.equal(t.last('asx_CreateDraft').Destinations[0].Folders[1].Name, 'Renamed');
    assert.equal(t.sent.filter(([k]) => k === 'navigate').at(-1)[1].data, 'access-ui20261006nav1');
    const d = await boot();
    await d.open();
    await d.press(d.find(d.$('destinations'), 'General'));
    await d.change(d.labelled(d.$('folderEditor'), 'Folder name'), 'Dropped');
    d.$('tab-settings').click();
    await d.document.settle();
    await d.press(d.find(d.$('tabPrompt'), 'Discard changes'));
    assert.equal(d.last('asx_CreateDraft'), undefined, 'Discard saves nothing');
    assert.equal(
      d.sent.filter(([k]) => k === 'navigate').at(-1)[1].data,
      'settings-ui20261006nav1',
    );
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
    // The ⋯ menu: Delete template confirms in the page; Schedule and Version history open panels.
    const t = await boot();
    await t.open();
    const menu = t.$('template-menu');
    assert.equal(menu.getAttribute('aria-haspopup'), 'menu');
    assert.equal(
      menu.getAttribute('aria-label'),
      'More actions for Account onboarding',
      'The ⋯ button names its template (spec 5.1)',
    );
    menu.key('ArrowDown');
    await t.document.settle();
    assert.equal(t.$('template-menu-list').hidden, false);
    assert.equal(t.document.activeElement.id, 'menu-history');
    t.document.activeElement.key('ArrowDown');
    assert.equal(t.document.activeElement.id, 'menu-schedule');
    t.document.activeElement.key('Escape');
    assert.equal(t.$('template-menu-list').hidden, true);
    assert.equal(t.document.activeElement, menu);
    await t.press(menu);
    await t.press(t.$('menu-delete'));
    assert.equal(
      t.document.activeElement.textContent,
      'Delete Account onboarding and all its versions? No new folder work starts for it. Folders, documents and access in SharePoint stay as they are. Work already sent to SharePoint may still finish.',
    );
    await t.press(t.find(t.$('template-bar').querySelector('.confirm'), 'Delete template'));
    assert.deepEqual(t.deleted, [['asx_template', TEMPLATE]]);
    // Schedule: Save schedule writes the template row and says so.
    const s = await boot();
    await s.open();
    await s.press(s.$('template-menu'));
    await s.press(s.$('menu-schedule'));
    assert.equal(s.$('schedule-panel').hidden, false);
    await s.press(s.$('saveAvailability'));
    assert.equal(s.sent.filter(([k]) => k === 'update').length, 1);
    assert.equal(s.$('fb-schedule').textContent, 'Schedule saved.');
    // Version history lists the versions; picking one opens it.
    const h = await boot();
    await h.open();
    await h.press(h.$('template-menu'));
    await h.press(h.$('menu-history'));
    assert.equal(h.$('history-panel').hidden, false);
    assert.equal(h.document.activeElement.textContent, 'v1 · Published');
    const loads = h.sent.filter(([k]) => k === 'asx_LoadDraft').length;
    await h.press(h.document.activeElement);
    assert.equal(h.sent.filter(([k]) => k === 'asx_LoadDraft').length, loads + 1);
    assert.equal(h.$('history-panel').hidden, true);
  }
  {
    // After Delete template, focus goes to the table's next template, else the Tables heading
    // (ruling 2), and the result stays visible.
    const list = [
      { asx_templateid: TEMPLATE, asx_name: 'Account onboarding', asx_table: 'account' },
      { asx_templateid: OTHER, asx_name: 'Contract documents', asx_table: 'account' },
    ];
    const t = await boot({ templates: list });
    await t.open();
    await t.press(t.$('template-menu'));
    await t.press(t.$('menu-delete'));
    list.splice(0, 1);
    await t.press(t.find(t.$('template-bar').querySelector('.confirm'), 'Delete template'));
    assert.equal(t.document.activeElement.textContent, 'Contract documents');
    assert.match(
      t.$('fb-templates').visibleText,
      /^Account onboarding deleted\. Nothing in SharePoint changed\.$/,
    );
    const only = [
      { asx_templateid: TEMPLATE, asx_name: 'Account onboarding', asx_table: 'account' },
    ];
    const last = await boot({ templates: only });
    await last.open();
    await last.press(last.$('template-menu'));
    await last.press(last.$('menu-delete'));
    only.splice(0);
    await last.press(
      last.find(last.$('template-bar').querySelector('.confirm'), 'Delete template'),
    );
    assert.equal(last.document.activeElement.id, 'tables-heading');
  }
  {
    // Tables rail: readiness words, Needs repair links to Settings, Remove table confirms with F-05 text.
    const paused = await boot({ runtime: { Enabled: false } });
    assert.match(paused.$('templateTree').visibleText, /Ready · automation paused/);
    const broken = await boot({
      runtime: {
        Registration: {
          Readiness: [{ Scope: 'account', Status: 'Missing' }],
          Pending: 1,
          Error: null,
        },
      },
    });
    await broken.press(broken.find(broken.$('templateTree'), 'Needs repair'));
    assert.deepEqual(
      broken.sent.filter(([k]) => k === 'navigate').at(-1)[1].data,
      'settings-ui20261006nav1',
    );
    const t = await boot();
    const remove = t
      .$('templateTree')
      .querySelectorAll('button')
      .find((b) => b.getAttribute('aria-label') === 'Remove Account');
    await t.press(remove);
    assert.equal(
      t.document.activeElement.textContent,
      'Stop creating folders for Account? Queued folder work for this table is cancelled. Templates are kept, and nothing in SharePoint is deleted.',
    );
    assert.ok(t.find(t.$('templateTree'), 'Keep table'));
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
    await t.press(t.find(t.$('templateTree'), 'Contract documents'));
    assert.equal(
      t.document.activeElement.textContent,
      'Open Contract documents? Your unsaved changes to Account onboarding are discarded.',
    );
    await t.press(t.find(t.$('templateTree'), 'Keep editing'));
    assert.equal(t.sent.filter(([k]) => k === 'asx_LoadDraft').length, loads);
    assert.equal(t.$('version-chip').textContent, 'Draft v2 · unsaved changes');
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
    await t.open();
    await t.press(t.$('template-menu'));
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
    assert.equal(nav.data, 'monitor-ui20261006nav1');
    // CountRecords reads Dataverse's daily snapshot, so its total is usually "about N" (ruling 5).
    const about = await boot({
      handle: (api, b) =>
        b.Command === 'CountRecords'
          ? { Status: 'Counted', Run: { Total: 1240, TotalEstimated: true } }
          : null,
    });
    await about.open();
    await about.press(about.$('template-menu'));
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
    await busy.open();
    await busy.press(busy.$('template-menu'));
    await busy.press(busy.$('menu-rerun'));
    assert.equal(busy.$('rerun-all').hidden, true);
    assert.match(
      busy.$('rerun-progress').visibleText,
      /Re-run in progress: 500 of 1,240 · Open in Monitor/,
    );
    // Preview re-run: more records than the traced bound is refused at the button.
    const many = await boot({ pickMany: true });
    await many.open();
    await many.press(many.$('template-menu'));
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
  console.log(
    'PASS Folder templates: empty states, version chip, Publish and its reasons, unsaved-changes prompts, menu, Delete and focus after it, Schedule, Version history, rail, folders and focus, Insert field, condition builder and its depth bound, lookup labels, preview of edits, Re-run all with exact and estimated totals. Fake DOM; browser QA separate.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
