// State/handler contract with a minimal DOM and mocked Dataverse. Not a browser or visual test.
const fs = require('fs'),
  vm = require('vm'),
  assert = require('assert/strict'),
  path = require('path');
class Node {
  constructor(tag) {
    this.tagName = tag.toUpperCase();
    this.children = [];
    this._text = '';
    this.value = '';
    this.style = {};
    this.dataset = {};
    this.attrs = {};
    this.className = '';
    this.classList = {
      add: (c) => (this.className += ' ' + c),
      toggle: (c, on) => {
        this.className = this.className
          .split(' ')
          .filter((x) => x !== c)
          .concat(on ? [c] : [])
          .join(' ');
      },
    };
  }
  set textContent(v) {
    this._text = String(v);
    this.children = [];
  }
  get textContent() {
    return this._text + this.children.map((c) => c.textContent).join('');
  }
  append(...nodes) {
    this.children.push(...nodes);
  }
  replaceChildren(...nodes) {
    this._text = '';
    this.children = [...nodes];
  }
  setAttribute(k, v) {
    this.attrs[k] = v;
  }
  get options() {
    return this.children;
  }
  get lastElementChild() {
    return this.children.at(-1);
  }
}
const base = path.resolve(__dirname, '../../client/admin'),
  html = fs.readFileSync(path.join(base, 'index.html'), 'utf8'),
  nodes = {};
for (const m of html.matchAll(/<([a-z]+)[^>]*\bid="([^"]+)"[^>]*>/g)) nodes[m[2]] = new Node(m[1]);
const document = {
  getElementById: (id) => nodes[id],
  createElement: (t) => new Node(t),
  createElementNS: (ns, t) => {
    const n = new Node(t);
    n.namespaceURI = ns;
    return n;
  },
  querySelectorAll: () => [],
};
const requests = [],
  table = {
    LogicalName: 'account',
    EntitySetName: 'accounts',
    PrimaryIdAttribute: 'accountid',
    PrimaryNameAttribute: 'name',
  };
const contact = {
  LogicalName: 'contact',
  EntitySetName: 'contacts',
  PrimaryIdAttribute: 'contactid',
  PrimaryNameAttribute: 'fullname',
  DisplayName: { UserLocalizedLabel: { Label: 'Contact' } },
};
const lead = {
  LogicalName: 'lead',
  EntitySetName: 'leads',
  PrimaryIdAttribute: 'leadid',
  PrimaryNameAttribute: 'fullname',
  DisplayName: { UserLocalizedLabel: { Label: 'Lead' } },
};
const libraries = [
  {
    asx_libraryid: 'lib-a',
    asx_name: 'General',
    _asx_siteid_value: 'site-a',
    asx_entryurl: 'https://example.test/a',
  },
  {
    asx_libraryid: 'lib-b',
    asx_name: 'Sensitive',
    _asx_siteid_value: 'site-b',
    asx_entryurl: 'https://example.test/b',
  },
];
const libraryQueries = [];
const xrm = {
  Utility: { getGlobalContext: () => ({ getClientUrl: () => 'https://example.test' }) },
  WebApi: {
    retrieveRecord: async (name, id) => ({
      asx_templateid: id,
      asx_name: 'Account onboarding',
      asx_table: 'account',
    }),
    retrieveMultipleRecords: async (name, options) => {
      if (name === 'asx_library') libraryQueries.push(options);
      return {
        entities:
          name === 'asx_library'
            ? libraries
            : name === 'asx_runtimetable'
              ? [{ asx_logicalname: 'account' }]
              : name === 'asx_site'
                ? [
                    { asx_siteid: 'site-a', asx_name: 'Delivery' },
                    { asx_siteid: 'site-b', asx_name: 'Commercial' },
                  ]
                : [],
      };
    },
    online: {
      execute: async (req) => {
        requests.push(req);
        return {
          ok: true,
          json: async () => ({
            Result: JSON.stringify(
              req.getMetadata().operationName === 'asx_PreviewTemplate'
                ? {
                    Folders: [
                      {
                        Section: 'destination_1',
                        Name: 'Resolved account',
                        RelativePath: 'Resolved account',
                        Node: 'root',
                      },
                    ],
                    Notices: [
                      "Folder 'destination_1/child' is waiting for 'root.code' to have a value.",
                    ],
                  }
                : {
                    TemplateId: 'template-1',
                    RevisionId: 'revision-1',
                    Status: 'Draft',
                    RowVersion: '1',
                  },
            ),
          }),
        };
      },
    },
  },
};
const fetch = async (url) => ({
  ok: true,
  json: async () => ({
    value: url.includes('/Attributes?')
      ? [{ LogicalName: 'name', AttributeType: 'String' }]
      : url.includes('LookupAttributeMetadata')
        ? [
            {
              LogicalName: 'parentaccountid',
              Targets: ['account'],
              DisplayName: { UserLocalizedLabel: { Label: 'Parent account' } },
            },
          ]
        : url.includes('/Attributes/')
          ? []
          : url.includes('/accounts?')
            ? [{ accountid: 'record-1', name: 'Example' }]
            : [table, contact, lead],
  }),
});
let active = 'templates';
const feedback = {},
  inits = {},
  runtimeListeners = [],
  navigations = [];
const ui = {
  BUILD: 'ui20261006nav1',
  activeTab: () => active,
  onTab: (tab, init) => (inits[tab] = init),
  feedback: (area, text, kind = 'success') => (feedback[area] = { text, kind }),
  clearFeedback: (area) => delete feedback[area],
  // A copy: objects made inside the vm context have its Object prototype, which deepEqual refuses.
  navigate: async (tab, link) => navigations.push([tab, link && { ...link }]),
  deeplink: () => null,
  can: () => true,
  // The shell loads the runtime once; it is not back yet when the tab starts.
  runtime: () => null,
  onRuntime: (fn) => runtimeListeners.push(fn),
  // As the shell: a shared result reaches every onRuntime listener.
  setRuntime: (result) => runtimeListeners.forEach((fn) => fn(result)),
  api: async (name, request) =>
    JSON.parse(
      (
        await (
          await xrm.WebApi.online.execute({
            Request: JSON.stringify(request),
            getMetadata: () => ({ operationName: name }),
          })
        ).json()
      ).Result,
    ),
};
// What the tab running the action said last, as its feedback line shows it.
const said = () => feedback[active]?.text ?? '';
const win = { parent: { Xrm: xrm }, addEventListener: () => {}, AsxdUi: ui };
vm.runInNewContext(fs.readFileSync(path.join(base, 'admin.js'), 'utf8'), {
  document,
  window: win,
  location: { hash: '' },
  fetch,
  console,
  crypto: require('crypto').webcrypto,
  URL,
});
const descendants = (n) => n.children.flatMap((c) => [c, ...descendants(c)]);
const find = (text, tag) =>
  descendants(nodes.folderEditor).find((n) => n.tagName === tag && n.textContent === text);
const control = (text) =>
  descendants(nodes.folderEditor).find((n) => n.tagName === 'LABEL' && n._text === text)
    .lastElementChild;
async function change(n, value) {
  assert(!n.disabled, 'Control must be enabled');
  n.value = value;
  await n.onchange();
}
(async () => {
  await inits.templates();
  // The tab reuses the shell's runtime Get instead of sending its own, and shows table readiness
  // when the shell's Get returns.
  assert.equal(
    requests.filter((r) => r.getMetadata().operationName === 'asx_RuntimeAdmin').length,
    0,
    'No second runtime Get',
  );
  runtimeListeners.forEach((fn) =>
    fn({ Registration: { Readiness: [{ Scope: 'account', Status: 'Pending' }] } }),
  );
  assert.match(nodes.templateTree.textContent, /Pending: not registered/);
  await win.AsxdAdmin.refreshCatalog();
  assert.equal(libraryQueries.length, 2, 'Initial load and refresh both read the library catalog');
  for (const query of libraryQueries) {
    assert.match(query, /\$filter=asx_approved eq true$/);
    assert.doesNotMatch(
      query,
      /asx_policyapplied/,
      'Templates list every approved library, whatever its access state',
    );
  }
  assert.equal(nodes.authorWorkspace.hidden, true);
  assert.equal(nodes.templateActions.hidden, true);
  await change(nodes.table, 'account');
  assert.equal(nodes.authorWorkspace.hidden, false);
  assert.equal(nodes.templateActions.hidden, false);
  nodes.addDestination.onclick();
  assert.match(nodes.destinations.textContent, /root.name/);
  assert(!control('Include folder').disabled);
  await change(control('Include folder'), 'conditional');
  find('Add condition', 'BUTTON').onclick();
  control('Destination name').value = 'Account onboarding';
  control('Destination name').oninput();
  assert.equal(nodes.save.disabled, false);
  await control('Destination name').onchange();
  assert.match(nodes.destinations.textContent, /Account onboarding/);
  await change(control('Site'), 'site-b');
  assert.deepEqual(
    control('Library').options.map((o) => o.value),
    ['', 'lib-b'],
    'Library choices follow selected site',
  );
  // The library link opens Sites & access through the shell, naming the library.
  find('View this library’s team access →', 'BUTTON').onclick();
  assert.deepEqual(navigations, [['access', { library: 'lib-b' }]]);
  await change(control('Readable folder name'), 'Sensitive root');
  descendants(nodes.destinations)
    .find((n) => n.tagName === 'BUTTON' && n.textContent === '＋ Add child folder')
    .onclick();
  await change(control('Readable folder name'), 'Review');
  await change(control('Include folder'), 'conditional');
  find('Add condition', 'BUTTON').onclick();
  const cond = descendants(nodes.folderEditor).find((n) => n.className === 'condition-line');
  const literal = cond.children.find((n) => n.tagName === 'INPUT');
  literal.value = 'Active';
  assert.equal(
    typeof literal.oninput,
    'function',
    'Condition typing must update draft before blur',
  );
  literal.oninput();
  await change(control('Get name fields from'), 'parentaccountid:account');
  find('Insert field', 'BUTTON').onclick();
  assert.equal(control('Get name fields from').value, 'parentaccountid:account');
  nodes.addDestination.onclick();
  assert.equal(
    control('Get name fields from').value,
    'root',
    'A new folder starts with its own current-record selection',
  );
  await nodes.save.onclick();
  const draft = JSON.parse(
    requests.find((r) => r.getMetadata().operationName === 'asx_CreateDraft').Request,
  );
  assert.equal(draft.Destinations[0].Name, 'Account onboarding');
  assert.equal(draft.Sources.length, 2);
  assert.equal(draft.Sources[1].Lookup, 'parentaccountid');
  assert.equal(draft.Destinations[0].Folders[1].Name, 'Review{lookup_1.name}');
  assert(
    draft.Destinations[0].Folders[0].Condition.Conditions.length === 1,
    'Root conditions are saved',
  );
  assert.equal(draft.Destinations.length, 2);
  assert.equal(draft.Destinations[0].LibraryId, 'lib-b');
  assert.equal(draft.Destinations[0].Key, 'destination_1');
  assert.equal(draft.Destinations[0].Folders[0].Name, 'Sensitive root');
  assert.equal(draft.Destinations[0].Folders[1].Condition.Conditions[0].Literal, 'Active');
  assert.equal(
    draft.Destinations[1].Folders[0].Name,
    '{root.name}',
    'Second destination retains independent root',
  );
  await change(nodes.record, 'record-1');
  await nodes.preview.onclick();
  assert.match(nodes.previewTrees.textContent, /Resolved account/);
  assert.match(nodes.previewTrees.textContent, /is waiting for 'root\.code' to have a value/);
  assert(!nodes.previewTrees.textContent.includes('BindingKey'));
  await change(control('Readable folder name'), 'Changed again');
  assert.match(nodes.previewTrees.textContent, /Draft changed/);
  assert.equal(nodes.preview.disabled, true, 'Editing invalidates saved preview');
  assert(
    requests.every(
      (r) =>
        ['asx_CreateDraft', 'asx_PreviewTemplate'].includes(r.getMetadata().operationName) ||
        (r.getMetadata().operationName === 'asx_RuntimeAdmin' &&
          JSON.parse(r.Request).Command === 'Get'),
    ),
    'No permission/runtime mutations',
  );
  nodes.useSchedule.checked = true;
  nodes.useSchedule.onchange();
  assert.equal(nodes.scheduleFields.hidden, false);
  nodes.useSchedule.checked = false;
  nodes.useSchedule.onchange();
  assert.equal(nodes.scheduleFields.hidden, true);
  await nodes.save.onclick();
  const repeated = JSON.parse(
    requests.filter((r) => r.getMetadata().operationName === 'asx_CreateDraft').at(-1).Request,
  );
  assert.equal(repeated.RevisionId, 'revision-1');
  assert.equal(repeated.RowVersion, '1', 'Dirty edits preserve the draft concurrency token');
  await change(nodes.table, '');
  assert.equal(nodes.record.disabled, true);
  assert.equal(nodes.record.value, '');
  assert.equal(nodes.save.disabled, true);
  assert.equal(nodes.queueBatch.disabled, true);

  xrm.WebApi.retrieveMultipleRecords = async (name) => ({
    entities:
      name === 'asx_template'
        ? [
            { asx_templateid: 'template-1', asx_name: 'Account onboarding', asx_table: 'account' },
            { asx_templateid: 'template-2', asx_name: 'Sales documents', asx_table: 'account' },
          ]
        : name === 'asx_revision'
          ? [
              { asx_revisionid: 'latest-3', asx_version: 3, asx_status: 'Draft' },
              { asx_revisionid: 'older-2', asx_version: 2, asx_status: 'Published' },
            ]
          : [],
  });
  const loadedIds = [];
  xrm.WebApi.online.execute = async (req) => {
    loadedIds.push(req.RevisionId);
    return {
      ok: true,
      json: async () => ({
        Result: JSON.stringify({
          RevisionId: req.RevisionId,
          RowVersion: '3',
          Status: req.RevisionId === 'latest-3' ? 'Draft' : 'Published',
          Draft: {
            ...draft,
            Destinations: draft.Destinations.map((d) => ({
              ...d,
              Folders: d.Folders.map((f) => ({ ...f, Condition: null })),
            })),
          },
        }),
      }),
    };
  };
  await change(nodes.table, 'account');
  assert.deepEqual(
    loadedIds,
    ['latest-3'],
    'Selecting a table automatically loads its highest revision',
  );
  assert.equal(nodes.savedRevision.value, 'latest-3');
  assert.match(nodes.destinations.textContent, /Account onboarding/);
  assert.equal(nodes.publish.disabled, false);
  {
    // Publishing a template whose lookup table is not enabled succeeds and shows the notice.
    const load = xrm.WebApi.online.execute;
    const notice =
      "Changes to contact records don't update folders until the account record changes or is replanned. Enable contact in the Tables panel to react to its changes.";
    xrm.WebApi.online.execute = async (req) =>
      req.getMetadata().operationName === 'asx_PublishTemplate'
        ? {
            ok: true,
            json: async () => ({
              Result: JSON.stringify({ Status: 'Published', Notices: [notice] }),
            }),
          }
        : load(req);
    await nodes.publish.onclick();
    assert.match(said(), /Revision published/);
    assert(said().includes(notice), 'The template editor shows publish notices');
    assert.equal(feedback.templates.kind, 'success', 'A publish notice is not an error');
    xrm.WebApi.online.execute = load;
    loadedIds.splice(1); // Publishing reloads the revision; later checks count explicit loads only.
  }
  await change(nodes.savedRevision, 'older-2');
  await nodes.loadRevision.onclick();
  assert.deepEqual(loadedIds, ['latest-3', 'older-2']);
  assert.equal(nodes.publish.disabled, true, 'Older published revisions remain manually loadable');
  assert.match(nodes.templateTree.textContent, /Account onboarding/);
  assert.match(nodes.templateTree.textContent, /Sales documents/);
  const secondTemplate = descendants(nodes.templateTree).find(
    (n) => n.tagName === 'BUTTON' && n.textContent === 'Sales documents',
  );
  await secondTemplate.onclick();
  assert.equal(nodes.templateName.value, 'Sales documents');
  {
    // Delete template asks in the page (no platform dialog); Keep template deletes nothing.
    const asked = [];
    const deleted = [];
    xrm.WebApi.deleteRecord = async (table, id) => deleted.push([table, id]);
    ui.confirmInline = async (invoker, options) => {
      asked.push([invoker, options]);
      return false;
    };
    await nodes.deleteTemplate.onclick();
    assert.equal(asked.length, 1, 'Delete template confirms in the page');
    assert.equal(asked[0][0], nodes.deleteTemplate);
    // The mocked retrieveRecord answers with Account onboarding for any template.
    assert.match(
      asked[0][1].text,
      /^Delete Account onboarding and all its versions\? .*SharePoint stay as they are/,
    );
    assert.equal(asked[0][1].confirm, 'Delete template');
    assert.equal(asked[0][1].keep, 'Keep template');
    assert.equal(asked[0][1].danger, true);
    assert.deepEqual(deleted, [], 'Keep template deletes nothing');
    // Delete template deletes the selected template once and says so.
    ui.confirmInline = async (invoker, options) => {
      asked.push([invoker, options]);
      return true;
    };
    await nodes.deleteTemplate.onclick();
    assert.equal(asked.length, 2);
    assert.deepEqual(deleted, [['asx_template', 'template-2']], 'Delete template deletes it');
    assert.match(said(), /Template deleted/);
    assert.equal(nodes.revision.textContent, 'Template deleted');
    delete xrm.WebApi.deleteRecord;
    delete ui.confirmInline;
  }
  await nodes.newTemplate.onclick();
  assert.equal(nodes.templateName.disabled, false);
  assert.equal(nodes.savedRevision.options.length, 1);
  assert.equal(
    nodes.destinations.textContent.includes('Sensitive root'),
    false,
    'New template has an independent empty tree',
  );
  {
    // Tables panel: enabled tables come from asx_runtimetable; add, remove and enable are server commands.
    let enabled = ['account'];
    const profile = () => ({
      WorkerId: 'worker-1',
      Tables: enabled,
      SharePointHosts: [],
      Enabled: false,
      ProcessRecordUpdates: false,
      RowVersion: '9',
      Registration: {
        Readiness: enabled
          .map((t) => ({ Scope: t, Status: t === 'lead' ? 'Pending' : 'Ready' }))
          .concat([{ Scope: 'team', Status: 'Ready' }]),
        ExtraSteps: 0,
        StepIds: ['step-1'],
      },
    });
    const commands = [];
    const retrieve = xrm.WebApi.retrieveMultipleRecords;
    xrm.WebApi.retrieveMultipleRecords = async (name, options) =>
      name === 'asx_runtimetable'
        ? { entities: enabled.map((asx_logicalname) => ({ asx_logicalname })) }
        : name === 'asx_runtime'
          ? { entities: [{ asx_runtimeid: 'runtime-1' }] }
          : name === 'systemuser'
            ? { entities: [{ systemuserid: 'worker-1', fullname: 'Documents worker' }] }
            : retrieve(name, options);
    xrm.WebApi.online.execute = async (req) => {
      assert.equal(req.getMetadata().operationName, 'asx_RuntimeAdmin');
      const payload = JSON.parse(req.Request);
      commands.push(payload);
      if (payload.Command === 'AddTable') enabled = [...new Set([...enabled, payload.Table])];
      if (payload.Command === 'RemoveTable') enabled = enabled.filter((t) => t !== payload.Table);
      return { ok: true, json: async () => ({ Result: JSON.stringify(profile()) }) };
    };
    // AddTable and RemoveTable results reach the shell's runtime (the chip and other tabs).
    const shared = [];
    runtimeListeners.push((result) => shared.push(result));
    const buttons = () => descendants(nodes.templateTree).filter((n) => n.tagName === 'BUTTON');
    const named = (label) => buttons().find((n) => n.attrs['aria-label'] === label);
    const text = (label) => buttons().find((n) => n.textContent === label);
    // Starting the tab with the shell's runtime shows readiness badges (and, here, enabled tables).
    ui.runtime = () => profile();
    await inits.templates();
    assert.match(nodes.templateTree.textContent, /Account onboarding/);
    assert.equal(nodes.templateTree.textContent.includes('Not enabled'), false);
    assert.match(nodes.templateTree.textContent, /account.*Ready/, 'Badge shows readiness text');
    assert.deepEqual(
      nodes.enableTable.options.map((o) => o.value),
      ['', 'contact', 'lead'],
      'Picker lists document-enabled tables that are not enabled',
    );
    nodes.enableTable.value = 'lead';
    await nodes.enableTable.onchange();
    assert.deepEqual(commands.at(-1), { Command: 'AddTable', Table: 'lead' });
    assert.match(nodes.templateTree.textContent, /Lead.*Pending/);
    assert.equal(
      shared.at(-1).Tables.includes('lead'),
      true,
      'The result is shared with the shell',
    );
    assert.deepEqual(
      nodes.enableTable.options.map((o) => o.value),
      ['', 'contact'],
    );
    const count = commands.length;
    await named('Remove account').onclick();
    assert.equal(commands.length, count, 'Remove waits for in-page confirmation');
    assert.match(
      nodes.templateTree.textContent,
      /Capture for this table stops\. Its templates are kept\./,
    );
    await text('Cancel').onclick();
    assert.equal(commands.length, count);
    await named('Remove account').onclick();
    await text('Remove table').onclick();
    assert.deepEqual(commands.at(-1), { Command: 'RemoveTable', Table: 'account' });
    assert.match(nodes.templateTree.textContent, /Not enabled/);
    assert.match(nodes.templateTree.textContent, /Sales documents/, 'Templates are kept');
    await named('Enable account').onclick();
    assert.deepEqual(commands.at(-1), { Command: 'AddTable', Table: 'account' });
    assert.equal(nodes.templateTree.textContent.includes('Not enabled'), false);
    xrm.WebApi.retrieveMultipleRecords = retrieve;
  }
  console.log(
    'PASS admin handler contracts: destination isolation, site filter, stable keys, child and root conditions, independent folder lookups, optional schedule, top-bar actions, table-first workspace, server preview, stale-preview invalidation and workspace navigation, table readiness from the shared runtime, Tables panel results shared with the shell, empty input validation and table reset. Mocked DOM/API; visual QA separate.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
