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
const tabs = ['author', 'access', 'administration'].map((view) => {
  const n = new Node('button');
  n.dataset.view = view;
  return n;
});
const document = {
  getElementById: (id) => nodes[id],
  createElement: (t) => new Node(t),
  createElementNS: (ns, t) => {
    const n = new Node(t);
    n.namespaceURI = ns;
    return n;
  },
  querySelectorAll: () => tabs,
};
const requests = [],
  table = {
    LogicalName: 'account',
    EntitySetName: 'accounts',
    PrimaryIdAttribute: 'accountid',
    PrimaryNameAttribute: 'name',
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
const xrm = {
  Utility: { getGlobalContext: () => ({ getClientUrl: () => 'https://example.test' }) },
  WebApi: {
    retrieveRecord: async (name, id) => ({
      asx_templateid: id,
      asx_name: 'Account onboarding',
      asx_table: 'account',
    }),
    retrieveMultipleRecords: async (name) => ({
      entities:
        name === 'asx_library'
          ? libraries
          : name === 'asx_site'
            ? [
                { asx_siteid: 'site-a', asx_name: 'Delivery' },
                { asx_siteid: 'site-b', asx_name: 'Commercial' },
              ]
            : [],
    }),
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
            : [table],
  }),
});
vm.runInNewContext(fs.readFileSync(path.join(base, 'admin.js'), 'utf8'), {
  document,
  window: { parent: { Xrm: xrm }, addEventListener: () => {} },
  location: { hash: '' },
  fetch,
  console,
  crypto: { randomUUID: () => 'test' },
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
  await new Promise(setImmediate);
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
  assert(!nodes.previewTrees.textContent.includes('BindingKey'));
  tabs[2].onclick();
  assert.equal(nodes.templateActions.hidden, true);
  assert.equal(nodes.recordTools.hidden, false);
  tabs[0].onclick();
  assert.equal(nodes.recordTools.hidden, true);
  tabs[1].onclick();
  assert.equal(nodes['author-view'].hidden, true);
  assert.equal(nodes.access.hidden, false);
  assert.equal(nodes.runtime.hidden, true);
  tabs[0].onclick();
  await change(control('Readable folder name'), 'Changed again');
  assert.match(nodes.previewTrees.textContent, /Draft changed/);
  assert.equal(nodes.preview.disabled, true, 'Editing invalidates saved preview');
  assert(
    requests.every((r) =>
      ['asx_CreateDraft', 'asx_PreviewTemplate'].includes(r.getMetadata().operationName),
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
  const before = requests.length;
  await nodes.loadRuntime.onclick();
  assert.match(nodes.status.textContent, /No runtime profile is installed/);
  await nodes.recoverOperation.onclick();
  assert.match(nodes.status.textContent, /Select the expired operation/);
  assert.equal(requests.length, before, 'Prerequisite errors must not submit product APIs');
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
  await nodes.newTemplate.onclick();
  assert.equal(nodes.templateName.disabled, false);
  assert.equal(nodes.savedRevision.options.length, 1);
  assert.equal(
    nodes.destinations.textContent.includes('Sensitive root'),
    false,
    'New template has an independent empty tree',
  );
  xrm.WebApi.retrieveMultipleRecords = async (name) => ({
    entities:
      name === 'asx_runtime'
        ? [{ asx_runtimeid: 'runtime-1' }]
        : name === 'systemuser'
          ? [{ systemuserid: 'worker-1', fullname: 'Documents worker' }]
          : [],
  });
  let runtimeProfile = {
    WorkerId: 'worker-1',
    Tables: ['account'],
    SharePointHosts: ['example.sharepoint.com'],
    Enabled: false,
    ProcessRecordUpdates: false,
    RowVersion: '1',
    Migrated: false,
    Registration: {
      Readiness: [
        { Scope: 'account', Status: 'Pending' },
        { Scope: 'team', Status: 'Pending' },
      ],
      ExtraSteps: 0,
      StepIds: [],
    },
  };
  const runtimeSaves = [];
  xrm.WebApi.online.execute = async (req) => {
    assert.equal(req.getMetadata().operationName, 'asx_RuntimeAdmin');
    const payload = JSON.parse(req.Request);
    if (payload.Command === 'Save') {
      runtimeSaves.push(payload);
      runtimeProfile = {
        ...payload,
        RowVersion: '2',
        Migrated: true,
        Registration: {
          Readiness: payload.Tables.map((t) => ({ Scope: t, Status: 'Ready' })).concat([
            { Scope: 'team', Status: 'Ready' },
          ]),
          ExtraSteps: 0,
          StepIds: ['step-1'],
        },
      };
    }
    if (payload.Command === 'Unregister') {
      runtimeSaves.push(payload);
      runtimeProfile.Registration.Readiness.forEach((r) => (r.Status = 'Pending'));
    }
    return { ok: true, json: async () => ({ Result: JSON.stringify(runtimeProfile) }) };
  };
  await nodes.loadRuntime.onclick();
  assert.match(nodes.runtimePending.textContent, /pending registration/i);
  assert.match(nodes.runtimeReadiness.textContent, /account.*Pending/);
  assert.deepEqual(
    nodes.runtimeTables.options.filter((o) => o.selected).map((o) => o.value),
    ['account'],
  );
  assert.equal(nodes.runtimeWorker.value, 'worker-1');
  nodes.runtimeRecordUpdates.checked = true;
  await nodes.saveRuntime.onclick();
  assert.equal(runtimeSaves.at(-1).ProcessRecordUpdates, true);
  assert.deepEqual(runtimeSaves.at(-1).Tables, ['account']);
  assert.match(nodes.runtimeReadiness.textContent, /account.*Ready/);
  assert.equal(nodes.runtimePending.hidden, true);
  await nodes.unregisterRuntime.onclick();
  assert.equal(nodes.unregisterConfirm.hidden, false);
  assert.equal(runtimeSaves.at(-1).Command, 'Save', 'Unregister waits for confirmation');
  await nodes.confirmUnregister.onclick();
  assert.equal(runtimeSaves.at(-1).Command, 'Unregister');
  console.log(
    'PASS admin handler contracts: destination isolation, site filter, stable keys, child and root conditions, independent folder lookups, optional schedule, top-bar actions, table-first workspace, server preview, stale-preview invalidation and workspace navigation, missing-probe/runtime setup, empty input validation and table reset. Mocked DOM/API; visual QA separate.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
