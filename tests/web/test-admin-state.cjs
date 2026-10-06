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
const win = { parent: { Xrm: xrm }, addEventListener: () => {} };
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
  await new Promise(setImmediate);
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
    assert.match(nodes.status.textContent, /Revision published/);
    assert(nodes.status.textContent.includes(notice), 'The template editor shows publish notices');
    assert.equal(nodes.status.className, '', 'A publish notice is not an error');
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
        Tables: runtimeProfile.Tables,
        RowVersion: '2',
        Migrated: true,
        Registration: {
          Readiness: runtimeProfile.Tables.map((t) => ({ Scope: t, Status: 'Ready' })).concat([
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
  assert.equal(nodes.runtimeTables, undefined, 'Runtime administration has no table picker');
  assert.equal(nodes.runtimeWorker.value, 'worker-1');
  nodes.runtimeRecordUpdates.checked = true;
  await nodes.saveRuntime.onclick();
  assert.equal(runtimeSaves.at(-1).ProcessRecordUpdates, true);
  assert.equal('Tables' in runtimeSaves.at(-1), false, 'Save no longer sends Tables');
  assert.doesNotMatch(
    nodes.status.textContent,
    /registration verified/,
    'A pause or resume does not verify registration',
  );
  assert.match(nodes.status.textContent, /Runtime profile saved/);
  assert.match(nodes.runtimeReadiness.textContent, /account.*Ready/);
  assert.equal(nodes.runtimePending.hidden, true);
  await nodes.unregisterRuntime.onclick();
  assert.equal(nodes.unregisterConfirm.hidden, false);
  assert.equal(runtimeSaves.at(-1).Command, 'Save', 'Unregister waits for confirmation');
  await nodes.confirmUnregister.onclick();
  assert.equal(runtimeSaves.at(-1).Command, 'Unregister');
  {
    // Pausing with a disabled or deleted worker: the worker list omits it, but Save must still
    // send the configured WorkerId so the server takes the toggle-only path.
    const select = nodes.runtimeWorker;
    let selected = '';
    Object.defineProperty(select, 'value', {
      configurable: true,
      get: () => selected,
      // Like a browser select: a value without a matching option selects nothing.
      set: (v) => (selected = select.children.some((o) => o.value === v) ? v : ''),
    });
    const retrieve = xrm.WebApi.retrieveMultipleRecords;
    xrm.WebApi.retrieveMultipleRecords = async (name) => ({
      entities: name === 'asx_runtime' ? [{ asx_runtimeid: 'runtime-1' }] : [],
    });
    runtimeProfile = { ...runtimeProfile, WorkerId: 'worker-gone', Enabled: true };
    await nodes.loadRuntime.onclick();
    assert.equal(select.value, 'worker-gone');
    assert.match(select.textContent, /Configured worker \(disabled or not found\)/);
    nodes.runtimeEnabled.checked = false;
    await nodes.saveRuntime.onclick();
    assert.equal(runtimeSaves.at(-1).Command, 'Save');
    assert.equal(runtimeSaves.at(-1).WorkerId, 'worker-gone', 'Pause keeps the configured worker');
    assert.equal(runtimeSaves.at(-1).Enabled, false);
    xrm.WebApi.retrieveMultipleRecords = retrieve;
    delete select.value;
    select.value = '';
  }
  {
    const failed = [
      {
        asyncoperationid: 'job-1',
        _regardingobjectid_value: 'rec-1',
        '_regardingobjectid_value@Microsoft.Dynamics.CRM.lookuplogicalname': 'account',
        message: 'Runtime identity/source scope is incomplete.',
        createdon: '2026-10-05T10:00:00Z',
      },
    ];
    xrm.WebApi.retrieveMultipleRecords = async (name, options) => {
      if (name === 'asyncoperation') {
        assert.match(options, /statuscode eq 31/);
        assert.match(options, /_owningextensionid_value eq step-1/);
        return { entities: failed };
      }
      return { entities: [] };
    };
    // Templates 'template-1' and 'template-2' (both account) were loaded earlier in this file.
    const queued = [];
    xrm.WebApi.online.execute = async (req) => {
      assert.equal(req.getMetadata().operationName, 'asx_ManageWork');
      queued.push(JSON.parse(req.Request));
      return { ok: true, json: async () => ({ Result: JSON.stringify({ Status: 'Pending' }) }) };
    };
    await nodes.loadFailedJobs.onclick();
    assert.match(nodes.failedJobs.textContent, /account.*rec-1.*incomplete/);
    nodes.failedJobs.children[0].children[0].checked = true;
    await nodes.replanFailed.onclick();
    assert.deepEqual(
      queued.map((q) => [q.Command, q.TemplateId, q.RecordId]),
      [
        ['Queue', 'template-1', 'rec-1'],
        ['Queue', 'template-2', 'rec-1'],
      ],
    );
    assert.notEqual(queued[0].RequestId, queued[1].RequestId);
  }
  {
    xrm.WebApi.retrieveMultipleRecords = async (name) =>
      name === 'asyncoperation'
        ? {
            entities: [
              {
                asyncoperationid: 'job-2',
                _regardingobjectid_value: 'rec-9',
                '_regardingobjectid_value@Microsoft.Dynamics.CRM.lookuplogicalname': 'contact',
                message: 'Failed.',
                createdon: '2026-10-05T11:00:00Z',
              },
            ],
          }
        : { entities: [] };
    let calls = 0;
    xrm.WebApi.online.execute = async () => {
      calls++;
      return { ok: true, json: async () => ({ Result: JSON.stringify({ Status: 'Pending' }) }) };
    };
    await nodes.loadFailedJobs.onclick();
    nodes.failedJobs.children[0].children[0].checked = true;
    await nodes.replanFailed.onclick();
    assert.equal(calls, 0);
    assert.match(nodes.status.textContent, /contact.*no template/i);
  }
  {
    // Blocked records: outbox rows whose planning failed, each with an in-page Retry.
    const blockedRows = [
      {
        asx_payload: JSON.stringify({
          Key: 'request:aaa',
          Status: 'Blocked',
          Table: 'account',
          RecordId: 'rec-7',
          Notices: ['Planning failed. Check that the site and library are approved.'],
        }),
        modifiedon: '2026-10-05T16:02:40Z',
      },
    ];
    const blockedQueries = [],
      waitingQueries = [];
    // Waiting records: record plans that skipped folders until the record changes.
    const waitingRows = [
      {
        asx_payload: JSON.stringify({
          Key: 'recordplan:t1:rec-8',
          Status: 'Waiting',
          Table: 'account',
          TemplateId: 'template-1',
          RecordId: 'rec-8',
          Notices: ["Folder name 'A:B' was adjusted to 'A-B' for SharePoint."],
          Waiting: [
            "Folder 'general/root' is waiting for 'root.name' to have a value. Fill in 'root.name', then replan the record.",
            "Folder 'general/b' has the same name 'A' as 'general/a'; it waits until the names differ. Change the record so the folder gets a usable name of its own, then replan the record.",
          ],
        }),
        modifiedon: '2026-10-05T16:04:00Z',
      },
    ];
    xrm.WebApi.retrieveMultipleRecords = async (name, options, size) => {
      assert.equal(name, 'asx_outbox');
      if (/Waiting|wait2/.test(options)) {
        waitingQueries.push([options, size]);
        return waitingQueries.length === 1
          ? { entities: waitingRows, nextLink: 'https://example.test/api?$skiptoken=wait2' }
          : {
              entities: [
                {
                  asx_payload: JSON.stringify({
                    Key: 'recordplan:t2:rec-9',
                    Status: 'Waiting',
                    Table: 'contact',
                    RecordId: 'rec-9',
                    Waiting: [],
                    Notices: [],
                  }),
                  modifiedon: '2026-10-05T16:05:00Z',
                },
              ],
            };
      }
      blockedQueries.push([options, size]);
      return blockedQueries.length === 1
        ? { entities: blockedRows, nextLink: 'https://example.test/api?$skiptoken=page2' }
        : {
            entities: [
              {
                asx_payload: JSON.stringify({ Key: 'team-event:x', Status: 'Blocked' }),
                modifiedon: '2026-10-05T16:03:00Z',
              },
            ],
          };
    };
    const retried = [];
    xrm.WebApi.online.execute = async (req) => {
      assert.equal(req.getMetadata().operationName, 'asx_ManageWork');
      retried.push(JSON.parse(req.Request));
      return { ok: true, json: async () => ({ Result: JSON.stringify({ Status: 'Pending' }) }) };
    };
    await nodes.loadBlockedRecords.onclick();
    assert.match(blockedQueries[0][0], /asx_status eq 'Blocked'/);
    assert.equal(blockedQueries[0][1], 50, 'Same page size as the failed-jobs list');
    assert.match(
      nodes.blockedRecords.textContent,
      /account.*rec-7.*site and library are approved.*2026-10-05T16:02:40Z/,
    );
    assert.equal(nodes.moreBlockedRecords.hidden, false);
    await nodes.moreBlockedRecords.onclick();
    assert.equal(blockedQueries[1][0], '?$skiptoken=page2');
    assert.equal(nodes.blockedRecords.children.length, 2, 'Load more appends');
    assert.match(nodes.blockedRecords.children[1].textContent, /team-event:x/);
    assert.equal(nodes.moreBlockedRecords.hidden, true);
    const retry = nodes.blockedRecords.children[0].children.find((n) => n.tagName === 'BUTTON');
    assert.equal(retry.textContent, 'Retry');
    await retry.onclick();
    assert.deepEqual(retried, [{ Command: 'RetryOutbox', Key: 'request:aaa' }]);
    assert.equal(retry.disabled, true);
    assert.match(nodes.status.textContent, /queued for planning again/i);
    // The Waiting section lists record plans with folders that wait, each with Replan.
    assert.match(html, /<h5>Waiting<\/h5>/);
    assert.match(waitingQueries[0][0], /asx_status eq 'Waiting'/);
    assert.match(waitingQueries[0][0], /orderby=modifiedon desc/);
    assert.equal(waitingQueries[0][1], 50, 'Same page size as the blocked-records list');
    assert.equal(
      nodes.waitingRecords.children[0].textContent,
      "account rec-8 · Folder 'general/root' is waiting for 'root.name' to have a value. Fill in 'root.name', then replan the record. (+1 more) · 2026-10-05T16:04:00Z Replan",
    );
    assert.equal(nodes.moreWaitingRecords.hidden, false);
    await nodes.moreWaitingRecords.onclick();
    assert.equal(waitingQueries[1][0], '?$skiptoken=wait2');
    assert.equal(nodes.waitingRecords.children.length, 2, 'Load more appends');
    assert.equal(nodes.moreWaitingRecords.hidden, true);
    const second = nodes.waitingRecords.children[1].children.find((n) => n.tagName === 'BUTTON');
    assert.equal(second.disabled, true, 'Replan needs the template and the record');
    const replan = nodes.waitingRecords.children[0].children.find((n) => n.tagName === 'BUTTON');
    assert.equal(replan.textContent, 'Replan');
    retried.length = 0;
    await replan.onclick();
    assert.equal(retried.length, 1);
    assert.deepEqual(
      [retried[0].Command, retried[0].TemplateId, retried[0].RecordId],
      ['Replan', 'template-1', 'rec-8'],
    );
    assert.ok(retried[0].RequestId, 'Replan carries a fresh request ID');
    assert.equal(replan.disabled, true);
    assert.match(nodes.status.textContent, /queued for replanning/i);
    xrm.WebApi.retrieveMultipleRecords = async () => ({ entities: [] });
    await nodes.loadBlockedRecords.onclick();
    assert.equal(nodes.blockedRecords.children.length, 0);
    assert.equal(nodes.moreBlockedRecords.hidden, true);
    assert.equal(nodes.waitingRecords.children.length, 0);
    assert.equal(nodes.moreWaitingRecords.hidden, true);
    assert.equal(nodes.status.textContent, 'No blocked records. No waiting records.');
  }
  {
    // Blocked jobs: Blocked asx_operation rows of every kind, each with the operator Retry in-page.
    assert.match(html, /<h4>Blocked jobs<\/h4>/);
    assert.doesNotMatch(html, /Blocked folder jobs/);
    const jobRows = [
      {
        asx_payload: JSON.stringify({
          Key: 'folderjob:abc',
          Status: 'Blocked',
          ErrorCode: 'FileAtExpectedFolderPath',
        }),
        modifiedon: '2026-10-05T17:00:00Z',
      },
    ];
    const jobQueries = [];
    xrm.WebApi.retrieveMultipleRecords = async (name, options, size) => {
      assert.equal(name, 'asx_operation');
      jobQueries.push([options, size]);
      return jobQueries.length === 1
        ? { entities: jobRows, nextLink: 'https://example.test/api?$skiptoken=jobs2' }
        : {
            entities: [
              {
                asx_payload: JSON.stringify({ Key: 'policywork:p', Status: 'Blocked' }),
                modifiedon: '2026-10-05T17:01:00Z',
              },
            ],
          };
    };
    const retried = [];
    xrm.WebApi.online.execute = async (req) => {
      assert.equal(req.getMetadata().operationName, 'asx_ManageWork');
      retried.push(JSON.parse(req.Request));
      return { ok: true, json: async () => ({ Result: JSON.stringify({ Status: 'Pending' }) }) };
    };
    await nodes.loadBlockedJobs.onclick();
    assert.match(jobQueries[0][0], /asx_status eq 'Blocked'/);
    assert.match(jobQueries[0][0], /orderby=modifiedon desc/);
    assert.equal(jobQueries[0][1], 50, 'Same page size as the blocked-records list');
    assert.match(
      nodes.blockedJobs.textContent,
      /folderjob:abc.*Folder job.*FileAtExpectedFolderPath.*2026-10-05T17:00:00Z/,
    );
    assert.equal(nodes.moreBlockedJobs.hidden, false);
    await nodes.moreBlockedJobs.onclick();
    assert.equal(jobQueries[1][0], '?$skiptoken=jobs2');
    assert.equal(nodes.blockedJobs.children.length, 2, 'Load more appends');
    assert.match(
      nodes.blockedJobs.children[1].textContent,
      /policywork:p.*Access policy.*No notice/,
    );
    assert.equal(nodes.moreBlockedJobs.hidden, true);
    const retry = nodes.blockedJobs.children[0].children.find((n) => n.tagName === 'BUTTON');
    assert.equal(retry.textContent, 'Retry');
    await retry.onclick();
    assert.deepEqual(retried, [{ Command: 'Retry', Key: 'folderjob:abc' }]);
    assert.equal(retry.disabled, true);
    assert.match(nodes.status.textContent, /queued to run again/i);
    xrm.WebApi.online.execute = async () => ({
      ok: true,
      json: async () => ({ Result: JSON.stringify({ Status: 'Applied' }) }),
    });
    const second = nodes.blockedJobs.children[1].children.find((n) => n.tagName === 'BUTTON');
    await second.onclick();
    assert.match(nodes.status.textContent, /Applied; nothing to retry/);
    // Cancel sits next to Retry and asks in the page first, like Remove in the Tables panel.
    const row = nodes.blockedJobs.children[0];
    const cancel = row.children.find((n) => n.tagName === 'BUTTON' && n.textContent === 'Cancel');
    assert(cancel, 'Blocked jobs offer Cancel next to Retry');
    const cancelled = [];
    xrm.WebApi.online.execute = async (req) => {
      assert.equal(req.getMetadata().operationName, 'asx_ManageWork');
      cancelled.push(JSON.parse(req.Request));
      return { ok: true, json: async () => ({ Result: JSON.stringify({ Status: 'Cancelled' }) }) };
    };
    cancel.onclick();
    assert.equal(cancelled.length, 0, 'Nothing is cancelled before the admin confirms');
    const ask = row.children.find((n) => n.className === 'callout');
    assert.equal(ask.hidden, false);
    assert.equal(
      ask.children[0].textContent,
      'Cancel folder job folderjob:abc? Nothing in SharePoint is undone or deleted.',
    );
    const [go, keep] = ask.children.filter((n) => n.tagName === 'BUTTON');
    assert.equal(go.textContent, 'Cancel job');
    keep.onclick();
    assert.equal(ask.hidden, true);
    assert.equal(cancelled.length, 0);
    cancel.onclick();
    await go.onclick();
    assert.deepEqual(cancelled, [{ Command: 'Cancel', Key: 'folderjob:abc' }]);
    assert.equal(ask.hidden, true);
    assert.equal(cancel.disabled, true);
    assert.equal(retry.disabled, true);
    assert.equal(
      nodes.status.textContent,
      'Job cancelled. Nothing in SharePoint was undone or deleted.',
    );
    xrm.WebApi.retrieveMultipleRecords = async () => ({ entities: [] });
    await nodes.loadBlockedJobs.onclick();
    assert.equal(nodes.blockedJobs.children.length, 0);
    assert.equal(nodes.moreBlockedJobs.hidden, true);
    assert.equal(nodes.status.textContent, 'No blocked jobs.');
  }
  {
    // Tables panel: enabled tables come from asx_runtimetable; add, remove and enable are server commands.
    let enabled = ['account'];
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
      return {
        ok: true,
        json: async () => ({
          Result: JSON.stringify({
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
          }),
        }),
      };
    };
    const buttons = () => descendants(nodes.templateTree).filter((n) => n.tagName === 'BUTTON');
    const named = (label) => buttons().find((n) => n.attrs['aria-label'] === label);
    const text = (label) => buttons().find((n) => n.textContent === label);
    // Loading the runtime profile refreshes readiness badges (and, here, enabled tables).
    await nodes.loadRuntime.onclick();
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
    assert.match(
      nodes.runtimeReadiness.textContent,
      /lead.*Pending/,
      'Runtime readiness refreshed',
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
    'PASS admin handler contracts: destination isolation, site filter, stable keys, child and root conditions, independent folder lookups, optional schedule, top-bar actions, table-first workspace, server preview, stale-preview invalidation and workspace navigation, missing-probe/runtime setup, blocked-record retry, waiting-record replan, blocked job retry and cancel, empty input validation and table reset. Mocked DOM/API; visual QA separate.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
