// Real workspace handlers with an isolated DOM/API contract; no live SharePoint calls.
const fs = require('fs'),
  vm = require('vm'),
  assert = require('assert/strict'),
  path = require('path');
class Node {
  constructor(tag) {
    this.tagName = tag;
    this.children = [];
    this.value = '';
    this.hidden = false;
    this.attrs = {};
    this._text = '';
  }
  set textContent(v) {
    this._text = String(v);
    this.children = [];
  }
  get textContent() {
    return this._text + this.children.map((c) => c.textContent).join('');
  }
  append(...n) {
    this.children.push(...n);
  }
  replaceChildren(...n) {
    this._text = '';
    this.children = n;
  }
  setAttribute(k, v) {
    this.attrs[k] = v;
  }
  get options() {
    return this.children;
  }
}
const base = path.resolve(__dirname, '../../client/admin'),
  html = fs.readFileSync(path.join(base, 'index.html'), 'utf8'),
  nodes = {};
for (const m of html.matchAll(/<([a-z]+)[^>]*\bid="([^"]+)"[^>]*>/g)) nodes[m[2]] = new Node(m[1]);
nodes.access.hidden = true;
const id = (n) => String(n).padStart(8, '0') + '-0000-0000-0000-000000000000';
const site = {
    asx_siteid: id(1),
    asx_name: 'Delivery',
    asx_approved: true,
    _asx_nativeid_value: id(2),
  },
  lib = {
    asx_libraryid: id(3),
    asx_name: 'General',
    _asx_siteid_value: id(1),
    asx_approved: true,
    asx_policyapplied: true,
  };
const requests = [],
  timers = [];
let discovery = false;
let operationStatus = 'Verified';
let inspectIssue = null,
  inspectFails = false;
let policy = { Status: 'Applied', RowVersion: '1', Policy: { Desired: [], Applied: [] } },
  refresh = 0;
const xrm = {
  Utility: { getGlobalContext: () => ({ getClientUrl: () => 'https://example.test' }) },
  Navigation: { openUrl: () => {} },
  WebApi: {
    retrieveMultipleRecords: async (table) => ({
      entities:
        table === 'asx_site'
          ? [site]
          : table === 'asx_library'
            ? [lib]
            : table === 'team'
              ? [{ teamid: id(4), name: 'Operations' }]
              : table === 'sharepointsite'
                ? [{ sharepointsiteid: id(2), name: 'Delivery' }]
                : [],
    }),
    retrieveRecord: async (table, key) =>
      table === 'asx_library' ? lib : table === 'asx_site' ? site : { name: 'Operations' },
    online: {
      execute: async (req) => {
        const command = JSON.parse(req.Request);
        requests.push(command);
        if (command.Command === 'Inspect' && inspectFails)
          throw new Error('Temporary status request failure');
        let result;
        if (command.Command === 'GetPolicy') result = policy;
        else if (command.Command === 'ApplyPolicy')
          result = policy = {
            Status: 'Queued',
            RowVersion: '2',
            Policy: { Desired: command.Entries, Applied: [], OperationKey: 'policywork:test' },
          };
        else if (command.Command === 'Inspect')
          result = discovery
            ? {
                Status: 'Discovered',
                Key: 'catalogprobe:discovery',
                RowVersion: '7',
                Observation: {
                  SiteId: id(1),
                  Libraries: [{ Id: id(8), Title: 'Archive' }],
                  NextLibraries: 'next',
                },
              }
            : { Status: operationStatus, Issue: inspectIssue, CatalogId: id(1) };
        else if (command.Command === 'DiscoverLibraries') {
          discovery = true;
          result = { Status: 'Pending', Key: 'catalogprobe:discovery' };
        } else
          result = {
            Status: 'Pending',
            Key: command.Command === 'CreateLibrary' ? 'librarycreate:test' : 'catalogprobe:test',
          };
        return { ok: true, json: async () => ({ Result: JSON.stringify(result) }) };
      },
    },
  },
};
const window = {
  Xrm: xrm,
  AsxdAdmin: {
    refreshCatalog: async () => {
      refresh++;
    },
  },
};
vm.runInNewContext(fs.readFileSync(path.join(base, 'sites-access.js'), 'utf8'), {
  window,
  document: { getElementById: (id) => nodes[id], createElement: (t) => new Node(t) },
  URL,
  crypto: { randomUUID: () => id(9) },
  setTimeout: (fn) => timers.push(fn),
});
(async () => {
  await window.AsxdSites.open();
  assert.equal(nodes['ad-site-title'].textContent, 'Delivery');
  assert.equal(nodes['ad-library-title'].textContent, 'General');
  nodes['ad-add-team'].onclick();
  nodes['ad-team-choice'].value = id(4);
  nodes['ad-team-access'].value = 'Read';
  nodes['ad-stage-team'].onclick();
  assert.equal(
    requests.filter((r) => r.Command === 'ApplyPolicy').length,
    0,
    'Staging does not onboard syncing',
  );
  assert.equal(nodes['ad-apply'].disabled, false);
  await nodes['ad-apply'].onclick();
  const applied = requests.find((r) => r.Command === 'ApplyPolicy');
  assert.equal(applied.LibraryId, id(3));
  assert.equal(applied.Entries[0].TeamId, id(4));
  assert.equal(applied.ReadRole, undefined, 'Role definitions stay server-owned');
  assert.equal(nodes['ad-apply'].disabled, true);
  policy = {
    Status: 'Applied',
    RowVersion: '3',
    Policy: { Desired: applied.Entries, Applied: applied.Entries },
  };
  policy.Policy.Notices = [
    "Team 'Operations': Integration App was not added to the library group because it is an application user.",
    '<b>SharePoint did not add ghost@example.com</b>',
  ];
  nodes.access.hidden = false;
  await timers.shift()();
  assert.match(nodes['ad-change-status'].textContent, /confirmed/);
  assert.equal(nodes['ad-access-notices'].hidden, false, 'Access notices are shown');
  assert.equal(nodes['ad-access-notices'].children.length, 2);
  assert.match(nodes['ad-access-notices'].textContent, /Integration App was not added/);
  assert.equal(
    nodes['ad-access-notices'].children[1].textContent,
    '<b>SharePoint did not add ghost@example.com</b>',
    'Notices are text, never markup',
  );
  delete policy.Policy.Notices;
  assert(refresh > 0);
  const select = nodes['ad-teams'].children[0].children[1].children[0];
  select.value = 'None';
  select.onchange();
  policy = {
    Status: 'Applied',
    RowVersion: '4',
    Policy: { Desired: applied.Entries, Applied: applied.Entries },
  };
  await nodes['ad-apply'].onclick();
  const removal = requests.filter((r) => r.Command === 'ApplyPolicy').at(-1);
  assert.equal(
    removal.RowVersion,
    '4',
    'Removal must refresh the policy version after background membership synchronization',
  );
  assert.equal(removal.Entries[0].Access, 'None');
  policy = {
    Status: 'Applied',
    RowVersion: '5',
    Policy: { Desired: removal.Entries, Applied: removal.Entries },
  };
  await timers.shift()();
  const edit = nodes['ad-teams'].children[0].children[1].children[0];
  edit.value = 'Read';
  edit.onchange();
  const beforeConflict = requests.filter((r) => r.Command === 'ApplyPolicy').length;
  policy = {
    Status: 'Applied',
    RowVersion: '6',
    Policy: { Desired: [{ TeamId: id(4), Access: 'Contribute' }], Applied: [] },
  };
  await nodes['ad-apply'].onclick();
  assert.equal(
    requests.filter((r) => r.Command === 'ApplyPolicy').length,
    beforeConflict,
    'Do not overwrite another administrator access change',
  );
  assert.match(nodes['ad-message'].textContent, /changed.*reload/i);
  policy = {
    Status: 'Applied',
    RowVersion: '7',
    Policy: { Desired: removal.Entries, Applied: removal.Entries, OperationKey: 'background' },
  };
  await nodes['ad-apply'].onclick();
  assert.equal(
    requests.filter((r) => r.Command === 'ApplyPolicy').length,
    beforeConflict,
    'Do not replace active worker generation',
  );
  assert.match(nodes['ad-message'].textContent, /in progress/i);
  policy = {
    Status: 'Applied',
    RowVersion: '8',
    Policy: { Desired: removal.Entries, Applied: removal.Entries },
  };
  await nodes['ad-apply'].onclick();
  policy = {
    Status: 'Applied',
    RowVersion: '9',
    Policy: { Desired: [{ TeamId: id(4), Access: 'Read' }], Applied: [] },
  };
  await timers.shift()();
  await nodes['ad-add-site'].onclick();
  nodes['ad-native'].value = id(2);
  await nodes['ad-validate'].onclick();
  assert.equal(requests.find((r) => r.Command === 'AddSite').NativeSiteId, id(2));
  await timers.shift()();
  assert.match(nodes['ad-provision-progress'].textContent, /Finishing setup/);
  const verifiedBox = nodes['ad-provision-progress'].children[0];
  await timers.shift()();
  assert.equal(
    nodes['ad-provision-progress'].children[0],
    verifiedBox,
    'Unchanged verification must not redraw',
  );
  operationStatus = 'ExternalUnknown';
  await timers.shift()();
  assert.match(nodes['ad-provision-progress'].textContent, /Waiting for confirmation/);
  assert.doesNotMatch(nodes['ad-provision-progress'].textContent, /Administration|Needs attention/);
  operationStatus = 'RecoveryRequired';
  inspectIssue = 'Library request outcome is unknown. Reconcile the original run before retry.';
  await timers.shift()();
  assert.match(nodes['ad-provision-progress'].textContent, /Needs attention/);
  assert.match(nodes['ad-provision-progress'].textContent, /Reconcile the original run/);
  operationStatus = 'RetryWait';
  inspectIssue = 'TransientReadFailure';
  await timers.shift()();
  assert.match(nodes['ad-provision-progress'].textContent, /Retrying automatically/);
  assert.doesNotMatch(nodes['ad-provision-progress'].textContent, /TransientReadFailure/);
  operationStatus = 'Blocked';
  inspectIssue = 'Grant access before retrying.';
  await timers.shift()();
  assert.match(nodes['ad-provision-progress'].textContent, /Needs attention/);
  assert.match(nodes['ad-provision-progress'].textContent, /Grant access before retrying/);
  inspectFails = true;
  await timers.shift()();
  assert.equal(nodes['ad-poll-status'].hidden, false);
  assert.match(nodes['ad-poll-status'].textContent, /may still be running/);
  inspectFails = false;
  operationStatus = 'Verified';
  inspectIssue = null;
  await timers.shift()();
  assert.equal(nodes['ad-poll-status'].hidden, true);
  operationStatus = 'Approved';
  await timers.shift()();
  assert.match(nodes['ad-message'].textContent, /ready/);
  assert.match(nodes['ad-provision-progress'].textContent, /Setup completed/);
  const readySite = nodes['ad-sites'].children[0],
    readyRequests = requests.length;
  await timers.shift()();
  assert.equal(requests.length, readyRequests, 'Idle workspace must not poll APIs');
  assert.equal(nodes['ad-sites'].children[0], readySite, 'Idle workspace preserves DOM and focus');
  nodes['ad-create'].onclick();
  nodes['ad-library-name'].value = 'Projects';
  nodes['ad-initial-team'].value = id(4);
  nodes['ad-initial-access'].value = 'Contribute';
  await nodes['ad-provision'].onclick();
  const created = requests.find((r) => r.Command === 'CreateLibrary');
  assert.equal(created.SiteId, id(1));
  assert.equal(created.Entries[0].Access, 'Contribute');
  await window.AsxdSites.selectLibrary(id(3));
  assert.equal(nodes['ad-library-title'].textContent, 'General');
  await nodes['ad-existing'].onclick();
  await timers.shift()();
  assert.equal(nodes['ad-existing-form'].hidden, false);
  assert.equal(nodes['ad-existing-more'].hidden, false);
  await nodes['ad-existing-choices'].children[0].onclick();
  const existing = requests.find((r) => r.Command === 'AddLibrary');
  assert.equal(existing.ListId, id(8));
  assert.equal(existing.NativeParentId, undefined, 'Native navigation is automatic');
  assert(!requests.some((r) => r.Command === 'RegisterTeam'), 'No manual team mapping step');
  lib.asx_policyapplied = false;
  await window.AsxdSites.selectLibrary(id(3));
  assert.equal(
    nodes['ad-library-status'].textContent,
    'Ready for folder templates',
    'An approved library is ready for templates before its access is applied',
  );
  assert.equal(nodes['ad-library-access'].textContent, 'Access setup pending');
  lib.asx_policyapplied = true;
  await window.AsxdSites.selectLibrary(id(3));
  assert.equal(nodes['ad-library-access'].textContent, 'Access applied');
  lib.asx_approved = false;
  await window.AsxdSites.selectLibrary(id(3));
  assert.equal(nodes['ad-library-status'].textContent, 'Needs attention');
  lib.asx_approved = true;
  console.log(
    'PASS Sites & access handlers: staging versus apply, automatic onboarding request, initial library teams, completion polling, and author deep link. Mocked APIs; connected acceptance pending.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
