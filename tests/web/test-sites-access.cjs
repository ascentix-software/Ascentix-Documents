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
// The refusal RemoveLibrary answers with, or null to remove.
let removalRefusal = null;
// The refusal RetrySetup answers with, or null to retry.
let retrySetupRefusal = null;
const libraryQueries = [],
  siteQueries = [];
const requests = [],
  // The API each request went to, in the same order as requests.
  apis = [],
  timers = [],
  teamQueries = [];
// Owner, Entra group and Microsoft 365 group teams. Dataverse fills group team members only as
// people sign in, so the group itself is granted and its members are never listed here.
const teams = [
  { teamid: id(4), name: 'Operations', teamtype: 0 },
  {
    teamid: id(5),
    name: 'Finance',
    teamtype: 2,
    membershiptype: 0,
    azureactivedirectoryobjectid: id(15),
  },
  {
    teamid: id(6),
    name: 'Project X',
    teamtype: 3,
    membershiptype: 2,
    azureactivedirectoryobjectid: id(16),
  },
  {
    teamid: id(7),
    name: 'Partners',
    teamtype: 3,
    membershiptype: 3,
    azureactivedirectoryobjectid: id(17),
  },
  { teamid: id(10), name: 'Unlinked', teamtype: 2, membershiptype: 1 },
  {
    teamid: id(11),
    name: 'Finance owners',
    teamtype: 2,
    membershiptype: 2,
    azureactivedirectoryobjectid: id(18),
  },
  {
    teamid: id(12),
    name: 'Project Y',
    teamtype: 3,
    membershiptype: 1,
    azureactivedirectoryobjectid: id(19),
  },
];
const lib2 = {
  asx_libraryid: id(13),
  asx_name: 'Contracts',
  _asx_siteid_value: id(1),
  asx_approved: true,
  asx_policyapplied: true,
};
let discovery = false;
// Inspect results by operation key, checked before the shared mock answers.
const inspectByKey = {};
let discovered = [{ Id: id(8), Title: 'Archive' }];
// List IDs of active catalog libraries, as Dataverse stores them (any case).
let registeredLists = [];
// Finds a rendered node depth-first.
const find = (n, test) => {
  if (test(n)) return n;
  for (const c of n.children || []) {
    const hit = find(c, test);
    if (hit) return hit;
  }
  return null;
};
let operationStatus = 'Verified';
let inspectIssue = null,
  inspectFails = false;
let policy = { Status: 'Applied', RowVersion: '1', Policy: { Desired: [], Applied: [] } },
  refresh = 0;
const xrm = {
  Utility: { getGlobalContext: () => ({ getClientUrl: () => 'https://example.test' }) },
  Navigation: { openUrl: () => {} },
  WebApi: {
    retrieveMultipleRecords: async (table, options) => {
      if (table === 'team') teamQueries.push(options);
      if (table === 'asx_library') libraryQueries.push(options);
      if (table === 'asx_site') siteQueries.push(options);
      if (table === 'asx_library' && /asx_listid eq/.test(options))
        return {
          entities: registeredLists
            .filter((v) => options.toLowerCase().includes("asx_listid eq '" + v.toLowerCase()))
            .map((v) => ({ asx_listid: v })),
        };
      return {
        entities:
          table === 'asx_site'
            ? [site]
            : table === 'asx_library'
              ? [lib]
              : table === 'team'
                ? teams
                : table === 'sharepointsite'
                  ? [{ sharepointsiteid: id(2), name: 'Delivery' }]
                  : [],
      };
    },
    retrieveRecord: async (table, key) =>
      table === 'asx_library'
        ? key === id(13)
          ? lib2
          : lib
        : table === 'asx_site'
          ? site
          : { name: 'Operations' },
    online: {
      execute: async (req) => {
        const command = JSON.parse(req.Request);
        requests.push(command);
        apis.push(req.getMetadata().operationName);
        if (command.Command === 'Inspect' && inspectFails)
          throw new Error('Temporary status request failure');
        if (command.Command === 'RemoveLibrary' && removalRefusal)
          return { ok: false, text: async () => removalRefusal };
        if (command.Command.startsWith('Remove'))
          return {
            ok: true,
            json: async () => ({
              Result: JSON.stringify({
                Status: 'Removed',
                Notices: ['Nothing was deleted or changed in SharePoint: the library stays.'],
              }),
            }),
          };
        if (command.Command === 'RetrySetup' && retrySetupRefusal)
          return { ok: false, text: async () => retrySetupRefusal };
        let result;
        if (command.Command === 'GetPolicy') result = policy;
        else if (command.Command === 'RetryAccessRun')
          result = policy = { ...policy, RunStatus: 'Pending', RunNotice: null };
        else if (command.Command === 'CancelAccessRun')
          result = policy = {
            Status: 'NeedsReview',
            RowVersion: '41',
            Policy: { ...policy.Policy, OperationKey: null },
          };
        else if (command.Command === 'RetrySetup') result = { Status: 'Pending', Key: command.Key };
        else if (command.Command === 'CancelSetup')
          result = {
            Status: 'Cancelled',
            Key: command.Key,
            Notices: ['Cancelled. Nothing in SharePoint was deleted.'],
          };
        else if (command.Command === 'ApplyPolicy')
          // A run a flow is working on is not replaced: the change waits for it.
          result = policy =
            policy?.Policy?.OperationKey === 'background'
              ? {
                  ...policy,
                  RowVersion: '2',
                  Policy: { ...policy.Policy, Desired: command.Entries, ApplyPending: true },
                }
              : {
                  Status: 'Queued',
                  RowVersion: '2',
                  Policy: {
                    Desired: command.Entries,
                    Applied: [],
                    OperationKey: 'policywork:test',
                  },
                };
        else if (command.Command === 'Inspect' && inspectByKey[command.Key])
          result = inspectByKey[command.Key];
        else if (command.Command === 'Inspect')
          result = discovery
            ? {
                Status: 'Discovered',
                Key: 'catalogprobe:discovery',
                RowVersion: '7',
                Observation: {
                  SiteId: id(1),
                  Libraries: discovered,
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
  AsxdUi: { onTab: () => {}, deeplink: () => null, activeTab: () => 'access', feedback: () => {} },
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
  const teamQuery = decodeURIComponent(teamQueries[0]);
  assert.match(teamQuery, /teamtype eq 0 or teamtype eq 2 or teamtype eq 3/);
  assert.match(teamQuery, /isdefault eq false/);
  assert.doesNotMatch(teamQuery, /teamtype eq 1/, 'Access teams stay out of the picker');
  for (const picker of ['ad-team-choice', 'ad-initial-team']) {
    const options = nodes[picker].children.slice(1),
      byId = (n) => options.find((o) => o.value === id(n));
    assert.equal(options.length, teams.length, picker + ' lists every eligible team');
    assert.equal(byId(4).textContent, 'Operations');
    assert.equal(byId(4).disabled, false);
    assert.equal(byId(5).textContent, 'Finance (Entra group)');
    assert.equal(byId(5).disabled, false);
    assert.equal(byId(6).textContent, 'Project X (Microsoft 365 group, owners)');
    assert.equal(byId(6).disabled, false);
    assert.equal(byId(7).disabled, true, 'A guests-only team cannot be chosen');
    assert.match(
      byId(7).textContent,
      /Partners \(Microsoft 365 group\).*SharePoint has no sign-in claim for only the guests/,
    );
    assert.equal(byId(10).disabled, true, 'A team with no group object ID cannot be chosen');
    assert.match(byId(10).textContent, /no Microsoft Entra group object ID/);
    // Teams whose group reaches more people than the team are labelled and need consent.
    assert.equal(byId(11).disabled, false);
    assert.equal(
      byId(11).textContent,
      'Finance owners (Entra group, owners; all members get access)',
    );
    assert.equal(byId(12).disabled, false);
    assert.equal(
      byId(12).textContent,
      'Project Y (Microsoft 365 group, members; guests also get access)',
    );
  }
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
  await timers.shift()();
  // While a run is in progress the teams stay editable and Apply is sent: the change waits for
  // the run and is applied right after it.
  assert.equal(nodes['ad-teams'].children[0].children[1].children[0].disabled, false);
  assert.equal(nodes['ad-add-team'].disabled, false);
  await nodes['ad-apply'].onclick();
  assert.equal(
    requests.filter((r) => r.Command === 'ApplyPolicy').length,
    beforeConflict + 1,
    'An apply during a run is sent and waits for that run',
  );
  assert.match(nodes['ad-message'].textContent, /applied right after the access run in progress/);
  assert.match(nodes['ad-change-status'].textContent, /applied right after the access run/);
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
  // A library already in Documents is not offered again, also after a rename in SharePoint:
  // it is matched by list ID, not by title. A removed one (no active row) is still offered.
  discovered = [
    { Id: id(8), Title: 'Archive' },
    { Id: id(27), Title: 'AcceptCRenamed' },
    { Id: id(28), Title: 'Removed earlier' },
  ];
  registeredLists = [id(27).toUpperCase()];
  await nodes['ad-existing'].onclick();
  await timers.shift()();
  assert.equal(nodes['ad-existing-form'].hidden, false);
  assert.equal(nodes['ad-existing-more'].hidden, false);
  assert.deepEqual(
    nodes['ad-existing-choices'].children.map((c) => c.textContent),
    ['Add Archive', 'Add Removed earlier'],
  );
  assert.match(libraryQueries.at(-1), /statecode eq 0 and \(asx_listid eq '00000008-.*' or /);
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
  lib.asx_policyapplied = true;
  lib.asx_approved = false;
  await window.AsxdSites.selectLibrary(id(3));
  assert.equal(nodes['ad-library-status'].textContent, 'Needs attention');
  lib.asx_approved = true;
  // Consent: a team whose group reaches more people is applied only after the admin confirms
  // the warning shown in the page.
  policy = { Status: 'Applied', RowVersion: '20', Policy: { Desired: [], Applied: [] } };
  await window.AsxdSites.selectLibrary(id(13));
  assert.equal(nodes['ad-library-title'].textContent, 'Contracts');
  assert.equal(nodes['ad-apply-warning'].hidden, true);
  nodes['ad-add-team'].onclick();
  nodes['ad-team-choice'].value = id(11);
  nodes['ad-team-access'].value = 'Read';
  nodes['ad-stage-team'].onclick();
  const applies = () => requests.filter((r) => r.Command === 'ApplyPolicy').length;
  const beforeConsent = applies();
  await nodes['ad-apply'].onclick();
  assert.equal(applies(), beforeConsent, 'Nothing is applied before the admin confirms');
  assert.equal(nodes['ad-apply-warning'].hidden, false);
  assert.match(
    nodes['ad-apply-warning'].textContent,
    /All members of the group will have access to this library, not only its owners\./,
  );
  assert.match(nodes['ad-apply'].textContent, /Confirm and apply/);
  await nodes['ad-apply'].onclick();
  const consented = requests.filter((r) => r.Command === 'ApplyPolicy').at(-1);
  assert.equal(applies(), beforeConsent + 1);
  assert.equal(consented.AcknowledgeBroaderAccess, true);
  assert.equal(consented.Entries[0].TeamId, id(11));
  assert.equal(nodes['ad-apply-warning'].hidden, true);
  assert.match(nodes['ad-apply'].textContent, /Apply access changes/);
  // The same for a new library's initial team.
  nodes['ad-create'].onclick();
  nodes['ad-library-name'].value = 'Partners';
  nodes['ad-initial-team'].value = id(12);
  nodes['ad-initial-access'].value = 'Read';
  const creates = () => requests.filter((r) => r.Command === 'CreateLibrary').length;
  const beforeCreate = creates();
  await nodes['ad-provision'].onclick();
  assert.equal(creates(), beforeCreate, 'No library is created before the admin confirms');
  assert.equal(nodes['ad-provision-warning'].hidden, false);
  assert.equal(
    nodes['ad-provision-warning'].textContent,
    "The group's guests will also have access to this library. Select Confirm and create to continue.",
  );
  await nodes['ad-provision'].onclick();
  const createdWithConsent = requests.filter((r) => r.Command === 'CreateLibrary').at(-1);
  assert.equal(creates(), beforeCreate + 1);
  assert.equal(createdWithConsent.AcknowledgeBroaderAccess, true);
  assert.equal(nodes['ad-provision-warning'].hidden, true);
  // An owner team needs no consent.
  nodes['ad-create'].onclick();
  nodes['ad-library-name'].value = 'Ledger';
  nodes['ad-initial-team'].value = id(4);
  await nodes['ad-provision'].onclick();
  assert.equal(creates(), beforeCreate + 2);
  assert.equal(
    requests.filter((r) => r.Command === 'CreateLibrary').at(-1).AcknowledgeBroaderAccess,
    undefined,
  );
  // A library that inherits the site's permissions is added only after the admin confirms, in
  // the page, that Documents stops the inheritance.
  const warning =
    'This library inherits permissions from the site. When you approve it, Documents stops the inheritance, keeps a copy of the current site permissions, and then manages team access on it.';
  discovered = [{ Id: id(14), Title: 'Shared', HasUniqueRoleAssignments: false }];
  await nodes['ad-existing'].onclick();
  await timers.shift()();
  const adds = () => requests.filter((r) => r.Command === 'AddLibrary');
  const beforeAdd = adds().length;
  await nodes['ad-existing-choices'].children[0].onclick();
  assert.equal(adds().length, beforeAdd, 'Nothing is added before the admin confirms');
  assert.equal(nodes['ad-existing-warning'].hidden, false);
  assert.equal(
    nodes['ad-existing-warning'].textContent,
    warning + ' Select Confirm and add to continue.',
  );
  assert.equal(nodes['ad-existing-choices'].children[0].textContent, 'Confirm and add Shared');
  await nodes['ad-existing-choices'].children[0].onclick();
  assert.equal(adds().length, beforeAdd + 1);
  assert.equal(adds().at(-1).ListId, id(14));
  assert.equal(adds().at(-1).BreakInheritance, true);
  // An approval refused for want of consent offers the consent on its card.
  inspectByKey['catalogprobe:test'] = {
    Status: 'Blocked',
    Key: 'catalogprobe:test',
    Issue: warning + ' Approve it with that acknowledgement to continue.',
    Observation: { Inherits: true, SiteId: id(1), ListId: id(14), WebUrl: undefined },
  };
  await timers.shift()();
  assert.match(nodes['ad-provision-progress'].textContent, /keeps a copy of the current site/);
  const stop = find(
    nodes['ad-provision-progress'],
    (n) => n.textContent === 'Stop inheritance and add',
  );
  assert(stop, 'The blocked card offers to stop the inheritance');
  await stop.onclick();
  assert.equal(adds().length, beforeAdd + 2);
  assert.equal(adds().at(-1).BreakInheritance, true);
  assert.equal(adds().at(-1).ListId, id(14));
  delete inspectByKey['catalogprobe:test'];
  // A library reset to inherit stops its access run (GetPolicy answers as the server does: the
  // run is Blocked with the inheritance notice). It needs attention like any stopped run, with
  // Retry and Cancel, and Apply access, which asks before stopping the inheritance again.
  const inheritsAgain =
    'This library inherits permissions again. Use Apply access to let Documents stop the inheritance again.';
  policy = {
    Status: 'Queued',
    RowVersion: '30',
    Policy: {
      Desired: [{ TeamId: id(4), Access: 'Read' }],
      Applied: [{ TeamId: id(4), Access: 'Read' }],
      OperationKey: 'policywork:stopped',
      Inherits: true,
    },
    RunStatus: 'Blocked',
    RunNotice: inheritsAgain,
  };
  // A library not opened before, so its policy is read fresh.
  lib2.asx_libraryid = id(19);
  lib2.asx_name = 'Reset';
  lib2.asx_policyapplied = false;
  await window.AsxdSites.selectLibrary(id(13));
  assert.equal(nodes['ad-change-status'].textContent, 'Needs attention: ' + inheritsAgain);
  assert.equal(nodes['ad-change-status'].className, 'ad-issue');
  assert.equal(nodes['ad-library-access'].textContent, 'Needs attention');
  assert.equal(nodes['ad-run-actions'].hidden, false, 'The stopped run offers Retry and Cancel');
  assert.equal(nodes['ad-run-retry'].disabled, false);
  assert.equal(nodes['ad-run-cancel'].disabled, false);
  assert.equal(nodes['ad-apply'].disabled, false, 'Apply access is offered for the reset library');
  lib2.asx_policyapplied = true;
  const beforeReapply = applies();
  await nodes['ad-apply'].onclick();
  assert.equal(applies(), beforeReapply, 'Nothing is applied before the admin confirms');
  assert.match(
    nodes['ad-apply-warning'].textContent,
    /When you apply access, Documents stops the inheritance, keeps a copy of the current site permissions/,
  );
  await nodes['ad-apply'].onclick();
  const reapplied = requests.filter((r) => r.Command === 'ApplyPolicy').at(-1);
  assert.equal(applies(), beforeReapply + 1);
  assert.equal(reapplied.BreakInheritance, true);
  assert.equal(reapplied.AcknowledgeBroaderAccess, undefined);
  // Re-point asks in the page, then shows what changed as text.
  const repoints = () => requests.filter((r) => r.Command === 'RepointLibrary');
  assert.equal(nodes['ad-confirm'].hidden, true);
  nodes['ad-repoint-library'].onclick();
  assert.equal(repoints().length, 0, 'Nothing is re-pointed before the admin confirms');
  assert.equal(nodes['ad-confirm'].hidden, false);
  assert.match(
    nodes['ad-confirm-text'].textContent,
    /Re-point Reset: .*Nothing in SharePoint changes/,
  );
  nodes['ad-confirm-cancel'].onclick();
  assert.equal(nodes['ad-confirm'].hidden, true);
  nodes['ad-repoint-library'].onclick();
  await nodes['ad-confirm-go'].onclick();
  assert.equal(repoints().length, 1);
  assert.equal(repoints()[0].CatalogId, id(19));
  assert.equal(nodes['ad-confirm'].hidden, true);
  inspectByKey['catalogprobe:test'] = {
    Status: 'Approved',
    Key: 'catalogprobe:test',
    Observation: {
      Changes: [
        'https://example.sharepoint.com/sites/proto/General → https://example.sharepoint.com/sites/proto/<b>Shared</b>',
      ],
    },
  };
  await timers.shift()();
  assert.match(nodes['ad-message'].textContent, /Reset re-pointed/);
  assert.equal(nodes['ad-changes'].hidden, false);
  assert.equal(
    nodes['ad-changes'].children[0].textContent,
    'https://example.sharepoint.com/sites/proto/General → https://example.sharepoint.com/sites/proto/<b>Shared</b>',
    'Changes are text, never markup',
  );
  // A library that no longer exists is reported on its card.
  nodes['ad-repoint-site'].onclick();
  assert.match(nodes['ad-confirm-text'].textContent, /Re-point Delivery: .*SharePoint site record/);
  await nodes['ad-confirm-go'].onclick();
  assert.equal(requests.at(-1).Command, 'RepointSite');
  assert.equal(requests.at(-1).CatalogId, id(1));
  inspectByKey['catalogprobe:test'] = {
    Status: 'Blocked',
    Key: 'catalogprobe:test',
    Issue: 'This library no longer exists on the site. Remove it, or register the new library.',
  };
  await timers.shift()();
  assert.match(nodes['ad-provision-progress'].textContent, /no longer exists on the site/);
  assert(
    !find(nodes['ad-provision-progress'], (n) => n.textContent === 'Retry after repair'),
    'A blocked re-point is not retried in place',
  );
  const again = find(nodes['ad-provision-progress'], (n) => n.textContent === 'Re-point again');
  assert(again, 'A blocked re-point offers to re-point again');
  const beforeAgain = requests.filter((r) => r.Command === 'RepointSite').length;
  delete inspectByKey['catalogprobe:test'];
  await again.onclick();
  const repointedAgain = requests.filter((r) => r.Command === 'RepointSite');
  assert.equal(repointedAgain.length, beforeAgain + 1);
  assert.equal(repointedAgain.at(-1).CatalogId, id(1));
  assert(!requests.some((r) => r.Command === 'Retry' && r.Key === 'catalogprobe:test'));
  // Remove asks in the page; a refusal lists the templates that use the library.
  const removes = () => requests.filter((r) => r.Command === 'RemoveLibrary');
  nodes['ad-remove-library'].onclick();
  assert.equal(removes().length, 0, 'Nothing is removed before the admin confirms');
  assert.match(nodes['ad-confirm-text'].textContent, /Nothing in SharePoint is deleted or changed/);
  assert.equal(nodes['ad-confirm-go'].textContent, 'Remove library');
  removalRefusal = "Used by template 'Accounts' (published). Change the template first.";
  await nodes['ad-confirm-go'].onclick();
  assert.equal(removes().length, 1);
  assert.equal(removes()[0].CatalogId, id(19));
  assert.match(nodes['ad-message'].textContent, /Used by template 'Accounts' \(published\)/);
  assert.equal(nodes['ad-message'].className, 'ad-issue');
  removalRefusal = null;
  nodes['ad-remove-library'].onclick();
  await nodes['ad-confirm-go'].onclick();
  assert.equal(removes().length, 2);
  assert.match(nodes['ad-message'].textContent, /Reset was removed from Documents/);
  assert.match(nodes['ad-changes'].textContent, /Nothing was deleted or changed in SharePoint/);
  assert.match(libraryQueries.at(-1), /statecode eq 0/, 'Removed libraries are hidden');
  nodes['ad-remove-site'].onclick();
  assert.equal(nodes['ad-confirm-go'].textContent, 'Remove site');
  await nodes['ad-confirm-go'].onclick();
  assert.equal(requests.at(-1).Command, 'RemoveSite');
  assert.equal(requests.at(-1).CatalogId, id(1));
  assert.match(siteQueries.at(-1), /statecode eq 0/, 'Removed sites are hidden');
  {
    // An access run that stopped or waits shows its notice on the library, with Retry and
    // Cancel in place; Apply access replaces a run that stopped.
    const stuck = {
      Status: 'Queued',
      RowVersion: '40',
      Policy: {
        Desired: [{ TeamId: id(4), Access: 'Read' }],
        Applied: [],
        OperationKey: 'policywork:stuck',
      },
      RunStatus: 'Blocked',
      RunNotice: 'SharePoint refused the write (HTTP 403).',
    };
    policy = stuck;
    lib2.asx_libraryid = id(20);
    lib2.asx_name = 'Stuck';
    await window.AsxdSites.selectLibrary(id(13));
    assert.equal(
      nodes['ad-change-status'].textContent,
      'Needs attention: SharePoint refused the write (HTTP 403).',
    );
    assert.equal(nodes['ad-change-status'].className, 'ad-issue');
    assert.equal(nodes['ad-library-access'].textContent, 'Needs attention');
    assert.equal(nodes['ad-run-actions'].hidden, false);
    assert.equal(nodes['ad-add-team'].disabled, false, 'A stopped run does not lock the teams');
    await nodes['ad-run-retry'].onclick();
    assert.deepEqual(requests.at(-1), {
      Command: 'RetryAccessRun',
      LibraryId: id(20),
      OperationKey: 'policywork:stuck',
    });
    assert.equal(apis.at(-1), 'asx_SecurityAdmin');
    assert.equal(nodes['ad-change-status'].textContent, 'Applying access and syncing members…');
    assert.equal(nodes['ad-library-access'].textContent, 'Applying access…');
    assert.equal(nodes['ad-run-actions'].hidden, true);
    assert.equal(
      nodes['ad-add-team'].disabled,
      false,
      'A running access run does not lock the teams',
    );
    policy = {
      ...stuck,
      RunStatus: 'RetryWait',
      RunNotice: 'Waiting to retry after a temporary error (HTTP 503); attempt 4.',
      RunNextAttemptUtc: '/Date(1791225000000)/',
    };
    await timers.shift()();
    assert.equal(
      nodes['ad-change-status'].textContent,
      'Needs attention: Waiting to retry after a temporary error (HTTP 503); attempt 4. Next check: 2026-10-05 18:30 UTC.',
    );
    assert.equal(nodes['ad-library-access'].textContent, 'Needs attention');
    const cancels = () => requests.filter((r) => r.Command === 'CancelAccessRun');
    nodes['ad-run-cancel'].onclick();
    assert.equal(cancels().length, 0, 'Nothing is cancelled before the admin confirms');
    assert.equal(nodes['ad-confirm'].hidden, false);
    assert.match(
      nodes['ad-confirm-text'].textContent,
      /^Cancel the access run for Stuck\? .*nothing is undone or deleted/,
    );
    assert.equal(nodes['ad-confirm-go'].textContent, 'Cancel access run');
    assert.equal(nodes['ad-confirm-cancel'].textContent, 'Keep it');
    await nodes['ad-confirm-go'].onclick();
    assert.deepEqual(cancels(), [
      { Command: 'CancelAccessRun', LibraryId: id(20), OperationKey: 'policywork:stuck' },
    ]);
    assert.match(nodes['ad-message'].textContent, /Nothing in SharePoint was undone or deleted/);
    assert.equal(
      nodes['ad-change-status'].textContent,
      'The access run was cancelled. Apply access to run it again.',
    );
    assert.equal(nodes['ad-library-access'].textContent, 'Needs attention');
    assert.equal(nodes['ad-apply'].disabled, false, 'Apply access starts a new run');
    // Apply access replaces a stopped run with the admin's newer access.
    policy = stuck;
    const team = nodes['ad-teams'].children[0].children[1].children[0];
    team.value = 'Contribute';
    team.onchange();
    const before = requests.filter((r) => r.Command === 'ApplyPolicy').length;
    await nodes['ad-apply'].onclick();
    const replaced = requests.filter((r) => r.Command === 'ApplyPolicy');
    assert.equal(replaced.length, before + 1, 'A stopped run does not refuse Apply');
    assert.equal(replaced.at(-1).Entries[0].Access, 'Contribute');
    // A change waiting behind a stopped run says the run must be retried or cancelled first.
    policy = { ...stuck, Policy: { ...stuck.Policy, ApplyPending: true } };
    lib2.asx_libraryid = id(21); // a library not loaded yet, so its access is read
    await window.AsxdSites.selectLibrary(id(13));
    assert.equal(
      nodes['ad-change-status'].textContent,
      'Needs attention: SharePoint refused the write (HTTP 403). The access run stopped: Retry or Cancel it, then your change applies.',
    );
    assert.equal(nodes['ad-change-status'].className, 'ad-issue');
    assert.equal(nodes['ad-library-access'].textContent, 'Needs attention');
    assert.equal(nodes['ad-run-actions'].hidden, false);
    // Applied access with team membership left unsynced is not shown as confirmed.
    policy = {
      Status: 'Applied',
      RowVersion: '42',
      Policy: {
        Desired: stuck.Policy.Desired,
        Applied: stuck.Policy.Desired,
        MembershipIncomplete: true,
        Notices: ["Team 'Big' has more people than one access run can store."],
      },
    };
    lib2.asx_libraryid = id(22);
    await window.AsxdSites.selectLibrary(id(13));
    assert.match(
      nodes['ad-change-status'].textContent,
      /^Needs attention: access is applied, but team membership was not synced. People removed from a team keep access, and people added get none/,
    );
    assert.equal(nodes['ad-change-status'].className, 'ad-issue');
    assert.equal(nodes['ad-library-access'].textContent, 'Needs attention');
    assert.doesNotMatch(nodes['ad-change-status'].textContent, /confirmed/);
    policy = { ...policy, Policy: { ...policy.Policy, MembershipIncomplete: false, Notices: [] } };
    lib2.asx_libraryid = id(23);
    await window.AsxdSites.selectLibrary(id(13));
    assert.equal(nodes['ad-change-status'].textContent, 'Access and team membership confirmed.');
    // The header follows the policy and its run, never the flag the library list was read
    // with: here the list still says access is pending while the policy is Applied.
    lib2.asx_policyapplied = false;
    lib2.asx_libraryid = id(24);
    await window.AsxdSites.selectLibrary(id(13));
    assert.equal(nodes['ad-change-status'].textContent, 'Access and team membership confirmed.');
    assert.equal(nodes['ad-library-access'].textContent, 'Access applied');
    // A run that is queued or running.
    policy = {
      Status: 'Queued',
      RowVersion: '43',
      Policy: { Desired: stuck.Policy.Desired, Applied: [], OperationKey: 'policywork:new' },
      RunStatus: 'Pending',
    };
    lib2.asx_policyapplied = true;
    lib2.asx_libraryid = id(25);
    await window.AsxdSites.selectLibrary(id(13));
    assert.equal(nodes['ad-library-access'].textContent, 'Applying access…');
    // It applies: the next poll shows it, with no reload.
    policy = {
      Status: 'Applied',
      RowVersion: '44',
      Policy: { Desired: stuck.Policy.Desired, Applied: stuck.Policy.Desired },
    };
    await timers.shift()();
    assert.equal(nodes['ad-change-status'].textContent, 'Access and team membership confirmed.');
    assert.equal(nodes['ad-library-access'].textContent, 'Access applied');
    // Access never applied.
    policy = { Status: 'Missing', RowVersion: '', Policy: null };
    lib2.asx_policyapplied = false;
    lib2.asx_libraryid = id(26);
    await window.AsxdSites.selectLibrary(id(13));
    assert.equal(nodes['ad-library-access'].textContent, 'Access setup pending');
    lib2.asx_policyapplied = true;
  }
  {
    // A library setup that needs attention offers Retry and Cancel through the catalog API, so
    // a Documents Security Administrator needs no Operator role.
    nodes['ad-create'].onclick();
    nodes['ad-library-name'].value = 'Stalled';
    nodes['ad-initial-team'].value = '';
    await nodes['ad-provision'].onclick();
    inspectByKey['librarycreate:test'] = {
      Status: 'RecoveryRequired',
      Key: 'librarycreate:test',
      Issue: 'The library create may have reached SharePoint and its answer was lost.',
    };
    await timers.shift()();
    const area = nodes['ad-provision-progress'],
      named = (text) => find(area, (n) => n.textContent === text);
    assert.match(area.textContent, /Stalled.*Needs attention/);
    assert.match(area.textContent, /answer was lost/);
    retrySetupRefusal =
      'The library create may have reached SharePoint and its answer was lost. Recover it with the original create response from the flow run, or Cancel the setup.';
    await named('Retry').onclick();
    assert.deepEqual(requests.at(-1), { Command: 'RetrySetup', Key: 'librarycreate:test' });
    assert.equal(apis.at(-1), 'asx_CatalogAdmin');
    assert.match(nodes['ad-message'].textContent, /original create response/);
    assert.equal(nodes['ad-message'].className, 'ad-issue');
    retrySetupRefusal = null;
    await named('Retry').onclick();
    assert.equal(apis.at(-1), 'asx_CatalogAdmin');
    assert.match(nodes['ad-message'].textContent, /queued to run again/);
    await timers.shift()();
    named('Cancel setup').onclick();
    assert.equal(
      requests.filter((r) => r.Command === 'CancelSetup').length,
      0,
      'Nothing is cancelled before the admin confirms',
    );
    assert.match(
      nodes['ad-confirm-text'].textContent,
      /^Cancel the setup of Stalled\? Nothing in SharePoint is deleted\./,
    );
    assert.equal(nodes['ad-confirm-go'].textContent, 'Cancel setup');
    await nodes['ad-confirm-go'].onclick();
    assert.deepEqual(requests.at(-1), { Command: 'CancelSetup', Key: 'librarycreate:test' });
    assert.equal(apis.at(-1), 'asx_CatalogAdmin');
    assert.doesNotMatch(area.textContent, /Stalled/, 'A cancelled setup leaves the list');
    assert.match(nodes['ad-changes'].textContent, /Nothing in SharePoint was deleted/);
    assert(!apis.includes('asx_ManageWork'), 'Sites never needs the Operator role');
    delete inspectByKey['librarycreate:test'];
  }
  {
    // A setup whose first access run was cancelled ends Ready with that notice instead of
    // "Setup completed.".
    nodes['ad-create'].onclick();
    nodes['ad-library-name'].value = 'Unsynced';
    nodes['ad-initial-team'].value = '';
    await nodes['ad-provision'].onclick();
    inspectByKey['librarycreate:test'] = {
      Status: 'Ready',
      Key: 'librarycreate:test',
      Issue: 'Library created. Its first access run was cancelled; apply access on the library.',
    };
    await timers.shift()();
    // Each card ends with Dismiss; the Unsynced card is the text from its name to that.
    const card = /Unsynced(?:(?!Dismiss)[^])*Dismiss/.exec(
      nodes['ad-provision-progress'].textContent,
    );
    assert(card, 'The finished setup stays listed');
    assert.match(card[0], /first access run was cancelled; apply access/);
    assert.doesNotMatch(card[0], /Setup completed/);
    assert.match(
      nodes['ad-message'].textContent,
      /Unsynced: Library created. Its first access run/,
    );
    delete inspectByKey['librarycreate:test'];
  }
  {
    // After a reload, a re-point still running or blocked is found again with the other setup
    // activity, and its command and ID come from its probe, so "Re-point again" still works.
    // Until a card is read it shows "Loading…", never the generic message or action. Cards of
    // a library that was removed or deleted are not shown at all.
    const fresh = {},
      later = [],
      sent = [],
      activity = [],
      lists = [],
      checks = [],
      // Active catalog rows; a removed or deleted one is not here.
      active = new Set([id(1), id(3)]),
      moved = { ...site, asx_url: 'https://example.sharepoint.com/sites/moved' };
    let release,
      inspecting = false;
    const gate = new Promise((resolve) => (release = resolve));
    for (const m of html.matchAll(/<([a-z]+)[^>]*\bid="([^"]+)"[^>]*>/g))
      fresh[m[2]] = new Node(m[1]);
    const inspected = {
      'catalogprobe:repoint:removed': {
        Status: 'Blocked',
        Issue: 'This library no longer exists on the site. Remove it, or register the new library.',
        Observation: { Repoint: true, CatalogId: id(30), SiteId: id(1), ListId: id(31) },
      },
      'librarycreate:deleted': {
        Status: 'RecoveryRequired',
        Issue: 'The library create may have reached SharePoint and its answer was lost.',
        CatalogId: id(32),
      },
    };
    const reloaded = {
      Xrm: {
        Utility: xrm.Utility,
        Navigation: xrm.Navigation,
        WebApi: {
          retrieveMultipleRecords: async (table, options) => {
            if (table === 'asx_operation') {
              activity.push(options);
              return {
                entities: [
                  {
                    asx_workkey: 'catalogprobe:repoint:old',
                    asx_workkind: 'Repoint',
                    asx_displayname: 'General',
                    asx_siteurl: 'https://example.sharepoint.com/sites/moved',
                    asx_status: 'Blocked',
                  },
                  {
                    asx_workkey: 'catalogprobe:repoint:removed',
                    asx_workkind: 'Repoint',
                    asx_displayname: 'AcceptC Chrome 1',
                    asx_siteurl: 'https://example.sharepoint.com/sites/moved',
                    asx_status: 'Blocked',
                  },
                  {
                    asx_workkey: 'librarycreate:deleted',
                    asx_workkind: 'LibrarySetup',
                    asx_displayname: 'Deleted later',
                    asx_siteurl: 'https://example.sharepoint.com/sites/moved',
                    asx_status: 'RecoveryRequired',
                  },
                ],
              };
            }
            const by = /(asx_libraryid|asx_siteid) eq ([0-9a-f-]{36})/.exec(options);
            if (by) {
              checks.push(options);
              return {
                entities: active.has(by[2]) && /statecode eq 0/.test(options) ? [{}] : [],
              };
            }
            if (table === 'asx_library') lists.push(options);
            // The list reads answer as they did before the Remove, like a read that does not
            // reflect it yet.
            return {
              entities: table === 'asx_site' ? [moved] : table === 'asx_library' ? [lib] : [],
            };
          },
          retrieveRecord: xrm.WebApi.retrieveRecord,
          online: {
            execute: async (req) => {
              const command = JSON.parse(req.Request);
              sent.push(command);
              if (command.Command === 'Inspect') {
                inspecting = true;
                await gate;
              }
              const result =
                command.Command === 'Inspect'
                  ? {
                      Key: command.Key,
                      ...(inspected[command.Key] || {
                        Status: 'Blocked',
                        Issue:
                          'This library no longer exists on the site. Remove it, or register the new library.',
                        Observation: {
                          Repoint: true,
                          CatalogId: id(3),
                          SiteId: id(1),
                          ListId: id(21),
                        },
                      }),
                    }
                  : command.Command === 'GetPolicy'
                    ? { Status: 'Applied', RowVersion: '1', Policy: { Desired: [], Applied: [] } }
                    : command.Command.startsWith('Remove')
                      ? { Status: 'Removed', CatalogId: command.CatalogId, Notices: [] }
                      : { Status: 'Pending', Key: 'catalogprobe:repoint:again' };
              if (command.Command.startsWith('Remove')) active.delete(command.CatalogId);
              return { ok: true, json: async () => ({ Result: JSON.stringify(result) }) };
            },
          },
        },
      },
      AsxdAdmin: { refreshCatalog: async () => {} },
      AsxdUi: {
        onTab: () => {},
        deeplink: () => null,
        activeTab: () => 'access',
        feedback: () => {},
      },
    };
    vm.runInNewContext(fs.readFileSync(path.join(base, 'sites-access.js'), 'utf8'), {
      window: reloaded,
      document: { getElementById: (key) => fresh[key], createElement: (t) => new Node(t) },
      URL,
      crypto: { randomUUID: () => id(22) },
      setTimeout: (fn) => later.push(fn),
    });
    const card = fresh['ad-provision-progress'],
      opening = reloaded.AsxdSites.open();
    for (let i = 0; i < 50 && !inspecting; i++) await new Promise((r) => setImmediate(r));
    // Before the first read: no generic message or action, and nothing of the removed ones.
    assert.doesNotMatch(card.textContent, /Setup needs review|Retry after repair/);
    assert.doesNotMatch(card.textContent, /Needs attention/);
    assert.doesNotMatch(card.textContent, /AcceptC Chrome 1|Deleted later/);
    assert.match(card.textContent, /Loading…/);
    release();
    await opening;
    assert.match(activity[0], /asx_workkind eq 'Repoint'/);
    // Read at once after the load, not at the first poll.
    assert.match(card.textContent, /General.*no longer exists on the site/);
    assert.doesNotMatch(card.textContent, /Loading…/);
    assert(checks.some((c) => c.includes('asx_libraryid eq ' + id(30))));
    assert(checks.some((c) => c.includes('asx_libraryid eq ' + id(32))));
    await later.shift()();
    assert.doesNotMatch(
      card.textContent,
      /AcceptC Chrome 1|Deleted later/,
      'Cards of a removed or deleted library are not shown',
    );
    assert.equal(card.children.length, 1);
    const again = find(card, (n) => n.textContent === 'Re-point again');
    assert(again, 'A blocked re-point found after a reload offers to re-point again');
    assert(!find(card, (n) => n.textContent === 'Retry after repair'));
    await again.onclick();
    assert.deepEqual(
      sent.filter((c) => c.Command === 'RepointLibrary').map((c) => c.CatalogId),
      [id(3)],
    );
    await later.shift()();
    assert.match(card.textContent, /General.*no longer exists on the site/);
    // Remove: the library leaves the list and its re-point card goes, with no reload.
    const listed = () => fresh['ad-libraries'].children.map((c) => c.textContent);
    assert.deepEqual(listed(), ['General']);
    const before = lists.length;
    fresh['ad-remove-library'].onclick();
    await fresh['ad-confirm-go'].onclick();
    assert.deepEqual(sent.at(-1), { Command: 'RemoveLibrary', CatalogId: id(3) });
    assert.match(fresh['ad-message'].textContent, /General was removed from Documents/);
    assert(lists.length > before, 'The library list is read again after Remove');
    assert.deepEqual(listed(), [], 'The removed library is no longer listed');
    assert.equal(fresh['ad-library-detail'].hidden, true);
    assert.doesNotMatch(card.textContent, /General/, 'Its re-point card goes with it');
    assert.equal(card.hidden, true);
    // A link to it does not bring it back.
    await reloaded.AsxdSites.selectLibrary(id(3));
    assert.match(fresh['ad-message'].textContent, /This library was removed from Documents/);
    assert.deepEqual(listed(), []);
    // Remove site: the site leaves the list the same way.
    fresh['ad-remove-site'].onclick();
    await fresh['ad-confirm-go'].onclick();
    assert.deepEqual(sent.at(-1), { Command: 'RemoveSite', CatalogId: id(1) });
    assert.deepEqual(
      fresh['ad-sites'].children.map((c) => c.textContent),
      [],
      'The removed site is no longer listed',
    );
    assert.equal(fresh['ad-site-title'].textContent, 'Select or add a site');
  }
  {
    // A team deleted in Dataverse while it has access to the library: Dataverse no longer has
    // the team, so reading it by ID fails. The library shows it as a deleted team by its last
    // known name, or its ID when none is known, with what happens next; status polling goes on;
    // and Apply access is enabled, sends the change and succeeds.
    const fresh = {},
      later = [],
      sent = [],
      teamReads = [],
      gone = id(40),
      unnamed = id(41),
      deletedNotice =
        'This team was deleted in Dataverse. Documents removes its access the next time access is applied.';
    for (const m of html.matchAll(/<([a-z]+)[^>]*\bid="([^"]+)"[^>]*>/g))
      fresh[m[2]] = new Node(m[1]);
    const entries = [
      { TeamId: gone, Access: 'Read' },
      { TeamId: unnamed, Access: 'Contribute' },
      { TeamId: id(4), Access: 'Read' },
    ];
    let current = {
      Status: 'Queued',
      RowVersion: '50',
      // The scheduled refresh already removed the unnamed team's access.
      Policy: {
        Desired: entries,
        Applied: [entries[0], { TeamId: unnamed, Access: 'None' }, entries[2]],
        OperationKey: 'policywork:refresh',
      },
      Teams: [
        { TeamId: gone, Name: 'AcceptC Team 1', Deleted: true },
        { TeamId: unnamed, Name: null, Deleted: true },
        { TeamId: id(4), Name: 'Operations', Deleted: false },
      ],
    };
    const deletedWindow = {
      Xrm: {
        Utility: xrm.Utility,
        Navigation: xrm.Navigation,
        WebApi: {
          retrieveMultipleRecords: async (table) => ({
            entities:
              table === 'asx_site'
                ? [site]
                : table === 'asx_library'
                  ? [lib]
                  : table === 'team'
                    ? [teams[0]]
                    : [],
          }),
          retrieveRecord: async (table, key) => {
            if (table === 'team') {
              teamReads.push(key);
              if (key === gone || key === unnamed)
                throw new Error('The requested record was not found.');
            }
            return xrm.WebApi.retrieveRecord(table, key);
          },
          online: {
            execute: async (req) => {
              const command = JSON.parse(req.Request);
              sent.push(command);
              let result = { Status: 'Pending', Key: 'catalogprobe:none' };
              if (command.Command === 'GetPolicy') result = current;
              else if (command.Command === 'ApplyPolicy')
                // The server leaves deleted teams out and queues the removal of their access.
                result = current = {
                  Status: 'Queued',
                  RowVersion: '51',
                  Policy: {
                    Desired: command.Entries.filter((e) => e.TeamId === id(4)),
                    Applied: entries,
                    OperationKey: 'policywork:apply',
                  },
                  Teams: current.Teams,
                };
              return { ok: true, json: async () => ({ Result: JSON.stringify(result) }) };
            },
          },
        },
      },
      AsxdAdmin: { refreshCatalog: async () => {} },
      AsxdUi: {
        onTab: () => {},
        deeplink: () => null,
        activeTab: () => 'access',
        feedback: () => {},
      },
    };
    vm.runInNewContext(fs.readFileSync(path.join(base, 'sites-access.js'), 'utf8'), {
      window: deletedWindow,
      document: { getElementById: (key) => fresh[key], createElement: (t) => new Node(t) },
      URL,
      crypto: { randomUUID: () => id(42) },
      setTimeout: (fn) => later.push(fn),
    });
    fresh.access.hidden = false;
    await deletedWindow.AsxdSites.open();
    assert.doesNotMatch(fresh['ad-message'].textContent, /not found/);
    const rows = () => fresh['ad-teams'].children.map((r) => r.children[0].textContent);
    assert.equal(rows()[0], 'Deleted team: AcceptC Team 1' + deletedNotice);
    assert.equal(
      rows()[1],
      'Deleted team: ' +
        unnamed +
        'Its access was removed. Apply access to clear it from this list.',
    );
    assert.equal(rows()[2], 'Operations');
    assert.deepEqual(teamReads, [], 'No team is read by ID once the policy names it');
    // Its access cannot be chosen: the next Apply removes it.
    assert.equal(fresh['ad-teams'].children[0].children[1].children[0].disabled, true);
    // Polling the queued run keeps working.
    for (let i = 0; i < 4; i++) await later.shift()();
    assert.equal(fresh['ad-poll-status'].hidden, true);
    assert.doesNotMatch(fresh['ad-poll-status'].textContent, /Status updates are unavailable/);
    // Apply is enabled with no other change, and succeeds.
    assert.equal(fresh['ad-apply'].disabled, false);
    await fresh['ad-apply'].onclick();
    const apply = sent.find((c) => c.Command === 'ApplyPolicy');
    assert(apply, 'Apply access sends the change');
    assert.doesNotMatch(fresh['ad-message'].textContent, /not found|changed since you opened/);
    assert.match(fresh['ad-message'].textContent, /Access submitted/);
    assert.deepEqual(rows(), ['Operations'], 'The deleted teams leave the list once applied');
    // The next Apply sees the list it shows.
    fresh['ad-teams'].children[0].children[1].children[0].value = 'Contribute';
    fresh['ad-teams'].children[0].children[1].children[0].onchange();
    await fresh['ad-apply'].onclick();
    assert.match(fresh['ad-message'].textContent, /Access submitted/);
    assert.equal(sent.filter((c) => c.Command === 'ApplyPolicy').length, 2);
  }
  console.log(
    'PASS Sites & access handlers: staging versus apply, automatic onboarding request, initial library teams, completion polling, stuck access runs (also the inheritance stop) and library setups with Retry and Cancel, the access label from the policy and its run, discovery matched by list ID, re-point found again after a reload with no generic flash, removed destinations hidden at once and after a reload, deleted Dataverse teams shown by name with polling and Apply still working, and author deep link. Mocked APIs; connected acceptance pending.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
