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
  // A library reset to inherit stops its access run; Apply access asks before stopping it again.
  policy = {
    Status: 'Queued',
    RowVersion: '30',
    Policy: {
      Desired: [{ TeamId: id(4), Access: 'Read' }],
      Applied: [{ TeamId: id(4), Access: 'Read' }],
      OperationKey: 'policywork:stopped',
      Inherits: true,
    },
  };
  // A library not opened before, so its policy is read fresh.
  lib2.asx_libraryid = id(19);
  lib2.asx_name = 'Reset';
  await window.AsxdSites.selectLibrary(id(13));
  assert.equal(
    nodes['ad-change-status'].textContent,
    'This library inherits permissions again. Use Apply access to let Documents stop the inheritance again.',
  );
  assert.equal(nodes['ad-apply'].disabled, false, 'Apply access is offered for the reset library');
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
  console.log(
    'PASS Sites & access handlers: staging versus apply, automatic onboarding request, initial library teams, completion polling, and author deep link. Mocked APIs; connected acceptance pending.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
