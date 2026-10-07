'use strict';
// Sites & access on the fake DOM with the real shell (shell.js, then sites-access.js) and mocked
// Dataverse APIs; no live SharePoint calls.
const fs = require('fs'),
  vm = require('vm'),
  assert = require('assert/strict'),
  path = require('path');
const { createDocument } = require('./fake-dom.cjs');
const base = path.resolve(__dirname, '../../client/admin'),
  html = fs.readFileSync(path.join(base, 'index.html'), 'utf8');
const read = (name) => fs.readFileSync(path.join(base, name), 'utf8');
const BUILD = /const BUILD = '([^']+)'/.exec(read('shell.js'))[1];
const id = (n) => String(n).padStart(8, '0') + '-0000-0000-0000-000000000000';

// Loads the page the way the app does: the shell, then this tab's script, on the access tab.
function boot(xrm, { timers, uuid, refreshCatalog = async () => {}, link = null }) {
  const document = createDocument(html);
  const session = new Map([['asxd.launched', '1']]);
  // A deep link another page stored before navigating here.
  if (link) session.set('asxd.deeplink', JSON.stringify(link));
  const window = {
    parent: { Xrm: xrm },
    location: { search: '?data=access-' + BUILD, hash: '' },
    sessionStorage: {
      getItem: (k) => session.get(k) ?? null,
      setItem: (k, v) => session.set(k, v),
      removeItem: (k) => session.delete(k),
    },
    AsxdAdmin: { refreshCatalog },
  };
  const context = vm.createContext({
    window,
    document,
    URL,
    crypto: { randomUUID: () => uuid },
    setTimeout: (fn) => timers.push(fn),
    Intl,
    URLSearchParams,
    navigator: {},
    console,
    clearTimeout: () => {},
    setInterval: () => 0,
    clearInterval: () => {},
  });
  for (const name of ['shell.js', 'sites-access.js']) vm.runInContext(read(name), context);
  const nodes = new Proxy({}, { get: (_, key) => document.getElementById(key) });
  const press = async (node) => {
    node.click();
    await document.settle();
  };
  // The button with this text inside the open confirmation of a container.
  const confirmIn = (container, label) =>
    [...container.querySelector('.confirm').querySelectorAll('button')].find(
      (b) => b.textContent === label,
    );
  const menuItem = (menu, label) =>
    [...nodes[menu].querySelectorAll('button')].find((b) => b.textContent === label);
  return { document, window, nodes, press, confirmIn, menuItem };
}
// A JSON error as Dataverse sends a refusal.
const refused = (message) => ({ ok: false, json: async () => ({ error: { message } }) });

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
  siteQueries = [],
  nativeQueries = [],
  navigations = [];
// The SharePoint sites Dataverse document management has.
let nativeSites = [{ sharepointsiteid: id(2), name: 'Delivery' }];
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
// The libraries the site's library list answers with, the destinations, revisions and templates
// that say which templates use them, and the Monitor problem summary.
let destinations = [],
  revisions = [],
  templateRows = [],
  listedLibraries = [lib],
  problemSummary = { BlockedJobs: 2, TemplateRuns: 1 };
// Page sizes of the library-count reads, team records read by ID, and GetPolicy reads in flight.
const countSizes = [],
  teamReadIds = [];
let slowPolicies = false,
  policiesInFlight = 0,
  maxPolicies = 0;
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
  Navigation: {
    openUrl: () => {},
    navigateTo: async (page) =>
      navigations.push({
        ...page,
        data: new URLSearchParams(page.webresourceName.split('?')[1] || '').get('data'),
      }),
  },
  WebApi: {
    retrieveMultipleRecords: async (table, options, size) => {
      if (table === 'asx_library' && options.includes('$select=_asx_siteid_value'))
        countSizes.push(size);
      if (table === 'team') teamQueries.push(options);
      if (table === 'asx_library') libraryQueries.push(options);
      if (table === 'asx_site') siteQueries.push(options);
      if (table === 'sharepointsite') nativeQueries.push(options);
      if (table === 'asx_library' && options.includes('$select=_asx_siteid_value'))
        return { entities: [lib, lib2] };
      if (table === 'asx_library' && options.includes('_asx_siteid_value eq'))
        return { entities: listedLibraries };
      if (table === 'asx_destination')
        return {
          entities: destinations.filter((d) => options.includes(d._asx_libraryid_value)),
        };
      if (table === 'asx_revision')
        return { entities: revisions.filter((r) => options.includes(r.asx_revisionid)) };
      if (table === 'asx_template') return { entities: templateRows };
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
                  ? nativeSites
                  : [],
      };
    },
    retrieveRecord: async (table, key) =>
      (table === 'team' && teamReadIds.push(key) && false) ||
      (table === 'asx_library'
        ? key === id(13)
          ? lib2
          : lib
        : table === 'asx_site'
          ? site
          : { name: 'Operations' }),
    online: {
      execute: async (req) => {
        const command = JSON.parse(req.Request);
        requests.push(command);
        apis.push(req.getMetadata().operationName);
        if (command.Command === 'Inspect' && inspectFails)
          throw new Error('Temporary status request failure');
        if (command.Command === 'RemoveLibrary' && removalRefusal) return refused(removalRefusal);
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
          return refused(retrySetupRefusal);
        let result;
        if (command.Command === 'Summary') result = { Status: 'Summary', Summary: problemSummary };
        else if (command.Command === 'GetPolicy') {
          result = policy;
          // Held for a turn while slowPolicies is on, to count the reads in flight together.
          if (slowPolicies) {
            policiesInFlight++;
            maxPolicies = Math.max(maxPolicies, policiesInFlight);
            await new Promise(setImmediate);
            policiesInFlight--;
          }
        } else if (command.Command === 'RetryAccessRun')
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
const page = boot(xrm, {
  timers,
  uuid: id(9),
  refreshCatalog: async () => {
    refresh++;
  },
});
const { document, window, nodes, press, confirmIn, menuItem } = page;
// Makes the page track a setup the way a reload finds one; trackOperation reads it at once.
async function trackSetup(key, name) {
  await window.AsxdSites.trackOperation(key, name, 'LibrarySetup');
  await document.settle();
}
// A click that takes focus first, as it does in a browser (the fake DOM's click() does not).
const pressFocused = async (node) => {
  node.focus();
  await press(node);
};
const teamRows = () => nodes['ad-team-rows'].querySelectorAll('.team-row');
const rowButton = (row, text) =>
  [...row.querySelectorAll('button')].find((b) => b.textContent === text);
// A team row's ✕, named by its aria-label.
const removeOf = (row) =>
  [...row.querySelectorAll('button')].find((b) =>
    /^Remove /.test(b.getAttribute('aria-label') || ''),
  );
const libraryRows = () => [...nodes['ad-libraries'].querySelectorAll('tr')];
const libraryButton = (name) =>
  [...nodes['ad-libraries'].querySelectorAll('button')].find((b) => b.textContent === name);
const rowNamed = (name) =>
  libraryRows().find((r) => r.querySelector('.library-name').textContent === name);
// The Access column of a library's row, as a person reads it.
const accessOf = (name) => rowNamed(name).querySelector('.access').visibleText;
// Opens the drawer of a setup that has no library row yet, by its operation key.
const openSetup = (key) =>
  pressFocused(nodes['ad-libraries'].querySelector('[data-focus-key="setup:' + key + '"]'));
const stage = (team, access) => {
  nodes['ad-add-team'].onclick();
  nodes['ad-team-choice'].value = team;
  nodes['ad-team-access'].value = access;
  nodes['ad-stage-team'].onclick();
};
(async () => {
  await document.fire('DOMContentLoaded');
  assert.equal(nodes['ad-site-title'].textContent, 'Delivery');
  await pressFocused(libraryButton('General'));
  // The site header links to Monitor with the problem count, read once at load.
  const problemPill = nodes['ad-site-problems'].querySelector('.problem-pill');
  assert.equal(problemPill.hidden, false);
  assert.equal(problemPill.textContent, 'Monitor · 2 problems');
  assert.equal(requests.filter((r) => r.Command === 'Summary').length, 1);
  assert.equal(nodes['ad-library-title'].textContent, 'General');
  const teamQuery = decodeURIComponent(teamQueries[0]);
  assert.match(teamQuery, /teamtype eq 0 or teamtype eq 2 or teamtype eq 3/);
  assert.match(teamQuery, /isdefault eq false/);
  assert.doesNotMatch(teamQuery, /teamtype eq 1/, 'Access teams stay out of the picker');
  for (const picker of ['ad-team-choice', 'ad-initial-team']) {
    const options = [...nodes[picker].options].slice(1),
      byId = (n) => options.find((o) => o.value === id(n));
    assert.equal(options.length, teams.length, picker + ' lists every eligible team');
    assert.equal(byId(4).textContent, 'Operations');
    assert.equal(byId(4).disabled, false);
    assert.equal(byId(5).textContent, 'Finance (Entra group)');
    assert.equal(byId(5).disabled, false);
    assert.equal(byId(6).textContent, 'Project X (Microsoft 365 group · owners)');
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
    assert.equal(byId(11).textContent, 'Finance owners (Entra group · all members)');
    assert.equal(byId(12).disabled, false);
    assert.equal(byId(12).textContent, 'Project Y (Microsoft 365 group · members + guests)');
  }
  stage(id(4), 'Read');
  assert.equal(
    requests.filter((r) => r.Command === 'ApplyPolicy').length,
    0,
    'Staging does not onboard syncing',
  );
  assert.equal(nodes['ad-apply'].disabled, false);
  await press(nodes['ad-apply']);
  const applied = requests.find((r) => r.Command === 'ApplyPolicy');
  assert.equal(applied.LibraryId, id(3));
  assert.equal(applied.Entries[0].TeamId, id(4));
  assert.equal(applied.ReadRole, undefined, 'Role definitions stay server-owned');
  assert.equal(nodes['ad-apply'].disabled, true);
  assert.equal(nodes['fb-access-library'].textContent, 'Access submitted.');
  policy = {
    Status: 'Applied',
    RowVersion: '3',
    Policy: { Desired: applied.Entries, Applied: applied.Entries },
  };
  policy.Policy.Notices = [
    "Team 'Operations': Integration App was not added to the library group because it is an application user.",
    '<b>SharePoint did not add ghost@example.com</b>',
  ];
  await timers.shift()();
  assert.match(nodes['ad-run-status'].textContent, /confirmed/);
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
  // Remove on the row, then Apply: the removal is confirmed in the page and sent as None.
  await press(removeOf(teamRows()[0]));
  policy = {
    Status: 'Applied',
    RowVersion: '4',
    Policy: { Desired: applied.Entries, Applied: applied.Entries },
  };
  await press(nodes['ad-apply']);
  assert.equal(
    requests.filter((r) => r.Command === 'ApplyPolicy').length,
    1,
    'Nothing is applied before the admin confirms the removal',
  );
  await press(confirmIn(nodes['ad-drawer'], 'Apply access'));
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
  // A team with no access is not listed; adding it again gives it access.
  assert.doesNotMatch(nodes['ad-team-rows'].visibleText, /Operations/);
  stage(id(4), 'Read');
  const beforeConflict = requests.filter((r) => r.Command === 'ApplyPolicy').length;
  policy = {
    Status: 'Applied',
    RowVersion: '6',
    Policy: { Desired: [{ TeamId: id(4), Access: 'Contribute' }], Applied: [] },
  };
  await press(nodes['ad-apply']);
  assert.equal(
    requests.filter((r) => r.Command === 'ApplyPolicy').length,
    beforeConflict,
    'Do not overwrite another administrator access change',
  );
  assert.match(nodes['fb-access-library'].textContent, /changed.*reload/i);
  assert.equal(nodes['fb-access-library'].className, 'feedback is-error');
  policy = {
    Status: 'Applied',
    RowVersion: '7',
    Policy: { Desired: removal.Entries, Applied: removal.Entries, OperationKey: 'background' },
  };
  await timers.shift()();
  // While a run is in progress the teams stay editable and Apply is sent: the change waits for
  // the run and is applied right after it.
  assert.equal(teamRows()[0].querySelector('select').disabled, false);
  assert.equal(nodes['ad-add-team'].disabled, false);
  await press(nodes['ad-apply']);
  assert.equal(
    requests.filter((r) => r.Command === 'ApplyPolicy').length,
    beforeConflict + 1,
    'An apply during a run is sent and waits for that run',
  );
  assert.equal(nodes['fb-access-library'].textContent, 'Access submitted.');
  assert.equal(nodes['ad-run-status'].textContent, 'Saved · applies after the current run');
  policy = {
    Status: 'Applied',
    RowVersion: '9',
    Policy: { Desired: [{ TeamId: id(4), Access: 'Read' }], Applied: [] },
  };
  await timers.shift()();
  // Add site: pick a SharePoint site in the combobox, then add and check it.
  await press(nodes['ad-add-site']);
  await press(nodes['ad-native-list'].querySelectorAll('[role=option]')[0]);
  assert.equal(nodes['ad-native-search'].value, 'Delivery');
  await pressFocused(nodes['ad-validate']);
  assert(document.activeElement === nodes['ad-add-site'], 'Add site returns focus to ＋ Add');
  assert.equal(requests.find((r) => r.Command === 'AddSite').NativeSiteId, id(2));
  assert.equal(requests.find((r) => r.Command === 'AddSite').Name, 'Delivery');
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
  assert.doesNotMatch(nodes['ad-provision-progress'].textContent, /Open Administration/);
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
  assert.match(nodes['fb-access'].textContent, /ready/);
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
  await pressFocused(nodes['ad-provision']);
  assert(
    document.activeElement === nodes['ad-create'],
    'Create returns focus to ＋ Create library',
  );
  const created = requests.find((r) => r.Command === 'CreateLibrary');
  assert.equal(created.SiteId, id(1));
  assert.equal(created.Entries[0].Access, 'Contribute');
  assert.equal(nodes['fb-access'].textContent, 'Creating Projects.');
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
  await press(nodes['ad-existing']);
  await timers.shift()();
  assert.equal(nodes['ad-existing-form'].hidden, false);
  assert.equal(nodes['ad-existing-more'].hidden, false);
  assert.deepEqual(
    [...nodes['ad-existing-choices'].children].map((c) => c.textContent),
    ['Add Archive', 'Add Removed earlier'],
  );
  assert.match(libraryQueries.at(-1), /statecode eq 0 and \(asx_listid eq '00000008-.*' or /);
  await pressFocused(nodes['ad-existing-choices'].children[0]);
  assert(
    document.activeElement === nodes['ad-existing'],
    'Add existing returns focus to Add existing library',
  );
  const existing = requests.find((r) => r.Command === 'AddLibrary');
  assert.equal(existing.ListId, id(8));
  assert.equal(existing.NativeParentId, undefined, 'Native navigation is automatic');
  assert(!requests.some((r) => r.Command === 'RegisterTeam'), 'No manual team mapping step');
  lib.asx_policyapplied = false;
  await window.AsxdSites.selectLibrary(id(3));
  assert.equal(
    nodes['ad-library-status'].hidden,
    true,
    'An approved library needs no attention before its access is applied',
  );
  assert.doesNotMatch(accessOf('General'), /Needs attention/);
  lib.asx_policyapplied = true;
  lib.asx_approved = false;
  await window.AsxdSites.selectLibrary(id(3));
  assert.equal(nodes['ad-library-status'].hidden, false);
  assert.equal(nodes['ad-library-status'].textContent, 'Needs attention');
  assert.equal(accessOf('General'), 'Needs attention');
  lib.asx_approved = true;
  // Consent: a team whose group reaches more people is applied only after the admin confirms
  // the warning shown in the page.
  policy = { Status: 'Applied', RowVersion: '20', Policy: { Desired: [], Applied: [] } };
  await window.AsxdSites.selectLibrary(id(13));
  assert.equal(nodes['ad-library-title'].textContent, 'Contracts');
  stage(id(11), 'Read');
  const applies = () => requests.filter((r) => r.Command === 'ApplyPolicy').length;
  const beforeConsent = applies();
  await press(nodes['ad-apply']);
  assert.equal(applies(), beforeConsent, 'Nothing is applied before the admin confirms');
  assert.match(
    document.activeElement.textContent,
    /All members of the group will have access to this library, not only its owners\./,
  );
  assert.equal(nodes['ad-apply'].textContent.includes('Confirm and apply'), false);
  await press(confirmIn(nodes['ad-drawer'], 'Apply access'));
  const consented = requests.filter((r) => r.Command === 'ApplyPolicy').at(-1);
  assert.equal(applies(), beforeConsent + 1);
  assert.equal(consented.AcknowledgeBroaderAccess, true);
  assert.equal(consented.Entries[0].TeamId, id(11));
  assert.equal(nodes['ad-drawer'].querySelector('.confirm'), null);
  assert.equal(nodes['ad-apply'].textContent, 'Apply access');
  // The same for a new library's initial team.
  nodes['ad-create'].onclick();
  nodes['ad-library-name'].value = 'Partners';
  nodes['ad-initial-team'].value = id(12);
  nodes['ad-initial-access'].value = 'Read';
  const creates = () => requests.filter((r) => r.Command === 'CreateLibrary').length;
  const beforeCreate = creates();
  await press(nodes['ad-provision']);
  assert.equal(creates(), beforeCreate, 'No library is created before the admin confirms');
  assert.equal(
    document.activeElement.textContent,
    "The group's guests will also have access to this library.",
  );
  await press(confirmIn(nodes['ad-library-form'], 'Create library'));
  const createdWithConsent = requests.filter((r) => r.Command === 'CreateLibrary').at(-1);
  assert.equal(creates(), beforeCreate + 1);
  assert.equal(createdWithConsent.AcknowledgeBroaderAccess, true);
  // An owner team needs no consent.
  nodes['ad-create'].onclick();
  nodes['ad-library-name'].value = 'Ledger';
  nodes['ad-initial-team'].value = id(4);
  await press(nodes['ad-provision']);
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
  await press(nodes['ad-existing']);
  await timers.shift()();
  const adds = () => requests.filter((r) => r.Command === 'AddLibrary');
  const beforeAdd = adds().length;
  assert.equal(nodes['ad-existing-choices'].children[0].textContent, 'Add Shared');
  await press(nodes['ad-existing-choices'].children[0]);
  assert.equal(adds().length, beforeAdd, 'Nothing is added before the admin confirms');
  assert.equal(document.activeElement.textContent, warning);
  await press(confirmIn(nodes['ad-existing-form'], 'Stop inheritance and add'));
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
    (n) => n.tagName === 'BUTTON' && n.textContent === 'Stop inheritance and add',
  );
  assert(stop, 'The blocked card offers to stop the inheritance');
  await press(stop);
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
  assert.equal(nodes['ad-run-status'].textContent, inheritsAgain);
  assert.equal(nodes['ad-run-status'].className, 'ad-issue');
  assert.equal(accessOf('Reset'), 'Needs attention');
  assert.equal(nodes['ad-run-actions'].hidden, false, 'The stopped run offers Retry and Cancel');
  assert.equal(nodes['ad-run-retry'].disabled, false);
  assert.equal(nodes['ad-run-cancel'].disabled, false);
  assert.equal(nodes['ad-apply'].disabled, false, 'Apply access is offered for the reset library');
  assert.equal(nodes['ad-apply'].textContent, 'Stop inheritance and apply');
  lib2.asx_policyapplied = true;
  const beforeReapply = applies();
  await press(nodes['ad-apply']);
  assert.equal(applies(), beforeReapply, 'Nothing is applied before the admin confirms');
  assert.match(
    document.activeElement.textContent,
    /When you apply access, Documents stops the inheritance, keeps a copy of the current site permissions/,
  );
  await press(confirmIn(nodes['ad-drawer'], 'Apply access'));
  const reapplied = requests.filter((r) => r.Command === 'ApplyPolicy').at(-1);
  assert.equal(applies(), beforeReapply + 1);
  assert.equal(reapplied.BreakInheritance, true);
  assert.equal(reapplied.AcknowledgeBroaderAccess, undefined);
  // Re-point asks in the page, under the library's header, then shows what changed as text.
  const repoints = () => requests.filter((r) => r.Command === 'RepointLibrary');
  await nodes['ad-library-menu'].onclick();
  await press(menuItem('ad-library-menu-list', 'Re-point library'));
  assert.equal(repoints().length, 0, 'Nothing is re-pointed before the admin confirms');
  assert.match(
    document.activeElement.textContent,
    /Re-point Reset: .*Nothing in SharePoint changes/,
  );
  await press(confirmIn(nodes['ad-drawer'], 'Keep current address'));
  assert.equal(nodes['ad-drawer'].querySelector('.confirm'), null);
  assert(document.activeElement === nodes['ad-library-menu'], 'Keep returns to the ⋯ button');
  await nodes['ad-library-menu'].onclick();
  await press(menuItem('ad-library-menu-list', 'Re-point library'));
  await press(confirmIn(nodes['ad-drawer'], 'Re-point'));
  assert.equal(repoints().length, 1);
  assert.equal(repoints()[0].CatalogId, id(19));
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
  assert.match(nodes['fb-access'].textContent, /Reset re-pointed/);
  assert.equal(nodes['ad-changes'].hidden, false);
  assert.equal(
    nodes['ad-changes'].children[0].textContent,
    'https://example.sharepoint.com/sites/proto/General → https://example.sharepoint.com/sites/proto/<b>Shared</b>',
    'Changes are text, never markup',
  );
  // A library that no longer exists is reported on its card.
  await nodes['ad-site-menu'].onclick();
  await press(menuItem('ad-site-menu-list', 'Re-point site'));
  assert.match(document.activeElement.textContent, /Re-point Delivery: .*SharePoint site record/);
  await press(confirmIn(nodes['ad-site-header'], 'Re-point'));
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
    !find(
      nodes['ad-provision-progress'],
      (n) => n.tagName === 'BUTTON' && n.textContent === 'Retry',
    ),
    'A blocked re-point is not retried in place',
  );
  const again = find(
    nodes['ad-provision-progress'],
    (n) => n.tagName === 'BUTTON' && n.textContent === 'Re-point again',
  );
  assert(again, 'A blocked re-point offers to re-point again');
  const beforeAgain = requests.filter((r) => r.Command === 'RepointSite').length;
  delete inspectByKey['catalogprobe:test'];
  await press(again);
  const repointedAgain = requests.filter((r) => r.Command === 'RepointSite');
  assert.equal(repointedAgain.length, beforeAgain + 1);
  assert.equal(repointedAgain.at(-1).CatalogId, id(1));
  assert(!requests.some((r) => r.Command === 'Retry' && r.Key === 'catalogprobe:test'));
  // Remove asks in the page; a refusal lists the templates that use the library.
  const removes = () => requests.filter((r) => r.Command === 'RemoveLibrary');
  await nodes['ad-library-menu'].onclick();
  await press(menuItem('ad-library-menu-list', 'Remove library'));
  assert.equal(removes().length, 0, 'Nothing is removed before the admin confirms');
  assert.match(document.activeElement.textContent, /Nothing in SharePoint is deleted or changed/);
  removalRefusal = "Used by template 'Accounts' (published). Change the template first.";
  await press(confirmIn(nodes['ad-drawer'], 'Remove library'));
  assert.equal(removes().length, 1);
  assert.equal(removes()[0].CatalogId, id(19));
  assert.equal(nodes['fb-access-library'].textContent, removalRefusal);
  assert.equal(nodes['fb-access-library'].className, 'feedback is-error');
  removalRefusal = null;
  await nodes['ad-library-menu'].onclick();
  await press(menuItem('ad-library-menu-list', 'Remove library'));
  await press(confirmIn(nodes['ad-drawer'], 'Remove library'));
  assert.equal(removes().length, 2);
  assert.match(nodes['fb-access'].textContent, /Reset was removed from Documents/);
  assert(
    document.activeElement === nodes['ad-libraries-heading'],
    'Remove library moves focus to the Libraries heading',
  );
  assert.equal(nodes['ad-libraries-heading'].getAttribute('tabindex'), '-1');
  assert.match(nodes['ad-changes'].textContent, /Nothing was deleted or changed in SharePoint/);
  assert.match(libraryQueries.at(-1), /statecode eq 0/, 'Removed libraries are hidden');
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
    assert.equal(nodes['ad-run-status'].textContent, 'SharePoint refused the write (HTTP 403).');
    assert.equal(nodes['ad-run-status'].className, 'ad-issue');
    assert.equal(accessOf('Stuck'), 'Needs attention');
    assert.equal(nodes['ad-run-actions'].hidden, false);
    assert.equal(nodes['ad-add-team'].disabled, false, 'A stopped run does not lock the teams');
    await press(nodes['ad-run-retry']);
    assert.deepEqual(requests.at(-1), {
      Command: 'RetryAccessRun',
      LibraryId: id(20),
      OperationKey: 'policywork:stuck',
    });
    assert.equal(apis.at(-1), 'asx_SecurityAdmin');
    assert.equal(nodes['ad-run-status'].textContent, 'Applying access and syncing members…');
    assert.equal(accessOf('Stuck'), 'Applying access…');
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
      nodes['ad-run-status'].textContent,
      'Waiting to retry after a temporary error (HTTP 503); attempt 4. Next check: 2026-10-05 18:30 UTC.',
    );
    assert.equal(accessOf('Stuck'), 'Needs attention');
    const cancels = () => requests.filter((r) => r.Command === 'CancelAccessRun');
    await press(nodes['ad-run-cancel']);
    assert.equal(cancels().length, 0, 'Nothing is cancelled before the admin confirms');
    assert.match(
      document.activeElement.textContent,
      /^Cancel the access run for Stuck\? .*nothing is undone or deleted/,
    );
    assert(confirmIn(nodes['ad-drawer'], 'Keep access run'));
    await press(confirmIn(nodes['ad-drawer'], 'Cancel access run'));
    assert.deepEqual(cancels(), [
      { Command: 'CancelAccessRun', LibraryId: id(20), OperationKey: 'policywork:stuck' },
    ]);
    assert.match(
      nodes['fb-access-library'].textContent,
      /Nothing in SharePoint was undone or deleted/,
    );
    assert.equal(
      nodes['ad-run-status'].textContent,
      'The access run was cancelled. Apply access to run it again.',
    );
    assert.equal(accessOf('Stuck'), 'Needs attention');
    assert.equal(nodes['ad-apply'].disabled, false, 'Apply access starts a new run');
    // Apply access replaces a stopped run with the admin's newer access.
    policy = stuck;
    const team = teamRows()[0].querySelector('select');
    team.value = 'Contribute';
    team.onchange();
    const before = requests.filter((r) => r.Command === 'ApplyPolicy').length;
    await press(nodes['ad-apply']);
    const replaced = requests.filter((r) => r.Command === 'ApplyPolicy');
    assert.equal(replaced.length, before + 1, 'A stopped run does not refuse Apply');
    assert.equal(replaced.at(-1).Entries[0].Access, 'Contribute');
    // A change waiting behind a stopped run: the status is the run's notice, with no tail.
    policy = { ...stuck, Policy: { ...stuck.Policy, ApplyPending: true } };
    lib2.asx_libraryid = id(21); // a library not loaded yet, so its access is read
    await window.AsxdSites.selectLibrary(id(13));
    assert.equal(nodes['ad-run-status'].textContent, 'SharePoint refused the write (HTTP 403).');
    assert.equal(nodes['ad-run-status'].className, 'ad-issue');
    assert.equal(accessOf('Stuck'), 'Needs attention');
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
      nodes['ad-run-status'].textContent,
      /^Needs attention: access is applied, but team membership was not synced. People removed from a team keep access, and people added get none/,
    );
    assert.equal(nodes['ad-run-status'].className, 'ad-issue');
    assert.equal(accessOf('Stuck'), 'Needs attention');
    assert.doesNotMatch(nodes['ad-run-status'].textContent, /confirmed/);
    policy = { ...policy, Policy: { ...policy.Policy, MembershipIncomplete: false, Notices: [] } };
    lib2.asx_libraryid = id(23);
    await window.AsxdSites.selectLibrary(id(13));
    assert.equal(nodes['ad-run-status'].textContent, 'Access and team membership confirmed.');
    // The header follows the policy and its run, never the flag the library list was read
    // with: here the list still says access is pending while the policy is Applied.
    lib2.asx_policyapplied = false;
    lib2.asx_libraryid = id(24);
    await window.AsxdSites.selectLibrary(id(13));
    assert.equal(nodes['ad-run-status'].textContent, 'Access and team membership confirmed.');
    assert.equal(accessOf('Stuck'), 'Access applied');
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
    assert.equal(accessOf('Stuck'), 'Applying access…');
    // It applies: the next poll shows it, with no reload.
    policy = {
      Status: 'Applied',
      RowVersion: '44',
      Policy: { Desired: stuck.Policy.Desired, Applied: stuck.Policy.Desired },
    };
    await timers.shift()();
    assert.equal(nodes['ad-run-status'].textContent, 'Access and team membership confirmed.');
    assert.equal(accessOf('Stuck'), 'Access applied');
    // Access never applied.
    policy = { Status: 'Missing', RowVersion: '', Policy: null };
    lib2.asx_policyapplied = false;
    lib2.asx_libraryid = id(26);
    await window.AsxdSites.selectLibrary(id(13));
    assert.equal(accessOf('Stuck'), 'Access setup pending');
    lib2.asx_policyapplied = true;
  }
  {
    // A library setup that needs attention offers Retry and Cancel through the catalog API, so
    // a Documents Security Administrator needs no Operator role.
    nodes['ad-create'].onclick();
    nodes['ad-library-name'].value = 'Stalled';
    nodes['ad-initial-team'].value = '';
    await press(nodes['ad-provision']);
    inspectByKey['librarycreate:test'] = {
      Status: 'RecoveryRequired',
      Key: 'librarycreate:test',
      Issue: 'The library create may have reached SharePoint and its answer was lost.',
    };
    await timers.shift()();
    // A setup with no library row yet has a row of its own; its stage card opens in the drawer.
    assert.match(accessOf('Stalled'), /^Needs attention · step \d of 4$/);
    assert.doesNotMatch(nodes['ad-provision-progress'].textContent, /Stalled/);
    await openSetup('librarycreate:test');
    const area = nodes['ad-drawer-progress'],
      named = (text) => find(area, (n) => n.tagName === 'BUTTON' && n.textContent === text);
    assert.equal(nodes['ad-library-title'].textContent, 'Stalled');
    assert.match(area.textContent, /Stalled.*Needs attention/);
    assert.match(area.textContent, /answer was lost/);
    retrySetupRefusal =
      'The library create may have reached SharePoint and its answer was lost. Recover it with the original create response from the flow run, or Cancel the setup.';
    await press(named('Retry'));
    assert.deepEqual(requests.at(-1), { Command: 'RetrySetup', Key: 'librarycreate:test' });
    assert.equal(apis.at(-1), 'asx_CatalogAdmin');
    // A card in the drawer reports in the drawer's footer, which a full-width drawer keeps in
    // view; the line under the page header does not get it.
    assert.equal(nodes['fb-access-library'].textContent, retrySetupRefusal);
    assert.equal(nodes['fb-access-library'].className, 'feedback is-error');
    assert.equal(nodes['ad-drawer-footer'].hidden, false);
    assert.equal(nodes['ad-apply'].hidden, true, 'A setup drawer has no Apply');
    assert.notEqual(nodes['fb-access'].textContent, retrySetupRefusal);
    retrySetupRefusal = null;
    await press(named('Retry'));
    assert.equal(apis.at(-1), 'asx_CatalogAdmin');
    assert.equal(nodes['fb-access-library'].textContent, 'Setup queued again.');
    await timers.shift()();
    await press(named('Cancel setup'));
    assert.equal(
      requests.filter((r) => r.Command === 'CancelSetup').length,
      0,
      'Nothing is cancelled before the admin confirms',
    );
    assert.match(
      document.activeElement.textContent,
      /^Cancel the setup of Stalled\? Nothing in SharePoint is deleted\./,
    );
    await press(confirmIn(area, 'Cancel setup'));
    assert.deepEqual(requests.at(-1), { Command: 'CancelSetup', Key: 'librarycreate:test' });
    assert.equal(apis.at(-1), 'asx_CatalogAdmin');
    assert.doesNotMatch(area.textContent, /Stalled/, 'A cancelled setup leaves the list');
    assert.match(nodes['ad-changes'].textContent, /Nothing in SharePoint was deleted/);
    assert.equal(nodes['ad-drawer'].hidden, true, 'The drawer of a cancelled setup closes');
    assert.equal(nodes['fb-access'].textContent, 'The setup of Stalled was cancelled.');
    assert(
      apis.every((a, i) => a !== 'asx_ManageWork' || requests[i].Command === 'Summary'),
      'Sites calls asx_ManageWork only for the header problem count',
    );
    delete inspectByKey['librarycreate:test'];
  }
  {
    // A setup whose first access run was cancelled ends Ready with that notice instead of
    // "Setup completed.".
    nodes['ad-create'].onclick();
    nodes['ad-library-name'].value = 'Unsynced';
    nodes['ad-initial-team'].value = '';
    await press(nodes['ad-provision']);
    inspectByKey['librarycreate:test'] = {
      Status: 'Ready',
      Key: 'librarycreate:test',
      Issue: 'Library created. Its first access run was cancelled; apply access on the library.',
    };
    await timers.shift()();
    await pressFocused(libraryButton('Unsynced'));
    const card = [...nodes['ad-drawer-progress'].querySelectorAll('section')].find((c) =>
      /Unsynced/.test(c.textContent),
    );
    assert(card, 'The finished setup stays listed');
    assert.match(card.textContent, /first access run was cancelled; apply access/);
    assert.doesNotMatch(card.textContent, /Setup completed/);
    assert(rowButton(card, 'Dismiss'));
    assert.match(nodes['fb-access'].textContent, /Unsynced: Library created. Its first access run/);
    delete inspectByKey['librarycreate:test'];
  }
  {
    // After a reload, a re-point still running or blocked is found again with the other setup
    // activity, and its command and ID come from its probe, so "Re-point again" still works.
    // Until a card is read it shows "Loading…", never the generic message or action. Cards of
    // a library that was removed or deleted are not shown at all.
    const later = [],
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
    const reloadedXrm = {
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
    };
    const reloaded = boot(reloadedXrm, { timers: later, uuid: id(22) });
    const fresh = reloaded.nodes,
      card = fresh['ad-provision-progress'];
    await reloaded.document.fire('DOMContentLoaded');
    assert(inspecting, 'The cards found again are being read');
    // Before the first read: no generic message or action, and nothing of the removed ones.
    assert.doesNotMatch(card.textContent, /Setup needs review|Retry/);
    assert.doesNotMatch(card.textContent, /Needs attention/);
    assert.doesNotMatch(card.textContent, /AcceptC Chrome 1|Deleted later/);
    assert.match(card.textContent, /Loading…/);
    release();
    await reloaded.document.settle();
    assert.match(activity[0], /asx_workkind eq 'Repoint'/);
    // Read at once after the load, not at the first poll. The re-point of General is a library
    // card: its row needs attention, and its drawer shows the card.
    assert.doesNotMatch(card.textContent, /Loading…/);
    const generalRow = () =>
      [...fresh['ad-libraries'].querySelectorAll('tr')].find(
        (r) => r.querySelector('.library-name').textContent === 'General',
      );
    assert.equal(generalRow().querySelector('.access').visibleText, 'Needs attention');
    await reloaded.press(generalRow().querySelector('.library-name'));
    const drawerCards = fresh['ad-drawer-progress'];
    assert.match(drawerCards.textContent, /General.*no longer exists on the site/);
    assert(checks.some((c) => c.includes('asx_libraryid eq ' + id(30))));
    assert(checks.some((c) => c.includes('asx_libraryid eq ' + id(32))));
    await later.shift()();
    assert.doesNotMatch(
      card.textContent + drawerCards.textContent,
      /AcceptC Chrome 1|Deleted later/,
      'Cards of a removed or deleted library are not shown',
    );
    assert.equal(drawerCards.children.length, 1);
    const again = find(
      drawerCards,
      (n) => n.tagName === 'BUTTON' && n.textContent === 'Re-point again',
    );
    assert(again, 'A blocked re-point found after a reload offers to re-point again');
    assert(!find(drawerCards, (n) => n.tagName === 'BUTTON' && n.textContent === 'Retry'));
    await reloaded.press(again);
    assert.deepEqual(
      sent.filter((c) => c.Command === 'RepointLibrary').map((c) => c.CatalogId),
      [id(3)],
    );
    await later.shift()();
    assert.match(drawerCards.textContent, /General.*no longer exists on the site/);
    // Remove: the library leaves the list and its re-point card goes, with no reload.
    const listed = () =>
      [...fresh['ad-libraries'].querySelectorAll('.library-name')].map((c) => c.textContent);
    assert.deepEqual(listed(), ['General']);
    const before = lists.length;
    await fresh['ad-library-menu'].onclick();
    await reloaded.press(reloaded.menuItem('ad-library-menu-list', 'Remove library'));
    await reloaded.press(reloaded.confirmIn(fresh['ad-drawer'], 'Remove library'));
    assert.deepEqual(sent.at(-1), { Command: 'RemoveLibrary', CatalogId: id(3) });
    assert.match(fresh['fb-access'].textContent, /General was removed from Documents/);
    assert(
      reloaded.document.activeElement === fresh['ad-site-title'],
      'With no library left, focus moves to the site heading',
    );
    assert(lists.length > before, 'The library list is read again after Remove');
    assert.deepEqual(listed(), [], 'The removed library is no longer listed');
    assert.equal(fresh['ad-drawer'].hidden, true);
    assert.doesNotMatch(
      card.textContent + drawerCards.textContent,
      /General/,
      'Its re-point card goes with it',
    );
    assert.equal(fresh['ad-site-cards'].hidden, true);
    // A link to it does not bring it back.
    await reloaded.window.AsxdSites.selectLibrary(id(3));
    assert.match(fresh['fb-access'].textContent, /This library was removed from Documents/);
    assert.deepEqual(listed(), []);
    // Remove site, now that it has no libraries: the site leaves the list the same way.
    await fresh['ad-site-menu'].onclick();
    const removeSite = reloaded.menuItem('ad-site-menu-list', 'Remove site');
    assert.equal(removeSite.getAttribute('aria-disabled'), null);
    await reloaded.press(removeSite);
    await reloaded.press(reloaded.confirmIn(fresh['ad-site-header'], 'Remove site'));
    assert.deepEqual(sent.at(-1), { Command: 'RemoveSite', CatalogId: id(1) });
    assert.deepEqual(
      [...fresh['ad-sites'].children].map((c) => c.textContent),
      [],
      'The removed site is no longer listed',
    );
    assert.equal(fresh['ad-site-title'].textContent, 'Select or add a site');
    assert(
      reloaded.document.activeElement === fresh['ad-sites-heading'],
      'Remove site moves focus to the Sites heading',
    );
  }
  {
    // A team deleted in Dataverse while it has access to the library: Dataverse no longer has
    // the team, so reading it by ID fails. The library shows it as a deleted team by its last
    // known name, with what happens next; one whose access is already removed is not listed;
    // status polling goes on; and Apply access is enabled, sends the change and succeeds.
    const later = [],
      sent = [],
      teamReads = [],
      gone = id(40),
      unnamed = id(41);
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
    const deletedXrm = {
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
    };
    const deleted = boot(deletedXrm, { timers: later, uuid: id(42) });
    const fresh = deleted.nodes;
    await deleted.document.fire('DOMContentLoaded');
    await deleted.press(fresh['ad-libraries'].querySelector('.library-name'));
    assert.doesNotMatch(fresh['fb-access-library'].textContent, /not found/);
    const teamRowsOf = () => [...fresh['ad-team-rows'].querySelectorAll('.team-row')];
    const rows = () => teamRowsOf().map((r) => r.querySelector('strong').textContent);
    assert.deepEqual(rows(), ['Deleted team: AcceptC Team 1', 'Operations']);
    assert.equal(
      teamRowsOf()[0].querySelector('.ad-issue').textContent,
      'Documents removes its access the next time access is applied.',
    );
    assert.doesNotMatch(fresh['ad-team-rows'].visibleText, /Apply access to clear it/);
    assert.deepEqual(teamReads, [], 'No team is read by ID once the policy names it');
    // Its access cannot be chosen: the next Apply removes it.
    assert.equal(fresh['ad-team-rows'].querySelector('select').disabled, true);
    // Polling the queued run keeps working.
    for (let i = 0; i < 4; i++) await later.shift()();
    assert.equal(fresh['ad-poll-status'].hidden, true);
    assert.doesNotMatch(fresh['ad-poll-status'].textContent, /Status updates are unavailable/);
    // Apply is enabled with no other change, and succeeds.
    assert.equal(fresh['ad-apply'].disabled, false);
    await deleted.press(fresh['ad-apply']);
    const apply = sent.find((c) => c.Command === 'ApplyPolicy');
    assert(apply, 'Apply access sends the change');
    assert.doesNotMatch(fresh['fb-access-library'].textContent, /not found|changed since/);
    assert.equal(fresh['fb-access-library'].textContent, 'Access submitted.');
    assert.deepEqual(rows(), ['Operations'], 'The deleted teams leave the list once applied');
    // The next Apply sees the list it shows.
    const access = fresh['ad-team-rows'].querySelector('select');
    access.value = 'Contribute';
    access.onchange();
    await deleted.press(fresh['ad-apply']);
    assert.equal(fresh['fb-access-library'].textContent, 'Access submitted.');
    assert.equal(sent.filter((c) => c.Command === 'ApplyPolicy').length, 2);
  }
  // The blocks below run on the first page again, on a library named General with Operations.
  policy = {
    Status: 'Applied',
    RowVersion: '60',
    Policy: {
      Desired: [{ TeamId: id(4), Access: 'Read' }],
      Applied: [{ TeamId: id(4), Access: 'Read' }],
    },
  };
  lib2.asx_libraryid = id(50);
  lib2.asx_name = 'General';
  await window.AsxdSites.selectLibrary(id(13));
  {
    // F-14: a refusal shows the server's sentence, not a raw body, at the form that failed.
    removalRefusal = 'Remove is refused: templates use this library: Account onboarding.';
    await nodes['ad-library-menu'].onclick();
    await [...nodes['ad-library-menu-list'].querySelectorAll('button')]
      .find((b) => b.textContent === 'Remove library')
      .onclick();
    await document.settle();
    const box = nodes['ad-drawer'].querySelector('.confirm');
    assert(box, 'The confirmation renders inside the library drawer (F-12)');
    assert.equal(
      document.activeElement.textContent.startsWith('Remove General from Documents?'),
      true,
    );
    [...box.querySelectorAll('button')].find((b) => b.textContent === 'Remove library').click();
    await document.settle();
    assert.equal(
      nodes['fb-access-library'].textContent,
      'Remove is refused: templates use this library: Account onboarding.',
    );
    assert.doesNotMatch(nodes['fb-access-library'].textContent, /\{|error/);
    removalRefusal = null;
  }
  {
    // F-21: Remove on a team row strikes it through with Undo; Apply with removals says what changes.
    const row = teamRows()[0];
    assert.deepEqual(
      [...row.querySelector('select').options].map((o) => o.value),
      ['Read', 'Contribute'],
    );
    assert.equal(row.querySelector('select').getAttribute('aria-label'), 'Access for Operations');
    // Focus keys (spec 5.2): each team-row control has a stable key, and a re-render keeps focus by it.
    const accessKey = row.querySelector('select').dataset.focusKey;
    assert.match(accessKey, /^team:[^:]+:access$/);
    const team = accessKey.replace(/:access$/, '');
    const removeButton = removeOf(row);
    assert.equal(removeButton.dataset.focusKey, team + ':remove');
    assert.equal(
      nodes['ad-team-rows'].closest('[data-focus-scope]') !== null,
      true,
      'The team rows are a focus scope',
    );
    removeButton.focus();
    removeButton.click();
    await document.settle();
    const removed = teamRows()[0];
    assert.match(removed.visibleText, /Removed · Undo/);
    assert(removed.classList.contains('is-removed'));
    assert.equal(
      document.activeElement.dataset.focusKey,
      team + ':undo',
      'Remove moves focus to Undo',
    );
    [...removed.querySelectorAll('button')].find((b) => b.textContent === 'Undo').click();
    await document.settle();
    assert.doesNotMatch(nodes['ad-team-rows'].visibleText, /Removed/);
    assert.equal(
      document.activeElement.dataset.focusKey,
      team + ':remove',
      'Undo moves focus back to Remove',
    );
    removeOf(teamRows()[0]).click();
    await document.settle();
    nodes['ad-apply'].click();
    await document.settle();
    assert.match(
      document.activeElement.textContent,
      /Removed teams lose the access Documents gave them\. Access given another way, such as sharing links or site membership, is not changed\./,
    );
    [...nodes['ad-drawer'].querySelector('.confirm').querySelectorAll('button')]
      .find((b) => b.textContent === 'Apply access')
      .click();
    await document.settle();
    assert.equal(
      requests.filter((r) => r.Command === 'ApplyPolicy').at(-1).Entries[0].Access,
      'None',
    );
    assert.equal(nodes['fb-access-library'].textContent, 'Access submitted.');
  }
  {
    // F-33: consent uses the shared confirmation; the old "Confirm and apply" relabel is gone.
    assert.equal(nodes['ad-apply'].textContent.includes('Confirm and apply'), false);
    assert.doesNotMatch(html, /Confirm and (apply|create)/);
  }
  {
    // Remove site waits for its libraries, with the reason; the site actions live in a menu.
    await nodes['ad-site-menu'].onclick();
    const remove = [...nodes['ad-site-menu-list'].querySelectorAll('button')].find(
      (b) => b.textContent === 'Remove site',
    );
    assert.equal(remove.getAttribute('aria-disabled'), 'true');
    assert.equal(
      document.getElementById(remove.getAttribute('aria-describedby')).textContent,
      'Remove its libraries first',
    );
  }
  {
    // The Add site combobox searches as you type, after 300 ms, and offers more pages.
    nodes['ad-add-site'].click();
    await document.settle();
    const box = nodes['ad-native-search'];
    assert.equal(box.getAttribute('role'), 'combobox');
    assert.equal(box.getAttribute('aria-autocomplete'), 'list');
    box.value = 'Deliv';
    box.oninput();
    // The debounce timer is the newest one queued.
    await timers.pop()();
    await document.settle();
    assert.match(nativeQueries.at(-1), /contains\(name,'Deliv'\)/);
    assert.equal(
      nodes['ad-native-list'].querySelectorAll('[role=option]')[0].textContent,
      'Delivery',
    );
    // No SharePoint sites in Dataverse: the empty result says how to set one up.
    nativeSites = [];
    box.value = '';
    box.oninput();
    await timers.pop()();
    await document.settle();
    assert.match(
      nodes['ad-native-list'].visibleText,
      /No SharePoint sites are set up in Dataverse document management yet\./,
    );
    assert.ok(nodes['ad-native-list'].querySelector('a'), 'How to set one up');
  }
  {
    // A setup whose create is unknown shows Documents' SharePoint check and its choices.
    inspectByKey['librarycreate:lost'] = {
      Key: 'librarycreate:lost',
      Status: 'RecoveryRequired',
      RowVersion: 'rv-9',
      Recovery: {
        State: 'Found',
        Candidates: [
          {
            ListId: id(30),
            Title: 'Projects',
            Url: '/sites/delivery/Projects',
            CreatedUtc: '2026-10-06T17:41:02Z',
            IsLibrary: true,
            TitleMatches: true,
            UrlMatches: true,
            CreatedAfterRequest: true,
            CatalogEntry: 'None',
          },
        ],
        Choices: ['UseLibrary', 'CheckAgain', 'Cancel'],
      },
    };
    await trackSetup('librarycreate:lost', 'Projects');
    await openSetup('librarycreate:lost');
    const card = [...nodes['ad-drawer-progress'].querySelectorAll('section')].find((c) =>
      /Projects/.test(c.textContent),
    );
    assert.match(
      card.visibleText,
      /SharePoint has a library Projects at \/sites\/delivery\/Projects, created .*\. It matches this request\./,
    );
    [...card.querySelectorAll('button')]
      .find((b) => b.textContent === 'Use the library that was created')
      .click();
    await document.settle();
    assert.deepEqual(
      [apis.at(-1), requests.at(-1)],
      [
        'asx_CatalogAdmin',
        {
          Command: 'ResolveSetup',
          Key: 'librarycreate:lost',
          Choice: 'UseLibrary',
          ListId: id(30),
          RowVersion: 'rv-9',
        },
      ],
    );
    assert.equal(
      nodes['fb-access-library'].textContent,
      'Using the existing library. Setup continues.',
    );
    [...card.querySelectorAll('button')].find((b) => b.textContent === 'Check again').click();
    await document.settle();
    assert.deepEqual([apis.at(-1), requests.at(-1).Command], ['asx_CatalogAdmin', 'RecheckSetup']);
    assert.equal(nodes['fb-access-library'].textContent, 'Checking SharePoint again.');
    // Another blocked setup links to Monitor instead of "Open Administration".
    inspectByKey['librarycreate:blocked'] = {
      Key: 'librarycreate:blocked',
      Status: 'Blocked',
      Issue: 'SharePoint refused the create.',
    };
    await trackSetup('librarycreate:blocked', 'Archive');
    await openSetup('librarycreate:blocked');
    const blocked = [...nodes['ad-drawer-progress'].querySelectorAll('section')].find((c) =>
      /Archive/.test(c.textContent),
    );
    assert.ok([...blocked.querySelectorAll('button')].find((b) => b.textContent === 'Retry'));
    [...blocked.querySelectorAll('button')]
      .find((b) => b.textContent === 'Open in Monitor')
      .click();
    await document.settle();
    assert.equal(navigations.at(-1).data, 'monitor-' + BUILD);
    assert.doesNotMatch(nodes['ad-drawer-progress'].visibleText, /Open Administration/);
  }
  {
    // A deleted team whose access is already removed is cleared on the next policy read.
    await window.AsxdSites.selectLibrary(id(13));
    policy = {
      Status: 'Applied',
      RowVersion: '20',
      Policy: { Desired: [{ TeamId: id(4), Access: 'Read' }], Applied: [] },
      Teams: [{ TeamId: id(4), Deleted: true, Name: 'Operations' }],
    };
    await timers.shift()();
    assert.doesNotMatch(nodes['ad-team-rows'].visibleText, /Operations/);
    assert.doesNotMatch(nodes['ad-drawer'].visibleText, /Apply access to clear it/);
  }
  {
    // Focus keys (spec 5.2): a redraw gives focus back to the new button with the same key.
    const keyed = (area, key) => area.querySelector('[data-focus-key="' + key + '"]');
    const siteKey = 'site:' + id(1);
    const siteButton = keyed(nodes['ad-sites'], siteKey);
    await pressFocused(siteButton);
    assert(keyed(nodes['ad-sites'], siteKey) !== siteButton, 'The site list was redrawn');
    assert(document.activeElement === keyed(nodes['ad-sites'], siteKey), 'focus');
    // A library's row opens its drawer with focus on the heading; closing it gives focus back to
    // the row's button, though the table was redrawn meanwhile.
    const libraryKey = 'library:' + id(3);
    const libraryButton = keyed(nodes['ad-libraries'], libraryKey);
    await pressFocused(libraryButton);
    assert(keyed(nodes['ad-libraries'], libraryKey) !== libraryButton, 'focus');
    assert(document.activeElement === nodes['ad-library-title'], 'focus');
    await press(nodes['ad-drawer-close']);
    assert.equal(nodes['ad-drawer'].hidden, true);
    assert(document.activeElement === keyed(nodes['ad-libraries'], libraryKey), 'focus');
    await openSetup('librarycreate:blocked');
    const cardKey = 'card:librarycreate:blocked:retry';
    const cardButton = keyed(nodes['ad-drawer-progress'], cardKey);
    cardButton.focus();
    inspectByKey['librarycreate:blocked'] = {
      ...inspectByKey['librarycreate:blocked'],
      Issue: 'SharePoint refused the create again.',
    };
    await timers.shift()();
    assert.match(nodes['ad-drawer-progress'].textContent, /refused the create again/);
    assert(keyed(nodes['ad-drawer-progress'], cardKey) !== cardButton, 'focus');
    assert(document.activeElement === keyed(nodes['ad-drawer-progress'], cardKey), 'focus');
  }
  {
    // No shared-connection section on this tab; no hint paragraphs.
    assert.equal(nodes['ad-manage-connection'], null);
    assert.doesNotMatch(
      nodes.access.visibleText,
      /Shared connection|Site owners retain administrative access|Adding a library validates access/,
    );
    // Team labels (kept text #6).
    const labels = [...nodes['ad-team-choice'].options].map((o) => o.textContent);
    assert(labels.includes('Finance (Entra group)'));
    assert(labels.includes('Project Y (Microsoft 365 group · members + guests)'));
    assert(labels.includes('Finance owners (Entra group · all members)'));
    assert(labels.includes('Project X (Microsoft 365 group · owners)'));
    assert(labels.includes('Operations'));
  }
  // The blocks below run on a site with two libraries: General, and Contracts.
  lib2.asx_libraryid = id(13);
  lib2.asx_name = 'Contracts';
  {
    // Rail: each site with its library count; "＋ Add" opens Add site.
    const rows = nodes['ad-sites'].querySelectorAll('.ad-site');
    assert.equal(rows[0].querySelector('strong').textContent, 'Delivery');
    assert.equal(rows[0].querySelector('.sub').textContent, '2 libraries');
    assert.equal(rows[0].getAttribute('aria-current'), 'true');
    // The counts come from all active libraries, read 5,000 at a time as other full reads are.
    assert(countSizes.length > 0 && countSizes.every((n) => n === 5000), String(countSizes));
    assert.equal(nodes['ad-add-site'].textContent, '＋ Add');
  }
  {
    // Libraries table: Used by follows the server's in-use rule (a Draft revision, or the
    // template's published one); Teams counts the listed teams; Access is a dot and a label.
    destinations = [
      { _asx_libraryid_value: lib.asx_libraryid, _asx_revisionid_value: 'rev-pub' },
      { _asx_libraryid_value: lib.asx_libraryid, _asx_revisionid_value: 'rev-old' },
      { _asx_libraryid_value: lib.asx_libraryid, _asx_revisionid_value: 'rev-draft' },
    ];
    revisions = [
      { asx_revisionid: 'rev-pub', asx_status: 'Published', _asx_templateid_value: 'tpl-1' },
      { asx_revisionid: 'rev-old', asx_status: 'Published', _asx_templateid_value: 'tpl-2' },
      { asx_revisionid: 'rev-draft', asx_status: 'Draft', _asx_templateid_value: 'tpl-3' },
    ];
    templateRows = [
      { asx_templateid: 'tpl-1', _asx_publishedrevisionid_value: 'rev-pub' },
      { asx_templateid: 'tpl-2', _asx_publishedrevisionid_value: 'rev-newer' },
      { asx_templateid: 'tpl-3', _asx_publishedrevisionid_value: null },
    ];
    policy = {
      Status: 'Applied',
      RowVersion: '1',
      Policy: {
        Desired: [{ TeamId: id(4), Access: 'Contribute' }],
        Applied: [{ TeamId: id(4), Access: 'Contribute' }],
      },
      // Operations, marked deleted by an earlier block, exists again.
      Teams: [{ TeamId: id(4), Name: 'Operations', Deleted: false }],
    };
    listedLibraries = [lib, lib2];
    await press(nodes['ad-sites'].querySelector('.ad-site'));
    const general = rowNamed('General');
    assert.equal(
      general.querySelector('.used-by').textContent,
      '2 templates',
      'A replaced revision does not count',
    );
    assert.equal(general.querySelector('.teams').textContent, '1 team');
    assert.equal(general.querySelector('.access').visibleText, 'Access applied');
    assert.equal(general.querySelector('.access .dot').dataset.tone, 'ok');
    assert.equal(rowNamed('Contracts').querySelector('.used-by').textContent, 'Not used');
    assert.equal(nodes['ad-drawer'].hidden, true, 'Choosing a site opens no drawer');
  }
  {
    // The drawer: a dialog with focus on its heading. A changed row is tinted and says what it
    // was; the footer and the table count the change; Discard resets it.
    await pressFocused(libraryButton('General'));
    assert.equal(nodes['ad-drawer'].hidden, false);
    assert.equal(nodes['ad-drawer'].getAttribute('role'), 'dialog');
    assert.equal(document.activeElement, nodes['ad-library-title']);
    assert.equal(
      nodes['ad-drawer'].querySelector('.drawer-sub').textContent,
      'Team access applies to every folder in this library',
    );
    const select = nodes['ad-team-rows'].querySelector('select');
    select.value = 'Read';
    select.onchange();
    await document.settle();
    const changed = nodes['ad-team-rows'].querySelector('.team-row');
    assert.ok(changed.classList.contains('is-changed'));
    assert.match(changed.querySelector('.sub').textContent, / · was Contribute$/);
    assert.equal(nodes['ad-change-status'].textContent, '1 change not applied');
    assert.equal(rowNamed('General').querySelector('.access').visibleText, '1 change not applied');
    await press(nodes['ad-discard']);
    assert.equal(nodes['ad-team-rows'].querySelector('select').value, 'Contribute');
    assert.equal(nodes['ad-change-status'].textContent, '');
    assert.equal(nodes['ad-apply'].textContent, 'Apply access');
  }
  {
    // Escape in a drawer confirmation answers it; Escape on the heading closes the drawer, and
    // focus returns to the library's row button even after the table was redrawn. (Here the
    // confirmation removes itself before the event bubbles, so the drawer staying open is not
    // the guard's proof: the browser flow is.)
    await press(nodes['ad-library-menu']);
    await press(menuItem('ad-library-menu-list', 'Remove library'));
    document.activeElement.key('Escape');
    await document.settle();
    assert.equal(nodes['ad-drawer'].hidden, false);
    assert.equal(nodes['ad-drawer'].querySelector('.confirm'), null);
    window.AsxdSites.redraw();
    nodes['ad-library-title'].focus();
    nodes['ad-library-title'].key('Escape');
    assert.equal(nodes['ad-drawer'].hidden, true);
    assert.equal(document.activeElement.dataset.focusKey, 'library:' + lib.asx_libraryid);
  }
  {
    // A library being set up: a four-segment bar under its name, its stage in Access, and its
    // row opens the stage card in the drawer.
    inspectByKey['librarycreate:board'] = { Key: 'librarycreate:board', Status: 'AccessPending' };
    await trackSetup('librarycreate:board', 'Board papers');
    const setup = rowNamed('Board papers');
    assert.equal(setup.querySelectorAll('.mini-progress span').length, 4);
    assert.equal(setup.querySelectorAll('.mini-progress .done').length, 2);
    assert.match(setup.querySelector('.access').visibleText, / · step \d of 4$/);
    assert.equal(setup.querySelector('.used-by').textContent, 'Not used');
    await press(setup.querySelector('button'));
    assert.ok(nodes['ad-drawer-progress'].querySelector('.ad-progress-card'));
    assert.equal(nodes['ad-drawer-access'].hidden, true, 'A setup drawer shows only its stage');
  }
  {
    // A setup that completes with its drawer open becomes a library row; closing the drawer then
    // gives focus to the libraries, since the row that opened it is gone.
    inspectByKey['librarycreate:board'] = {
      Key: 'librarycreate:board',
      Status: 'Ready',
      CatalogId: lib2.asx_libraryid,
    };
    await timers.shift()();
    assert.equal(
      nodes['ad-libraries'].querySelector('[data-focus-key="setup:librarycreate:board"]'),
      null,
    );
    assert.equal(nodes['ad-drawer'].hidden, false);
    nodes['ad-library-title'].focus();
    nodes['ad-library-title'].key('Escape');
    await document.settle();
    assert.equal(nodes['ad-drawer'].hidden, true);
    assert.equal(document.activeElement.id, 'ad-libraries-heading');
    delete inspectByKey['librarycreate:board'];
  }
  {
    // Opening a site reads each library's access four at a time, and reads no team by ID and
    // refreshes no catalog: the drawer does that for the library it opens.
    const extra = [60, 61, 62, 63].map((n) => ({
      asx_libraryid: id(n),
      asx_name: 'Extra ' + n,
      _asx_siteid_value: id(1),
      asx_approved: true,
      asx_policyapplied: true,
    }));
    listedLibraries = [lib, lib2, ...extra];
    policy = {
      Status: 'Applied',
      RowVersion: '64',
      Policy: {
        Desired: [{ TeamId: id(64), Access: 'Read' }],
        Applied: [{ TeamId: id(64), Access: 'Read' }],
      },
    };
    const teamReads = teamReadIds.length,
      refreshes = refresh;
    slowPolicies = true;
    maxPolicies = 0;
    await press(nodes['ad-sites'].querySelector('.ad-site'));
    slowPolicies = false;
    assert.equal(rowNamed('Extra 63').querySelector('.teams').textContent, '1 team');
    assert.equal(teamReadIds.length, teamReads, 'Opening a site reads no team by ID');
    assert.equal(refresh, refreshes, 'Opening a site refreshes no catalog');
    assert(maxPolicies > 1 && maxPolicies <= 4, 'GetPolicy reads in flight: ' + maxPolicies);
    listedLibraries = [lib, lib2];
  }
  {
    // A stopped access run waits for the admin: it is not read again every 5 seconds unless its
    // library's drawer is open.
    policy = {
      Status: 'Queued',
      RowVersion: '65',
      Policy: {
        Desired: [{ TeamId: id(4), Access: 'Read' }],
        Applied: [],
        OperationKey: 'policywork:stopped',
      },
      RunStatus: 'Blocked',
      RunNotice: 'SharePoint refused the write (HTTP 403).',
    };
    await press(nodes['ad-sites'].querySelector('.ad-site'));
    assert.equal(nodes['ad-drawer'].hidden, true);
    assert.equal(accessOf('General'), 'Needs attention');
    const reads = () => requests.filter((r) => r.Command === 'GetPolicy').length,
      before = reads();
    await timers.shift()();
    assert.equal(reads(), before, 'A stopped run of a closed library is not polled');
  }
  {
    // A link to a library on another site opens that site and the library's drawer, without
    // opening the first site on the way.
    const other = {
        asx_siteid: id(70),
        asx_name: 'Archive site',
        asx_approved: true,
        _asx_nativeid_value: id(2),
      },
      archived = {
        asx_libraryid: id(71),
        asx_name: 'Old papers',
        _asx_siteid_value: id(70),
        asx_approved: true,
        asx_policyapplied: true,
        statecode: 0,
      },
      reads = [];
    const linkedXrm = {
      Utility: xrm.Utility,
      Navigation: xrm.Navigation,
      WebApi: {
        retrieveMultipleRecords: async (table, options) => {
          reads.push(table + options);
          return {
            entities:
              table === 'asx_site'
                ? [site, other]
                : table === 'asx_library' && options.includes('_asx_siteid_value eq ' + id(70))
                  ? [archived]
                  : [],
          };
        },
        retrieveRecord: async (table) =>
          table === 'asx_library' ? archived : table === 'asx_site' ? other : { name: 'Team' },
        online: {
          execute: async (req) => ({
            ok: true,
            json: async () => ({
              Result: JSON.stringify(
                JSON.parse(req.Request).Command === 'GetPolicy'
                  ? { Status: 'Applied', RowVersion: '1', Policy: { Desired: [], Applied: [] } }
                  : { Status: 'Pending' },
              ),
            }),
          }),
        },
      },
    };
    const linked = boot(linkedXrm, {
      timers: [],
      uuid: id(72),
      link: { tab: 'access', library: id(71) },
    });
    await linked.document.fire('DOMContentLoaded');
    assert.equal(linked.nodes['ad-site-title'].textContent, 'Archive site');
    assert.equal(linked.nodes['ad-library-title'].textContent, 'Old papers');
    assert.equal(linked.nodes['ad-drawer'].hidden, false);
    assert(
      !reads.some((r) => r.includes('_asx_siteid_value eq ' + id(1))),
      'The first site is not opened on the way',
    );
  }
  {
    // Add site opens as a side panel; Cancel closes it and returns focus to "＋ Add".
    await pressFocused(nodes['ad-add-site']);
    assert.equal(nodes['ad-site-form'].hidden, false);
    assert.equal(nodes['ad-site-form'].getAttribute('role'), 'dialog');
    assert.equal(nodes['ad-drawer'].hidden, true, 'One panel at a time');
    await press(nodes['ad-cancel-site']);
    assert.equal(nodes['ad-site-form'].hidden, true);
    assert.equal(document.activeElement, nodes['ad-add-site']);
  }
  {
    await press(nodes['ad-site-problems'].querySelector('.problem-pill'));
    assert.equal(navigations.at(-1).data, 'monitor-' + BUILD);
  }
  console.log(
    'PASS Sites & access on the shell: the sites rail with library counts, the libraries table (Used by from the in-use rule, teams, access with a dot, the four-segment bar of a setup), the access drawer as a dialog (changed rows, the change count, Discard, Escape and focus back to the redrawn row), the Monitor problem count in the header, staging versus apply, removals with Undo and their confirmation, consent in the page, server refusals at the form that failed, the Add site combobox, site and library menus with focus keys, library setups with Documents’ SharePoint check and Open in Monitor, stuck access runs (also the inheritance stop), the access label from the policy and its run, discovery matched by list ID, re-point found again after a reload, removed destinations hidden at once and after a reload, deleted Dataverse teams, and where focus goes after actions that close what was focused; fix round 1: card actions in the drawer report in its footer, library counts read 5,000 a page, opening a site reads access four at a time with no team reads or catalog refresh, a stopped run of a closed library is not polled, focus after a setup that became a library row, and a link to a library on another site. Mocked APIs; connected acceptance pending.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
