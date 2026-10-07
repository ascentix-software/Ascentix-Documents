// Static source contract only: no browser, rendering, HTTP or connected UI execution.
const fs = require('fs'),
  path = require('path'),
  assert = require('assert/strict'),
  root = path.resolve(__dirname, '../..');
const read = (p) => fs.readFileSync(path.join(root, p), 'utf8');
const html = read('client/admin/index.html');
const scripts = ['shell.js', 'admin.js', 'sites-access.js'];
const js = scripts.map((name) => read('client/admin/' + name)).join('\n');
const ids = [...html.matchAll(/\bid="([^"]+)"/g)].map((m) => m[1]);
if (new Set(ids).size !== ids.length) throw new Error('Duplicate admin element ID');
for (const match of js.matchAll(/\$\(['"]([^'"]+)['"]\)/g))
  if (!ids.includes(match[1])) throw new Error('Missing admin element: ' + match[1]);
for (const match of html.matchAll(/(?:src|href)="([^"]+)"/g)) {
  if (match[1].startsWith('#')) continue;
  if (
    /^[a-z]+:/i.test(match[1]) ||
    !fs.existsSync(path.join(root, 'client/admin', match[1].split('?')[0]))
  )
    throw new Error('Nonlocal or missing admin resource: ' + match[1]);
}
assert.doesNotMatch(html, /id="status"/, 'The global banner is removed');
assert.doesNotMatch(html, /<footer/, 'The footer is removed');
assert.equal((html.match(/<main\b/g) || []).length, 1, 'Exactly one main');
assert.equal((html.match(/<h1\b/g) || []).length, 1, 'Exactly one h1');
const tabs = [...html.matchAll(/<button[^>]*\brole="tab"[^>]*>/g)].map((m) => m[0]);
assert.deepEqual(
  tabs.map((t) => /\bid="([^"]+)"/.exec(t)[1]),
  ['tab-templates', 'tab-access', 'tab-monitor', 'tab-settings'],
);
for (const tab of tabs) {
  const panel = /aria-controls="([^"]+)"/.exec(tab)[1];
  assert.match(
    html,
    new RegExp('id="' + panel + '"[^>]*role="tabpanel"|role="tabpanel"[^>]*id="' + panel + '"'),
  );
}
assert.doesNotMatch(html, /aria-pressed/);
assert.doesNotMatch(
  js,
  /\binnerHTML\b|outerHTML\s*=|insertAdjacentHTML|document\.write|openConfirmDialog|openAlertDialog|window\.confirm|\balert\(/,
  'textContent only and no browser or platform dialogs',
);
for (const match of html.matchAll(/aria-(?:describedby|labelledby)="([^"]+)"/g))
  for (const id of match[1].split(/\s+/)) assert(ids.includes(id), 'Broken ARIA reference: ' + id);
const build = /const BUILD = '([^']+)'/.exec(read('client/admin/shell.js'))[1];
for (const match of html.matchAll(/\?v=([\w]+)/g))
  assert.equal(match[1], build, 'Cache-buster differs');
for (const file of ['AppModuleSiteMap.xml', 'AppModuleSiteMap_managed.xml']) {
  const map = read('solution/AscentixDocuments/src/AppModuleSiteMaps/asx_DocumentsAdmin/' + file);
  const subs = [
    ...map.matchAll(/<SubArea Id="([^"]+)" Title="([^"]+)" Url="[^"]*data=([a-z]+)-([\w]+)"/g),
  ];
  assert.deepEqual(
    subs.map((s) => [s[1], s[2], s[3], s[4]]),
    [
      ['asx_studio', 'Folder templates', 'templates', build],
      ['asx_policyadmin', 'Sites &amp; access', 'access', build],
      ['asx_operationsadmin', 'Monitor', 'monitor', build],
      ['asx_runtimeadmin', 'Settings', 'settings', build],
    ],
    file,
  );
}
for (const removed of [
  'recoveryPanel',
  'recoveryRun',
  'recoveryToken',
  'recoveryEvidence',
  'recoveryResponse',
  'recoverOperation',
  'asx_RecoverWorker',
  'Open recovery',
])
  if (html.includes(removed) || js.includes(removed))
    throw new Error('Evidence recovery is removed: ' + removed);
console.log(
  'PASS admin static contract: ' +
    ids.length +
    ' unique IDs, tabs, no banner, sitemap and build. Visual/connected QA NOT RUN.',
);
