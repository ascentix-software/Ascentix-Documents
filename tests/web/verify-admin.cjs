// Static source contract only: no browser, rendering, HTTP or connected UI execution.
const fs = require('fs'),
  path = require('path'),
  assert = require('assert/strict'),
  root = path.resolve(__dirname, '../..');
const read = (p) => fs.readFileSync(path.join(root, p), 'utf8');
const html = read('client/admin/index.html');
const scripts = ['shell.js', 'admin.js', 'sites-access.js', 'operations.js'];
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
const css = read('client/admin/admin.css');
const h1s = [...html.matchAll(/<h1\b[^>]*\bid="([^"]+)"/g)].map((m) => m[1]);
assert.deepEqual(h1s, [
  'overview-title',
  'editor-title',
  'ad-site-title',
  'monitor-title',
  'settings-title',
]);
assert.equal((html.match(/<h1\b/g) || []).length, h1s.length, 'Every h1 has an id');
assert.doesNotMatch(html, /\brole="tab(list|panel)?"/, 'No in-page tab bar');
assert.doesNotMatch(html, /class="[^"]*\bshell\b|automationChip|monitorBadge|tabPrompt/);
assert.doesNotMatch(css, /#access\b/, 'Sites & access uses the shared controls');
assert.doesNotMatch(html + js, /\bad-primary\b/, 'One primary button class');
const FEEDBACK = [
  'fb-templates',
  'fb-schedule',
  'fb-rerun',
  'fb-access',
  'fb-access-add',
  'fb-access-library',
  'fb-monitor',
  'fb-check',
  'fb-advanced',
  'fb-settings',
  'fb-settings-save',
];
assert.deepEqual(
  [...html.matchAll(/\bid="(fb-[^"]+)"/g)].map((m) => m[1]).sort(),
  [...FEEDBACK].sort(),
  'One feedback line per page header, plus panel and footer lines',
);
assert.doesNotMatch(js, /\.id\s*=\s*'fb-/, 'Feedback lines are page markup, not made by script');
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
const planning = read('src/Ascentix.Documents.Domain/Planning.cs');
const bounds = /public static class Bounds\s*\{([\s\S]*?)\n\}/.exec(planning)[1];
const shellBounds = /const BOUNDS = \{([\s\S]*?)\};/.exec(read('client/admin/shell.js'))?.[1];
assert(shellBounds, 'AsxdUi.BOUNDS is missing from shell.js');
for (const [, name, value] of bounds.matchAll(/public const int (\w+) = (\d+);/g)) {
  const key = name[0].toLowerCase() + name.slice(1);
  const client = new RegExp('\\b' + key + ':\\s*(\\d+)').exec(shellBounds);
  assert(client, 'AsxdUi.BOUNDS lacks ' + key);
  assert.equal(client[1], value, 'AsxdUi.BOUNDS.' + key + ' differs from Bounds.' + name);
}
// Helper-text budget (spec 4.3, 4.4, decision D9): exactly these elements have class "help".
const KEPT = [
  'help-automation-settings',
  'help-record-updates',
  'help-stop-tracking',
  'help-publish',
  'help-include-root',
  'help-folder-access',
];
const helpInHtml = [...html.matchAll(/<p class="help" id="([^"]+)"/g)].map((m) => m[1]);
const helpInJs = [...js.matchAll(/\bhelp\(\s*'([^']+)'/g)].map((m) => m[1]);
assert.deepEqual(
  [...helpInHtml, ...helpInJs].sort(),
  [...KEPT].sort(),
  'Only the kept texts are help',
);
assert.doesNotMatch(html + js, /class(Name)?\s*=\s*["']hint|'hint'/, 'No hint class remains');
// AsxdUi.help itself builds its node with el('p', text, 'help') (shell.js, Task 1). The pattern
// skips exactly that expression (negative lookahead), and the two checks after it keep the
// exemption to that one place: it appears once in all scripts, and it is in shell.js.
const HELPER = "el('p', text, 'help')";
assert.doesNotMatch(
  js,
  /el\((?!'p', text, 'help'\))[^)]*,\s*'help'\)|className\s*=\s*'help'/,
  'help elements come only from AsxdUi.help',
);
assert.equal(js.split(HELPER).length - 1, 1, 'Only AsxdUi.help uses ' + HELPER);
assert(read('client/admin/shell.js').includes(HELPER), 'AsxdUi.help is defined in shell.js');
for (const [, text] of html.matchAll(/placeholder="([^"]*)"/g))
  assert(text.split(/\s+/).length <= 4, 'Placeholder longer than 4 words: ' + text);
assert.doesNotMatch(
  html,
  /<(input|select|textarea|button)[^>]*\stitle=/,
  'No title tooltips on controls',
);
assert.doesNotMatch(js, /\.title\s*=/, 'No title tooltips set from script');
for (const removed of [
  'recoveryRun',
  'recoveryToken',
  'recoveryEvidence',
  'recoveryResponse',
  'asx_RecoverWorker',
])
  assert(!html.includes(removed) && !js.includes(removed), removed);
console.log(
  'PASS admin static contract: ' +
    ids.length +
    ' unique IDs, one h1 per page, no tab bar or banner, shared controls, the feedback budget, sitemap and build, the helper-text budget. Visual/connected QA NOT RUN.',
);
