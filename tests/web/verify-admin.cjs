// Static source contract only: no browser, rendering, HTTP or connected UI execution.
const fs = require('fs'),
  path = require('path'),
  root = path.resolve(__dirname, '../..');
const html = fs.readFileSync(path.join(root, 'client/admin/index.html'), 'utf8');
const js = ['admin.js', 'sites-access.js']
  .map((name) => fs.readFileSync(path.join(root, 'client/admin', name), 'utf8'))
  .join('\n');
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
    ' unique IDs and local resources. Visual/connected QA NOT RUN.',
);
