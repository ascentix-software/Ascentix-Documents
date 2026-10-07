'use strict';
// Documents tab form script with a mocked Xrm and formContext and fake timers: a new record stays
// hidden until saved, a found location shows Documents, a missing one hides it with a notice and
// polls, the 2-minute timeout, a refused read, named tabs and sections, the exact query, and
// stopping on unload or a new record. Not a browser test.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(
  path.resolve(__dirname, '../../client/form/documents-tab.js'),
  'utf8',
);
const ID = '0a1b2c3d-0000-4000-8000-00000000abcd';
const PENDING =
  "Folders are being created in SharePoint. Documents appears here when they're ready.";
const LATE =
  "Folders aren't ready yet. Reopen this record later, or ask an administrator to check Monitor in Ascentix Documents.";
const QUERY =
  '?$select=sharepointdocumentlocationid&$filter=_regardingobjectid_value eq ' +
  ID +
  ' and statecode eq 0 and statuscode eq 1 and servicetype eq 0 and locationtype eq 0' +
  ' and sitecollectionid ne null&$top=1';

const flush = () => new Promise((resolve) => setImmediate(resolve));

function control(name) {
  return {
    name,
    visible: true,
    setVisible(value) {
      this.visible = value;
    },
  };
}

// A form with a Documents nav item, a Files tab with a Docs section, and a record id (null: new).
// `answer(call)` returns the rows a location read finds, or throws to refuse it.
function load(options = {}) {
  const { id = '{' + ID.toUpperCase() + '}', answer = () => [], nav = true } = options;
  const timers = new Map();
  let clock = 0;
  let nextTimer = 1;
  const listeners = {};
  const calls = [];
  const notices = new Map();
  const postSave = [];
  const navItem = control('navSPDocuments');
  const docs = control('Docs');
  const files = Object.assign(control('Files'), {
    sections: { get: (name) => (name === 'Docs' ? docs : null) },
  });
  const record = { id };
  const formContext = {
    ui: {
      navigation: { items: { get: (name) => (nav && name === 'navSPDocuments' ? navItem : null) } },
      tabs: { get: (name) => (name === 'Files' ? files : null) },
      setFormNotification(message, level, unique) {
        notices.set(unique, { message, level });
        return true;
      },
      clearFormNotification(unique) {
        notices.delete(unique);
        return true;
      },
      getFormType: () => (record.id ? 2 : 1),
    },
    data: {
      entity: {
        getId: () => record.id || '',
        addOnPostSave: (handler) => postSave.push(handler),
        removeOnPostSave: (handler) => {
          const at = postSave.indexOf(handler);
          if (at >= 0) postSave.splice(at, 1);
        },
      },
    },
  };
  const refuse = (name) => () => {
    throw new Error(name + ' is not allowed');
  };
  const sandbox = {
    alert: refuse('alert'),
    confirm: refuse('confirm'),
    prompt: refuse('prompt'),
    setTimeout: (fn, ms) => {
      const handle = nextTimer++;
      timers.set(handle, { fn, at: clock + ms });
      return handle;
    },
    clearTimeout: (handle) => timers.delete(handle),
    addEventListener: (type, fn) => (listeners[type] ||= []).push(fn),
    Xrm: {
      WebApi: {
        retrieveMultipleRecords: async (table, query) => {
          calls.push({ table, query });
          return { entities: answer(calls.length) };
        },
      },
    },
  };
  sandbox.window = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(source, sandbox);
  const onLoad = (...params) =>
    sandbox.AscentixDocuments.Form.onLoad({ getFormContext: () => formContext }, ...params);
  // Runs every timer due within `ms`, letting each read settle.
  const advance = async (ms) => {
    const until = clock + ms;
    for (;;) {
      const due = [...timers].filter(([, t]) => t.at <= until).sort((a, b) => a[1].at - b[1].at);
      if (!due.length) break;
      const [handle, timer] = due[0];
      timers.delete(handle);
      clock = timer.at;
      timer.fn();
      await flush();
    }
    clock = until;
  };
  const saved = async (newId) => {
    record.id = newId;
    await Promise.all(postSave.map((handler) => handler({ getFormContext: () => formContext })));
  };
  const shown = () => [navItem.visible, files.visible, docs.visible];
  const notice = () => notices.get('asx-documents-pending') || null;
  const unload = () => (listeners.pagehide || []).forEach((fn) => fn({}));
  return {
    sandbox,
    onLoad,
    advance,
    saved,
    shown,
    notice,
    notices,
    unload,
    calls,
    timers,
    postSave,
    record,
    navItem,
    files,
    docs,
  };
}

(async () => {
  {
    // The global keeps what it had and gains Form.onLoad.
    const sandbox = { setTimeout, clearTimeout, AscentixDocuments: { Other: 1 } };
    sandbox.window = sandbox;
    vm.createContext(sandbox);
    vm.runInContext(source, sandbox);
    assert.equal(sandbox.AscentixDocuments.Other, 1);
    assert.equal(typeof sandbox.AscentixDocuments.Form.onLoad, 'function');
  }
  {
    // A new record is hidden with no notice and no read, then checked again after save.
    const f = load({ id: null, answer: () => [{ sharepointdocumentlocationid: 'l1' }] });
    await f.onLoad('Files');
    assert.deepEqual(f.shown(), [false, false, true]);
    assert.equal(f.notice(), null);
    assert.equal(f.calls.length, 0);
    assert.equal(f.timers.size, 0);
    assert.equal(f.postSave.length, 1);
    await f.saved('{' + ID + '}');
    assert.equal(f.calls.length, 1);
    assert.deepEqual(f.shown(), [true, true, true]);
    assert.equal(f.notice(), null);
  }
  {
    // A found location shows Documents at once and starts no timer.
    const f = load({ answer: () => [{ sharepointdocumentlocationid: 'l1' }] });
    f.navItem.visible = false;
    await f.onLoad();
    assert.equal(f.navItem.visible, true);
    assert.equal(f.notice(), null);
    assert.equal(f.timers.size, 0);
  }
  {
    // The query is the platform's own, for the record's id without braces, in lower case.
    const f = load();
    await f.onLoad();
    assert.equal(f.calls[0].table, 'sharepointdocumentlocation');
    assert.equal(f.calls[0].query, QUERY);
  }
  {
    // Not found: hidden with the notice, re-checked every 5 seconds, shown once a location exists.
    let ready = false;
    const f = load({ answer: () => (ready ? [{ sharepointdocumentlocationid: 'l1' }] : []) });
    await f.onLoad();
    assert.equal(f.navItem.visible, false);
    assert.deepEqual(f.notice(), { message: PENDING, level: 'INFO' });
    await f.advance(4999);
    assert.equal(f.calls.length, 1);
    await f.advance(1);
    assert.equal(f.calls.length, 2);
    await f.advance(10000);
    assert.equal(f.calls.length, 4);
    assert.equal(f.navItem.visible, false);
    ready = true;
    await f.advance(5000);
    assert.equal(f.calls.length, 5);
    assert.equal(f.navItem.visible, true);
    assert.equal(f.notice(), null);
    assert.equal(f.timers.size, 0);
    await f.advance(60000);
    assert.equal(f.calls.length, 5);
  }
  {
    // After 2 minutes: still hidden, the notice changes, and polling stops.
    const f = load();
    await f.onLoad('Files');
    await f.advance(115000);
    assert.equal(f.calls.length, 24);
    assert.deepEqual(f.notice(), { message: PENDING, level: 'INFO' });
    await f.advance(5000);
    assert.equal(f.calls.length, 25);
    assert.deepEqual(f.shown(), [false, false, true]);
    assert.deepEqual(f.notice(), { message: LATE, level: 'INFO' });
    assert.equal(f.notices.size, 1);
    assert.equal(f.timers.size, 0);
    await f.advance(60000);
    assert.equal(f.calls.length, 25);
  }
  {
    // A refused read shows everything with no notice.
    const f = load({
      answer: () => {
        throw new Error('Principal user is missing prvReadSharePointDocumentLocation privilege.');
      },
    });
    f.navItem.visible = false;
    await f.onLoad('Files');
    assert.deepEqual(f.shown(), [true, true, true]);
    assert.equal(f.notice(), null);
    assert.equal(f.timers.size, 0);
  }
  {
    // A read that fails while polling shows everything and clears the notice.
    const f = load({
      answer: (n) => {
        if (n === 3) throw new Error('Network error');
        return [];
      },
    });
    await f.onLoad();
    await f.advance(10000);
    assert.equal(f.calls.length, 3);
    assert.equal(f.navItem.visible, true);
    assert.equal(f.notice(), null);
    assert.equal(f.timers.size, 0);
  }
  {
    // Named tabs and "tab.section" names hide and show with the nav item; missing names are
    // ignored; one comma-separated parameter works as well as several.
    let ready = false;
    const f = load({ answer: () => (ready ? [{ sharepointdocumentlocationid: 'l1' }] : []) });
    await f.onLoad('Files.Docs, Missing', 'Gone.Docs', 'Files.Nope');
    assert.deepEqual(f.shown(), [false, true, false]);
    ready = true;
    await f.advance(5000);
    assert.deepEqual(f.shown(), [true, true, true]);

    const g = load();
    await g.onLoad('Files', 'Files.Docs');
    assert.deepEqual(g.shown(), [false, false, false]);
  }
  {
    // A form without the nav item still handles its named tabs.
    const f = load({ nav: false });
    await f.onLoad('Files');
    assert.equal(f.files.visible, false);
  }
  {
    // Unload stops polling.
    const f = load();
    await f.onLoad();
    assert.equal(f.timers.size, 1);
    f.unload();
    assert.equal(f.timers.size, 0);
    await f.advance(120000);
    assert.equal(f.calls.length, 1);
  }
  {
    // A different record on the form stops the old poll; a repeated load keeps one post-save
    // handler and one timer.
    const f = load();
    await f.onLoad();
    f.record.id = '{11111111-2222-3333-4444-555555555555}';
    await f.advance(5000);
    assert.equal(f.calls.length, 1);
    assert.equal(f.timers.size, 0);

    const g = load();
    await g.onLoad();
    await g.onLoad();
    assert.equal(g.postSave.length, 1);
    assert.equal(g.timers.size, 1);
    await g.advance(5000);
    assert.equal(g.calls.length, 3);
  }
  {
    // Without the execution context the script does nothing and throws nothing.
    const f = load();
    await f.sandbox.AscentixDocuments.Form.onLoad(undefined);
    assert.equal(f.calls.length, 0);
    assert.equal(f.navItem.visible, true);
  }
  console.log(
    'PASS Documents tab form script: a new record hidden until saved and checked after save, a found location shown, not found hidden with the notice and polled every 5 seconds until found, the 2-minute timeout message, a refused or failed read shown with no notice, named tabs and sections with the nav item (missing names ignored), the exact query, unload and a new record stopping the poll, one post-save handler per form, no dialogs. Mocked Xrm, fake timers.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
