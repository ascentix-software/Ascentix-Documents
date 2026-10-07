'use strict';
// The fake DOM's own contract: settle() returns while a handler waits on an in-page
// confirmation and resumes after the answer, still waits for work that is moving, and
// descendant selectors match the way a browser matches them.
const assert = require('node:assert/strict');
const { createDocument } = require('./fake-dom.cjs');

// Fails instead of hanging when settle() never returns. The timer is not unref'd: an unref'd
// timer lets Node exit silently with code 0 when settle() waits on a promise that never settles.
const within = (promise, ms, what) => {
  let timer;
  return Promise.race([
    promise,
    new Promise((_, reject) => {
      timer = setTimeout(
        () => reject(new Error(what + ' did not return within ' + ms + ' ms')),
        ms,
      );
    }),
  ]).finally(() => clearTimeout(timer));
};

(async () => {
  {
    // A click whose handler awaits an unanswered confirmation: settle() returns, the work stays
    // tracked, and the settle() after the answer finishes it.
    const d = createDocument(
      '<html><body><button id="go">Delete</button><button id="yes">Delete template</button></body></html>',
    );
    let answer;
    let done = false;
    d.getElementById('go').onclick = async () => {
      const ok = await new Promise((resolve) => (answer = resolve));
      await Promise.resolve();
      done = ok;
    };
    d.getElementById('yes').onclick = () => answer(true);
    d.getElementById('go').click();
    await within(d.settle(), 1000, 'settle() with an open confirmation');
    assert.equal(done, false);
    assert.equal(d.pending.size, 1, 'The waiting handler stays tracked');
    d.getElementById('yes').click();
    await within(d.settle(), 1000, 'settle() after the answer');
    assert.equal(done, true);
    assert.equal(d.pending.size, 0);
  }
  {
    // Work that keeps moving is waited for: a long microtask chain, then a few event-loop turns.
    const d = createDocument('<html><body><button id="go">Go</button></body></html>');
    const steps = [];
    d.getElementById('go').onclick = async () => {
      for (let i = 0; i < 20; i++) await Promise.resolve();
      steps.push('microtasks');
      for (let i = 0; i < 3; i++) await new Promise(setImmediate);
      steps.push('turns');
    };
    d.getElementById('go').click();
    await d.settle();
    assert.deepEqual(steps, ['microtasks', 'turns']);
  }
  {
    // A handler that throws does not break settle(); document events are tracked like clicks.
    const d = createDocument('<html><body><button id="bad">Bad</button></body></html>');
    d.getElementById('bad').onclick = async () => {
      throw new Error('refused');
    };
    d.getElementById('bad').click();
    await d.settle();
    assert.equal(d.pending.size, 0);
    let loaded = false;
    d.addEventListener('DOMContentLoaded', async () => {
      await Promise.resolve();
      loaded = true;
    });
    await d.fire('DOMContentLoaded');
    assert.equal(loaded, true);
  }
  {
    // Descendant selectors, at any depth, in selector lists, in matches() and closest().
    const d = createDocument(
      '<html><body><table id="t"><thead><tr><th>Name</th></tr></thead><tbody id="rows">' +
        '<tr class="row"><td><button class="act">Retry</button></td></tr><tr class="row"><td><span>Done</span></td></tr>' +
        '</tbody></table><div class="actions"><p><span class="note">x</span></p></div></body></html>',
    );
    assert.equal(d.querySelectorAll('tbody tr').length, 2, 'tbody tr skips the header row');
    assert.equal(d.querySelectorAll('tr').length, 3);
    assert.equal(d.getElementById('t').querySelectorAll('tbody tr').length, 2);
    assert.equal(d.querySelectorAll('table .row button.act').length, 1);
    assert.equal(d.querySelectorAll('div span').length, 1, 'Any depth, not only children');
    assert.equal(d.querySelectorAll('thead button').length, 0);
    assert.equal(d.querySelectorAll('tbody tr, div span').length, 3, 'Selector lists combine');
    assert.equal(d.querySelector('button.act').closest('tbody tr').className, 'row');
    assert.equal(d.querySelector('span.note').matches('div p span'), true);
    assert.throws(() => d.querySelectorAll('tbody > tr'), /Unsupported selector/);
    assert.throws(() => d.querySelectorAll('tr + tr'), /Unsupported selector/);
  }
  {
    // querySelectorAll, children, childNodes and a select's options are browser collections:
    // indexed and iterable, with no Array methods.
    const d = createDocument(
      '<html><body><ul id="l">One<li>a</li><li>b</li></ul><select id="s"><option>x</option></select></body></html>',
    );
    const list = d.getElementById('l');
    const items = list.querySelectorAll('li');
    assert.equal(items.length, 2);
    assert.equal(items[1].textContent, 'b');
    assert.equal(items.item(0).textContent, 'a');
    assert.equal(items.item(5), null);
    assert.deepEqual(
      [...items].map((i) => i.textContent),
      ['a', 'b'],
    );
    const seen = [];
    items.forEach((i, n) => seen.push(n + i.textContent));
    assert.deepEqual(seen, ['0a', '1b']);
    assert.deepEqual([...items.keys()], [0, 1]);
    assert.equal([...items.entries()][1][1], items[1]);
    assert.equal([...items.values()][0], items[0]);
    for (const method of ['map', 'filter', 'find', 'findIndex', 'some', 'every', 'slice', 'at'])
      assert.equal(items[method], undefined, 'NodeList has no ' + method);
    assert.equal(list.children.length, 2, 'children leaves out text');
    assert.equal(list.childNodes.length, 3, 'childNodes keeps text');
    assert.equal(list.children.forEach, undefined, 'children is an HTMLCollection');
    assert.equal(list.children.findIndex, undefined);
    assert.equal(d.getElementById('s').options.map, undefined);
    assert.equal(d.getElementById('s').options[0].textContent, 'x');
  }
  console.log(
    'PASS fake DOM: settle() returns while a confirmation is open and resumes after the answer, waits for moving work, survives a throwing handler; descendant selectors and selector lists; collections without Array methods. Harness only.',
  );
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
