'use strict';
// A small DOM for handler tests. It parses index.html into an element tree with attributes and
// supports what the admin scripts call: IDs, simple selectors joined by descendant combinators, attributes, focus, events and
// text. It is not a browser: no layout, no CSS, and events reach the target and its ancestors only.
const VOID = new Set([
  'area',
  'base',
  'br',
  'col',
  'embed',
  'hr',
  'img',
  'input',
  'link',
  'meta',
  'source',
  'track',
  'wbr',
]);
const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', '#39': "'", nbsp: ' ' };
const decode = (text) => text.replace(/&(amp|lt|gt|quot|#39|nbsp);/g, (_, name) => ENTITIES[name]);
const camel = (name) => name.replace(/-([a-z])/g, (_, c) => c.toUpperCase());
// Event-loop turns in a row with no tracked work settling or starting, after which settle()
// treats the page as idle. Mocked APIs answer in microtasks, which all run inside one turn.
const QUIET = 5;

class FakeEvent {
  constructor(type, init = {}) {
    Object.assign(this, init);
    this.type = type;
    this.defaultPrevented = false;
    this.propagationStopped = false;
  }
  preventDefault() {
    this.defaultPrevented = true;
  }
  stopPropagation() {
    this.propagationStopped = true;
  }
}

// One compound selector: tag, #id, .class, [attr], [attr=value], [attr="value"], :not([...]).
function compoundParts(selector) {
  const parts =
    selector.match(/^[a-z0-9*]+|#[\w-]+|\.[\w-]+|\[[^\]]+\]|:not\(\[[^\]]+\]\)/gi) || [];
  if (parts.join('') !== selector) throw new Error('Unsupported selector: ' + selector);
  return parts;
}
// Splits on descendant combinators (whitespace outside [...] and (...)): 'tbody tr' → ['tbody', 'tr'].
// Child, sibling and other combinators come out as their own part and are refused by compoundParts.
function compounds(selector) {
  const out = [];
  let current = '';
  let depth = 0;
  for (const ch of selector.trim()) {
    if (ch === '[' || ch === '(') depth++;
    else if (ch === ']' || ch === ')') depth--;
    if (depth === 0 && /\s/.test(ch)) {
      if (current) out.push(current);
      current = '';
    } else current += ch;
  }
  if (current) out.push(current);
  return out;
}
// The last compound matches the node; each earlier one matches an ancestor, in order, taking the
// nearest ancestor that matches (correct for descendant-only chains, as in browsers).
function matchSelector(node, selector) {
  const chain = compounds(selector);
  if (!chain.length) throw new Error('Unsupported selector: ' + selector);
  chain.forEach(compoundParts);
  if (!matchCompound(node, chain[chain.length - 1])) return false;
  let at = node.parentNode;
  for (let i = chain.length - 2; i >= 0; i--) {
    while (at && !matchCompound(at, chain[i])) at = at.parentNode;
    if (!at) return false;
    at = at.parentNode;
  }
  return true;
}
function matchCompound(node, selector) {
  if (node.nodeType !== 1) return false;
  return compoundParts(selector).every((part) => {
    if (part.startsWith('#')) return node.id === part.slice(1);
    if (part.startsWith('.')) return node.classList.contains(part.slice(1));
    if (part.startsWith(':not(')) return !matchAttribute(node, part.slice(5, -1));
    if (part.startsWith('[')) return matchAttribute(node, part);
    return part === '*' || node.tagName === part.toUpperCase();
  });
}
function matchAttribute(node, part) {
  const [, name, value] = /^\[([\w-]+)(?:=["']?([^"'\]]*)["']?)?\]$/.exec(part) || [];
  if (!name) throw new Error('Unsupported attribute selector: ' + part);
  if (name === 'hidden') return value === undefined ? node.hidden : false;
  if (name === 'disabled' && value === undefined) return node.disabled;
  const actual = node.getAttribute(name);
  return value === undefined ? actual !== null : actual === value;
}

// What querySelectorAll, children and a select's options return in a browser: indexed, iterable
// and array-like, but without Array methods, so a script that calls one on them fails here too.
class FakeHTMLCollection {
  constructor(nodes) {
    nodes.forEach((node, i) => (this[i] = node));
    Object.defineProperty(this, 'length', { value: nodes.length });
  }
  item(index) {
    return this[index] ?? null;
  }
  *[Symbol.iterator]() {
    for (let i = 0; i < this.length; i++) yield this[i];
  }
}
// A NodeList adds forEach, entries, keys and values to that.
class FakeNodeList extends FakeHTMLCollection {
  forEach(fn, thisArg) {
    for (let i = 0; i < this.length; i++) fn.call(thisArg, this[i], i, this);
  }
  *entries() {
    for (let i = 0; i < this.length; i++) yield [i, this[i]];
  }
  *keys() {
    for (let i = 0; i < this.length; i++) yield i;
  }
  *values() {
    yield* this;
  }
}

class FakeElement {
  constructor(document, tag) {
    this.ownerDocument = document;
    this.tagName = tag.toUpperCase();
    this.nodeType = tag === '#text' ? 3 : 1;
    // Every child, text nodes included; children and childNodes are read-only views of it.
    this._nodes = [];
    this.parentNode = null;
    this.attributes = new Map();
    this.dataset = {};
    this.style = {};
    this.listeners = new Map();
    this._text = '';
    this.value = '';
    this.checked = false;
    this.disabled = false;
    this.selected = false;
    this.multiple = false;
    this.hidden = false;
    this.open = false;
  }
  get id() {
    return this.getAttribute('id') || '';
  }
  set id(value) {
    this.setAttribute('id', value);
  }
  get className() {
    return this.getAttribute('class') || '';
  }
  set className(value) {
    this.setAttribute('class', value);
  }
  get type() {
    return this.getAttribute('type') || '';
  }
  set type(value) {
    this.setAttribute('type', value);
  }
  get tabIndex() {
    const value = this.getAttribute('tabindex');
    return value === null ? -1 : Number(value);
  }
  set tabIndex(value) {
    this.setAttribute('tabindex', String(value));
  }
  get classList() {
    const names = () => this.className.split(/\s+/).filter(Boolean);
    return {
      contains: (name) => names().includes(name),
      add: (...add) => (this.className = [...new Set([...names(), ...add])].join(' ')),
      remove: (...drop) =>
        (this.className = names()
          .filter((n) => !drop.includes(n))
          .join(' ')),
      toggle: (name, on = !names().includes(name)) => {
        if (on) this.classList.add(name);
        else this.classList.remove(name);
        return on;
      },
    };
  }
  setAttribute(name, value) {
    this.attributes.set(name, String(value));
    if (name.startsWith('data-')) this.dataset[camel(name.slice(5))] = String(value);
  }
  getAttribute(name) {
    if (name.startsWith('data-') && this.dataset[camel(name.slice(5))] !== undefined)
      return String(this.dataset[camel(name.slice(5))]);
    return this.attributes.has(name) ? this.attributes.get(name) : null;
  }
  hasAttribute(name) {
    return this.getAttribute(name) !== null;
  }
  removeAttribute(name) {
    this.attributes.delete(name);
    if (name.startsWith('data-')) delete this.dataset[camel(name.slice(5))];
  }
  get textContent() {
    return this._text + this._nodes.map((c) => c.textContent).join('');
  }
  set textContent(value) {
    this.replaceChildren();
    this._text = String(value ?? '');
  }
  // The text a person sees: hidden subtrees and closed <details> bodies are left out.
  get visibleText() {
    if (this.hidden) return '';
    const kids =
      this.tagName === 'DETAILS' && !this.open
        ? this._nodes.filter((c) => c.tagName === 'SUMMARY')
        : this._nodes;
    return this._text + kids.map((c) => c.visibleText).join('');
  }
  get isConnected() {
    return this.ownerDocument.documentElement.contains(this);
  }
  append(...nodes) {
    for (const node of nodes) this.insert(node, this._nodes.length);
  }
  prepend(...nodes) {
    nodes.forEach((node, i) => this.insert(node, i));
  }
  insert(node, index) {
    const child = typeof node === 'string' ? this.ownerDocument.createTextNode(node) : node;
    child.remove();
    child.parentNode = this;
    this._nodes.splice(Math.min(index, this._nodes.length), 0, child);
  }
  replaceChildren(...nodes) {
    for (const child of [...this._nodes]) child.remove();
    this._text = '';
    this.append(...nodes);
  }
  remove() {
    if (!this.parentNode) return;
    const siblings = this.parentNode._nodes;
    siblings.splice(siblings.indexOf(this), 1);
    this.parentNode = null;
  }
  after(...nodes) {
    const parent = this.parentNode;
    if (!parent) return;
    let index = parent._nodes.indexOf(this) + 1;
    for (const node of nodes) parent.insert(node, index++);
  }
  before(...nodes) {
    const parent = this.parentNode;
    if (!parent) return;
    let index = parent._nodes.indexOf(this);
    for (const node of nodes) parent.insert(node, index++);
  }
  contains(node) {
    for (let n = node; n; n = n.parentNode) if (n === this) return true;
    return false;
  }
  get children() {
    return new FakeHTMLCollection(this._nodes.filter((c) => c.nodeType === 1));
  }
  get childNodes() {
    return new FakeNodeList(this._nodes);
  }
  get firstElementChild() {
    return this._nodes.find((c) => c.nodeType === 1) || null;
  }
  get lastElementChild() {
    return [...this._nodes].reverse().find((c) => c.nodeType === 1) || null;
  }
  get nextElementSibling() {
    const siblings = this.parentNode?._nodes || [];
    return siblings.slice(siblings.indexOf(this) + 1).find((c) => c.nodeType === 1) || null;
  }
  get options() {
    return new FakeHTMLCollection(this.descendants().filter((n) => n.tagName === 'OPTION'));
  }
  get selectedOptions() {
    return new FakeHTMLCollection([...this.options].filter((o) => o.selected));
  }
  descendants() {
    return this._nodes.flatMap((c) => [c, ...c.descendants()]);
  }
  matches(selector) {
    return selector.split(',').some((s) => matchSelector(this, s.trim()));
  }
  querySelectorAll(selector) {
    return new FakeNodeList(
      this.descendants().filter((n) => n.nodeType === 1 && n.matches(selector)),
    );
  }
  querySelector(selector) {
    return this.querySelectorAll(selector)[0] || null;
  }
  closest(selector) {
    for (let n = this; n && n.nodeType === 1; n = n.parentNode) if (n.matches(selector)) return n;
    return null;
  }
  addEventListener(type, fn) {
    if (!this.listeners.has(type)) this.listeners.set(type, []);
    this.listeners.get(type).push(fn);
  }
  removeEventListener(type, fn) {
    this.listeners.set(
      type,
      (this.listeners.get(type) || []).filter((f) => f !== fn),
    );
  }
  dispatchEvent(event) {
    event.target = event.target || this;
    for (let n = this; n && !event.propagationStopped; n = n.parentNode) {
      event.currentTarget = n;
      const results = [...(n.listeners.get(event.type) || [])].map((fn) => fn.call(n, event));
      const handler = n['on' + event.type];
      if (typeof handler === 'function') results.push(handler.call(n, event));
      results.forEach((r) => this.ownerDocument.track(r));
    }
    return !event.defaultPrevented;
  }
  click() {
    if (!this.disabled) this.dispatchEvent(new FakeEvent('click'));
  }
  key(key, init = {}) {
    return this.dispatchEvent(new FakeEvent('keydown', { key, ...init }));
  }
  focus() {
    this.ownerDocument.activeElement = this;
  }
  blur() {
    if (this.ownerDocument.activeElement === this)
      this.ownerDocument.activeElement = this.ownerDocument.body;
  }
  scrollIntoView() {}
  getBoundingClientRect() {
    return { x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 };
  }
}

function parse(document, html) {
  const root = new FakeElement(document, 'document');
  const stack = [root];
  const pattern =
    /<!--[\s\S]*?-->|<!doctype[^>]*>|<\/([a-zA-Z][\w-]*)\s*>|<([a-zA-Z][\w-]*)((?:\s+[^\s=>/]+(?:\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+))?)*)\s*(\/?)>|([^<]+)/gi;
  for (const [, close, open, attrs, selfClose, text] of html.matchAll(pattern)) {
    const top = stack[stack.length - 1];
    if (close) {
      const at = stack.map((n) => n.tagName).lastIndexOf(close.toUpperCase());
      if (at > 0) stack.length = at;
    } else if (open) {
      const node = new FakeElement(document, open);
      for (const [, name, a, b, c] of (attrs || '').matchAll(
        /([^\s=/]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+)))?/g,
      )) {
        const value = decode(a ?? b ?? c ?? '');
        node.setAttribute(name, value);
        if (['hidden', 'disabled', 'checked', 'selected', 'open', 'multiple'].includes(name))
          node[name] = true;
        if (name === 'value') node.value = value;
      }
      top.append(node);
      if (!selfClose && !VOID.has(open.toLowerCase())) stack.push(node);
    } else if (text && text.trim())
      top.append(document.createTextNode(decode(text.replace(/\s+/g, ' '))));
  }
  return root;
}

function createDocument(html) {
  const document = {
    // Work that events started: each handler's promise, removed once it settles either way.
    pending: new Set(),
    track: (value) => {
      if (!value || typeof value.then !== 'function') return;
      const work = Promise.resolve(value).then(
        () => document.pending.delete(work),
        () => document.pending.delete(work),
      );
      document.pending.add(work);
    },
    listeners: new Map(),
    visibilityState: 'visible',
    createElement: (tag) => new FakeElement(document, tag),
    createElementNS: (ns, tag) =>
      Object.assign(new FakeElement(document, tag), { namespaceURI: ns }),
    createTextNode: (text) =>
      Object.assign(new FakeElement(document, '#text'), { _text: String(text) }),
    getElementById: (id) =>
      document.documentElement.descendants().find((n) => n.nodeType === 1 && n.id === id) || null,
    querySelectorAll: (s) => document.documentElement.querySelectorAll(s),
    querySelector: (s) => document.documentElement.querySelector(s),
    addEventListener: (type, fn) => {
      if (!document.listeners.has(type)) document.listeners.set(type, []);
      document.listeners.get(type).push(fn);
    },
    removeEventListener: (type, fn) =>
      document.listeners.set(
        type,
        (document.listeners.get(type) || []).filter((f) => f !== fn),
      ),
    dispatchEvent: (event) => {
      for (const fn of document.listeners.get(event.type) || []) document.track(fn(event));
      return true;
    },
    // Runs every listener of a document event and waits until the work it started settles.
    fire: async (type) => {
      document.dispatchEvent(new FakeEvent(type));
      await document.settle();
    },
    // Lets the work that events started run until it stops moving. It never awaits a handler
    // outright: a click handler that awaits an in-page confirmation stays pending until the test
    // presses one of the confirmation's buttons, so awaiting it would hang. The event loop turns
    // until QUIET turns in a row settle nothing and start nothing; unsettled work stays in
    // `pending`, and the settle() after the test answers picks it up.
    settle: async () => {
      for (let turn = 0, quiet = 0; turn < 200 && quiet < QUIET; turn++) {
        const before = new Set(document.pending);
        await new Promise(setImmediate);
        const changed =
          before.size !== document.pending.size ||
          [...before].some((w) => !document.pending.has(w));
        quiet = changed ? 0 : quiet + 1;
      }
    },
  };
  document.documentElement = parse(document, html);
  document.body = document.documentElement.querySelector('body') || document.documentElement;
  document.activeElement = document.body;
  return document;
}

module.exports = { createDocument, FakeEvent };
