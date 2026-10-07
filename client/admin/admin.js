'use strict';
// Folder templates: the templates list, the read-first overview of a template with its Schedule
// and Versions panels, and the editor (template bar with its version chip, the folder tree and
// folder settings, the condition builder, the preview of the current edits), and re-runs for
// existing records. Text is only ever set with textContent.
(() => {
  const $ = (id) => document.getElementById(id);
  const state = {
    tables: [],
    root: null,
    lookups: [],
    sources: [],
    sites: [],
    libraries: [],
    sections: [],
    saved: null,
    // The template open in the editor (null for a new one until its first save).
    template: null,
    templates: [],
    // Every template's revisions: the list's states and the overview's Versions read them.
    revisions: [],
    enabledTables: [],
    // 'overview' shows a template read-only; 'edit' shows the editor.
    view: 'overview',
    step: 1,
    // The template the overview shows: its row, the revision it reads, that revision's sources and
    // destinations, its last re-run and the policies of its libraries.
    overview: null,
    // One GetPolicy per library per page load, shared by every reader.
    policyCache: new Map(),
    listQuery: '',
    // The overview's Folders groups that are open, by destination key.
    openGroups: new Set(),
    // An earlier version opened from Versions: it can be read, not saved or published.
    readOnly: false,
    editBase: null,
    busy: false,
    // The list's first load is done: its empty state is real, not "still loading".
    loaded: false,
    // Edits since the last load or save: autosave, Publish and the guards follow it.
    unsaved: false,
    // Autosave: 'idle', 'saving', 'saved' or 'error', when it last saved, why it failed, the save
    // in flight, whether an edit made during it needs one more save, the number of edits so far,
    // and the pause before the next save.
    saveState: 'idle',
    savedAt: null,
    saveError: null,
    saving: null,
    saveAgain: false,
    edits: 0,
    saveTimer: null,
    // Which open editor work belongs to: a save that answers after another template opened
    // leaves the new one alone.
    session: 0,
    // Publish is under way: autosave waits for it.
    publishing: false,
    // The published revision the draft is compared with ({ version, sources, sections }), and
    // what changed since it (changesSince).
    publishedSnapshot: null,
    changes: null,
    // The destination open in step 1, by key.
    expanded: null,
    // The last autosave found invalid conditions.
    invalid: false,
    run: null,
    runTotal: null,
    batch: null,
    previewRecord: null,
    previewed: false,
    // The data-focus-key that takes focus after the next render.
    focusKey: null,
    // Related tables whose fields could not be read, even after a retry.
    failedTables: new Set(),
    // Which template open the background preload belongs to.
    preloadId: 0,
    pickersStale: false,
    // A list redraw waits for the list's open confirmation.
    railStale: false,
  };
  const OPERATOR = 'prvCreateasx_operatorcommand';
  const READ_ONLY = 'Viewing an earlier version.';
  // The states of a re-run that has not ended.
  const ACTIVE = ['Running', 'Waiting', 'Retrying', 'Paused', 'Blocked'];
  // The open side panel (Schedule, Version history or Re-run): one at a time.
  let panel = null;
  // Dataverse IDs compare without case.
  const same = (a, b) => !!a && !!b && String(a).toLowerCase() === String(b).toLowerCase();
  const metadata = new Map();
  const xrm = window.parent?.Xrm || window.Xrm;
  const ui = window.AsxdUi;
  // Results of an action show in the feedback line under the header of the page that ran it.
  const message = (text, error = false) =>
    ui.feedback(ui.activeTab(), text, error ? 'error' : 'success');
  const el = (tag, text, css) => {
    const node = document.createElement(tag);
    if (text != null) node.textContent = text;
    if (css) node.className = css;
    return node;
  };
  const option = (value, label) => {
    const o = el('option', label);
    o.value = value;
    return o;
  };
  function select(items, value, change) {
    const s = el('select');
    items.forEach((i) => s.append(option(i.value, i.label)));
    s.value = value;
    s.onchange = () => {
      change(s.value);
      dirty();
    };
    return s;
  }
  function label(text, control) {
    const l = el('label', text);
    l.append(control);
    return l;
  }
  function button(text, action, css = 'secondary') {
    const b = el('button', text, css);
    b.type = 'button';
    b.onclick = action;
    return b;
  }
  const keyed = (node, key) => {
    node.dataset.focusKey = key;
    return node;
  };
  const plural = ui.plural;

  // The name the unsaved-changes prompt and Retry use.
  const templateLabel = () => $('templateName').value.trim() || 'this template';
  const changeCount = (step) => (state.readOnly ? 0 : state.changes?.byStep[step] || 0);
  const current = () => ({ sources: state.sources, sections: state.sections });
  // Publish needs the Publisher role and something not yet published; unsaved edits are saved
  // first.
  function publishReason() {
    if (!ui.can('prvCreateasx_publication')) return ui.needs('prvCreateasx_publication');
    if (state.saved && state.saved.Status !== 'Draft' && !state.unsaved) return 'Already published';
    return null;
  }
  // The header's first row: the table, the name (typed in for a new template), the version the
  // draft becomes, which version stays live, and the save status.
  function renderHeader() {
    if (!state.root) return;
    $('editor-table').textContent = tableName(state.root.LogicalName) + ' ›';
    const name = $('templateName');
    const naming = !state.template;
    // A new template's name becomes the heading once it is saved; focus stays on the heading.
    if (!naming && !name.hidden && document.activeElement === name) $('editor-title').focus();
    name.hidden = !naming;
    name.disabled = state.busy;
    $('editor-name').textContent = naming
      ? ''
      : state.template.asx_name || tableName(state.root.LogicalName);
    const base = state.editBase,
      pill = $('editor-pill');
    if (state.readOnly && base) {
      const live = same(base.RevisionId, state.template?._asx_publishedrevisionid_value);
      pill.textContent =
        'v' +
        base.Version +
        ' · ' +
        (live ? 'Live' : base.Status === 'Draft' ? 'Draft' : 'Replaced');
      pill.dataset.tone = 'muted';
    } else {
      pill.textContent =
        'Draft v' + (!base ? 1 : base.Status === 'Draft' ? base.Version : base.Version + 1);
      pill.dataset.tone = 'warning';
    }
    const live = state.readOnly
      ? null
      : (state.publishedSnapshot?.version ??
        (state.template ? templateState(state.template).live : null));
    $('editor-note').hidden = live == null;
    $('editor-note').textContent = live == null ? '' : 'v' + live + ' stays live until you publish';
  }
  const clock = new Intl.DateTimeFormat(undefined, { hour: '2-digit', minute: '2-digit' });
  function renderSaveStatus() {
    const status = $('save-status');
    const failed = state.saveState === 'error';
    status.classList.toggle('is-error', failed);
    if (failed) {
      status.textContent = "Couldn't save · ";
      const retry = button('Retry', () => autosave(), 'link');
      retry.setAttribute('aria-label', 'Retry saving ' + templateLabel());
      status.append(retry);
      return;
    }
    status.textContent =
      state.saveState === 'saving'
        ? 'Saving…'
        : state.saveState === 'saved'
          ? 'Saved ' + clock.format(state.savedAt)
          : '';
  }
  // The stepper: the current step, a dot on a step with changes, and the Folders step marked when
  // its conditions are not valid. The name says it too.
  const STEPS = ['Destinations', 'Folders', 'Review and publish'];
  function renderStepper() {
    STEPS.forEach((label, index) => {
      const n = index + 1,
        tab = $('step-tab-' + n),
        selected = n === state.step;
      tab.setAttribute('aria-selected', String(selected));
      tab.tabIndex = selected ? 0 : -1;
      tab.classList.toggle('is-current', selected);
      const error = n === 2 && state.invalid && !state.readOnly;
      const changed = n < 3 && changeCount(n) > 0;
      tab.classList.toggle('has-error', error);
      tab.querySelector('.step-dot').hidden = error || !changed;
      if (error || changed)
        tab.setAttribute('aria-label', label + (error ? ', has errors' : ', has changes'));
      else tab.removeAttribute('aria-label');
    });
  }
  function renderFooter() {
    $('step-back').hidden = state.step === 1;
    $('step-next').hidden = state.step === 3;
    $('step-next').textContent = state.step === 1 ? 'Next: Folders' : 'Next: Review';
    const n = state.step < 3 ? changeCount(state.step) : 0;
    $('step-count').textContent = n ? plural(n, 'change', 'changes') + ' in this step' : '';
  }
  // Publish and ＋ Add destination follow the draft, the role, the bound and read-only.
  function renderActions() {
    const reason = publishReason();
    const publish = $('publish');
    publish.hidden = state.readOnly || (!state.saved && !state.sections.length);
    ui.disable(publish, 'publish-reason', reason);
    publish.classList.toggle('primary', !reason);
    publish.classList.toggle('secondary', !!reason);
    const add = $('add-destination');
    add.hidden = state.readOnly || !state.libraries.length;
    $('destination-limit-note').hidden = add.hidden;
    $('destination-limit-note').textContent = 'Up to ' + ui.BOUNDS.destinations + ' per template';
    $('no-library').hidden = !!state.libraries.length || state.readOnly;
    ui.disable(
      add,
      'destination-reason',
      state.sections.length >= ui.BOUNDS.destinations
        ? 'A template can have up to ' +
            ui.BOUNDS.destinations +
            ' destinations, because each record plans all of them in one step that Dataverse stops after 2 minutes.'
        : null,
    );
  }
  // Everything around the steps: header, save status, stepper, footer and actions.
  function chrome() {
    if (!state.root) return;
    renderHeader();
    renderSaveStatus();
    renderStepper();
    renderFooter();
    renderActions();
  }
  // Shows a step. With focus, its first heading or control takes focus (Back and Next); a tab
  // keeps the focus it has.
  function goStep(n, focus = true) {
    state.step = n;
    for (const i of [1, 2, 3]) $('step-' + i).hidden = i !== n;
    renderStepper();
    renderFooter();
    if (focus) stepStart(n)?.focus();
  }
  const stepStart = (n) =>
    n === 1
      ? $('destination-list').querySelector('.destination-card input') || $('add-destination')
      : n === 2
        ? $('folders-heading')
        : $('preview-heading');
  // An edit: change tracking, the stepper and footer follow it, the preview dims until refreshed,
  // and the autosave waits for the pause after it.
  function dirty() {
    state.unsaved = true;
    state.edits++;
    state.changes = changesSince(state.publishedSnapshot, current());
    $('preview-stale').hidden = !state.previewed;
    $('previewTrees').classList.toggle('is-stale', state.previewed);
    $('refreshPreview').hidden = !state.previewRecord;
    renderStepper();
    renderFooter();
    renderActions();
    scheduleSave();
  }
  function scheduleSave() {
    if (state.readOnly) return;
    clearTimeout(state.saveTimer);
    state.saveTimer = setTimeout(autosave, 1500);
  }
  // Saves the draft once edits pause. One save at a time: an edit made while one runs is saved
  // by one more save after it, with the row version the first one returned.
  async function autosave() {
    if (state.readOnly || !state.root || state.publishing) return undefined;
    if (state.saving) {
      state.saveAgain = true;
      return state.saving;
    }
    if (!state.unsaved) return undefined;
    // A new template is saved once it has a name and a destination.
    if (!state.template && (!$('templateName').value.trim() || !state.sections.length))
      return undefined;
    state.invalid = !validate({ focus: false });
    renderStepper();
    if (state.invalid) return undefined;
    const edits = state.edits,
      session = state.session;
    state.saveState = 'saving';
    renderSaveStatus();
    const saving = (async () => {
      try {
        let refresh = null;
        try {
          await save();
        } catch (error) {
          // The draft was saved; only the reads after it failed.
          if (!error.saved) throw error;
          refresh = error.message;
        }
        if (session !== state.session) return;
        // Edits made while this save was in flight are still unsaved.
        state.unsaved = state.edits !== edits;
        state.saveState = 'saved';
        state.savedAt = new Date();
        state.saveError = null;
        if (refresh) ui.feedback('editor', refresh, 'error');
        else ui.clearFeedback('editor');
      } catch (error) {
        if (session !== state.session) return;
        state.unsaved = true;
        state.saveState = 'error';
        state.saveError = error.message || String(error);
        state.saveAgain = false;
        ui.feedback('editor', state.saveError, 'error');
      } finally {
        if (state.saving === saving) state.saving = null;
        if (session === state.session) {
          renderSaveStatus();
          renderHeader();
          renderActions();
        }
      }
      if (session === state.session && state.saveAgain && state.saveState === 'saved') {
        state.saveAgain = false;
        await autosave();
      }
    })();
    state.saving = saving;
    return saving;
  }
  // Saves now: Save draft in the unsaved-changes prompt, and Publish. Refuses with the reason
  // when the draft could not be saved.
  async function flushSave() {
    clearTimeout(state.saveTimer);
    await state.saving;
    await autosave();
    await state.saving;
    if (state.invalid) {
      goStep(2, false);
      validate();
      throw new Error('Fix the highlighted conditions first.');
    }
    if (state.saveState === 'error') throw new Error(state.saveError);
  }
  async function task(action) {
    if (state.busy) return;
    state.busy = true;
    chrome();
    try {
      await action();
    } catch (error) {
      message(error.message || String(error), true);
    } finally {
      state.busy = false;
      chrome();
    }
  }
  async function get(path) {
    const response = await fetch(
      xrm.Utility.getGlobalContext().getClientUrl() + '/api/data/v9.2/' + path,
      {
        credentials: 'same-origin',
        headers: { Accept: 'application/json', 'OData-Version': '4.0' },
      },
    );
    const data = await response.json();
    if (!response.ok) throw new Error(data.error?.message || 'Dataverse request failed.');
    return data;
  }
  async function all(path) {
    const data = await get(path);
    if (data['@odata.nextLink'])
      throw new Error('Metadata/catalog exceeded the current completeness bound.');
    return data.value;
  }
  const display = (item) => item.DisplayName?.UserLocalizedLabel?.Label || item.LogicalName;
  const tableInfo = (table) =>
    state.tables.find((t) => t.LogicalName === table) || metadata.get(table)?.info || null;
  const tableName = (table) => display(tableInfo(table) || { LogicalName: table });
  function kind(attribute) {
    if (attribute.IsSecured || attribute.IsValidForRead === false) return null;
    const map = {
      String: 'Text',
      Memo: 'Text',
      Decimal: 'Number',
      Double: 'Number',
      Integer: 'Number',
      BigInt: 'Number',
      Money: 'Number',
      Boolean: 'Boolean',
      Picklist: 'Choice',
      State: 'Choice',
      Status: 'Choice',
      Lookup: 'Lookup',
      Customer: 'Lookup',
      Owner: 'Lookup',
      Uniqueidentifier: 'Lookup',
    };
    return map[attribute.AttributeType] || null;
  }
  // A table's readable columns with their kinds, choice labels and lookup targets. Read once; a
  // read in progress is shared, and a failed one is forgotten so it can be tried again.
  const reading = new Map();
  function fields(table) {
    if (metadata.has(table)) return Promise.resolve(metadata.get(table));
    if (!reading.has(table))
      reading.set(
        table,
        readFields(table).finally(() => reading.delete(table)),
      );
    return reading.get(table);
  }
  async function readFields(table) {
    const base = "EntityDefinitions(LogicalName='" + table + "')/Attributes";
    const [values, dates, choices, singleChoices, lookups, info] = await Promise.all([
      all(base + '?$select=LogicalName,DisplayName,AttributeType,IsSecured,IsValidForRead'),
      all(
        base +
          '/Microsoft.Dynamics.CRM.DateTimeAttributeMetadata?$select=LogicalName,DateTimeBehavior',
      ),
      all(
        base +
          '/Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet',
      ),
      Promise.all(
        ['Picklist', 'State', 'Status'].map((type) =>
          all(
            base +
              '/Microsoft.Dynamics.CRM.' +
              type +
              'AttributeMetadata?$select=LogicalName&$expand=OptionSet',
          ),
        ),
      ).then((lists) => lists.flat()),
      all(
        base +
          '/Microsoft.Dynamics.CRM.LookupAttributeMetadata?$select=LogicalName,DisplayName,Targets,IsSecured,IsValidForRead',
      ),
      // A related table outside the document-enabled list still needs its display name.
      state.tables.some((t) => t.LogicalName === table)
        ? null
        : get(
            "EntityDefinitions(LogicalName='" +
              table +
              "')?$select=LogicalName,DisplayName,PrimaryNameAttribute",
          ).catch(() => null),
    ]);
    const columns = values.map((a) => ({
      Name: a.LogicalName,
      Label: display(a),
      Kind: kind(a),
      secured: a.IsSecured || a.IsValidForRead === false,
    }));
    for (const d of dates) {
      const c = columns.find((c) => c.Name === d.LogicalName);
      if (c && !c.secured)
        c.Kind =
          d.DateTimeBehavior?.Value === 'DateOnly'
            ? 'DateOnly'
            : d.DateTimeBehavior?.Value === 'TimeZoneIndependent'
              ? null
              : 'DateTime';
    }
    for (const d of choices) {
      const c = columns.find((c) => c.Name === d.LogicalName);
      if (c && !c.secured) c.Kind = 'MultiChoice';
    }
    for (const d of [...choices, ...singleChoices]) {
      const c = columns.find((c) => c.Name === d.LogicalName);
      if (!c) continue;
      // Labels only; the value is added only where two labels in one set are the same.
      const options = (d.OptionSet?.Options || []).map((o) => ({
        value: String(o.Value),
        label: o.Label?.UserLocalizedLabel?.Label || String(o.Value),
      }));
      const seen = new Map();
      options.forEach((o) => seen.set(o.label, (seen.get(o.label) || 0) + 1));
      c.Options = options.map((o) =>
        seen.get(o.label) > 1 ? { ...o, label: o.label + ' (' + o.value + ')' } : o,
      );
    }
    const readable = lookups.filter((l) => !l.IsSecured && l.IsValidForRead !== false);
    for (const l of readable) {
      const c = columns.find((c) => c.Name === l.LogicalName);
      if (c) c.Targets = l.Targets;
    }
    const result = {
      columns: columns.filter((c) => c.Kind).sort((a, b) => a.Label.localeCompare(b.Label)),
      lookups: readable,
      info: info?.LogicalName ? info : null,
    };
    metadata.set(table, result);
    return result;
  }
  const deprecated = (c) => c.Label.startsWith('(Deprecated)');
  const lookupFor = (source) =>
    state.lookups.find((l) => l.Lookup === source.Lookup && l.Table === source.Table);
  // How a related source reads before its field: "Primary Contact".
  const via = (source) => lookupFor(source)?.Label.split(' → ')[0] || source.Lookup;
  function availableFields() {
    return state.sources.flatMap((s) =>
      s.columns.map((c) => ({
        value: s.Alias + '.' + c.Name,
        label: (s.Alias === 'root' ? '' : via(s) + ' › ') + c.Label,
        kind: c.Kind,
        options: c.Options,
      })),
    );
  }
  // How a template's field reads, by alias and column: "Account Name" on the record itself,
  // "Primary Contact › Full Name" on a related one; null when its sources do not have it.
  function labelOf(sources) {
    const root = sources.find((s) => s.Alias === 'root');
    const lookups = metadata.get(root?.Table)?.lookups || [];
    return (alias, column) => {
      const source = sources.find((s) => s.Alias === alias);
      const field = source?.columns.find((c) => c.Name === column);
      if (!source || !field) return null;
      if (alias === 'root') return field.Label;
      const lookup = lookups.find((l) => l.LogicalName === source.Lookup);
      return (lookup ? display(lookup) : source.Lookup) + ' › ' + field.Label;
    };
  }
  // A folder name with its fields shown by label: {root.name} → [Account Name].
  function readable(name) {
    const label = labelOf(state.sources);
    return String(name || '').replace(
      /\{([a-z0-9_]+)\.([a-z0-9_]+)\}/gi,
      (match, alias, column) => {
        const text = label(alias, column);
        return text ? '[' + text + ']' : match;
      },
    );
  }
  // Picker values: 'root.<column>' and, for a related record, 'lookup:<lookup>:<table>:<column>'
  // until a field of it is used, when the source is added under its alias.
  function pickerValue(field) {
    const [alias, column] = String(field || '').split('.');
    if (alias === 'root') return field;
    const source = state.sources.find((s) => s.Alias === alias);
    return source ? 'lookup:' + source.Lookup + ':' + source.Table + ':' + column : field;
  }
  function resolveField(value) {
    if (!value.startsWith('lookup:')) return value;
    const [, lookup, table, column] = value.split(':');
    let source = state.sources.find((s) => s.Lookup === lookup && s.Table === table);
    if (!source) {
      if (state.sources.filter((s) => s.Alias !== 'root').length >= ui.BOUNDS.relatedRecords)
        throw new Error(relatedReason());
      let i = 1;
      while (state.sources.some((s) => s.Alias === 'lookup_' + i)) i++;
      source = {
        Alias: 'lookup_' + i,
        Table: table,
        Lookup: lookup,
        columns: metadata.get(table)?.columns || [],
      };
      state.sources.push(source);
    }
    return source.Alias + '.' + column;
  }
  const relatedReason = () =>
    'A template can use up to ' +
    ui.BOUNDS.relatedRecords +
    ' related records, because Documents tracks each one so it can re-run records when it changes.';
  // One picker: this record's fields, then one group per related record. The primary name comes
  // first and deprecated fields last. At the related-record bound, the groups of related records
  // not used yet are disabled.
  function fieldSelect(keep) {
    const picker = el('select');
    const order = (columns, primary) =>
      columns
        .filter(keep)
        .sort(
          (a, b) =>
            deprecated(a) - deprecated(b) ||
            (a.Name === primary ? -1 : b.Name === primary ? 1 : a.Label.localeCompare(b.Label)),
        );
    const group = (text, columns, valueOf, disabled = false) => {
      const g = el('optgroup');
      g.setAttribute('label', text);
      g.disabled = disabled;
      g.append(...columns.map((c) => option(valueOf(c), c.Label)));
      picker.append(g);
    };
    const root = state.sources.find((s) => s.Alias === 'root');
    group('This record', order(root.columns, state.root.PrimaryNameAttribute), (c) => {
      return 'root.' + c.Name;
    });
    const full = state.sources.filter((s) => s.Alias !== 'root').length >= ui.BOUNDS.relatedRecords;
    const listed = new Set();
    for (const lookup of state.lookups) {
      const meta = metadata.get(lookup.Table);
      if (!meta) {
        // Still loading in the background, or failed twice: shown, never silently left out.
        const g = el('optgroup');
        g.setAttribute('label', lookup.Label);
        g.disabled = true;
        const note = option(
          '',
          state.failedTables.has(lookup.Table) ? "Couldn't load these fields" : 'Loading fields…',
        );
        note.disabled = true;
        g.append(note);
        picker.append(g);
        listed.add(lookup.Lookup + ':' + lookup.Table);
        continue;
      }
      const used = state.sources.some(
        (s) => s.Lookup === lookup.Lookup && s.Table === lookup.Table,
      );
      listed.add(lookup.Lookup + ':' + lookup.Table);
      group(
        lookup.Label,
        order(meta.columns, tableInfo(lookup.Table)?.PrimaryNameAttribute),
        (c) => 'lookup:' + lookup.Lookup + ':' + lookup.Table + ':' + c.Name,
        full && !used,
      );
    }
    // A related record the draft uses whose lookup is no longer offered keeps its fields.
    for (const source of state.sources)
      if (source.Alias !== 'root' && !listed.has(source.Lookup + ':' + source.Table))
        group(
          source.Lookup + ' → ' + tableName(source.Table),
          order(source.columns, null),
          (c) => 'lookup:' + source.Lookup + ':' + source.Table + ':' + c.Name,
        );
    return { picker, full };
  }
  const NAME_KINDS = ['Text', 'Number', 'Choice', 'DateOnly'];
  function fieldPicker(into) {
    const { picker, full } = fieldSelect((c) => NAME_KINDS.includes(c.Kind));
    picker.setAttribute('aria-label', 'Insert field');
    keyed(picker, 'edit:insert-field');
    picker.value = 'root.' + state.root.PrimaryNameAttribute;
    into.append(picker);
    if (full) {
      // The reason names why some groups are off; the picker itself stays usable.
      const note = el('p', relatedReason(), 'reason');
      note.id = 'related-reason';
      picker.setAttribute('aria-describedby', 'related-reason');
      into.after(note);
    }
    return picker;
  }
  const operators = (kind) =>
    ['Equal', 'NotEqual', 'IsNull', 'IsNotNull'].concat(
      ['Text', 'MultiChoice'].includes(kind)
        ? ['Contains', 'DoesNotContain']
        : ['Number', 'DateOnly', 'DateTime'].includes(kind)
          ? ['Greater', 'GreaterOrEqual', 'Less', 'LessOrEqual']
          : [],
    );
  const operatorLabel = {
    Equal: 'equals',
    NotEqual: 'does not equal',
    IsNull: 'is empty',
    IsNotNull: 'has a value',
    Contains: 'contains',
    DoesNotContain: 'does not contain',
    Greater: 'greater than',
    GreaterOrEqual: 'at least',
    Less: 'less than',
    LessOrEqual: 'at most',
  };
  const unary = (operator) => ['IsNull', 'IsNotNull'].includes(operator);
  // Operators as rule sentences say them: "Status is Active".
  const SHORT = {
    Equal: 'is',
    NotEqual: 'is not',
    Contains: 'contains',
    DoesNotContain: "doesn't contain",
    IsNull: 'is empty',
    IsNotNull: 'has a value',
    Greater: 'greater than',
    GreaterOrEqual: 'at least',
    Less: 'less than',
    LessOrEqual: 'at most',
  };
  // A date as the version chip shows it: "3 Oct".
  const day = (value) => ui.time(value).textContent.split(',')[0];
  // One condition as words: its field, the operator, and the value by its label (an option's
  // label, Yes or No, the record's name, the date, or the other field).
  function clause(c, sources) {
    const label = labelOf(sources);
    const named = (value) => {
      const [alias, column] = String(value).split('.');
      return label(alias, column) || 'Unavailable field';
    };
    const text = named(c.field) + ' ' + (SHORT[c.Operator] || c.Operator);
    if (unary(c.Operator)) return text;
    if (c.right) return text + ' ' + named(c.right);
    const [alias, column] = String(c.field).split('.');
    const field = sources.find((s) => s.Alias === alias)?.columns.find((f) => f.Name === column);
    const option = (value) => field.Options?.find((o) => o.value === value)?.label ?? value;
    switch (field?.Kind) {
      case 'Choice':
        return text + ' ' + option(c.Literal);
      case 'MultiChoice':
        return text + ' ' + String(c.Literal).split(',').map(option).join(', ');
      case 'Boolean':
        return text + ' ' + (c.Literal === 'false' ? 'No' : 'Yes');
      case 'Lookup':
        return text + ' ' + (c.label || 'Record not available');
      case 'DateOnly':
        // A date without a time is read in local time, so it is the same day everywhere.
        return text + ' ' + (c.Literal ? day(c.Literal + 'T00:00') : '');
      case 'DateTime':
        return text + ' ' + (c.Literal ? day(c.Literal) : '');
      default:
        return text + ' ' + c.Literal;
    }
  }
  // A folder's rule: "Always", or "◆ " and its rule sentence.
  function ruleText(folder, sources) {
    const sentence = folder.Condition
      ? ui.ruleSentence(folder.Condition, (c) => clause(c, sources))
      : '';
    return sentence
      ? { text: '◆ ' + sentence, conditional: true }
      : { text: 'Always', conditional: false };
  }
  // The targets of the lookup a condition compares; a record ID column is its own table.
  function lookupTargets(field) {
    const [alias, column] = String(field).split('.');
    const source = state.sources.find((s) => s.Alias === alias);
    const found = source?.columns.find((c) => c.Name === column);
    return found?.Targets || [source?.Table || state.root.LogicalName];
  }
  // An ISO time as YYYY-MM-DDTHH:mm in local time, for datetime-local inputs.
  function localTime(value) {
    if (!value) return '';
    const d = new Date(value);
    if (isNaN(d)) return '';
    return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 16);
  }
  const lowerFirst = (text) => text.charAt(0).toLowerCase() + text.slice(1);
  // One condition group as a fieldset. Groups are numbered in reading order, so every name is
  // unique: "Conditions for General", then "Group 2 conditions", "Group 3 conditions".
  function conditionGroup(group, folder, path, depth, count) {
    const number = ++count.n;
    const box = el('fieldset', null, 'condition-group');
    const legend = el(
      'legend',
      depth === 0 ? 'Conditions for ' + readable(folder.Name) : 'Group ' + number + ' conditions',
    );
    legend.tabIndex = -1;
    keyed(legend, 'group:' + folder.Key + ':' + path + ':legend');
    box.append(legend);
    const match = select(
      [
        { value: 'true', label: 'all conditions' },
        { value: 'false', label: 'any condition' },
      ],
      String(group.All),
      (v) => (group.All = v === 'true'),
    );
    box.append(label('Match', match));
    const name = (n) => 'condition ' + n + (depth === 0 ? '' : ' in group ' + number);
    group.Conditions.forEach((condition, index) =>
      box.append(conditionRow(group, condition, index, folder, path, name)),
    );
    const groupKey = 'group:' + folder.Key + ':' + path;
    const add = button('Add condition', () => {
      group.Conditions.push({ field: firstField(), Operator: 'Equal', Literal: '' });
      state.focusKey = 'cond:' + folder.Key + ':' + path + '.' + group.Conditions.length + ':field';
      dirty();
      render();
    });
    add.setAttribute('aria-label', 'Add condition to ' + lowerFirst(legend.textContent));
    keyed(add, groupKey + ':add');
    const nested = button('Add group', () => {
      if (ui.blocked(nested)) return;
      group.Groups.push({ All: true, Conditions: [], Groups: [] });
      state.focusKey = 'group:' + folder.Key + ':' + path + '.' + group.Groups.length + ':legend';
      dirty();
      render();
    });
    nested.setAttribute('aria-label', 'Add group to ' + lowerFirst(legend.textContent));
    keyed(nested, groupKey + ':addgroup');
    const toolbar = el('div', null, 'condition-toolbar row');
    toolbar.append(add, nested);
    box.append(toolbar);
    // Level depth + 1 holds this group; a group inside it would be one level deeper.
    if (depth + 1 >= ui.BOUNDS.conditionDepth)
      ui.disable(
        nested,
        'depth-reason-' + (folder.Key + '-' + path).replace(/[^\w-]/g, '-'),
        'Condition groups can be nested up to ' +
          ui.BOUNDS.conditionDepth +
          " levels deep, because deeper templates can't be sent to Dataverse.",
      );
    group.Groups.forEach((child, index) => {
      const childNumber = count.n + 1;
      const inner = conditionGroup(child, folder, path + '.' + (index + 1), depth + 1, count);
      const remove = button(
        'Remove group',
        () => {
          group.Groups.splice(index, 1);
          state.focusKey = groupKey + ':addgroup';
          dirty();
          render();
        },
        'remove',
      );
      remove.setAttribute('aria-label', 'Remove group ' + childNumber);
      keyed(remove, 'group:' + folder.Key + ':' + path + '.' + (index + 1) + ':remove');
      inner.append(remove);
      box.append(inner);
    });
    return box;
  }
  function firstField() {
    const primary = 'root.' + state.root.PrimaryNameAttribute;
    const fields = availableFields();
    return (fields.find((f) => f.value === primary) || fields[0])?.value || primary;
  }
  function conditionRow(group, c, index, folder, path, names) {
    const n = index + 1;
    const name = names(n);
    const key = 'cond:' + folder.Key + ':' + path + '.' + n;
    const row = el('div', null, 'condition-line');
    const available = availableFields();
    const selected = available.find((f) => f.value === c.field);
    const { picker: field } = fieldSelect(() => true);
    field.value = pickerValue(c.field);
    field.setAttribute('aria-label', 'Field, ' + name);
    keyed(field, key + ':field');
    field.onchange = () => {
      try {
        c.field = resolveField(field.value);
      } catch (error) {
        message(error.message || String(error), true);
        return;
      }
      Object.assign(c, { Operator: 'Equal', Literal: '', label: null, table: null, right: null });
      state.focusKey = key + ':field';
      dirty();
      render();
    };
    row.append(field);
    if (!selected) {
      // Named by its label when Documents still knows it, never by its internal name.
      const shown = readable('{' + c.field + '}');
      row.append(
        el(
          'p',
          'Unavailable field' +
            (shown.startsWith('{') ? '' : ': ' + shown) +
            '. Select a replacement.',
          'error',
        ),
      );
    }
    const operator = select(
      operators(selected?.kind).map((op) => ({ value: op, label: operatorLabel[op] })),
      c.Operator,
      (v) => {
        c.Operator = v;
        if (unary(v)) c.right = null;
        state.focusKey = key + ':operator';
        render();
      },
    );
    operator.setAttribute('aria-label', 'Operator, ' + name);
    keyed(operator, key + ':operator');
    row.append(operator);
    if (!unary(c.Operator)) {
      const same = available.filter((f) => f.kind === selected?.kind && f.value !== c.field);
      const compare = select(
        [
          { value: 'literal', label: 'a value' },
          { value: 'field', label: 'another field' },
        ],
        c.right ? 'field' : 'literal',
        (v) => {
          c.right = v === 'field' ? same[0]?.value || null : null;
          state.focusKey = key + ':compare';
          render();
        },
      );
      compare.setAttribute('aria-label', 'Compare to, ' + name);
      keyed(compare, key + ':compare');
      row.append(compare);
      if (c.right) {
        const other = select(
          same.map((f) => ({ value: f.value, label: f.label })),
          c.right,
          (v) => (c.right = v),
        );
        other.setAttribute('aria-label', 'Other field, ' + name);
        keyed(other, key + ':other');
        row.append(other);
      } else row.append(valueControl(c, selected?.kind, selected?.options || [], name, key));
    }
    const remove = button(
      'Remove',
      () => {
        group.Conditions.splice(index, 1);
        state.focusKey = group.Conditions[index]
          ? key + ':field'
          : 'group:' + folder.Key + ':' + path + ':add';
        dirty();
        render();
      },
      'remove',
    );
    remove.setAttribute('aria-label', 'Remove ' + name);
    keyed(remove, key + ':remove');
    row.append(remove);
    return row;
  }
  // The value control for a field kind. key is the row's focus-key prefix.
  function valueControl(condition, kind, options, name, key) {
    const named = (control) => {
      control.setAttribute('aria-label', 'Value, ' + name);
      return keyed(control, key + ':value');
    };
    if (kind === 'Boolean')
      return named(
        select(
          [
            { value: 'true', label: 'Yes' },
            { value: 'false', label: 'No' },
          ],
          condition.Literal || 'true',
          (v) => (condition.Literal = v),
        ),
      );
    if (kind === 'Choice')
      return named(
        select(
          [{ value: '', label: 'Choose a value' }, ...options],
          condition.Literal,
          (v) => (condition.Literal = v),
        ),
      );
    if (kind === 'MultiChoice') {
      const picker = el('select');
      picker.multiple = true;
      const chosen = new Set(String(condition.Literal || '').split(','));
      options.forEach((o) => {
        const item = option(o.value, o.label);
        item.selected = chosen.has(o.value);
        picker.append(item);
      });
      picker.onchange = () => {
        condition.Literal = Array.from(picker.selectedOptions, (o) => o.value).join(',');
        dirty();
      };
      return named(picker);
    }
    if (kind === 'Lookup') {
      const wrap = el('span', null, 'lookup-value');
      if (condition.Literal) {
        const shown = condition.label || 'Record not available';
        const chip = el('span', shown, 'chip');
        const clear = button(
          '✕',
          () => {
            condition.Literal = '';
            condition.label = null;
            condition.table = null;
            state.focusKey = key + ':choose';
            dirty();
            render();
          },
          'remove',
        );
        clear.setAttribute('aria-label', 'Remove ' + shown + ', ' + name);
        keyed(clear, key + ':clear');
        chip.append(clear);
        wrap.append(chip);
      }
      const choose = button('Choose record…', async () => {
        const targets = lookupTargets(condition.field);
        const picked = await xrm.Utility.lookupObjects({
          entityTypes: targets,
          defaultEntityType: targets[0],
          allowMultiSelect: false,
        });
        if (!picked?.length) return;
        condition.Literal = picked[0].id.replace(/[{}]/g, '').toLowerCase();
        condition.label = picked[0].name;
        condition.table = picked[0].entityType;
        dirty();
        ui.withFocus(render);
      });
      choose.setAttribute('aria-label', 'Choose record, ' + name);
      keyed(choose, key + ':choose');
      wrap.append(choose);
      return wrap;
    }
    const field = el('input');
    field.type = { Number: 'number', DateOnly: 'date', DateTime: 'datetime-local' }[kind] || 'text';
    if (kind === 'Number') {
      field.setAttribute('step', 'any');
      field.setAttribute('inputmode', 'decimal');
    }
    field.value =
      kind === 'DateTime' && condition.Literal
        ? localTime(condition.Literal)
        : condition.Literal || '';
    field.oninput = () => {
      condition.Literal =
        kind === 'Number'
          ? // As typed: reformatting would turn large or small numbers into exponent form.
            field.value.trim()
          : kind === 'DateTime'
            ? field.value
              ? new Date(field.value).toISOString()
              : ''
            : field.value;
      dirty();
    };
    field.onchange = field.oninput;
    return named(field);
  }
  // A number as the server reads it (decimal.Parse with a leading sign and a decimal point, in
  // the invariant culture): no exponent, no thousands separators.
  const DECIMAL = /^[+-]?(\d+(\.\d*)?|\.\d+)$/;
  // Checks every condition at its control. With focus, the first invalid control takes it, or the
  // footer names a problem in a folder that is not open; the autosave checks without focus.
  function validate({ focus = true } = {}) {
    document.querySelectorAll('.condition-error').forEach((n) => n.remove());
    document
      .querySelectorAll('#folderEditor [aria-invalid]')
      .forEach((n) => n.removeAttribute('aria-invalid'));
    const problems = [];
    const visit = (group, folder, path) => {
      if (!group.Conditions.length && !group.Groups.length)
        problems.push([
          'group:' + folder.Key + ':' + path + ':legend',
          'Add a condition or remove this group',
        ]);
      group.Conditions.forEach((c, i) => {
        const kind = availableFields().find((f) => f.value === c.field)?.kind;
        if (unary(c.Operator) || c.right) return;
        const message =
          kind === 'Number' && !DECIMAL.test(c.Literal)
            ? 'Enter a number'
            : (kind === 'DateOnly' || kind === 'DateTime') && !c.Literal
              ? 'Pick a date'
              : kind === 'Lookup' && !c.Literal
                ? 'Choose a record'
                : (kind === 'Choice' || kind === 'MultiChoice') && !c.Literal
                  ? 'Choose a value'
                  : null;
        const part = kind === 'Lookup' ? ':choose' : ':value';
        if (message)
          problems.push(['cond:' + folder.Key + ':' + path + '.' + (i + 1) + part, message]);
      });
      group.Groups.forEach((g, i) => visit(g, folder, path + '.' + (i + 1)));
    };
    for (const section of state.sections)
      for (const folder of section.Folders)
        if (folder.Condition) visit(folder.Condition, folder, '1');
    if (!problems.length) return true;
    for (const [key, text] of problems) {
      // The editor shows one folder at a time; problems in other folders have no control here.
      const control = document.querySelector('[data-focus-key="' + key + '"]');
      if (!control) continue;
      const error = el('p', text, 'error condition-error');
      error.id = key.replace(/[^\w-]/g, '_') + '-error';
      error.setAttribute('role', 'alert');
      control.after(error);
      control.setAttribute('aria-invalid', 'true');
      control.setAttribute('aria-describedby', error.id);
    }
    if (!focus) return false;
    const first = document.querySelector('[data-focus-key="' + problems[0][0] + '"]');
    if (first) first.focus();
    else
      ui.feedback(
        'editor',
        problems[0][1] + ' in a folder that is not open. Select it in Folders.',
        'error',
      );
    return false;
  }
  let selectedSection = null,
    selectedFolder = null;
  function folderIcon() {
    const icon = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    icon.setAttribute('class', 'folder-icon');
    icon.setAttribute('viewBox', '0 0 24 24');
    icon.setAttribute('aria-hidden', 'true');
    for (const [fill, d] of [
      ['#e8b84d', 'M2 5.5h7l2 2h11v12H2z'],
      ['#f8d77e', 'M2 9h20l-2 10.5H2z'],
    ]) {
      const p = document.createElementNS(icon.namespaceURI, 'path');
      p.setAttribute('fill', fill);
      p.setAttribute('stroke', '#ad7b16');
      p.setAttribute('stroke-width', '1.2');
      p.setAttribute('d', d);
      icon.append(p);
    }
    return icon;
  }
  const libraryFor = (section) =>
    state.libraries.find((l) => l.asx_libraryid === section.LibraryId);
  const siteName = (library) =>
    state.sites.find((s) => s.asx_siteid === library?._asx_siteid_value)?.asx_name ||
    'Unavailable site';
  // A destination's name, else its place in the template: "Destination 2".
  const sectionName = (section, index) =>
    section.Name && section.Name !== section.Key ? section.Name : 'Destination ' + (index + 1);
  const destinationName = (section) => sectionName(section, state.sections.indexOf(section));
  // A destination's folders in tree order with their depth; a folder whose parent is missing
  // starts its own tree after the others.
  function treeOrder(section) {
    const order = [],
      visited = new Set();
    const visit = (folder, depth) => {
      if (visited.has(folder)) return;
      visited.add(folder);
      order.push([folder, depth]);
      section.Folders.filter((f) => f.Parent === folder.Key).forEach((f) => visit(f, depth + 1));
    };
    section.Folders.filter((f) => !f.Parent).forEach((f) => visit(f, 0));
    section.Folders.filter((f) => !visited.has(f)).forEach((f) => visit(f, 0));
    return order;
  }
  const nodeKey = (section, folder) => 'node:' + section.Key + ':' + folder.Key;
  // The published revision's destination with this key, if it has one.
  const publishedSection = (key) =>
    state.publishedSnapshot?.sections.find((s) => s.Key === key) || null;
  function addChild(section, folder) {
    // Keys of the published revision are skipped too: a new folder never takes the key of one
    // removed since, so the change reads as removed plus added.
    const used = new Set(
      [...section.Folders, ...(publishedSection(section.Key)?.Folders || [])].map((f) => f.Key),
    );
    const child = {
      Key: nextKey('folder_', used),
      Parent: folder.Key,
      Name: 'New folder',
      Condition: null,
    };
    section.Folders.push(child);
    selectedSection = section;
    selectedFolder = child;
    state.focusKey = nodeKey(section, child);
    dirty();
    render();
  }
  function renderPreview(plan) {
    const trees = $('previewTrees');
    trees.replaceChildren();
    // Adjusted folder names and folders waiting for a value; text, never markup.
    for (const notice of plan.Notices || []) trees.append(el('p', notice, 'callout warning'));
    for (const section of state.sections) {
      const rows = (plan.Folders || []).filter((f) => f.Section === section.Key);
      const card = el('div', null, 'preview-card'),
        library = libraryFor(section);
      card.append(
        el('div', destinationName(section), 'eyebrow'),
        el('h3', library?.asx_name || 'Unavailable library'),
        el('p', siteName(library), 'muted'),
      );
      if (!rows.length)
        card.append(
          el('p', 'No folders for this record: its conditions exclude this destination.'),
        );
      for (const row of rows) {
        const line = el('div', null, 'preview-node'),
          text = el('span');
        text.append(el('strong', readable(row.Name)), el('span', row.RelativePath, 'preview-note'));
        line.append(folderIcon(), text);
        card.append(line);
      }
      trees.append(card);
    }
  }
  // Selection: the chosen destination's folder, else the first destination's top folder.
  function normalizeSelection() {
    if (!state.sections.includes(selectedSection)) {
      selectedSection = state.sections[0] || null;
      selectedFolder = null;
    }
    if (selectedSection && !selectedSection.Folders.includes(selectedFolder))
      selectedFolder =
        selectedSection.Folders.find((f) => !f.Parent) || selectedSection.Folders[0] || null;
  }
  function renderFolders() {
    const list = $('destinations');
    list.replaceChildren();
    if (!state.sections.length) list.append(el('p', 'No folders yet', 'empty'));
    for (const section of state.sections) {
      const card = el(
        'section',
        null,
        'destination-section' + (section === selectedSection ? ' selected-section' : ''),
      );
      card.append(el('h3', destinationName(section), 'section-heading'));
      const tree = el('div', null, 'tree');
      for (const [folder, depth] of treeOrder(section)) {
        const row = el('div', null, 'node-row');
        row.style.paddingLeft = depth * 16 + 'px';
        const node = button(
          readable(folder.Name) || 'Unnamed folder',
          () => {
            selectedSection = section;
            selectedFolder = folder;
            ui.withFocus(render);
          },
          'node',
        );
        node.prepend(folderIcon());
        keyed(node, nodeKey(section, folder));
        node.dataset.nav = 'true';
        if (section === selectedSection && folder === selectedFolder)
          node.setAttribute('aria-current', 'true');
        row.append(node);
        if (folder.Condition) {
          // A labelled marker, not colour only; the hidden text describes the node.
          const mark = el('span', '◆', 'conditional-mark');
          mark.setAttribute('aria-hidden', 'true');
          const text = el('span', ' (conditional)', 'sr-only');
          text.id = 'conditional-' + section.Key + '-' + folder.Key;
          node.setAttribute('aria-describedby', text.id);
          row.append(mark, text);
        }
        tree.append(row);
      }
      card.append(tree);
      const parent =
        section === selectedSection ? selectedFolder : section.Folders.find((f) => !f.Parent);
      if (parent) {
        const add = button(
          '＋ Add folder inside ' + (readable(parent.Name) || 'Unnamed folder'),
          () => {
            if (!ui.blocked(add)) addChild(section, parent);
          },
          'secondary wide',
        );
        keyed(add, 'add:' + section.Key);
        card.append(add);
        list.append(card);
        ui.disable(
          add,
          'folder-reason-' + section.Key,
          section.Folders.length >= ui.BOUNDS.foldersPerDestination
            ? 'A destination can have up to ' +
                ui.BOUNDS.foldersPerDestination +
                " folders, because a record's folders for one destination are kept in one Dataverse row."
            : null,
        );
      } else list.append(card);
    }
  }
  function renderEditor() {
    const editor = $('folderEditor');
    editor.replaceChildren();
    if (!selectedSection || !selectedFolder) return;
    const section = selectedSection,
      folder = selectedFolder,
      library = libraryFor(section);
    // The folder's destination; its name, site and library are edited in step 1.
    editor.append(el('h3', destinationName(section)));
    // A top folder without a field gives every record the same folder (row 46).
    const same = el('p', 'Every record will use this same folder.', 'warning');
    const fixed = () => !folder.Parent && !/\{[^}]+\}/.test(folder.Name);
    same.hidden = !fixed();
    const name = el('input');
    name.value = folder.Name;
    name.id = 'folder-name';
    keyed(name, 'edit:name');
    name.setAttribute('aria-describedby', 'folder-shows-as help-folder-access');
    const shows = el('p', 'Shows as: ' + readable(folder.Name), 'shows-as');
    shows.id = 'folder-shows-as';
    name.oninput = () => {
      folder.Name = name.value;
      shows.textContent = 'Shows as: ' + readable(folder.Name);
      same.hidden = !fixed();
      const legend = editor.querySelector('legend');
      if (legend) legend.textContent = 'Conditions for ' + readable(folder.Name);
      dirty();
      renderFolders();
    };
    name.onchange = name.oninput;
    editor.append(
      same,
      label('Folder name', name),
      shows,
      ui.help(
        'help-folder-access',
        'Everyone with access to ' +
          (library?.asx_name || 'this library') +
          ' can open this folder. To restrict a folder, use a separate library.',
      ),
    );
    const tokens = el('div', null, 'row token-row');
    editor.append(tokens);
    const picker = fieldPicker(tokens);
    tokens.append(
      keyed(
        button('Insert field', () => {
          if (!picker.value) return;
          let field;
          try {
            field = resolveField(picker.value);
          } catch (error) {
            message(error.message || String(error), true);
            return;
          }
          folder.Name += '{' + field + '}';
          dirty();
          ui.withFocus(render);
        }),
        'edit:insert',
      ),
    );
    const include = select(
      [
        { value: 'always', label: 'Always' },
        { value: 'conditional', label: 'When conditions match' },
      ],
      folder.Condition ? 'conditional' : 'always',
      (v) => {
        if (v === 'always') {
          folder.Condition = null;
          state.focusKey = 'edit:include';
        } else {
          folder.Condition = {
            All: true,
            Conditions: [{ field: firstField(), Operator: 'Equal', Literal: '' }],
            Groups: [],
          };
          state.focusKey = 'cond:' + folder.Key + ':1.1:field';
        }
        render();
      },
    );
    keyed(include, 'edit:include');
    editor.append(label('When should this folder appear?', include));
    if (!folder.Parent) {
      include.setAttribute('aria-describedby', 'help-include-root');
      editor.append(
        ui.help(
          'help-include-root',
          "If these conditions don't match, Documents skips this whole destination for the record. Folders it already created stay.",
        ),
      );
    }
    if (folder.Condition) editor.append(conditionGroup(folder.Condition, folder, '1', 0, { n: 0 }));
    const actions = el('div', null, 'row wrap folder-actions');
    actions.setAttribute('data-actions', '');
    editor.append(actions);
    const children = section.Folders.some((f) => f.Parent === folder.Key);
    // The top folder is the destination itself; it is removed with Remove destination.
    if (folder.Parent || children) {
      const removeFolder = button('Remove folder', () => {
        if (ui.blocked(removeFolder)) return;
        section.Folders.splice(section.Folders.indexOf(folder), 1);
        selectedFolder = section.Folders.find((f) => f.Key === folder.Parent) || null;
        if (selectedFolder) state.focusKey = nodeKey(section, selectedFolder);
        dirty();
        render();
      });
      keyed(removeFolder, 'edit:remove-folder');
      actions.append(removeFolder);
      ui.disable(
        removeFolder,
        'remove-folder-reason',
        children ? 'Remove its folders first' : null,
      );
    }
  }
  // Step 1: one destination open at a time, with its name, site, library and who can open its
  // folders; the others as one line each.
  function renderDestinations() {
    const list = $('destination-list');
    if (!state.root) {
      list.replaceChildren();
      return;
    }
    if (!state.sections.some((s) => s.Key === state.expanded))
      state.expanded = state.sections[0]?.Key ?? null;
    list.replaceChildren(
      ...state.sections.map((s) =>
        s.Key === state.expanded ? destinationCard(s) : destinationRow(s),
      ),
    );
  }
  // The Library list's last option goes to Sites & access to set one up.
  const SETUP = '__setup';
  function destinationCard(section) {
    const key = 'dest:' + section.Key,
      library = libraryFor(section),
      siteId = library?._asx_siteid_value || '';
    const card = el('section', null, 'destination-card');
    const main = el('div', null, 'destination-main');
    const top = el('div', null, 'card-top');
    top.setAttribute('data-actions', '');
    const remove = button(
      'Remove',
      async () => {
        const ok = await ui.confirmInline(remove, {
          text:
            'Remove ' +
            destinationName(section) +
            ' and its ' +
            plural(section.Folders.length, 'folder', 'folders') +
            ' from this draft? Nothing changes in SharePoint until you publish.',
          confirm: 'Remove destination',
          keep: 'Keep destination',
          danger: true,
        });
        if (!ok) return;
        state.sections.splice(state.sections.indexOf(section), 1);
        dirty();
        render();
        stepStart(1)?.focus();
      },
      'link danger',
    );
    remove.setAttribute('aria-label', 'Remove ' + destinationName(section));
    keyed(remove, key + ':remove');
    top.append(
      el('span', 'Destination ' + (state.sections.indexOf(section) + 1), 'eyebrow'),
      remove,
    );
    const name = el('input');
    name.value = section.Name === section.Key ? '' : section.Name || '';
    name.maxLength = 200;
    name.setAttribute('maxlength', '200');
    keyed(name, key + ':name');
    name.oninput = () => {
      section.Name = name.value;
      remove.setAttribute('aria-label', 'Remove ' + destinationName(section));
      dirty();
      renderFolders();
    };
    name.onchange = name.oninput;
    const sites = state.sites.filter((s) =>
      state.libraries.some((l) => l._asx_siteid_value === s.asx_siteid),
    );
    const site = select(
      [
        { value: '', label: 'Choose a site' },
        ...sites.map((s) => ({ value: s.asx_siteid, label: s.asx_name })),
      ],
      siteId,
      (v) => {
        section.LibraryId =
          state.libraries.find((l) => l._asx_siteid_value === v)?.asx_libraryid || '';
        ui.withFocus(render);
      },
    );
    keyed(site, key + ':site');
    const chosen = library?.asx_libraryid || '';
    const libraries = el('select');
    [
      { value: '', label: 'Choose a library' },
      ...state.libraries
        .filter((l) => l._asx_siteid_value === siteId)
        .map((l) => ({ value: l.asx_libraryid, label: l.asx_name })),
      { value: SETUP, label: 'Set up a library in Sites & access…' },
    ].forEach((o) => libraries.append(option(o.value, o.label)));
    libraries.value = chosen;
    keyed(libraries, key + ':library');
    const setUp = () => {
      libraries.value = chosen;
      return ui.navigate('access');
    };
    // Arrow keys move through the list without leaving the page; Enter or a pick with the
    // pointer goes to Sites & access. Leaving the list keeps the library.
    let browsing = false;
    libraries.addEventListener('keydown', (event) => {
      browsing = event.key !== 'Enter';
      if (event.key === 'Enter' && libraries.value === SETUP) {
        event.preventDefault();
        setUp();
      }
    });
    libraries.addEventListener('pointerdown', () => (browsing = false));
    libraries.addEventListener('focusout', () => {
      browsing = false;
      if (libraries.value === SETUP) libraries.value = chosen;
    });
    libraries.onchange = () => {
      if (libraries.value === SETUP) return browsing ? undefined : setUp();
      section.LibraryId = libraries.value;
      dirty();
      ui.withFocus(render);
      return undefined;
    };
    const pickers = el('div', null, 'picker-grid');
    pickers.append(label('Site', site), label('Library', libraries));
    main.append(top, label('Name', name), pickers);
    const who = el('div', null, 'who-can-open');
    const change = button(
      'Change in Sites & access',
      () => ui.navigate('access', { library: section.LibraryId }),
      'link',
    );
    change.dataset.nav = 'true';
    who.append(el('h3', 'Who can open these folders'), whoCanOpen(section), change);
    card.append(main, who);
    return card;
  }
  // The teams a library's policy lists, as Sites & access lists them: a deleted team shows while
  // it still has applied access.
  function listedTeams(policy) {
    const teams = new Map((policy.Teams || []).map((t) => [String(t.TeamId).toLowerCase(), t]));
    const applied = (teamId) =>
      policy.Policy?.Applied?.find((a) => same(a.TeamId, teamId))?.Access || 'None';
    return (policy.Policy?.Desired || []).flatMap((entry) => {
      const team = teams.get(String(entry.TeamId).toLowerCase());
      const access = team?.Deleted ? applied(entry.TeamId) : entry.Access;
      if (access === 'None') return [];
      const name = team?.Name || entry.TeamId;
      return [{ name: team?.Deleted ? 'Deleted team: ' + name : name, access }];
    });
  }
  // Who can open a destination's folders, from its library's policy. Without the Security
  // Administrator role the policy is not read, and the panel says what is needed.
  function whoCanOpen(section) {
    const box = el('div', null, 'team-lines');
    if (!ui.can('prvCreateasx_policy')) {
      box.append(el('p', ui.needs('prvCreateasx_policy'), 'reason'));
      return box;
    }
    if (!section.LibraryId) return box;
    box.append(el('div', null, 'skeleton'));
    policyFor(section.LibraryId).then((policy) =>
      box.replaceChildren(
        ...(policy ? listedTeams(policy) : []).map((t) =>
          el('p', t.name + ' · ' + t.access, 'team-line'),
        ),
      ),
    );
    return box;
  }
  // A destination not open: its number, name, site and library, its teams, and Edit.
  function destinationRow(section) {
    const name = destinationName(section),
      library = libraryFor(section);
    const row = el('div', null, 'destination-row');
    const teams = el('span', null, 'teams');
    const edit = button(
      'Edit',
      () => {
        state.expanded = section.Key;
        state.focusKey = 'dest:' + section.Key + ':name';
        render();
      },
      'link',
    );
    edit.setAttribute('aria-label', 'Edit ' + name);
    keyed(edit, 'dest:' + section.Key + ':edit');
    edit.dataset.nav = 'true';
    row.append(
      el('span', 'Destination ' + (state.sections.indexOf(section) + 1), 'eyebrow'),
      el('strong', name),
      el('span', siteName(library) + ' › ' + (library?.asx_name || 'Unavailable library'), 'where'),
      teams,
      edit,
    );
    if (ui.can('prvCreateasx_policy') && section.LibraryId)
      policyFor(section.LibraryId).then((policy) => {
        if (!policy) return;
        const listed = listedTeams(policy);
        teams.textContent =
          plural(listed.length, 'team', 'teams') +
          (listed.length
            ? ' · ' + listed.map((t) => t.name + ' (' + t.access + ')').join(', ')
            : '');
      });
    return row;
  }
  // A version opened read-only: its fields and actions are off; opening a folder or another
  // destination still works.
  function lockEditor() {
    for (const control of [
      ...$('destination-list').querySelectorAll('input, select, button'),
      ...$('authorWorkspace').querySelectorAll('input, select, button'),
    ])
      if (!control.dataset.nav) control.disabled = true;
  }
  function render() {
    normalizeSelection();
    renderDestinations();
    renderFolders();
    renderEditor();
    chrome();
    if (state.readOnly) lockEditor();
    if (state.focusKey) {
      const target = document.querySelector('[data-focus-key="' + state.focusKey + '"]');
      state.focusKey = null;
      target?.focus();
    }
  }
  // Serializes the selected table, sources, destinations, folders, and conditions for draft APIs.
  function payload() {
    const fieldMap = new Map(availableFields().map((f) => [f.value, f]));
    const used = new Set();
    const group = (g) => ({
      All: g.All,
      Groups: g.Groups.map(group),
      Conditions: g.Conditions.map((c) => {
        const f = fieldMap.get(c.field);
        if (!f) throw new Error('A condition field is unavailable.');
        used.add(c.field);
        const [Source, Column] = c.field.split('.');
        const single = unary(c.Operator);
        const right = !single && c.right ? fieldMap.get(c.right) : null;
        if (c.right && !single && (!right || right.kind !== f.kind))
          throw new Error('Comparison fields must have the same available type.');
        if (right) used.add(c.right);
        const [RightSource, RightColumn] = right ? c.right.split('.') : [null, null];
        return {
          Source,
          Column,
          Operator: c.Operator,
          LiteralKind: single || right ? null : f.kind,
          Literal: single || right ? null : f.kind === 'Boolean' ? c.Literal || 'true' : c.Literal,
          RightSource,
          RightColumn,
        };
      }),
    });
    const destinations = state.sections.map((section) => ({
      Key: section.Key,
      Name: section.Name || null,
      LibraryId: section.LibraryId,
      Folders: section.Folders.map((f) => {
        for (const token of f.Name.matchAll(/\{([a-z][a-z0-9_]*\.[a-z][a-z0-9_]*)\}/g)) {
          if (!fieldMap.has(token[1])) throw new Error('Naming field is unavailable: ' + token[1]);
          used.add(token[1]);
        }
        return {
          Key: f.Key,
          Parent: f.Parent,
          Name: f.Name,
          Condition: f.Condition ? group(f.Condition) : null,
        };
      }),
    }));
    return {
      TemplateId: state.template?.asx_templateid || null,
      Name: $('templateName').value.trim(),
      RevisionId: state.editBase?.RevisionId || null,
      RowVersion: state.editBase?.RowVersion || null,
      Table: state.root.LogicalName,
      Sources: state.sources
        .filter(
          (s) => s.Alias === 'root' || s.columns.some((c) => used.has(s.Alias + '.' + c.Name)),
        )
        .map((s) => ({
          Alias: s.Alias,
          Table: s.Table,
          Lookup: s.Lookup,
          Columns: s.columns
            .filter((c) => used.has(s.Alias + '.' + c.Name))
            .map((c) => ({ Name: c.Name, Kind: c.Kind })),
        })),
      Destinations: destinations,
    };
  }
  // The template APIs that take named parameters (LoadDraft, PublishTemplate) or one Request.
  async function api(name, parameters) {
    const request = {
      ...parameters,
      getMetadata: () => ({
        boundParameter: null,
        operationType: 0,
        operationName: name,
        parameterTypes: Object.fromEntries(
          Object.keys(parameters).map((p) => [
            p,
            { typeName: 'Edm.String', structuralProperty: 1 },
          ]),
        ),
      }),
    };
    const response = await xrm.WebApi.online.execute(request);
    const body = await response.json();
    if (!response.ok) throw new Error(body.error?.message || 'Server operation failed.');
    return body.Result;
  }
  // The template row as the list, the overview, the bar, the chip and Schedule read it. The one
  // copy of this $select.
  const TEMPLATE_SELECT =
    '?$select=asx_templateid,asx_name,asx_table,asx_disabled,asx_startsutc,asx_endsutc,_asx_publishedrevisionid_value';
  // Every revision of every template, newest first. The publisher is the last writer of a
  // published revision: publishing updates it, and the server freezes it afterwards.
  const REVISION_SELECT =
    '?$select=asx_revisionid,asx_version,asx_status,_asx_templateid_value,modifiedon,_modifiedby_value&$orderby=asx_version desc';
  const BY = '_modifiedby_value@OData.Community.Display.V1.FormattedValue';
  const reloadTemplate = (templateId) =>
    xrm.WebApi.retrieveRecord('asx_template', templateId, TEMPLATE_SELECT);
  async function latestRevision() {
    const rows = await xrm.WebApi.retrieveMultipleRecords(
      'asx_revision',
      '?$select=asx_revisionid&$filter=_asx_templateid_value eq ' +
        state.template.asx_templateid +
        '&$orderby=asx_version desc&$top=1',
    );
    return rows.entities[0]?.asx_revisionid;
  }
  const revisionsOf = (templateId) =>
    state.revisions.filter((r) => same(r._asx_templateid_value, templateId));
  // A template's state from its published pointer, its Draft revision and its switch.
  function templateState(t) {
    const rows = revisionsOf(t.asx_templateid);
    const live =
      rows.find((r) => same(r.asx_revisionid, t._asx_publishedrevisionid_value))?.asx_version ??
      null;
    const draft = rows.find((r) => r.asx_status === 'Draft')?.asx_version ?? null;
    const off = !!t.asx_disabled;
    return {
      live,
      draft,
      off,
      label: off ? 'Off' : live ? 'Live v' + live : draft ? 'Draft v' + draft : 'Draft',
      tone: off ? 'muted' : live ? 'ok' : 'warning',
    };
  }
  // A library's policy, read once per page load and shared; null without the Security
  // Administrator role (no call) or when the read fails, which a later caller tries again.
  function policyFor(libraryId) {
    if (!libraryId || !ui.can('prvCreateasx_policy')) return Promise.resolve(null);
    if (!state.policyCache.has(libraryId)) {
      const read = ui
        .api('asx_SecurityAdmin', { Command: 'GetPolicy', LibraryId: libraryId })
        .catch(() => {
          if (state.policyCache.get(libraryId) === read) state.policyCache.delete(libraryId);
          return null;
        });
      state.policyCache.set(libraryId, read);
    }
    return state.policyCache.get(libraryId);
  }
  // The number of teams a policy gives access (listedTeams).
  const teamCount = (policy) => listedTeams(policy).length;
  // A template's re-run from Monitor's TemplateRuns list (running, or ended in the last day);
  // null without the Operator role (no call), when there is none, or when the read fails.
  async function lastRun(templateId) {
    if (!ui.can(OPERATOR)) return null;
    try {
      const runs = await ui.api('asx_ManageWork', {
        Command: 'ListProblems',
        List: 'TemplateRuns',
      });
      return (runs.Problems || []).find((p) => same(p.Run?.TemplateId, templateId))?.Run || null;
    } catch {
      return null;
    }
  }
  // A loaded revision as the editor and the overview hold it. Pure: columnsOf(table) answers the
  // columns of each source table, which the caller has read with fields() first.
  function toModel(loaded, columnsOf) {
    const group = (g) => ({
      All: g.All,
      Groups: (g.Groups || []).map(group),
      Conditions: (g.Conditions || []).map((c) => ({
        field: c.Source + '.' + c.Column,
        Operator: c.Operator,
        Literal: c.Literal || '',
        right: c.RightSource ? c.RightSource + '.' + c.RightColumn : null,
        label: c.LiteralLabel || null,
        table: c.LiteralTable || null,
      })),
    });
    return {
      sources: loaded.Draft.Sources.map((source) => ({
        Alias: source.Alias,
        Table: source.Table,
        Lookup: source.Lookup,
        columns: columnsOf(source.Table),
      })),
      sections: loaded.Draft.Destinations.map((d) => ({
        ...d,
        Folders: d.Folders.map((f) => ({
          ...f,
          Condition: f.Condition ? group(f.Condition) : null,
        })),
      })),
    };
  }

  // Change tracking: the draft against the published revision, by destination and folder Key.
  // A field as the same field whatever alias a revision gave its related record:
  // 'root.<column>', or '<lookup>:<table>.<column>'.
  function identity(sources, field) {
    const [alias, column] = String(field || '').split('.');
    if (alias === 'root') return field;
    const source = sources.find((s) => s.Alias === alias);
    return source ? source.Lookup + ':' + source.Table + '.' + column : field;
  }
  // A folder name with its field tokens as identities.
  const canonicalName = (sources, name) =>
    String(name ?? '').replace(
      /\{([a-z0-9_]+\.[a-z0-9_]+)\}/gi,
      (match, field) => '{' + identity(sources, field) + '}',
    );
  const GUID_LITERAL = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  // A literal as compared: a date and time as one ISO form, a record ID in lower case.
  function canonicalLiteral(value) {
    const text = String(value ?? '');
    if (/^\d{4}-\d{2}-\d{2}T/.test(text) && !isNaN(new Date(text)))
      return new Date(text).toISOString();
    return GUID_LITERAL.test(text) ? text.toLowerCase() : text;
  }
  // A condition group as compared: identities instead of aliases, canonical literals, and no
  // display-only parts (a chosen record's name and table).
  function canonicalGroup(sources, group) {
    if (!group) return 'null';
    const walk = (g) => ({
      All: !!g.All,
      Conditions: (g.Conditions || []).map((c) => {
        const single = unary(c.Operator);
        return {
          field: identity(sources, c.field),
          Operator: c.Operator,
          right: !single && c.right ? identity(sources, c.right) : null,
          Literal: single || c.right ? null : canonicalLiteral(c.Literal),
        };
      }),
      Groups: (g.Groups || []).map(walk),
    });
    return JSON.stringify(walk(group));
  }
  // The first key with this prefix that is not used: 'folder_1', 'folder_2', ...
  function nextKey(prefix, used) {
    let n = 1;
    while (used.has(prefix + n)) n++;
    return prefix + n;
  }
  // What changed from snapshot to current ({ sources, sections } each): destinations added,
  // removed, renamed or moved to another library; folders added, removed, renamed, moved or with
  // other conditions. Each list holds the added, then the removed, then the edited, in tree
  // order. marks tells the draft's new and edited folders by 'destinationKey/folderKey';
  // removed keeps each removed folder as it was; byStep counts the changes of steps 1 and 2.
  function changesSince(snapshot, current) {
    const changes = {
      destinations: [],
      folders: [],
      marks: new Map(),
      removed: [],
      byStep: { 1: 0, 2: 0 },
    };
    if (!snapshot || !current) return changes;
    const before = new Map(snapshot.sections.map((s) => [s.Key, s]));
    const after = new Set(current.sections.map((s) => s.Key));
    const named = (section, list) => sectionName(section, list.sections.indexOf(section));
    const stored = (section) => (section.Name && section.Name !== section.Key ? section.Name : '');
    for (const section of current.sections)
      if (!before.has(section.Key)) {
        changes.destinations.push({
          kind: 'added',
          key: section.Key,
          name: named(section, current),
        });
        for (const folder of section.Folders)
          changes.marks.set(section.Key + '/' + folder.Key, 'new');
      }
    for (const section of snapshot.sections)
      if (!after.has(section.Key))
        changes.destinations.push({
          kind: 'removed',
          key: section.Key,
          name: named(section, snapshot),
        });
    const added = [],
      removed = [],
      edited = [];
    for (const section of current.sections) {
      const old = before.get(section.Key);
      if (!old) continue;
      const name = named(section, current);
      if (stored(section) !== stored(old))
        changes.destinations.push({
          kind: 'renamed',
          key: section.Key,
          name,
          before: named(old, snapshot),
          after: name,
        });
      if (section.LibraryId !== old.LibraryId)
        changes.destinations.push({
          kind: 'library',
          key: section.Key,
          name,
          before: old.LibraryId,
          after: section.LibraryId,
        });
      const was = new Map(old.Folders.map((f) => [f.Key, f]));
      const kept = new Set(section.Folders.map((f) => f.Key));
      for (const [folder] of treeOrder(section)) {
        const entry = { destination: section.Key, key: folder.Key, name: folder.Name };
        const prior = was.get(folder.Key);
        const mark = section.Key + '/' + folder.Key;
        if (!prior) {
          added.push({ kind: 'added', ...entry });
          changes.marks.set(mark, 'new');
          continue;
        }
        const edits = [];
        if (
          canonicalName(snapshot.sources, prior.Name) !==
          canonicalName(current.sources, folder.Name)
        )
          edits.push({ kind: 'renamed', ...entry, before: prior.Name, after: folder.Name });
        if ((prior.Parent ?? null) !== (folder.Parent ?? null))
          edits.push({ kind: 'moved', ...entry, before: prior.Parent, after: folder.Parent });
        if (
          canonicalGroup(snapshot.sources, prior.Condition) !==
          canonicalGroup(current.sources, folder.Condition)
        )
          edits.push({ kind: 'condition', ...entry });
        if (edits.length) changes.marks.set(mark, 'edited');
        edited.push(...edits);
      }
      for (const [folder] of treeOrder(old))
        if (!kept.has(folder.Key)) {
          removed.push({
            kind: 'removed',
            destination: section.Key,
            key: folder.Key,
            name: folder.Name,
          });
          changes.removed.push({ destination: section.Key, folder });
        }
    }
    changes.folders = [...added, ...removed, ...edited];
    changes.byStep = { 1: changes.destinations.length, 2: changes.folders.length };
    return changes;
  }

  // The overview: one template read top to bottom (destinations, folders, schedule, versions),
  // with Edit template to open the editor.
  async function showOverview(templateId) {
    state.view = 'overview';
    panel?.close(false);
    $('template-overview').hidden = false;
    $('template-editor').hidden = true;
    // The re-run read needs nothing else, so it runs beside the others.
    const reading = lastRun(templateId);
    try {
      const template = await reloadTemplate(templateId);
      const index = state.templates.findIndex((t) => same(t.asx_templateid, templateId));
      if (index >= 0) state.templates[index] = template;
      const rows = revisionsOf(templateId);
      const revisionId =
        template._asx_publishedrevisionid_value ||
        rows.find((r) => r.asx_status === 'Draft')?.asx_revisionid ||
        rows[0]?.asx_revisionid ||
        null;
      let model = { sources: [], sections: [] };
      if (revisionId) {
        const loaded = JSON.parse(await api('asx_LoadDraft', { RevisionId: revisionId }));
        // A table whose fields cannot be read shows its field tokens as written.
        await Promise.all(loaded.Draft.Sources.map((s) => fields(s.Table).catch(() => null)));
        model = toModel(loaded, (table) => metadata.get(table)?.columns || []);
      }
      const [run, policies] = await Promise.all([
        reading,
        Promise.all(model.sections.map(async (d) => [d.LibraryId, await policyFor(d.LibraryId)])),
      ]);
      // Another template opens with its first destination's folders shown.
      if (!same(state.overview?.template.asx_templateid, templateId))
        state.openGroups = new Set(model.sections.slice(0, 1).map((d) => d.Key));
      state.overview = { template, revisionId, model, run, policies: new Map(policies) };
    } catch (error) {
      state.overview = null;
      throw error;
    } finally {
      renderList();
      renderOverview();
    }
  }
  function renderOverview() {
    const o = state.overview;
    $('overview-actions').hidden = !o;
    $('overview-pill').hidden = !o;
    if (!o) {
      $('overview-title').textContent = 'Folder templates';
      $('overview-meta').textContent = '';
      const empty = el('div', null, 'empty-panel');
      empty.append(el('h2', 'No template selected'));
      $('overview-cards').replaceChildren(empty);
      return;
    }
    const t = o.template,
      s = templateState(t),
      name = t.asx_name || tableName(t.asx_table);
    $('overview-title').textContent = name;
    $('overview-pill').textContent =
      !s.off && s.live && s.draft ? 'Live v' + s.live + ' · Draft v' + s.draft : s.label;
    $('overview-pill').dataset.tone = s.tone;
    const published = state.revisions.find((r) =>
      same(r.asx_revisionid, t._asx_publishedrevisionid_value),
    );
    $('overview-meta').textContent =
      tableName(t.asx_table) +
      ' table' +
      (published?.modifiedon
        ? ' · published ' +
          day(published.modifiedon) +
          (published[BY] ? ' by ' + published[BY] : '')
        : '');
    $('overview-menu').setAttribute('aria-label', 'More actions for ' + name);
    // Re-run needs a published version; its separator goes with it.
    $('menu-rerun').hidden = $('menu-rerun-separator').hidden = !t._asx_publishedrevisionid_value;
    $('overview-edit').textContent = s.draft ? 'Continue Draft v' + s.draft : 'Edit template';
    $('overview-cards').replaceChildren(
      destinationsCard(o),
      foldersCard(o),
      scheduleCard(o),
      versionsCard(o),
    );
  }
  // A card whose head action is "Edit": named for what it edits, for screen readers.
  function card(options, editLabel) {
    const made = ui.card(options);
    if (editLabel) made.head.querySelector('.card-action').setAttribute('aria-label', editLabel);
    return made;
  }
  const edit = (step) => () => task(() => openEditor(step));
  function destinationsCard(o) {
    const sections = o.model.sections;
    const made = card(
      {
        title: 'Destinations',
        summary: plural(sections.length, 'library', 'libraries'),
        action: { label: 'Edit', onClick: edit(1), key: 'overview:destinations' },
      },
      'Edit destinations',
    );
    sections.forEach((section, index) => {
      const library = libraryFor(section),
        row = el('div', null, 'card-row dest-row');
      row.append(
        el('strong', sectionName(section, index)),
        el(
          'span',
          siteName(library) + ' › ' + (library?.asx_name || 'Unavailable library'),
          'muted',
        ),
      );
      const policy = o.policies.get(section.LibraryId);
      if (policy) row.append(el('span', plural(teamCount(policy), 'team', 'teams'), 'teams'));
      made.body.append(row);
    });
    return made.card;
  }
  function foldersCard(o) {
    const { sources, sections } = o.model;
    const folders = sections.flatMap((d) => d.Folders);
    const conditional = folders.filter((f) => ruleText(f, sources).conditional).length;
    const made = card(
      {
        title: 'Folders',
        summary: plural(folders.length, 'folder', 'folders') + ' · ' + conditional + ' conditional',
        action: { label: 'Edit', onClick: edit(2), key: 'overview:folders' },
      },
      'Edit folders',
    );
    const label = labelOf(sources);
    sections.forEach((section, index) => {
      const open = state.openGroups.has(section.Key);
      const toggle = button(
        '',
        () => {
          if (open) state.openGroups.delete(section.Key);
          else state.openGroups.add(section.Key);
          ui.withFocus(renderOverview);
        },
        'group-row',
      );
      toggle.setAttribute('aria-expanded', String(open));
      keyed(toggle, 'folders:' + section.Key);
      const glyph = el('span', open ? '▾' : '▸', 'glyph');
      glyph.setAttribute('aria-hidden', 'true');
      toggle.append(glyph, el('span', sectionName(section, index), 'group-name'));
      if (!open)
        toggle.append(el('span', plural(section.Folders.length, 'folder', 'folders'), 'count'));
      made.body.append(toggle);
      if (!open) return;
      for (const [folder, depth] of treeOrder(section)) {
        const row = el('div', null, 'card-row folder-row'),
          rule = ruleText(folder, sources);
        row.style.paddingLeft = 36 + 20 * depth + 'px';
        row.append(
          folderIcon(),
          ui.tokens(folder.Name, label),
          el('span', rule.text, rule.conditional ? 'rule conditional' : 'rule'),
        );
        made.body.append(row);
      }
    });
    return made.card;
  }
  // "Off", or "On" with its start (when still ahead) and its end: "On · no end date".
  function scheduleText(t) {
    if (t.asx_disabled) return 'Off';
    return (
      'On' +
      (t.asx_startsutc && new Date(t.asx_startsutc) > new Date()
        ? ' · starts ' + day(t.asx_startsutc)
        : '') +
      (t.asx_endsutc ? ' · ends ' + day(t.asx_endsutc) : ' · no end date')
    );
  }
  function scheduleCard(o) {
    const made = card(
      {
        title: 'Schedule and runs',
        summary: scheduleText(o.template),
        action: {
          label: 'Edit',
          onClick: () => openSchedule(made.head.querySelector('.card-action')),
          key: 'overview:schedule',
        },
      },
      'Edit schedule',
    );
    const run = o.run;
    if (run) {
      const row = el('div', null, 'card-row run-row');
      row.append(
        el('span', ACTIVE.includes(run.State) ? 'Re-run in progress' : 'Last re-run'),
        el(
          'strong',
          (run.StartedUtc ? day(run.StartedUtc) + ' · ' : '') +
            Number(run.Planned || 0).toLocaleString('en-US') +
            ' of ' +
            (run.TotalEstimated ? 'about ' : '') +
            Number(run.Total || 0).toLocaleString('en-US') +
            ' records',
        ),
      );
      made.body.append(row);
    }
    return made.card;
  }
  // One version: its number, Live, Replaced or Draft, when and by whom it was last saved, and
  // View, which opens it read-only in the editor.
  function versionRow(r, t, before = null) {
    const [text, tone] = same(r.asx_revisionid, t._asx_publishedrevisionid_value)
      ? ['Live', 'ok']
      : r.asx_status === 'Draft'
        ? ['Draft', 'muted']
        : ['Replaced', 'muted'];
    const row = el('div', null, 'card-row version-row'),
      status = el('span', text, 'state');
    status.dataset.tone = tone;
    const view = button(
      'View',
      () => {
        before?.();
        return task(() => openVersion(r.asx_revisionid));
      },
      'link',
    );
    view.setAttribute('aria-label', 'View v' + r.asx_version);
    keyed(view, 'version:' + r.asx_revisionid);
    row.append(
      el('span', 'v' + r.asx_version, 'v'),
      status,
      el(
        'span',
        [r.modifiedon ? day(r.modifiedon) : '', r[BY]].filter(Boolean).join(' · '),
        'muted',
      ),
      view,
    );
    return row;
  }
  // The draft, the live version and the one it replaced; All versions lists every one.
  function versionsCard(o) {
    const t = o.template,
      rows = revisionsOf(t.asx_templateid);
    const live = rows.find((r) => same(r.asx_revisionid, t._asx_publishedrevisionid_value));
    const shown = [
      rows.find((r) => r.asx_status === 'Draft'),
      live,
      live && rows.find((r) => r.asx_version < live.asx_version),
    ].filter(Boolean);
    const made = ui.card({
      title: 'Versions',
      summary: String(rows.length),
      action: {
        label: 'All versions',
        onClick: () => openHistory(made.head.querySelector('.card-action')),
        key: 'overview:versions',
      },
    });
    [...new Set(shown)]
      .sort((a, b) => b.asx_version - a.asx_version)
      .forEach((r) => made.body.append(versionRow(r, t)));
    return made.card;
  }

  // The editor: Edit template opens the latest revision of the overview's template; View opens
  // one version read-only; Close returns to the overview, asking first about unsaved edits.
  async function openEditor(step = 1) {
    const t = state.overview?.template;
    if (t && (!state.root || !same(state.template?.asx_templateid, t.asx_templateid)))
      await selectTemplate(t.asx_table, t.asx_templateid);
    state.step = step;
    showEditor();
  }
  function showEditor() {
    state.view = 'edit';
    panel?.close(false);
    $('template-overview').hidden = true;
    $('template-editor').hidden = false;
    goStep(state.step, false);
    chrome();
    renderList();
    $('editor-title').focus();
  }
  async function openVersion(revisionId) {
    const t = state.overview.template;
    if (!state.root || !same(state.template?.asx_templateid, t.asx_templateid))
      await selectTemplate(t.asx_table, t.asx_templateid);
    // The viewed version replaces the latest one selectTemplate loaded.
    await reloadVersion(revisionId);
    state.readOnly = true;
    state.changes = null;
    render();
    await openEditor(1);
  }
  // A new editor session: nothing saving, nothing to compare yet, step 1.
  function resetSession() {
    clearTimeout(state.saveTimer);
    state.session++;
    Object.assign(state, {
      saveState: 'idle',
      savedAt: null,
      saveError: null,
      saving: null,
      saveAgain: false,
      edits: 0,
      publishedSnapshot: null,
      changes: null,
      expanded: null,
      invalid: false,
      step: 1,
    });
    ui.clearFeedback('editor');
  }
  // Clears the editor, so the next Edit template loads the saved version again.
  function resetEditor() {
    resetSession();
    Object.assign(state, {
      saved: null,
      editBase: null,
      sections: [],
      template: null,
      root: null,
      unsaved: false,
      readOnly: false,
      run: null,
      batch: null,
    });
    // The background reads of the closed template take no more tables.
    state.preloadId++;
    selectedSection = null;
    selectedFolder = null;
    $('templateName').value = '';
    resetPreview();
    render();
  }
  async function closeEditor() {
    if (!(await ui.confirmLeave())) return;
    const id = state.template?.asx_templateid || state.overview?.template.asx_templateid;
    panel?.close(false);
    resetEditor();
    if (id) await task(() => showOverview(id));
    else {
      state.view = 'overview';
      $('template-overview').hidden = false;
      $('template-editor').hidden = true;
      renderList();
      renderOverview();
    }
    (state.overview ? $('overview-edit') : $('overview-title')).focus();
  }
  $('editor-close').onclick = closeEditor;
  $('overview-edit').onclick = edit(1);

  // The steps are tabs: a click shows its step; arrow keys, Home and End move focus only, and
  // Enter or Space shows the focused one.
  const stepTabs = [1, 2, 3].map((n) => $('step-tab-' + n));
  stepTabs.forEach((tab, index) => (tab.onclick = () => goStep(index + 1, false)));
  $('editor-steps').addEventListener('keydown', (event) => {
    const index = stepTabs.indexOf(document.activeElement);
    const target = {
      ArrowRight: index + 1,
      ArrowLeft: index - 1,
      Home: 0,
      End: stepTabs.length - 1,
    }[event.key];
    if (index < 0 || target === undefined) return;
    event.preventDefault();
    const next = stepTabs[(target + stepTabs.length) % stepTabs.length];
    stepTabs.forEach((tab) => (tab.tabIndex = tab === next ? 0 : -1));
    next.focus();
  });
  $('step-back').onclick = () => goStep(state.step - 1);
  $('step-next').onclick = () => goStep(state.step + 1);

  // The overview's ⋯ menu: keyboard, focus and closing come from AsxdUi.menu; choosing an item
  // closes it.
  ui.menu($('overview-menu'), $('overview-menu-list'));
  $('menu-rerun').onclick = () => openRerun($('overview-menu'));
  $('menu-delete').onclick = deleteTemplate;
  for (const close of document.querySelectorAll('.panel-close'))
    close.onclick = () => panel?.close();

  $('publish').onclick = async () => {
    if (ui.blocked($('publish'))) return;
    // What is published is what is saved: edits not saved yet are saved first.
    try {
      await flushSave();
    } catch (error) {
      ui.feedback('editor', error.message || String(error), 'error');
      return;
    }
    // A new template without a name cannot be saved yet.
    if (!state.saved) {
      $('templateName').focus();
      return;
    }
    const next = state.editBase.Version;
    const paused = ui.runtime() && !ui.runtime().Enabled;
    const ok = await ui.confirmInline($('publish'), {
      text: 'Publish v' + next + '?',
      details: ui.help(
        'help-publish',
        'Documents uses v' +
          next +
          ' for records created from now on, and for changed records when record updates are on. Existing records keep their folders until you re-run them.' +
          (paused ? " Automation is paused, so no folders are created until it's on." : ''),
      ),
      confirm: 'Publish v' + next,
      keep: 'Not now',
    });
    if (!ok) return;
    await ui.busy($('publish'), 'Publishing…', 'templates', async () => {
      // Edits made while the confirmation was open.
      await flushSave();
      state.publishing = true;
      try {
        const result = JSON.parse(
          await api('asx_PublishTemplate', {
            RevisionId: state.saved.RevisionId,
            RowVersion: state.saved.RowVersion,
          }),
        );
        // Publishing changes the revision row, so its row version is read back with it; the
        // next save would be refused with the old one. Edits made during the publish stay.
        // The published version is what the draft is compared with from now on.
        if (state.unsaved) {
          const loaded = await reloadBasis(state.saved.RevisionId);
          state.publishedSnapshot = {
            version: next,
            ...toModel(loaded, (table) => metadata.get(table)?.columns || []),
          };
        } else {
          await reloadVersion(state.saved.RevisionId);
          state.publishedSnapshot = structuredClone({ version: next, ...current() });
        }
        state.changes = changesSince(state.publishedSnapshot, current());
        state.template = await reloadTemplate(state.template.asx_templateid);
        // The list's state and the re-run's version follow the new published revision.
        await loadTemplates();
        chrome();
        ui.feedback(
          'templates',
          ['Published v' + next + '.', ...(result.Notices || [])].join(' '),
          'success',
          {
            label: 'Re-run existing records…',
            onClick: (event) => openRerun(event?.currentTarget),
          },
        );
      } finally {
        state.publishing = false;
        if (state.unsaved) scheduleSave();
      }
    });
    chrome();
  };

  // Every version of the overview's template, in a side panel.
  function openHistory(invoker) {
    const t = state.overview.template;
    $('history-list').replaceChildren(
      ...revisionsOf(t.asx_templateid).map((r) => versionRow(r, t, () => panel?.close(false))),
    );
    panel = ui.sidePanel($('history-panel'), invoker);
  }

  function openSchedule(invoker) {
    const t = state.overview.template;
    $('schedule-on').setAttribute('aria-checked', String(!t.asx_disabled));
    $('templateStart').value = localTime(t.asx_startsutc);
    $('templateEnd').value = localTime(t.asx_endsutc);
    ui.clearFeedback('schedule');
    panel = ui.sidePanel($('schedule-panel'), invoker);
  }
  $('schedule-on').onclick = () =>
    $('schedule-on').setAttribute(
      'aria-checked',
      String($('schedule-on').getAttribute('aria-checked') !== 'true'),
    );
  $('saveAvailability').onclick = () => {
    if (ui.blocked($('saveAvailability'))) return undefined;
    return ui.busy($('saveAvailability'), 'Saving…', 'schedule', async () => {
      const start = $('templateStart').value
        ? new Date($('templateStart').value).toISOString()
        : null;
      const end = $('templateEnd').value ? new Date($('templateEnd').value).toISOString() : null;
      if (start && end && end <= start) throw new Error('The end must be after the start.');
      const id = state.overview.template.asx_templateid;
      await xrm.WebApi.updateRecord('asx_template', id, {
        asx_disabled: $('schedule-on').getAttribute('aria-checked') !== 'true',
        asx_startsutc: start,
        asx_endsutc: end,
      });
      const template = await reloadTemplate(id);
      state.overview.template = template;
      const index = state.templates.findIndex((t) => same(t.asx_templateid, id));
      if (index >= 0) state.templates[index] = template;
      if (same(state.template?.asx_templateid, id)) state.template = template;
      renderList();
      renderOverview();
      chrome();
      ui.feedback('schedule', 'Schedule saved.');
    });
  };

  async function deleteTemplate() {
    const t = state.overview.template;
    const name = t.asx_name || tableName(t.asx_table);
    const ok = await ui.confirmInline($('overview-menu'), {
      text:
        'Delete ' +
        name +
        ' and all its versions? No new folder work starts for it. Folders, documents and access in SharePoint stay as they are. Work already sent to SharePoint may still finish.',
      confirm: 'Delete template',
      keep: 'Keep template',
      danger: true,
    });
    if (!ok) return;
    // Its place in the list: focus goes to the template that takes it.
    const place = $('template-groups')
      .querySelectorAll('.list-row')
      .findIndex((r) => r.dataset.focusKey === 'template:' + t.asx_templateid);
    const done = await ui.busy($('overview-menu'), 'Deleting…', 'templates', async () => {
      await xrm.WebApi.deleteRecord('asx_template', t.asx_templateid);
      panel?.close(false);
      if (same(state.template?.asx_templateid, t.asx_templateid)) resetEditor();
      state.overview = null;
      await loadTemplates();
      renderOverview();
      ui.feedback('templates', name + ' deleted. Nothing in SharePoint changed.');
      return true;
    });
    if (!done) return;
    const next = place < 0 ? null : $('template-groups').querySelectorAll('.list-row')[place];
    (next || $('template-search')).focus();
  }

  // Unsaved edits ask before another template replaces them; true when it may. A save in
  // flight finishes first.
  async function mayDiscard(control, what) {
    await state.saving;
    if (!state.unsaved || !state.root) return true;
    return railAsk(control, {
      text:
        'Open ' +
        what +
        '? Your unsaved changes to ' +
        ($('templateName').value.trim() || 'this template') +
        ' are discarded.',
      confirm: 'Open ' + what,
      keep: 'Keep editing',
    });
  }

  // Loads a saved version into the editor with its version number, edit identity and row version.
  async function reloadVersion(revisionId) {
    const loaded = JSON.parse(await api('asx_LoadDraft', { RevisionId: revisionId }));
    if (loaded.Draft.Table !== state.root.LogicalName)
      throw new Error('This version belongs to another table.');
    await Promise.all(loaded.Draft.Sources.map((source) => fields(source.Table)));
    const model = toModel(loaded, (table) => metadata.get(table).columns);
    state.sources = model.sources;
    state.sections = model.sections;
    state.saved = {
      RevisionId: loaded.RevisionId,
      RowVersion: loaded.RowVersion,
      Status: loaded.Status,
    };
    state.editBase = { ...state.saved, Version: loaded.Version };
    state.unsaved = false;
    selectedSection = null;
    selectedFolder = null;
    render();
    if (state.previewRecord) await runPreview();
  }
  // Reads a version's row version and status again and leaves the editor's edits as they are.
  async function reloadBasis(revisionId) {
    const loaded = JSON.parse(await api('asx_LoadDraft', { RevisionId: revisionId }));
    state.saved = {
      RevisionId: loaded.RevisionId,
      RowVersion: loaded.RowVersion,
      Status: loaded.Status,
    };
    state.editBase = { ...state.saved, Version: loaded.Version };
    return loaded;
  }
  // The published revision the draft is compared with: the loaded one when it is the published
  // one, else read; null without one, or when it cannot be read (no change marks then).
  async function publishedModel() {
    const id = state.template?._asx_publishedrevisionid_value;
    if (!id) return null;
    if (same(id, state.saved?.RevisionId))
      return structuredClone({ version: state.editBase.Version, ...current() });
    try {
      const loaded = JSON.parse(await api('asx_LoadDraft', { RevisionId: id }));
      await Promise.all(loaded.Draft.Sources.map((s) => fields(s.Table).catch(() => null)));
      return {
        version: loaded.Version,
        ...toModel(loaded, (table) => metadata.get(table)?.columns || []),
      };
    } catch {
      return null;
    }
  }
  // Each table's fields take about seven metadata requests at once, and Dataverse serves 52
  // concurrent requests per user before it answers 429. Four tables at a time stay well inside
  // that, beside the page's other reads.
  const PRELOAD_TABLES = 4;
  async function preload(tables, id) {
    const queue = tables.filter((t) => !metadata.has(t));
    queue.forEach((t) => state.failedTables.delete(t));
    // A worker stops taking tables once another template opens (a newer preload id).
    const worker = async () => {
      while (queue.length && id === state.preloadId) {
        const table = queue.shift();
        try {
          await fields(table);
        } catch {
          // One retry: a 429 or a dropped request usually passes the second time.
          try {
            await fields(table);
          } catch {
            state.failedTables.add(table);
          }
        }
        if (id === state.preloadId) refreshPickers();
      }
    };
    await Promise.all(Array.from({ length: Math.min(PRELOAD_TABLES, queue.length) }, worker));
  }
  // Redraws the pickers as related tables arrive, unless the admin is typing in the editor or
  // answering a confirmation there. A skipped redraw marks them stale, and it runs once focus
  // moves: out of the field, or back to the invoker when the confirmation closes.
  function refreshPickers() {
    if (!state.root) return;
    const editor = $('template-editor');
    const active = document.activeElement;
    state.pickersStale =
      !!editor.querySelector('.confirm[role=group]') ||
      (editor.contains(active) && active.tagName === 'INPUT');
    if (!state.pickersStale) ui.withFocus(render);
  }
  // After the focus change settles: during focusout, focus has not reached its next element yet.
  for (const type of ['focusout', 'focusin'])
    $('template-editor').addEventListener(type, () => {
      if (state.pickersStale) setTimeout(refreshPickers, 0);
    });
  function resetPreview() {
    state.previewRecord = null;
    state.previewed = false;
    $('preview-record-name').textContent = '';
    $('preview-stale').hidden = true;
    $('refreshPreview').hidden = true;
    $('previewTrees').classList.remove('is-stale');
    $('previewTrees').replaceChildren(el('p', 'Pick a record to preview', 'empty'));
    $('preview-status').textContent = '';
  }
  async function selectTemplate(table, templateId = null) {
    const root = state.tables.find((t) => t.LogicalName === table);
    if (!root) throw new Error(tableName(table) + ' is not available for document management.');
    state.template = templateId ? await reloadTemplate(templateId) : null;
    $('templateName').value = state.template?.asx_name || 'New template';
    resetSession();
    Object.assign(state, {
      root,
      editBase: null,
      saved: null,
      unsaved: false,
      readOnly: false,
      sections: [],
      run: null,
      batch: null,
    });
    selectedSection = null;
    selectedFolder = null;
    panel?.close(false);
    ui.clearFeedback('templates');
    resetPreview();
    const meta = await fields(table);
    state.sources = [{ Alias: 'root', Table: table, Lookup: null, columns: meta.columns }];
    // The label reads the target's display name once its metadata is in.
    state.lookups = meta.lookups.flatMap((l) =>
      l.Targets.map((Table) => ({
        Lookup: l.LogicalName,
        Table,
        get Label() {
          return display(l) + ' → ' + tableName(Table);
        },
      })),
    );
    const id = ++state.preloadId;
    // The tables the draft uses load before the first render (reloadVersion); the other related
    // tables load after it, in the background.
    const latest = state.template ? await latestRevision() : null;
    if (latest) {
      await reloadVersion(latest);
      state.publishedSnapshot = await publishedModel();
      state.changes = changesSince(state.publishedSnapshot, current());
    }
    render();
    preload([...new Set(meta.lookups.flatMap((l) => l.Targets))], id);
  }
  $('go-access').onclick = () => ui.navigate('access');
  // Tables are added, removed and repaired in Settings.
  $('manage-tables').onclick = () => ui.navigate('settings');
  // Every row of a paged read.
  async function pages(table, query) {
    const rows = [];
    let options = query;
    do {
      const result = await xrm.WebApi.retrieveMultipleRecords(table, options);
      rows.push(...result.entities);
      options = result.nextLink
        ? new URL(result.nextLink, xrm.Utility.getGlobalContext().getClientUrl()).search
        : null;
    } while (options);
    return rows;
  }
  async function loadEnabledTables() {
    const rows = await pages(
      'asx_runtimetable',
      '?$select=asx_logicalname&$orderby=asx_logicalname',
    );
    state.enabledTables = [...new Set(rows.map((r) => r.asx_logicalname).filter(Boolean))].sort();
  }
  // A confirmation in the list. A redraw asked for while it is open waits, and runs once it closes.
  async function railAsk(control, options) {
    try {
      return await ui.confirmInline(control, options);
    } finally {
      if (state.railStale) renderList();
    }
  }
  async function loadTemplates() {
    const [templates, revisions] = await Promise.all([
      pages('asx_template', TEMPLATE_SELECT + '&$orderby=asx_name'),
      pages('asx_revision', REVISION_SELECT),
    ]);
    state.templates = templates;
    state.revisions = revisions;
    renderList();
  }
  // Opens a template's overview from the list, asking first when editor changes would be lost.
  async function pick(control, t) {
    if (state.view === 'edit') {
      if (!(await mayDiscard(control, t.asx_name || tableName(t.asx_table)))) return;
      resetEditor();
    }
    $('new-template-table').hidden = true;
    // The list redraws with the overview; focus moves to the redrawn row it was on.
    const focused = document.activeElement === control;
    await task(() => showOverview(t.asx_templateid));
    if (focused)
      document.querySelector('[data-focus-key="template:' + t.asx_templateid + '"]')?.focus();
  }
  // The list's tables in order: enabled tables by name, then tables no longer enabled.
  function listTables() {
    const byName = (a, b) => tableName(a).localeCompare(tableName(b));
    const tables = [...new Set(state.templates.map((t) => t.asx_table).filter(Boolean))];
    return {
      enabled: tables.filter((t) => state.enabledTables.includes(t)).sort(byName),
      disabled: tables.filter((t) => !state.enabledTables.includes(t)).sort(byName),
    };
  }
  // The template the list marks as selected: the one the editor or the overview shows.
  const selectedId = () =>
    state.view === 'edit'
      ? state.template?.asx_templateid
      : state.overview?.template.asx_templateid;
  function listRow(t) {
    const s = templateState(t),
      status = el('span', s.label, 'row-state');
    status.dataset.tone = s.tone;
    const row = button('', () => pick(row, t), 'list-row');
    row.append(el('span', t.asx_name || tableName(t.asx_table), 'row-name'), status);
    keyed(row, 'template:' + t.asx_templateid);
    if (same(t.asx_templateid, selectedId())) row.setAttribute('aria-current', 'true');
    return row;
  }
  // The templates list: one group per table, each template with its state; the search filters
  // by name.
  function renderList() {
    const host = $('template-groups');
    // Never under an open confirmation (as Monitor's lists): railAsk redraws when it closes.
    state.railStale = !!host.querySelector('.confirm[role=group]');
    if (state.railStale) return;
    host.replaceChildren();
    $('new-template').hidden = !state.enabledTables.length;
    if (!state.loaded) return;
    const { enabled, disabled } = listTables();
    if (!enabled.length && !disabled.length && !state.enabledTables.length) {
      host.append(el('p', 'No tables yet', 'empty'));
      return;
    }
    const query = state.listQuery.trim().toLowerCase();
    const group = (table) => {
      const rows = state.templates.filter(
        (t) =>
          t.asx_table === table && (t.asx_name || tableName(table)).toLowerCase().includes(query),
      );
      if (!rows.length) return null;
      const box = el('div', null, 'list-group');
      box.append(el('span', tableName(table), 'group-label'), ...rows.map(listRow));
      return box;
    };
    host.append(...enabled.map(group).filter(Boolean));
    const others = disabled.map(group).filter(Boolean);
    if (others.length) host.append(el('p', 'Not enabled', 'rail-group'), ...others);
  }
  $('template-search').oninput = () => {
    state.listQuery = $('template-search').value;
    renderList();
  };
  // ＋ New: the table first when there is more than one. Arrow keys move through the tables
  // without opening one; Enter, or a pick with the pointer, opens it.
  const tablePicker = $('new-template-table');
  let browsing = false;
  $('new-template').onclick = () => {
    const tables = [...state.enabledTables].sort((a, b) =>
      tableName(a).localeCompare(tableName(b)),
    );
    if (tables.length === 1) return newTemplate($('new-template'), tables[0]);
    tablePicker.replaceChildren(
      option('', 'Choose a table'),
      ...tables.map((t) => option(t, tableName(t))),
    );
    tablePicker.value = '';
    tablePicker.hidden = false;
    tablePicker.focus();
    return undefined;
  };
  // Leaving the picker ends a keyboard browse, so a later pick with the pointer opens.
  for (const type of ['pointerdown', 'focusout'])
    tablePicker.addEventListener(type, () => (browsing = false));
  tablePicker.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      browsing = false;
      tablePicker.hidden = true;
      $('new-template').focus();
      return;
    }
    browsing = event.key !== 'Enter';
    if (event.key !== 'Enter') return;
    event.preventDefault();
    chooseTable();
  });
  tablePicker.onchange = () => {
    if (!browsing) chooseTable();
  };
  function chooseTable() {
    if (!tablePicker.value) return;
    tablePicker.hidden = true;
    newTemplate($('new-template'), tablePicker.value);
  }
  async function newTemplate(control, table) {
    if (state.view === 'edit' && !(await mayDiscard(control, 'a new template'))) return;
    await task(async () => {
      resetEditor();
      await selectTemplate(table, null);
      showEditor();
    });
  }
  $('add-destination').onclick = () => {
    if (ui.blocked($('add-destination')) || !state.libraries.length || state.readOnly) return;
    // Keys of the published revision are skipped too, as for folders.
    const used = new Set(
      [...state.sections, ...(state.publishedSnapshot?.sections || [])].map((s) => s.Key),
    );
    const library = state.libraries[0];
    const section = {
      Key: nextKey('destination_', used),
      // A new destination is named after its library.
      Name: library.asx_name,
      LibraryId: library.asx_libraryid,
      Folders: [
        {
          Key: 'root',
          Parent: null,
          Name: '{root.' + state.root.PrimaryNameAttribute + '}',
          Condition: null,
        },
      ],
    };
    state.sections.push(section);
    // The new destination opens in step 1, and its top folder in step 2.
    state.expanded = section.Key;
    selectedSection = section;
    selectedFolder = section.Folders[0];
    state.focusKey = 'dest:' + section.Key + ':name';
    dirty();
    render();
  };
  $('templateName').oninput = () => {
    if (!state.template) dirty();
  };

  // Saves the draft. The version follows CreateDraftApi: a Draft is saved in place, anything
  // else starts the next version; a new template starts at v1. It never redraws the editor, so
  // the field being typed in stays as it is.
  async function save() {
    const base = state.editBase,
      session = state.session;
    const name = $('templateName').value.trim();
    const saved = JSON.parse(await api('asx_CreateDraft', { Request: JSON.stringify(payload()) }));
    // Another template opened meanwhile: this answer is not about it.
    if (session !== state.session) return;
    // What the save made is recorded at once, so a failed read below cannot make the next save
    // create a second template or a second draft.
    state.saved = saved;
    state.template = state.template || {
      asx_templateid: saved.TemplateId,
      asx_name: name,
      asx_table: state.root.LogicalName,
    };
    state.editBase = {
      ...saved,
      Version: !base ? 1 : base.Status === 'Draft' ? base.Version : base.Version + 1,
    };
    // A draft saved in place changes nothing else; a new template or version is read back.
    if (base?.Status === 'Draft') return;
    try {
      // The saved version's own number: a version opened from Version history is not the latest.
      const revision = await xrm.WebApi.retrieveRecord(
        'asx_revision',
        saved.RevisionId,
        '?$select=asx_version',
      );
      if (session !== state.session) return;
      if (revision?.asx_version) state.editBase.Version = revision.asx_version;
      const template = await reloadTemplate(saved.TemplateId);
      if (session !== state.session) return;
      state.template = template;
      await loadTemplates();
    } catch (error) {
      const refused = new Error(
        'Draft saved, but the page could not refresh: ' + (error.message || String(error)),
      );
      refused.saved = true;
      throw refused;
    }
  }
  // Unsaved edits and a save in flight ask before the page changes; Save draft saves them first.
  // A version opened read-only has nothing to save: leaving it never asks.
  ui.setDirtyGuard(() =>
    state.root && !state.readOnly && (state.unsaved || state.saving || state.saveState === 'error')
      ? {
          template: templateLabel(),
          save: flushSave,
          discard: () => {
            clearTimeout(state.saveTimer);
            state.unsaved = false;
          },
        }
      : null,
  );

  $('chooseRecord').onclick = async () => {
    const picked = await xrm.Utility.lookupObjects({
      entityTypes: [state.root.LogicalName],
      defaultEntityType: state.root.LogicalName,
      allowMultiSelect: false,
    });
    if (!picked?.length) return;
    state.previewRecord = {
      id: picked[0].id.replace(/[{}]/g, '').toLowerCase(),
      name: picked[0].name,
    };
    $('preview-record-name').textContent = state.previewRecord.name;
    $('chooseRecord').focus();
    await runPreview();
  };
  $('refreshPreview').onclick = () => runPreview();
  // Previews the saved revision, or the current edits when there are any.
  async function runPreview() {
    if (!state.previewRecord) return;
    const trees = $('previewTrees');
    trees.setAttribute('aria-busy', 'true');
    trees.classList.remove('is-stale');
    trees.replaceChildren(...[0, 1, 2].map(() => el('div', null, 'skeleton')));
    try {
      const request =
        state.unsaved || !state.saved
          ? { RevisionId: '', Draft: payload(), RecordId: state.previewRecord.id }
          : { RevisionId: state.saved.RevisionId, RecordId: state.previewRecord.id };
      const plan = JSON.parse(
        await api('asx_PreviewTemplate', { Request: JSON.stringify(request) }),
      );
      renderPreview(plan);
      state.previewed = true;
      $('preview-stale').hidden = true;
      $('refreshPreview').hidden = true;
      const folders = plan.Folders || [];
      const sections = new Set(folders.map((f) => f.Section)).size;
      $('preview-status').textContent =
        'Preview updated: ' +
        plural(sections, 'destination', 'destinations') +
        ', ' +
        plural(folders.length, 'folder', 'folders') +
        '.';
    } catch (error) {
      trees.replaceChildren(
        el('p', 'Preview failed: ' + (error.message || String(error)), 'error'),
      );
      $('refreshPreview').hidden = false;
    } finally {
      trees.removeAttribute('aria-busy');
    }
  }

  // Re-run for existing records, of the template the editor or the overview shows.
  const rerunTemplate = () => (state.view === 'edit' ? state.template : state.overview?.template);
  async function openRerun(invoker = document.activeElement) {
    const t = rerunTemplate();
    panel = ui.sidePanel($('rerun-panel'), invoker);
    $('rerun-impact').replaceChildren();
    $('rerun-these').hidden = true;
    ui.clearFeedback('rerun');
    preventError(null);
    const [count, runs] = await Promise.all([
      ui
        .api('asx_ManageWork', {
          Command: 'CountRecords',
          TemplateId: t.asx_templateid,
        })
        .catch(() => null),
      ui
        .api('asx_ManageWork', { Command: 'ListProblems', List: 'TemplateRuns' })
        .catch(() => ({ Problems: [] })),
    ]);
    state.run =
      (runs?.Problems || []).find(
        (p) => same(p.Run?.TemplateId, t.asx_templateid) && ACTIVE.includes(p.Run.State),
      ) || null;
    const total = count?.Run
      ? (count.Run.TotalEstimated ? 'about ' : '') + count.Run.Total.toLocaleString('en-US')
      : null;
    state.runTotal = total;
    $('rerun-all').textContent = total ? 'Re-run all ' + total + ' records' : 'Re-run all records';
    $('rerun-all').hidden = !!state.run;
    $('rerun-progress').hidden = !state.run;
    if (state.run) {
      const run = state.run.Run;
      $('rerun-progress').replaceChildren(
        el(
          'span',
          'Re-run in progress: ' +
            run.Planned.toLocaleString('en-US') +
            ' of ' +
            (run.TotalEstimated ? 'about ' : '') +
            run.Total.toLocaleString('en-US') +
            ' · ',
        ),
        button('Open in Monitor', () => ui.navigate('monitor', { run: state.run.Key }), 'link'),
      );
    }
    // Every re-run action needs the Operator role; one reason line names it for all three.
    const role = ui.needs(OPERATOR);
    for (const id of ['rerun-all', 'rerun-preview', 'rerun-these'])
      ui.disable($(id), 'rerun-role-reason', role);
    ($('rerun-all').hidden ? $('rerun-preview') : $('rerun-all')).focus();
  }
  $('rerun-all').onclick = async () => {
    if (ui.blocked($('rerun-all'))) return;
    const t = rerunTemplate();
    // A re-run applies the published version.
    const version = templateState(t).live ?? state.editBase?.Version;
    const ok = await ui.confirmInline($('rerun-all'), {
      text:
        'Re-run v' +
        version +
        ' for all ' +
        (state.runTotal ? state.runTotal + ' ' : '') +
        tableName(t.asx_table) +
        ' records? Documents works through them in the background, after other work, so this can take a while. You can close this page and follow it in Monitor.',
      confirm: 'Re-run all records',
      keep: 'Not now',
    });
    if (!ok) return;
    await ui.busy($('rerun-all'), 'Starting…', 'rerun', async () => {
      const started = await ui.api('asx_ManageWork', {
        Command: 'StartTemplateRun',
        TemplateId: t.asx_templateid,
        RequestId: crypto.randomUUID(),
      });
      ui.feedback('rerun', ['Re-run started.', ...(started.Notices || [])].join(' '), 'success', {
        label: 'Follow it in Monitor →',
        onClick: () => ui.navigate('monitor', { run: started.Key }),
      });
    });
  };
  // A pick over the preview bound is refused at the button, with the traced reason (row 18).
  function preventError(text) {
    const control = $('rerun-preview');
    // Made on demand, so it is looked up directly rather than as page markup.
    let note = document.getElementById('rerun-reason');
    if (!text) {
      note?.remove();
      describe(control, null);
      return;
    }
    if (!note) {
      note = el('p', null, 'error');
      note.id = 'rerun-reason';
      note.setAttribute('role', 'alert');
      control.closest('[data-actions]').after(note);
    }
    note.textContent = text;
    describe(control, 'rerun-reason');
  }
  // Adds or drops the preview error in the button's description, keeping any role reason.
  function describe(control, id) {
    const refs = (control.getAttribute('aria-describedby') || '')
      .split(' ')
      .filter((r) => r && r !== 'rerun-reason');
    if (id) refs.push(id);
    if (refs.length) control.setAttribute('aria-describedby', refs.join(' '));
    else control.removeAttribute('aria-describedby');
  }
  $('rerun-preview').onclick = async () => {
    if (ui.blocked($('rerun-preview'))) return;
    const t = rerunTemplate();
    const picked = await xrm.Utility.lookupObjects({
      entityTypes: [t.asx_table],
      defaultEntityType: t.asx_table,
      allowMultiSelect: true,
    });
    if (!picked?.length) return;
    const bound = ui.BOUNDS.previewRecords;
    if (picked.length > bound) {
      preventError(
        'Preview covers up to ' +
          bound +
          " records at a time so it finishes within Dataverse's 2-minute limit. To re-run every record, use Re-run all.",
      );
      return;
    }
    preventError(null);
    await ui.busy($('rerun-preview'), 'Previewing…', 'rerun', async () => {
      state.batch = await ui.api('asx_ManageWork', {
        Command: 'PreviewBatch',
        TemplateId: t.asx_templateid,
        RecordIds: picked.map((r) => r.id.replace(/[{}]/g, '').toLowerCase()),
        RequestId: crypto.randomUUID(),
      });
      const names = new Map(picked.map((r) => [r.id.replace(/[{}]/g, '').toLowerCase(), r.name]));
      $('rerun-impact').replaceChildren(
        ...(state.batch.Batch?.Records || []).map((record) => {
          const card = el('section', null, 'impact');
          card.append(el('h4', names.get(String(record.Id).toLowerCase()) || 'Record'));
          const lines = el('ul');
          lines.append(
            ...(record.Impact || []).map((line) =>
              el('li', line.replace(/^Resolve or create expected path: [^/]+\//, '')),
            ),
          );
          card.append(lines);
          return card;
        }),
      );
      $('rerun-these').hidden = false;
    });
  };
  $('rerun-these').onclick = () =>
    ui.blocked($('rerun-these')) ||
    ui.busy($('rerun-these'), 'Re-running…', 'rerun', async () => {
      await ui.api('asx_ManageWork', {
        Command: 'QueueBatch',
        Key: state.batch.Key,
        RowVersion: state.batch.RowVersion,
      });
      $('rerun-these').hidden = true;
      ui.feedback(
        'rerun',
        'Re-run queued for ' +
          plural((state.batch.Batch?.Records || []).length, 'record', 'records') +
          '.',
      );
    });

  window.AsxdAdmin = {
    refreshCatalog: async () => {
      const rows = await xrm.WebApi.retrieveMultipleRecords(
        'asx_library',
        '?$select=asx_libraryid,asx_name,_asx_siteid_value,asx_entryurl&$filter=asx_approved eq true',
      );
      if (rows.nextLink) throw new Error('Use a narrower destination catalog.');
      state.libraries = rows.entities;
      const sites = await xrm.WebApi.retrieveMultipleRecords(
        'asx_site',
        '?$select=asx_siteid,asx_name&$filter=asx_approved eq true',
      );
      state.sites = sites.entities;
      chrome();
    },
    changesSince,
    nextKey,
  };
  // The first template listed, in the list's order.
  function firstListed() {
    const { enabled, disabled } = listTables();
    for (const table of [...enabled, ...disabled]) {
      const found = state.templates.find((t) => t.asx_table === table);
      if (found) return found;
    }
    return null;
  }
  async function start() {
    render();
    resetPreview();
    if (!xrm?.WebApi || !xrm?.Utility) return;
    ui.problemPill($('overview-problems'));
    renderOverview();
    await task(async () => {
      state.tables = (
        await all(
          'EntityDefinitions?$select=LogicalName,DisplayName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,IsDocumentManagementEnabled&$filter=IsDocumentManagementEnabled eq true',
        )
      )
        .filter((t) => t.PrimaryNameAttribute)
        .sort((a, b) => display(a).localeCompare(display(b)));
      const catalog = await xrm.WebApi.retrieveMultipleRecords(
        'asx_library',
        '?$select=asx_libraryid,asx_name,_asx_siteid_value,asx_entryurl&$filter=asx_approved eq true',
      );
      if (catalog.nextLink)
        throw new Error('Approved library catalog exceeded the current completeness bound.');
      state.libraries = catalog.entities;
      const sites = await xrm.WebApi.retrieveMultipleRecords(
        'asx_site',
        '?$select=asx_siteid,asx_name&$filter=asx_approved eq true',
      );
      if (sites.nextLink) throw new Error('Approved site catalog is incomplete.');
      state.sites = sites.entities;
      await loadEnabledTables();
      state.loaded = true;
      await loadTemplates();
      // A link to a template shows it; otherwise the first one listed.
      const linked = ui.deeplink()?.template;
      const shown = state.templates.find((t) => same(t.asx_templateid, linked)) || firstListed();
      if (shown) await showOverview(shown.asx_templateid);
      else renderOverview();
      // A link to a step opens the template's editor there.
      const step = Number(ui.deeplink()?.step);
      if (shown && same(shown.asx_templateid, linked) && [1, 2, 3].includes(step))
        await openEditor(step);
    });
  }
  ui.onTab('templates', start);
})();
