'use strict';
// Folder templates (spec 3.1): the Tables rail, the template bar with its version chip and
// actions, the folder tree and folder settings, the condition builder, the preview of the current
// edits, and re-runs for existing records. Text is only ever set with textContent.
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
    template: null,
    templates: [],
    enabledTables: [],
    readiness: null,
    expandedTables: new Set(),
    sectionSequence: 0,
    editBase: null,
    busy: false,
    // The rail's first load is done: its empty state is real, not "still loading".
    loaded: false,
    // Edits since the last load or save: the chip, Save draft, Publish and the guards follow it.
    unsaved: false,
    run: null,
    runTotal: null,
    batch: null,
    previewRecord: null,
    previewed: false,
    // The data-focus-key that takes focus after the next render (spec 5.2).
    focusKey: null,
    // Related tables whose fields could not be read, even after a retry.
    failedTables: new Set(),
    // Which template open the background preload belongs to.
    preloadId: 0,
    // A rail redraw waits for the rail's open confirmation.
    railStale: false,
  };
  const metadata = new Map();
  const xrm = window.parent?.Xrm || window.Xrm;
  const ui = window.AsxdUi;
  // The rail's ＋ Add table moves into the rail's empty state and back to the rail header.
  const addTableButton = $('addTable');
  // Results of an action show in the feedback line of the tab that ran it (spec 2.5).
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
  const plural = (n, one, many) => n.toLocaleString('en-US') + ' ' + (n === 1 ? one : many);

  // The version chip (spec 3.1): from the loaded revision's number and status and the schedule.
  function versionText() {
    const base = state.editBase;
    if (!base) return state.unsaved ? 'Draft v1 · unsaved changes' : 'Draft v1';
    if (base.Status === 'Draft')
      return 'Draft v' + base.Version + (state.unsaved ? ' · unsaved changes' : '');
    if (state.unsaved) return 'Draft v' + (base.Version + 1) + ' · unsaved changes';
    const t = state.template;
    const text = 'Published v' + base.Version;
    if (t?.asx_disabled) return text + ' · Off';
    if (t?.asx_startsutc && new Date(t.asx_startsutc) > new Date())
      return text + ' · starts ' + ui.time(t.asx_startsutc).textContent.split(',')[0];
    if (t?.asx_endsutc && new Date(t.asx_endsutc) <= new Date()) return text + ' · ended';
    return text;
  }
  function publishReason() {
    if (!ui.can('prvCreateasx_publication')) return ui.needs('prvCreateasx_publication');
    if (state.unsaved || !state.saved) return 'Save your changes first';
    if (state.saved.Status !== 'Draft') return 'Already published';
    return null;
  }
  function controls() {
    const open = !!state.root;
    $('template-bar').hidden = !open;
    $('authorWorkspace').hidden = !open;
    $('no-template').hidden = open;
    $('templateName').disabled = state.busy || !!state.template;
    const version = versionText();
    $('version-chip').textContent = version;
    $('version-chip-desc').textContent =
      state.editBase?.Status === 'Published' && !state.unsaved
        ? 'Published v' +
          state.editBase.Version +
          ' keeps running while you edit. Saving creates Draft v' +
          (state.editBase.Version + 1) +
          '.'
        : version + '.';
    const save = $('save');
    save.classList.toggle('has-dot', state.unsaved);
    if (state.unsaved) save.setAttribute('aria-label', 'Save draft, unsaved changes');
    else save.removeAttribute('aria-label');
    save.disabled = state.busy || !open || !state.sections.length;
    // A symbol button names its object (spec 5.1): "More actions for Account onboarding".
    $('template-menu').setAttribute(
      'aria-label',
      'More actions for ' + ($('templateName').value.trim() || 'this template'),
    );
    const reason = publishReason();
    ui.disable($('publish'), 'publish-reason', reason);
    $('publish').classList.toggle('primary', !reason);
    $('publish').classList.toggle('secondary', !!reason);
    const atLimit = state.sections.length >= ui.BOUNDS.destinations;
    $('addDestination').hidden = !state.libraries.length;
    $('no-library').hidden = !!state.libraries.length || !open;
    ui.disable(
      $('addDestination'),
      'destination-reason',
      atLimit
        ? 'A template can have up to ' +
            ui.BOUNDS.destinations +
            ' destinations, because each record plans all of them in one step that Dataverse stops after 2 minutes.'
        : null,
    );
    $('menu-history').hidden = !state.template;
    $('menu-delete').hidden = !state.template;
    $('menu-rerun').hidden =
      !state.template?._asx_publishedrevisionid_value && state.editBase?.Status !== 'Published';
    ui.disable(
      $('saveAvailability'),
      'schedule-reason',
      state.template ? null : 'Save the template first.',
    );
  }
  // An edit: the chip, Save draft and Publish follow it, and the preview dims until refreshed.
  function dirty() {
    state.unsaved = true;
    $('preview-stale').hidden = !state.previewed;
    $('previewTrees').classList.toggle('is-stale', state.previewed);
    $('refreshPreview').hidden = !state.previewRecord;
    controls();
  }
  async function task(action) {
    if (state.busy) return;
    state.busy = true;
    controls();
    try {
      await action();
    } catch (error) {
      message(error.message || String(error), true);
    } finally {
      state.busy = false;
      controls();
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
  // A folder name with its fields shown by label: {root.name} → [Account Name] (F-23).
  function readable(name) {
    return String(name || '').replace(
      /\{([a-z0-9_]+)\.([a-z0-9_]+)\}/gi,
      (match, alias, column) => {
        const source = state.sources.find((s) => s.Alias === alias);
        const field = source?.columns.find((c) => c.Name === column);
        if (!source || !field) return match;
        return '[' + (alias === 'root' ? '' : via(source) + ' › ') + field.Label + ']';
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
  // not used yet are disabled (spec 6.9).
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
  // The value control for a field kind (spec 3.1). key is the row's focus-key prefix.
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
  // Checks every condition at its control; the first invalid control takes focus (spec 5.2).
  function validate() {
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
    const first = document.querySelector('[data-focus-key="' + problems[0][0] + '"]');
    if (first) first.focus();
    else
      ui.feedback(
        'templates',
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
  function destinationName(section) {
    return section.Name && section.Name !== section.Key
      ? section.Name
      : 'Destination ' + (state.sections.indexOf(section) + 1);
  }
  const nodeKey = (section, folder) => 'node:' + section.Key + ':' + folder.Key;
  function addChild(section, folder) {
    do {
      section.nodeSequence++;
    } while (section.Folders.some((f) => f.Key === 'folder_' + section.nodeSequence));
    const child = {
      Key: 'folder_' + section.nodeSequence,
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
    $('destinationCount').textContent = plural(
      state.sections.length,
      'destination',
      'destinations',
    );
    if (!state.sections.length) list.append(el('p', 'No folders yet', 'empty'));
    for (const section of state.sections) {
      const card = el(
        'section',
        null,
        'destination-section' + (section === selectedSection ? ' selected-section' : ''),
      );
      card.append(el('h3', destinationName(section), 'section-heading'));
      const tree = el('div', null, 'tree'),
        visited = new Set();
      const addNode = (folder, depth) => {
        if (visited.has(folder)) return;
        visited.add(folder);
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
        if (section === selectedSection && folder === selectedFolder)
          node.setAttribute('aria-current', 'true');
        row.append(node);
        if (folder.Condition) {
          // A labelled marker, not colour only (F-16); the hidden text describes the node.
          const mark = el('span', '◆', 'conditional-mark');
          mark.setAttribute('aria-hidden', 'true');
          const text = el('span', ' (conditional)', 'sr-only');
          text.id = 'conditional-' + section.Key + '-' + folder.Key;
          node.setAttribute('aria-describedby', text.id);
          row.append(mark, text);
        }
        tree.append(row);
        section.Folders.filter((f) => f.Parent === folder.Key).forEach((f) =>
          addNode(f, depth + 1),
        );
      };
      section.Folders.filter((f) => !f.Parent).forEach((f) => addNode(f, 0));
      section.Folders.filter((f) => !visited.has(f)).forEach((f) => addNode(f, 0));
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
      library = libraryFor(section),
      siteId = library?._asx_siteid_value || '';
    const heading = el('h3', destinationName(section));
    editor.append(heading);
    const destination = el('input');
    destination.value = section.Name === section.Key ? '' : section.Name || '';
    destination.maxLength = 200;
    destination.setAttribute('maxlength', '200');
    keyed(destination, 'edit:destination');
    destination.oninput = () => {
      section.Name = destination.value;
      heading.textContent = destinationName(section);
      dirty();
      renderFolders();
    };
    destination.onchange = destination.oninput;
    editor.append(label('Destination name', destination));
    const pickers = el('div', null, 'picker-grid'),
      sites = state.sites.filter((s) =>
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
    keyed(site, 'edit:site');
    const libraryPicker = select(
      [
        { value: '', label: 'Choose a library' },
        ...state.libraries
          .filter((l) => l._asx_siteid_value === siteId)
          .map((l) => ({ value: l.asx_libraryid, label: l.asx_name })),
      ],
      library?.asx_libraryid || '',
      (v) => {
        section.LibraryId = v;
        ui.withFocus(render);
      },
    );
    keyed(libraryPicker, 'edit:library');
    pickers.append(label('Site', site), label('Library', libraryPicker));
    editor.append(pickers);
    editor.append(
      button(
        'View this library’s team access →',
        () => ui.navigate('access', { library: section.LibraryId }),
        'link',
      ),
    );
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
    const removeDestination = button(
      'Remove destination',
      async () => {
        const count = section.Folders.length;
        const ok = await ui.confirmInline(removeDestination, {
          text:
            'Remove ' +
            destinationName(section) +
            ' and its ' +
            plural(count, 'folder', 'folders') +
            ' from this draft? Nothing changes in SharePoint until you publish.',
          confirm: 'Remove destination',
          keep: 'Keep destination',
          danger: true,
        });
        if (!ok) return;
        state.sections.splice(state.sections.indexOf(section), 1);
        selectedSection = null;
        selectedFolder = null;
        dirty();
        render();
        $('folders-heading').focus();
      },
      'danger',
    );
    keyed(removeDestination, 'edit:remove-destination');
    actions.append(removeDestination);
  }
  function render() {
    normalizeSelection();
    renderFolders();
    renderEditor();
    controls();
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
  // The template row as the bar, chip and Schedule read it. The one copy of this $select.
  const TEMPLATE_SELECT =
    '?$select=asx_templateid,asx_name,asx_table,asx_disabled,asx_startsutc,asx_endsutc,_asx_publishedrevisionid_value';
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

  // The ⋯ menu (spec 5.6): Enter, Space or Down opens; Up and Down move; Escape closes.
  const menuItems = () =>
    [...$('template-menu-list').querySelectorAll('[role=menuitem]')].filter((i) => !i.hidden);
  function openMenu() {
    $('template-menu-list').hidden = false;
    $('template-menu').setAttribute('aria-expanded', 'true');
    menuItems()[0]?.focus();
  }
  function closeMenu() {
    $('template-menu-list').hidden = true;
    $('template-menu').setAttribute('aria-expanded', 'false');
    $('template-menu').focus();
  }
  $('template-menu').onclick = () => {
    if ($('template-menu').dataset.busy) return;
    if ($('template-menu-list').hidden) openMenu();
    else closeMenu();
  };
  $('template-menu').addEventListener('keydown', (event) => {
    if (event.key !== 'ArrowDown') return;
    event.preventDefault();
    openMenu();
  });
  $('template-menu-list').addEventListener('keydown', (event) => {
    const items = menuItems();
    const index = items.indexOf(document.activeElement);
    if (event.key === 'Escape') {
      event.preventDefault();
      closeMenu();
    } else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      event.preventDefault();
      items[(index + (event.key === 'ArrowDown' ? 1 : items.length - 1)) % items.length].focus();
    }
  });
  const fromMenu = (fn) => () => {
    $('template-menu-list').hidden = true;
    $('template-menu').setAttribute('aria-expanded', 'false');
    return fn();
  };
  $('menu-history').onclick = fromMenu(() => task(openHistory));
  $('menu-schedule').onclick = fromMenu(openSchedule);
  $('menu-rerun').onclick = fromMenu(openRerun);
  $('menu-delete').onclick = fromMenu(deleteTemplate);
  // One panel under the template bar at a time; Close returns to the ⋯ button.
  const PANELS = ['history-panel', 'schedule-panel', 'rerun-panel'];
  function showPanel(id) {
    for (const panel of PANELS) $(panel).hidden = panel !== id;
  }
  for (const close of document.querySelectorAll('.panel-close'))
    close.onclick = () => {
      $(close.dataset.close).hidden = true;
      $('template-menu').focus();
    };

  $('publish').onclick = async () => {
    if (ui.blocked($('publish'))) return;
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
      const result = JSON.parse(
        await api('asx_PublishTemplate', {
          RevisionId: state.saved.RevisionId,
          RowVersion: state.saved.RowVersion,
        }),
      );
      // Publishing changes the revision row, so its row version is read back with it; the next
      // Save draft would be refused with the old one.
      await reloadVersion(state.saved.RevisionId);
      state.template = await reloadTemplate(state.template.asx_templateid);
      controls();
      ui.feedback(
        'templates',
        ['Published v' + next + '.', ...(result.Notices || [])].join(' '),
        'success',
        {
          label: 'Re-run existing records…',
          onClick: openRerun,
        },
      );
    });
    controls();
  };

  async function openHistory() {
    const rows = await xrm.WebApi.retrieveMultipleRecords(
      'asx_revision',
      '?$select=asx_revisionid,asx_version,asx_status&$filter=_asx_templateid_value eq ' +
        state.template.asx_templateid +
        '&$orderby=asx_version desc',
    );
    $('history-list').replaceChildren(
      ...rows.entities.map((r) => {
        const item = el('li');
        const pick = button('v' + r.asx_version + ' · ' + r.asx_status, () =>
          openVersion(pick, r.asx_revisionid, r.asx_version),
        );
        if (r.asx_revisionid === state.saved?.RevisionId) pick.setAttribute('aria-current', 'true');
        item.append(pick);
        return item;
      }),
    );
    showPanel('history-panel');
    $('history-list').querySelector('button')?.focus();
  }
  async function openVersion(control, revisionId, version) {
    if (state.unsaved) {
      const ok = await ui.confirmInline(control, {
        text: 'Open v' + version + '? Your unsaved changes to this draft are discarded.',
        confirm: 'Open v' + version,
        keep: 'Keep editing',
      });
      if (!ok) return;
    }
    await task(() => reloadVersion(revisionId));
    $('history-panel').hidden = true;
    $('version-chip').focus();
  }

  function openSchedule() {
    const t = state.template;
    $('schedule-on').setAttribute('aria-checked', String(!t?.asx_disabled));
    $('templateStart').value = localTime(t?.asx_startsutc);
    $('templateEnd').value = localTime(t?.asx_endsutc);
    ui.clearFeedback('schedule');
    showPanel('schedule-panel');
    $('schedule-on').focus();
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
      await xrm.WebApi.updateRecord('asx_template', state.template.asx_templateid, {
        asx_disabled: $('schedule-on').getAttribute('aria-checked') !== 'true',
        asx_startsutc: start,
        asx_endsutc: end,
      });
      state.template = await reloadTemplate(state.template.asx_templateid);
      controls();
      ui.feedback('schedule', 'Schedule saved.');
    });
  };

  async function deleteTemplate() {
    const name = state.template.asx_name || $('templateName').value;
    const table = state.root.LogicalName;
    const ok = await ui.confirmInline($('template-menu'), {
      text:
        'Delete ' +
        name +
        ' and all its versions? No new folder work starts for it. Folders, documents and access in SharePoint stay as they are. Work already sent to SharePoint may still finish.',
      confirm: 'Delete template',
      keep: 'Keep template',
      danger: true,
    });
    if (!ok) return;
    const done = await ui.busy($('template-menu'), 'Deleting…', 'templates', async () => {
      await xrm.WebApi.deleteRecord('asx_template', state.template.asx_templateid);
      Object.assign(state, {
        saved: null,
        editBase: null,
        sections: [],
        template: null,
        root: null,
        unsaved: false,
      });
      for (const panel of PANELS) $(panel).hidden = true;
      await loadTemplates();
      render();
      ui.feedback('templates', name + ' deleted. Nothing in SharePoint changed.');
      return true;
    });
    if (!done) return;
    // Focus stays in the rail (ruling 2): the table's next template, else the Tables heading.
    const next = $('templateTree').querySelector('[data-table="' + table + '"] .template-item');
    (next || $('tables-heading')).focus();
  }

  // Unsaved edits ask before another template replaces them; true when it may.
  async function mayDiscard(control, what) {
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
    const sources = await Promise.all(
      loaded.Draft.Sources.map(async (source) => ({
        Alias: source.Alias,
        Table: source.Table,
        Lookup: source.Lookup,
        columns: (await fields(source.Table)).columns,
      })),
    );
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
    state.sources = sources;
    state.sections = loaded.Draft.Destinations.map((d) => ({
      ...d,
      nodeSequence: 0,
      Folders: d.Folders.map((f) => ({ ...f, Condition: f.Condition ? group(f.Condition) : null })),
    }));
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
  // Each table's fields take about seven metadata requests at once, and Dataverse serves 52
  // concurrent requests per user before it answers 429. Four tables at a time stay well inside
  // that, beside the page's other reads.
  const PRELOAD_TABLES = 4;
  async function preload(tables, id) {
    const queue = tables.filter((t) => !metadata.has(t));
    queue.forEach((t) => state.failedTables.delete(t));
    const worker = async () => {
      while (queue.length) {
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
  // answering a confirmation there; the next render picks the fields up then.
  function refreshPickers() {
    if (!state.root) return;
    const editor = $('folderEditor');
    const active = document.activeElement;
    if (
      editor.querySelector('.confirm[role=group]') ||
      (editor.contains(active) && active.tagName === 'INPUT')
    )
      return;
    ui.withFocus(render);
  }
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
    state.expandedTables.add(table);
    Object.assign(state, {
      root,
      editBase: null,
      saved: null,
      unsaved: false,
      sections: [],
      run: null,
      batch: null,
    });
    selectedSection = null;
    selectedFolder = null;
    for (const panel of PANELS) $(panel).hidden = true;
    ui.clearFeedback('templates');
    resetPreview();
    renderTemplateTree();
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
    if (latest) await reloadVersion(latest);
    else render();
    preload([...new Set(meta.lookups.flatMap((l) => l.Targets))], id);
  }
  addTableButton.onclick = () => {
    renderEnablePicker();
    $('tablePicker').hidden = !$('tablePicker').hidden;
    if (!$('tablePicker').hidden) $('enableTable').focus();
  };
  function renderEnablePicker() {
    $('enableTable').replaceChildren(
      option('', 'Choose a table'),
      ...state.tables
        .filter((t) => !state.enabledTables.includes(t.LogicalName))
        .map((t) => option(t.LogicalName, display(t))),
    );
    $('enableTable').value = '';
  }
  $('enableTable').onchange = () => {
    const name = $('enableTable').value;
    if (name) return task(() => enableTable(name));
    return undefined;
  };
  $('go-access').onclick = () => ui.navigate('access');
  async function loadEnabledTables() {
    const rows = [];
    let options = '?$select=asx_logicalname&$orderby=asx_logicalname';
    do {
      const result = await xrm.WebApi.retrieveMultipleRecords('asx_runtimetable', options);
      rows.push(...result.entities);
      options = result.nextLink
        ? new URL(result.nextLink, xrm.Utility.getGlobalContext().getClientUrl()).search
        : null;
    } while (options);
    state.enabledTables = [...new Set(rows.map((r) => r.asx_logicalname).filter(Boolean))].sort();
  }
  const runtimeCommand = (payload) => ui.api('asx_RuntimeAdmin', payload);
  function setReadiness(result) {
    const readiness = result?.Registration?.Readiness;
    state.readiness = Array.isArray(readiness)
      ? Object.fromEntries(readiness.map((r) => [r.Scope, r.Status]))
      : null;
  }
  // One of three readiness words per table (spec 3.1 Tables rail).
  function readiness(table) {
    const status = state.readiness?.[table];
    if (!status) return null;
    if (status !== 'Ready') return 'repair';
    return ui.runtime() && !ui.runtime().Enabled ? 'Ready · automation paused' : 'Ready';
  }
  // Applies a Get-shaped result from AddTable/RemoveTable to the panel, and shares it with the
  // chip and the other tabs through the shell's runtime.
  async function tableChanged(result) {
    setReadiness(result);
    ui.setRuntime(result);
    await loadEnabledTables();
    renderEnablePicker();
    $('tablePicker').hidden = true;
    renderTemplateTree();
  }
  async function enableTable(name) {
    await tableChanged(await runtimeCommand({ Command: 'AddTable', Table: name }));
    message(tableName(name) + ' added.');
  }
  async function disableTable(name) {
    await tableChanged(await runtimeCommand({ Command: 'RemoveTable', Table: name }));
    message(tableName(name) + ' removed. Its templates are kept.');
  }
  // A confirmation in the rail. A redraw asked for while it is open waits, and runs once it closes.
  async function railAsk(control, options) {
    try {
      return await ui.confirmInline(control, options);
    } finally {
      if (state.railStale) renderTemplateTree();
    }
  }
  async function removeTable(table, control) {
    const ok = await railAsk(control, {
      text:
        'Stop creating folders for ' +
        tableName(table) +
        '? Queued folder work for this table is cancelled. Templates are kept, and nothing in SharePoint is deleted.',
      confirm: 'Remove table',
      keep: 'Keep table',
      danger: true,
    });
    if (ok) await task(() => disableTable(table));
  }
  async function loadTemplates() {
    const rows = [];
    let options = '?$select=asx_templateid,asx_name,asx_table&$orderby=asx_name';
    do {
      const result = await xrm.WebApi.retrieveMultipleRecords('asx_template', options);
      rows.push(...result.entities);
      options = result.nextLink
        ? new URL(result.nextLink, xrm.Utility.getGlobalContext().getClientUrl()).search
        : null;
    } while (options);
    state.templates = rows;
    renderTemplateTree();
  }
  // Opens a template (or a new one) from the rail, asking first when edits would be lost.
  async function pick(control, table, templateId, what) {
    if (!(await mayDiscard(control, what))) return;
    await task(() => selectTemplate(table, templateId));
  }
  function renderTemplateTree() {
    const tree = $('templateTree');
    // Never under an open confirmation (as Monitor's lists): railAsk redraws when it closes.
    state.railStale = !!tree.querySelector('.confirm[role=group]');
    if (state.railStale) return;
    // ＋ Add table goes home before the rail is redrawn, so it is never dropped with it.
    $('tables-header').append(addTableButton);
    tree.replaceChildren();
    if (!state.loaded) return;
    const templated = [...new Set(state.templates.map((t) => t.asx_table))].filter(Boolean);
    const enabled = [...state.enabledTables].sort();
    const disabled = templated.filter((t) => !enabled.includes(t)).sort();
    renderNoTemplate(!enabled.length && !disabled.length);
    if (!enabled.length && !disabled.length) {
      const empty = el('div', null, 'empty');
      empty.append(el('p', 'No tables yet'));
      tree.append(empty);
      empty.append(addTableButton);
      return;
    }
    const entry = (table, isEnabled) => {
      const wrap = el('div', null, 'table-entry'),
        section = el('details'),
        summary = el('summary');
      wrap.dataset.table = table;
      summary.append(el('span', tableName(table), 'table-name'));
      const ready = isEnabled ? readiness(table) : null;
      if (ready && ready !== 'repair') summary.append(el('span', ready, 'status'));
      section.open = state.expandedTables.has(table);
      section.ontoggle = () => {
        if (section.open) state.expandedTables.add(table);
        else state.expandedTables.delete(table);
      };
      section.append(summary);
      const list = el('ul', null, 'template-list');
      // A confirmation from a template button renders after the list, not inside it.
      list.setAttribute('data-actions', '');
      for (const template of state.templates.filter((t) => t.asx_table === table)) {
        const item = el('li');
        const open = button(
          template.asx_name || tableName(table),
          () => pick(open, table, template.asx_templateid, template.asx_name || 'this template'),
          'template-item',
        );
        keyed(open, 'template:' + template.asx_templateid);
        if (template.asx_templateid === state.template?.asx_templateid)
          open.setAttribute('aria-current', 'true');
        item.append(open);
        list.append(item);
      }
      if (isEnabled) {
        const item = el('li');
        const create = button(
          '＋ New template',
          () => pick(create, table, null, 'a new template'),
          'new-template',
        );
        item.append(create);
        list.append(item);
      }
      section.append(list);
      wrap.append(section);
      if (ready === 'repair')
        wrap.append(button('Needs repair', () => ui.navigate('settings', { table }), 'link'));
      const action = isEnabled
        ? button('Remove', () => removeTable(table, action), 'link danger')
        : button('Enable', () => task(() => enableTable(table)), 'link');
      action.setAttribute('aria-label', (isEnabled ? 'Remove ' : 'Enable ') + tableName(table));
      wrap.append(action);
      tree.append(wrap);
    };
    enabled.forEach((table) => entry(table, true));
    if (disabled.length) {
      tree.append(el('h3', 'Not enabled', 'rail-group'));
      disabled.forEach((table) => entry(table, false));
    }
  }
  // The main area when no template is open (spec 3.1, rows 3, 6, 42, 66).
  function renderNoTemplate(noTables) {
    const panel = $('no-template');
    panel.replaceChildren(el('h2', 'No template selected'));
    if (noTables) panel.append(button('＋ Add table', () => addTableButton.click()));
    else panel.append(el('p', 'Pick a template in Tables, or create one with ＋ New template.'));
  }
  $('addDestination').onclick = () => {
    if (ui.blocked($('addDestination')) || !state.libraries.length) return;
    do {
      state.sectionSequence++;
    } while (state.sections.some((s) => s.Key === 'destination_' + state.sectionSequence));
    const library = state.libraries[0];
    state.sections.push({
      Key: 'destination_' + state.sectionSequence,
      // A new destination is named after its library (row 45).
      Name: library.asx_name,
      LibraryId: library.asx_libraryid,
      nodeSequence: 0,
      Folders: [
        {
          Key: 'root',
          Parent: null,
          Name: '{root.' + state.root.PrimaryNameAttribute + '}',
          Condition: null,
        },
      ],
    });
    selectedSection = state.sections[state.sections.length - 1];
    selectedFolder = selectedSection.Folders[0];
    state.focusKey = nodeKey(selectedSection, selectedFolder);
    dirty();
    render();
  };
  $('templateName').oninput = () => {
    if (!state.template) dirty();
  };

  // Saves the draft. The version follows CreateDraftApi: a Draft is saved in place, anything
  // else starts the next version; a new template starts at v1.
  async function save() {
    if (!validate()) throw new Error('Fix the highlighted conditions first.');
    const base = state.editBase;
    const name = $('templateName').value.trim();
    const saved = JSON.parse(await api('asx_CreateDraft', { Request: JSON.stringify(payload()) }));
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
    state.unsaved = false;
    controls();
    try {
      // The saved version's own number: a version opened from Version history is not the latest.
      const revision = await xrm.WebApi.retrieveRecord(
        'asx_revision',
        saved.RevisionId,
        '?$select=asx_version',
      );
      if (revision?.asx_version) state.editBase.Version = revision.asx_version;
      state.template = await reloadTemplate(saved.TemplateId);
      await loadTemplates();
    } catch (error) {
      controls();
      throw new Error(
        'Draft saved, but the page could not refresh: ' + (error.message || String(error)),
      );
    }
    controls();
    if (state.previewRecord) await runPreview();
  }
  $('save').onclick = async () => {
    await ui.busy($('save'), 'Saving…', 'templates', async () => {
      await save();
      ui.feedback('templates', 'Draft saved.');
    });
    controls();
  };
  ui.setDirtyGuard(() =>
    state.unsaved && state.root
      ? {
          template: $('templateName').value.trim() || 'this template',
          save: () => save(),
          discard: () => {
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
  // Previews the saved revision, or the current edits when there are any (spec 6.3).
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

  // Re-run for existing records (spec 3.1, 6.5).
  async function openRerun() {
    showPanel('rerun-panel');
    $('rerun-impact').replaceChildren();
    $('rerun-these').hidden = true;
    ui.clearFeedback('rerun');
    preventError(null);
    const [count, runs] = await Promise.all([
      ui
        .api('asx_ManageWork', {
          Command: 'CountRecords',
          TemplateId: state.template.asx_templateid,
        })
        .catch(() => null),
      ui
        .api('asx_ManageWork', { Command: 'ListProblems', List: 'TemplateRuns' })
        .catch(() => ({ Problems: [] })),
    ]);
    state.run =
      (runs?.Problems || []).find(
        (p) =>
          p.Run?.TemplateId === state.template.asx_templateid &&
          ['Running', 'Waiting', 'Retrying', 'Paused', 'Blocked'].includes(p.Run.State),
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
    const role = ui.needs('prvCreateasx_operatorcommand');
    for (const id of ['rerun-all', 'rerun-preview', 'rerun-these'])
      ui.disable($(id), 'rerun-role-reason', role);
    ($('rerun-all').hidden ? $('rerun-preview') : $('rerun-all')).focus();
  }
  $('rerun-all').onclick = async () => {
    if (ui.blocked($('rerun-all'))) return;
    const version =
      state.editBase?.Status === 'Published' ? state.editBase.Version : state.editBase?.Version - 1;
    const ok = await ui.confirmInline($('rerun-all'), {
      text:
        'Re-run v' +
        version +
        ' for all ' +
        (state.runTotal ? state.runTotal + ' ' : '') +
        display(state.root) +
        ' records? Documents works through them in the background, after other work, so this can take a while. You can close this page and follow it in Monitor.',
      confirm: 'Re-run all records',
      keep: 'Not now',
    });
    if (!ok) return;
    await ui.busy($('rerun-all'), 'Starting…', 'rerun', async () => {
      const started = await ui.api('asx_ManageWork', {
        Command: 'StartTemplateRun',
        TemplateId: state.template.asx_templateid,
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
    const picked = await xrm.Utility.lookupObjects({
      entityTypes: [state.root.LogicalName],
      defaultEntityType: state.root.LogicalName,
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
        TemplateId: state.template.asx_templateid,
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
      controls();
    },
  };
  async function start() {
    render();
    resetPreview();
    if (!xrm?.WebApi || !xrm?.Utility) return;
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
      // The shell's runtime Get, if it is back; onRuntime below fills in the rest. A caller
      // who cannot read the runtime gets null and sees no readiness.
      setReadiness(ui.runtime());
      renderEnablePicker();
      renderTemplateTree();
      render();
    });
  }
  // Table readiness follows the shell's runtime result: its one Get per load, and Turn on.
  ui.onRuntime((result) => {
    setReadiness(result);
    renderTemplateTree();
  });
  ui.onTab('templates', start);
})();
