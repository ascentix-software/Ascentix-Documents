'use strict';
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
    removing: null,
    expandedTables: new Set(),
    sectionSequence: 0,
    editBase: null,
    busy: false,
  };
  const metadata = new Map();
  const xrm = window.parent?.Xrm || window.Xrm;
  const message = (text, error = false) => {
    $('status').textContent = text;
    $('status').className = error ? 'error' : '';
  };
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
  function input(value, change) {
    const i = el('input');
    i.value = value;
    i.onchange = () => {
      change(i.value);
      dirty();
    };
    return i;
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
  function controls() {
    $('newTemplate').disabled =
      state.busy || !state.root || !state.enabledTables.includes(state.root.LogicalName);
    $('templateName').disabled = state.busy || !state.root || !!state.template;
    $('revisionControls').hidden = !state.root;
    $('templateActions').hidden = !state.root || $('author-view').hidden;
    $('authorWorkspace').hidden = !state.root;
    $('tablePrompt').hidden = !!state.root;
    $('maintenanceContext').textContent = state.root
      ? 'Selected table: ' +
        display(state.root) +
        '. Record: ' +
        (Array.from($('record').options).find((o) => o.value === $('record').value)?.textContent ||
          $('record').value ||
          'Choose a preview record on Folder templates.')
      : 'Choose a table on Folder templates first.';
    $('table').disabled = state.busy || !state.tables.length;
    $('record').disabled = state.busy || !state.root;
    $('addDestination').disabled =
      state.busy || !state.root || !state.libraries.length || state.sections.length >= 10;
    $('chooseRecord').disabled = state.busy || !state.root;
    $('loadRevision').disabled = state.busy || !$('savedRevision').value;
    $('save').disabled = state.busy || !state.root || !state.sections.length;
    $('preview').disabled = state.busy || !state.saved || !$('record').value;
    $('publish').disabled = state.busy || !state.saved || state.saved.Status !== 'Draft';
  }
  // Invalidates the saved preview and updates controls after an authoring edit.
  function dirty() {
    state.saved = null;
    $('previewTrees').replaceChildren(
      el('div', 'Draft changed. Save and preview again to refresh record results.', 'callout'),
    );
    $('revision').textContent =
      state.editBase?.Status === 'Draft'
        ? 'Unsaved changes to this draft.'
        : 'Unsaved changes. Saving starts the next draft revision.';
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
  async function fields(table) {
    if (metadata.has(table)) return metadata.get(table);
    const base = "EntityDefinitions(LogicalName='" + table + "')/Attributes";
    const values = await all(
      base + '?$select=LogicalName,DisplayName,AttributeType,IsSecured,IsValidForRead',
    );
    const dates = await all(
      base +
        '/Microsoft.Dynamics.CRM.DateTimeAttributeMetadata?$select=LogicalName,DateTimeBehavior',
    );
    const choices = await all(
      base +
        '/Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet',
    );
    const singleChoices = (
      await Promise.all(
        ['Picklist', 'State', 'Status'].map((type) =>
          all(
            base +
              '/Microsoft.Dynamics.CRM.' +
              type +
              'AttributeMetadata?$select=LogicalName&$expand=OptionSet',
          ),
        ),
      )
    ).flat();
    const lookups = await all(
      base +
        '/Microsoft.Dynamics.CRM.LookupAttributeMetadata?$select=LogicalName,DisplayName,Targets,IsSecured,IsValidForRead',
    );
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
      if (c)
        c.Options = (d.OptionSet?.Options || []).map((o) => ({
          value: String(o.Value),
          label: (o.Label?.UserLocalizedLabel?.Label || String(o.Value)) + ' (' + o.Value + ')',
        }));
    }
    const result = {
      columns: columns.filter((c) => c.Kind).sort((a, b) => a.Label.localeCompare(b.Label)),
      lookups: lookups.filter((l) => !l.IsSecured && l.IsValidForRead !== false),
    };
    metadata.set(table, result);
    return result;
  }
  function availableFields() {
    return state.sources.flatMap((s) =>
      s.columns.map((c) => ({
        value: s.Alias + '.' + c.Name,
        label:
          (s.Alias === 'root'
            ? 'Current record'
            : state.lookups.find((l) => l.Lookup === s.Lookup && l.Table === s.Table)?.Label ||
              s.Lookup ||
              s.Alias) +
          ' › ' +
          c.Label,
        kind: c.Kind,
        options: c.Options,
      })),
    );
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
  function conditionGroup(group, depth = 0) {
    const box = el('div', null, 'condition-group'),
      toolbar = el('div', null, 'condition-toolbar');
    toolbar.append(
      select(
        [
          { value: 'true', label: 'All (AND)' },
          { value: 'false', label: 'Any (OR)' },
        ],
        String(group.All),
        (v) => (group.All = v === 'true'),
      ),
    );
    toolbar.append(
      button('Add condition', () => {
        const first = availableFields()[0];
        if (!first) return;
        group.Conditions.push({ field: first.value, Operator: 'Equal', Literal: '' });
        dirty();
        render();
      }),
    );
    if (depth < 3)
      toolbar.append(
        button('Add group', () => {
          group.Groups.push({ All: true, Conditions: [], Groups: [] });
          dirty();
          render();
        }),
      );
    box.append(toolbar);
    group.Conditions.forEach((c, index) => {
      const row = el('div', null, 'condition-line'),
        available = availableFields(),
        selected = available.find((f) => f.value === c.field);
      row.append(
        select(available, c.field, (v) => {
          c.field = v;
          c.Operator = 'Equal';
          c.Literal = '';
          c.right = null;
          render();
        }),
      );
      row.append(
        select(
          operators(selected?.kind).map((op) => ({ value: op, label: operatorLabel[op] })),
          c.Operator,
          (v) => {
            c.Operator = v;
            if (['IsNull', 'IsNotNull'].includes(v)) c.right = null;
            render();
          },
        ),
      );
      if (!selected)
        row.append(
          el('span', 'Unavailable field: ' + c.field + '. Select a replacement.', 'error'),
        );
      const unary = ['IsNull', 'IsNotNull'].includes(c.Operator);
      if (!unary)
        row.append(
          select(
            [
              { value: 'literal', label: 'Compare to value' },
              { value: 'field', label: 'Compare to field' },
            ],
            c.right ? 'field' : 'literal',
            (v) => {
              c.right =
                v === 'field' ? available.find((f) => f.kind === selected?.kind)?.value : null;
              render();
            },
          ),
        );
      if (c.right && !unary)
        row.append(
          select(
            available.filter((f) => f.kind === selected?.kind),
            c.right,
            (v) => (c.right = v),
          ),
        );
      else if (unary) row.append(el('span', 'No comparison value', 'hint'));
      else if (selected?.kind === 'Boolean')
        row.append(
          select(
            [
              { value: 'true', label: 'Yes' },
              { value: 'false', label: 'No' },
            ],
            c.Literal || 'true',
            (v) => (c.Literal = v),
          ),
        );
      else if (selected?.kind === 'Choice' && selected.options?.length)
        row.append(
          select(
            [{ value: '', label: 'Select a value' }, ...selected.options],
            c.Literal,
            (v) => (c.Literal = v),
          ),
        );
      else if (selected?.kind === 'MultiChoice' && selected.options?.length) {
        const picker = el('select');
        picker.multiple = true;
        const chosen = new Set((c.Literal || '').split(','));
        selected.options.forEach((o) => {
          const item = option(o.value, o.label);
          item.selected = chosen.has(o.value);
          picker.append(item);
        });
        picker.onchange = () => {
          c.Literal = Array.from(picker.selectedOptions, (o) => o.value).join(',');
          dirty();
        };
        row.append(picker);
      } else {
        const editor = input(c.Literal, (v) => (c.Literal = v));
        editor.oninput = () => {
          c.Literal = editor.value;
          dirty();
        };
        editor.placeholder =
          {
            Number: 'Invariant number, e.g. 1.25',
            Choice: 'Choice numeric value',
            MultiChoice: 'Choice values, e.g. 1,3',
            Lookup: 'Record GUID',
            DateOnly: 'YYYY-MM-DD',
            DateTime: 'ISO timestamp with offset',
          }[selected?.kind] || 'Text';
        row.append(editor);
      }
      row.append(
        button(
          'Remove',
          () => {
            group.Conditions.splice(index, 1);
            dirty();
            render();
          },
          'remove',
        ),
      );
      box.append(row);
    });
    group.Groups.forEach((child, index) => {
      const nested = conditionGroup(child, depth + 1);
      nested.append(
        button(
          'Remove group',
          () => {
            group.Groups.splice(index, 1);
            dirty();
            render();
          },
          'remove',
        ),
      );
      box.append(nested);
    });
    if (!group.Conditions.length && !group.Groups.length)
      box.append(el('p', 'Add at least one condition. Empty groups cannot be saved.', 'hint'));
    return box;
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
  function showView(view) {
    const target = ['access', 'administration', 'runtime', 'operations'].includes(view)
      ? view
      : 'author';
    const admin = ['administration', 'runtime', 'operations'].includes(target);
    $('author-view').hidden = target !== 'author';
    $('templateActions').hidden = target !== 'author' || !state.root;
    for (const id of ['access', 'runtime', 'operations', 'recordTools']) {
      const panel = $(id);
      panel.classList.add('admin-panel');
      panel.hidden =
        id === 'runtime' || id === 'operations' || id === 'recordTools' ? !admin : id !== target;
      if (!panel.hidden) panel.open = true;
    }
    if (target === 'access') window.AsxdSites?.open();
    document.querySelectorAll('[data-view]').forEach((tab) => {
      const active = tab.dataset.view === (admin ? 'administration' : target);
      tab.classList.toggle('active', active);
      tab.setAttribute('aria-pressed', String(active));
    });
  }
  function destinationName(section) {
    return section.Name && section.Name !== section.Key
      ? section.Name
      : 'Destination ' + (state.sections.indexOf(section) + 1);
  }
  function addChild(section, folder) {
    if (section.Folders.length >= 100) {
      message('This destination has reached its 100-folder limit.', true);
      return;
    }
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
    dirty();
    render();
  }
  function renderPreview(plan) {
    $('previewTrees').replaceChildren();
    if (!plan.Folders?.length) {
      $('previewTrees').append(el('div', "No folders match this record's conditions.", 'callout'));
      return;
    }
    for (const section of state.sections) {
      const rows = plan.Folders.filter((f) => f.Section === section.Key);
      const card = el('div', null, 'preview-card'),
        library = libraryFor(section);
      card.append(
        el('div', destinationName(section), 'eyebrow'),
        el('h3', library?.asx_name || 'Unavailable library'),
        el('p', siteName(library), 'hint'),
      );
      if (!rows.length)
        card.append(el('p', 'This destination is excluded by its conditions.', 'hint'));
      for (const row of rows) {
        const line = el('div', null, 'preview-node'),
          text = el('span');
        text.append(el('strong', row.Name), el('span', row.RelativePath, 'preview-note'));
        line.append(folderIcon(), text);
        card.append(line);
      }
      card.append(el('p', 'Desired names only · Library access is inherited', 'hint'));
      $('previewTrees').append(card);
    }
  }
  function render() {
    $('destinations').replaceChildren();
    $('destinationCount').textContent =
      state.sections.length + (state.sections.length === 1 ? ' destination' : ' destinations');
    if (!state.sections.includes(selectedSection)) {
      selectedSection = state.sections[0] || null;
      selectedFolder = null;
    }
    if (selectedSection && !selectedSection.Folders.includes(selectedFolder))
      selectedFolder = selectedSection.Folders.find((f) => !f.Parent) || selectedSection.Folders[0];
    if (!state.sections.length)
      $('destinations').append(
        el(
          'div',
          state.root
            ? 'No destination sections yet. Add a library with an approved, applied access policy.'
            : 'Select a business table to begin.',
          'empty',
        ),
      );
    state.sections.forEach((section, index) => {
      const card = el(
          'div',
          null,
          'destination-section' + (section === selectedSection ? ' selected-section' : ''),
        ),
        library = libraryFor(section);
      card.append(
        button(
          destinationName(section),
          () => {
            selectedSection = section;
            selectedFolder = section.Folders.find((f) => !f.Parent);
            render();
          },
          'quiet section-heading',
        ),
        el('p', library?.asx_name || 'Unavailable library', 'hint'),
      );
      const tree = el('div', null, 'tree'),
        visited = new Set();
      const addNode = (folder, depth) => {
        if (visited.has(folder)) return;
        visited.add(folder);
        const node = button(
          '',
          () => {
            selectedSection = section;
            selectedFolder = folder;
            render();
          },
          'node' + (section === selectedSection && folder === selectedFolder ? ' active' : ''),
        );
        node.style.paddingLeft = 8 + depth * 12 + 'px';
        node.setAttribute(
          'aria-pressed',
          String(section === selectedSection && folder === selectedFolder),
        );
        node.append(folderIcon(), el('span', folder.Name || 'Unnamed folder', 'name'));
        if (folder.Condition) {
          const dot = el('span', null, 'dot');
          dot.title = 'Conditional folder';
          node.append(dot);
        }
        tree.append(node);
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
        card.append(
          button('＋ Add child folder', () => addChild(section, parent), 'secondary wide'),
        );
        card.append(el('p', 'Add inside: ' + (parent.Name || 'Unnamed folder'), 'hint'));
      }
      $('destinations').append(card);
    });
    const editor = $('folderEditor');
    editor.replaceChildren();
    if (!selectedSection || !selectedFolder) {
      editor.append(
        el('div', 'Select a destination and folder to edit its name and conditions.', 'empty'),
      );
      controls();
      return;
    }
    const section = selectedSection,
      folder = selectedFolder,
      library = libraryFor(section),
      siteId = library?._asx_siteid_value || '';
    editor.append(el('h2', destinationName(section)));
    const destinationInput = input(section.Name === section.Key ? '' : section.Name || '', (v) => {
      section.Name = v.trim();
      render();
    });
    destinationInput.oninput = () => {
      section.Name = destinationInput.value;
      dirty();
    };
    destinationInput.maxLength = 200;
    destinationInput.placeholder = 'e.g. Account onboarding';
    editor.append(label('Destination name', destinationInput));
    const pickers = el('div', null, 'picker-grid'),
      sites = state.sites.filter((s) =>
        state.libraries.some((l) => l._asx_siteid_value === s.asx_siteid),
      );
    pickers.append(
      label(
        'Site',
        select(
          [
            { value: '', label: 'Choose a site' },
            ...sites.map((s) => ({ value: s.asx_siteid, label: s.asx_name })),
          ],
          siteId,
          (v) => {
            const next = state.libraries.find((l) => l._asx_siteid_value === v);
            section.LibraryId = next?.asx_libraryid || '';
            render();
          },
        ),
      ),
    );
    pickers.append(
      label(
        'Library',
        select(
          [
            { value: '', label: 'Choose a library' },
            ...state.libraries
              .filter((l) => l._asx_siteid_value === siteId)
              .map((l) => ({ value: l.asx_libraryid, label: l.asx_name })),
          ],
          library?.asx_libraryid || '',
          (v) => {
            section.LibraryId = v;
            render();
          },
        ),
      ),
    );
    editor.append(pickers);
    editor.append(
      el(
        'p',
        'Existing folders at the resolved path are reused. Use unique naming fields when records need separate folders.',
        'hint',
      ),
    );
    editor.append(
      button(
        'View this library’s team access →',
        () => {
          showView('access');
          window.AsxdSites?.selectLibrary(section.LibraryId);
        },
        'quiet small',
      ),
      el('div', null, 'divider'),
    );
    const heading = el('div', null, 'row between');
    heading.append(
      el('h3', folder.Parent ? 'Folder details' : 'Root folder'),
      el('span', folder.Parent ? 'Child folder' : 'Section root', 'badge gray'),
    );
    editor.append(heading);
    editor.append(
      label(
        'Readable folder name',
        input(folder.Name, (v) => {
          folder.Name = v;
          render();
        }),
      ),
    );
    const descendants = new Set([folder.Key]);
    let growing = true;
    while (growing) {
      growing = false;
      for (const f of section.Folders)
        if (descendants.has(f.Parent) && !descendants.has(f.Key)) {
          descendants.add(f.Key);
          growing = true;
        }
    }
    if (folder.Parent)
      editor.append(
        label(
          'Parent folder',
          select(
            section.Folders.filter((f) => !descendants.has(f.Key)).map((f) => ({
              value: f.Key,
              label: f.Name || 'Unnamed folder',
            })),
            folder.Parent,
            (v) => {
              folder.Parent = v;
              render();
            },
          ),
        ),
      );
    const activeAlias =
      folder.NameSource || folder.Name.match(/\{([a-z][a-z0-9_]*)\./)?.[1] || 'root';
    const activeSource = state.sources.find((s) => s.Alias === activeAlias) || state.sources[0];
    const sourceKey =
      activeSource.Alias === 'root' ? 'root' : activeSource.Lookup + ':' + activeSource.Table;
    const sourcePicker = select(
      [
        { value: 'root', label: 'This record' },
        ...state.lookups.map((l) => ({ value: l.Lookup + ':' + l.Table, label: l.Label })),
      ],
      sourceKey,
      () => {},
    );
    sourcePicker.onchange = () =>
      task(async () => {
        const key = sourcePicker.value;
        let source = state.sources[0];
        if (key !== 'root') {
          const lookup = state.lookups.find((l) => l.Lookup + ':' + l.Table === key);
          if (!lookup) throw new Error('This related record is no longer available.');
          source = state.sources.find(
            (s) => s.Lookup === lookup.Lookup && s.Table === lookup.Table,
          );
          if (!source) {
            if (state.sources.length >= 6)
              throw new Error('A template can use up to five related records.');
            const meta = await fields(lookup.Table);
            let i = 1;
            while (state.sources.some((s) => s.Alias === 'lookup_' + i)) i++;
            source = {
              Alias: 'lookup_' + i,
              Table: lookup.Table,
              Lookup: lookup.Lookup,
              columns: meta.columns,
            };
            state.sources.push(source);
          }
        }
        folder.NameSource = source.Alias;
        render();
      });
    editor.append(
      label('Get name fields from', sourcePicker),
      el(
        'p',
        'Choose this record or a related record, then insert a field into this folder’s name. Each folder can use different records.',
        'hint',
      ),
    );
    const nameFields = availableFields().filter((f) =>
      f.value.startsWith(activeSource.Alias + '.'),
    );
    const tokenRow = el('div', null, 'token-row'),
      tokenPicker = select(nameFields, nameFields[0]?.value, () => {});
    tokenPicker.setAttribute('aria-label', 'Field to insert');
    tokenPicker.onchange = () => {};
    tokenRow.append(
      tokenPicker,
      button('Insert field', () => {
        if (tokenPicker.value) {
          folder.Name += '{' + tokenPicker.value + '}';
          dirty();
          render();
        }
      }),
    );
    editor.append(label('Insert into folder name', tokenRow));
    editor.append(
      el('div', null, 'divider'),
      el('h3', 'When should this folder appear?'),
      label(
        'Include folder',
        select(
          [
            { value: 'always', label: 'Always' },
            { value: 'conditional', label: 'When conditions match' },
          ],
          folder.Condition ? 'conditional' : 'always',
          (v) => {
            folder.Condition = v === 'always' ? null : { All: true, Conditions: [], Groups: [] };
            render();
          },
        ),
      ),
    );
    if (!folder.Parent)
      editor.append(
        el(
          'p',
          'If this condition does not match, this destination and all its child folders are skipped. Existing folders are kept.',
          'hint',
        ),
      );
    if (folder.Condition) editor.append(conditionGroup(folder.Condition));
    editor.append(
      el(
        'div',
        'Library-level inheritance: every child inherits the selected library’s access. Sensitive placement alone gives no team access.',
        'notice',
      ),
    );
    const actions = el('div', null, 'row wrap');
    actions.style.marginTop = '16px';
    if (folder.Parent)
      actions.append(
        button(
          'Remove folder',
          () => {
            if (section.Folders.some((f) => f.Parent === folder.Key)) {
              message('Remove child folders first.', true);
              return;
            }
            section.Folders.splice(section.Folders.indexOf(folder), 1);
            selectedFolder = null;
            dirty();
            render();
          },
          'remove',
        ),
      );
    actions.append(
      button(
        'Remove destination',
        () => {
          state.sections.splice(state.sections.indexOf(section), 1);
          selectedSection = null;
          selectedFolder = null;
          dirty();
          render();
        },
        'remove',
      ),
    );
    editor.append(actions);
    controls();
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
        const unary = ['IsNull', 'IsNotNull'].includes(c.Operator);
        const right = !unary && c.right ? fieldMap.get(c.right) : null;
        if (c.right && !unary && (!right || right.kind !== f.kind))
          throw new Error('Comparison fields must have the same available type.');
        if (right) used.add(c.right);
        const [RightSource, RightColumn] = right ? c.right.split('.') : [null, null];
        return {
          Source,
          Column,
          Operator: c.Operator,
          LiteralKind: unary || right ? null : f.kind,
          Literal: unary || right ? null : f.kind === 'Boolean' ? c.Literal || 'true' : c.Literal,
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
  async function revisions() {
    $('savedRevision').replaceChildren(option('', 'Select a saved revision'));
    if (state.template)
      state.template = await xrm.WebApi.retrieveRecord(
        'asx_template',
        state.template.asx_templateid,
        '?$select=asx_templateid,asx_name,asx_table,asx_disabled,asx_startsutc,asx_endsutc',
      );
    renderAvailability();
    if (!state.template) return;
    const rows = await xrm.WebApi.retrieveMultipleRecords(
      'asx_revision',
      '?$select=asx_revisionid,asx_version,asx_status&$filter=_asx_templateid_value eq ' +
        state.template.asx_templateid +
        '&$orderby=asx_version desc',
    );
    if (rows.nextLink) throw new Error('Revision list exceeds the completeness bound.');
    rows.entities.forEach((r) =>
      $('savedRevision').append(
        option(r.asx_revisionid, 'Revision ' + r.asx_version + ' · ' + r.asx_status),
      ),
    );
    return rows.entities[0]?.asx_revisionid;
  }
  function renderAvailability() {
    const t = state.template;
    const local = (v) => {
      if (!v) return '';
      const d = new Date(v);
      return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 16);
    };
    $('templateActive').checked = !t?.asx_disabled;
    $('templateStart').value = local(t?.asx_startsutc);
    $('templateEnd').value = local(t?.asx_endsutc);
    $('useSchedule').checked = !!(t?.asx_startsutc || t?.asx_endsutc);
    $('scheduleFields').hidden = !$('useSchedule').checked;
    $('saveAvailability').disabled = !t;
    $('deleteTemplate').disabled = !t;
    $('templateAvailability').textContent = !t
      ? 'Save a template first.'
      : t.asx_disabled
        ? 'Deactivated'
        : t.asx_endsutc && new Date(t.asx_endsutc) <= new Date()
          ? 'Ended'
          : t.asx_startsutc && new Date(t.asx_startsutc) > new Date()
            ? 'Scheduled'
            : 'Active (requires a published revision)';
  }
  $('useSchedule').onchange = () => {
    $('scheduleFields').hidden = !$('useSchedule').checked;
  };
  $('saveAvailability').onclick = () =>
    task(async () => {
      if (!state.template) throw new Error('Select a saved template first.');
      const start =
          $('useSchedule').checked && $('templateStart').value
            ? new Date($('templateStart').value).toISOString()
            : null,
        end =
          $('useSchedule').checked && $('templateEnd').value
            ? new Date($('templateEnd').value).toISOString()
            : null;
      if (start && end && end <= start) throw new Error('End date must be after start date.');
      await xrm.WebApi.updateRecord('asx_template', state.template.asx_templateid, {
        asx_disabled: !$('templateActive').checked,
        asx_startsutc: start,
        asx_endsutc: end,
      });
      await revisions();
      message('Template availability saved. Existing SharePoint content is unchanged.');
    });
  $('deleteTemplate').onclick = () =>
    task(async () => {
      if (!state.template) throw new Error('Select a saved template first.');
      const answer = await xrm.Navigation.openConfirmDialog({
        title: 'Delete template?',
        text: 'This deletes the template and all its revisions. No new work will start. Existing SharePoint folders, documents and access remain unchanged. A request already sent to SharePoint may finish.',
        confirmButtonLabel: 'Delete template',
      });
      if (!answer.confirmed) return;
      await xrm.WebApi.deleteRecord('asx_template', state.template.asx_templateid);
      state.saved = null;
      state.editBase = null;
      state.sections = [];
      state.template = null;
      await loadTemplates();
      await revisions();
      render();
      $('revision').textContent = 'Template deleted';
      message('Template deleted. Existing SharePoint content is unchanged.');
    });
  $('savedRevision').onchange = controls;
  // Loads a saved revision into the authoring state and records its edit identity and row version.
  async function loadRevision() {
    const loaded = JSON.parse(await api('asx_LoadDraft', { RevisionId: $('savedRevision').value }));
    if (loaded.Draft.Table !== state.root.LogicalName)
      throw new Error('Revision belongs to another table.');
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
      Groups: g.Groups.map(group),
      Conditions: g.Conditions.map((c) => ({
        field: c.Source + '.' + c.Column,
        Operator: c.Operator,
        Literal: c.Literal || '',
        right: c.RightSource ? c.RightSource + '.' + c.RightColumn : null,
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
    state.editBase = { ...state.saved };
    $('revision').textContent =
      loaded.Status +
      ' revision ' +
      loaded.RevisionId +
      (loaded.Status === 'Draft'
        ? '. Saves update this draft until you publish it.'
        : '. Editing starts the next draft; this published version keeps running.');
    $('previewTrees').replaceChildren(
      el('div', 'Revision loaded. Choose a record and run preview.', 'callout'),
    );
    render();
    message('Saved revision loaded. Stable section and folder identities are preserved.');
  }
  $('loadRevision').onclick = () => task(loadRevision);
  async function selectTemplate(table, templateId = null, create = false) {
    $('table').value = table;
    state.template = create
      ? null
      : state.templates.find((t) => t.asx_templateid === templateId) ||
        state.templates.find((t) => t.asx_table === table) ||
        null;
    $('templateName').value = state.template?.asx_name || 'New template';
    state.expandedTables.add(table);
    renderTemplateTree();
    state.editBase = null;
    state.root = state.tables.find((t) => t.LogicalName === $('table').value);
    batchReview = null;
    $('queueBatch').disabled = true;
    if (!state.root) {
      state.template = null;
      renderAvailability();
      state.sources = [];
      state.sections = [];
      state.saved = null;
      $('record').replaceChildren(option('', 'Select a table first'));
      $('record').value = '';
      state.lookups = [];
      $('savedRevision').replaceChildren(option('', 'Select a table first'));
      $('savedRevision').value = '';
      dirty();
      render();
      message('Select a business table to begin.');
      return;
    }
    const meta = await fields(state.root.LogicalName);
    state.sources = [
      { Alias: 'root', Table: state.root.LogicalName, Lookup: null, columns: meta.columns },
    ];
    state.sections = [];
    state.saved = null;
    state.lookups = meta.lookups.flatMap((l) =>
      l.Targets.map((Table) => ({
        Lookup: l.LogicalName,
        Table,
        Label: display(l) + ' → ' + Table,
      })),
    );
    const rows = await get(
      state.root.EntitySetName +
        '?$select=' +
        state.root.PrimaryIdAttribute +
        ',' +
        state.root.PrimaryNameAttribute +
        '&$top=20',
    );
    $('record').replaceChildren(option('', 'Select a record'));
    rows.value.forEach((r) =>
      $('record').append(
        option(
          r[state.root.PrimaryIdAttribute],
          r[state.root.PrimaryNameAttribute] || r[state.root.PrimaryIdAttribute],
        ),
      ),
    );
    $('record').disabled = false;
    const latest = await revisions();
    if (latest) {
      $('savedRevision').value = latest;
      await loadRevision();
      return;
    }
    dirty();
    render();
    message(
      state.libraries.length
        ? 'Choose approved destinations and define folders.'
        : 'No approved, applied-policy libraries are available. A security administrator must approve the catalog first.',
    );
  }
  $('table').onchange = () =>
    task(async () => {
      await loadTemplates();
      await selectTemplate($('table').value);
      $('tablePicker').hidden = true;
    });
  $('addTable').onclick = () => {
    renderEnablePicker();
    $('tablePicker').hidden = !$('tablePicker').hidden;
  };
  function renderEnablePicker() {
    $('enableTable').replaceChildren(
      option('', 'Select a table to enable'),
      ...state.tables
        .filter((t) => !state.enabledTables.includes(t.LogicalName))
        .map((t) => option(t.LogicalName, display(t))),
    );
    $('enableTable').value = '';
  }
  $('enableTable').onchange = () => {
    const name = $('enableTable').value;
    if (name) return task(() => enableTable(name));
  };
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
  function setReadiness(result) {
    const readiness = result?.Registration?.Readiness;
    state.readiness = Array.isArray(readiness)
      ? Object.fromEntries(readiness.map((r) => [r.Scope, r.Status]))
      : null;
  }
  // Applies a Get-shaped result from AddTable/RemoveTable to the panel and the loaded runtime card.
  async function tableChanged(result) {
    setReadiness(result);
    if (security.runtime) runtimeResult(result);
    state.removing = null;
    await loadEnabledTables();
    renderEnablePicker();
    $('tablePicker').hidden = true;
    renderTemplateTree();
  }
  async function enableTable(name) {
    await tableChanged(await runtimeCommand({ Command: 'AddTable', Table: name }));
    message('Table enabled for Documents.');
  }
  async function disableTable(name) {
    await tableChanged(await runtimeCommand({ Command: 'RemoveTable', Table: name }));
    message('Table removed from Documents. Its templates are kept.');
  }
  $('newTemplate').onclick = () => task(() => selectTemplate(state.root.LogicalName, null, true));
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
  function renderTemplateTree() {
    const tree = $('templateTree');
    tree.replaceChildren();
    const name = (table) =>
      display(state.tables.find((t) => t.LogicalName === table) || { LogicalName: table });
    const templated = [...new Set(state.templates.map((t) => t.asx_table))].filter(Boolean);
    const enabled = [...state.enabledTables].sort();
    const disabled = templated.filter((t) => !enabled.includes(t)).sort();
    const entry = (table, isEnabled) => {
      const wrap = el('div', null, 'table-entry'),
        section = el('details'),
        summary = el('summary', name(table));
      if (isEnabled && state.readiness?.[table])
        summary.append(
          el(
            'span',
            ' ' + (readinessText[state.readiness[table]] || state.readiness[table]),
            'hint',
          ),
        );
      section.open = state.expandedTables.has(table);
      section.ontoggle = () => {
        if (section.open) state.expandedTables.add(table);
        else state.expandedTables.delete(table);
      };
      section.append(summary);
      for (const template of state.templates.filter((t) => t.asx_table === table)) {
        const item = button(
          template.asx_name || table,
          () => task(() => selectTemplate(table, template.asx_templateid)),
          'template-item',
        );
        item.setAttribute(
          'aria-pressed',
          String(template.asx_templateid === state.template?.asx_templateid),
        );
        section.append(item);
      }
      if (isEnabled)
        section.append(
          button(
            '＋ New template',
            () => task(() => selectTemplate(table, null, true)),
            'template-item',
          ),
        );
      wrap.append(section);
      if (isEnabled && state.removing === table) {
        const confirm = el('div', null, 'callout');
        confirm.append(
          el('p', 'Capture for this table stops. Its templates are kept.'),
          button('Remove table', () => task(() => disableTable(table)), 'danger'),
          button('Cancel', () => {
            state.removing = null;
            renderTemplateTree();
          }),
        );
        wrap.append(confirm);
      } else {
        const action = isEnabled
          ? button('Remove', () => {
              state.removing = table;
              renderTemplateTree();
            })
          : button('Enable', () => task(() => enableTable(table)));
        action.setAttribute('aria-label', (isEnabled ? 'Remove ' : 'Enable ') + name(table));
        wrap.append(action);
      }
      tree.append(wrap);
    };
    enabled.forEach((table) => entry(table, true));
    if (disabled.length) {
      tree.append(el('h4', 'Not enabled'));
      disabled.forEach((table) => entry(table, false));
    }
    if (!enabled.length && !disabled.length)
      tree.append(el('p', 'Add a table to create its first template.', 'hint'));
  }
  $('addDestination').onclick = () => {
    if (state.sections.length >= 10) return;
    do {
      state.sectionSequence++;
    } while (state.sections.some((s) => s.Key === 'destination_' + state.sectionSequence));
    state.sections.push({
      Key: 'destination_' + state.sectionSequence,
      LibraryId: state.libraries[0].asx_libraryid,
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
    dirty();
    render();
  };
  $('chooseRecord').onclick = () =>
    task(async () => {
      if (!state.root) throw new Error('Select a business table first.');
      const selected = await xrm.Utility.lookupObjects({
        allowMultiSelect: false,
        defaultEntityType: state.root.LogicalName,
        entityTypes: [state.root.LogicalName],
      });
      if (!selected?.length) return;
      const record = selected[0],
        id = record.id.replace(/[{}]/g, '');
      if (record.entityType !== state.root.LogicalName || !/^[0-9a-f-]{36}$/i.test(id))
        throw new Error('The selected record does not match this table.');
      if (![...$('record').options].some((o) => o.value === id))
        $('record').append(option(id, record.name || id));
      $('record').value = id;
      $('record').onchange();
    });
  $('inspectRecord').onclick = () =>
    task(async () => {
      if (!state.root || !$('record').value) throw new Error('Select a table and record first.');
      if (!state.template) throw new Error('Select a saved template first.');
      const result = JSON.parse(
        await api('asx_ManageWork', {
          Request: JSON.stringify({
            Command: 'InspectRecord',
            TemplateId: state.template.asx_templateid,
            RecordId: $('record').value,
          }),
        }),
      );
      $('result').textContent = [
        'Status: ' + result.Status,
        ...(result.Record?.OperationStates || []).map((v) => 'Folder operation: ' + v),
        ...(result.Notices || []),
      ].join('\n');
      message('Current record result: ' + result.Status);
    });
  function readableObservation(value, depth = 0) {
    if (value == null) return 'None';
    if (typeof value !== 'object') return String(value);
    return Object.entries(value)
      .map(
        ([key, item]) =>
          '  '.repeat(depth) +
          (Array.isArray(value)
            ? 'Item ' + (Number(key) + 1)
            : key.replace(/([a-z])([A-Z])/g, '$1 $2')) +
          ': ' +
          (item && typeof item === 'object'
            ? '\n' + readableObservation(item, depth + 1)
            : readableObservation(item, depth + 1)),
      )
      .join('\n');
  }
  function batchSummary(batch) {
    return (
      (batch?.Records || [])
        .map(
          (record, i) =>
            'Record ' + (i + 1) + ' (' + record.Id + ')\n' + (record.Impact || []).join('\n'),
        )
        .join('\n\n') || 'No records in this review.'
    );
  }
  let batchReview = null;
  $('previewBatch').onclick = () =>
    task(async () => {
      if (!state.root) throw new Error('Select a table first.');
      $('queueBatch').disabled = true;
      batchReview = null;
      const selected = await xrm.Utility.lookupObjects({
        allowMultiSelect: true,
        defaultEntityType: state.root.LogicalName,
        entityTypes: [state.root.LogicalName],
      });
      if (!selected?.length) return;
      if (selected.length > 5 || selected.some((r) => r.entityType !== state.root.LogicalName))
        throw new Error('Select one to five records from this table.');
      if (!state.template) throw new Error('Select a saved template first.');
      batchReview = JSON.parse(
        await api('asx_ManageWork', {
          Request: JSON.stringify({
            Command: 'PreviewBatch',
            TemplateId: state.template.asx_templateid,
            RecordIds: selected.map((r) => r.id.replace(/[{}]/g, '')),
            RequestId: crypto.randomUUID(),
          }),
        }),
      );
      $('batchReview').textContent = batchSummary(batchReview.Batch);
      $('queueBatch').disabled = false;
    });
  $('queueBatch').onclick = () =>
    task(async () => {
      if (!batchReview) throw new Error('Review a batch first.');
      const result = JSON.parse(
        await api('asx_ManageWork', {
          Request: JSON.stringify({
            Command: 'QueueBatch',
            Key: batchReview.Key,
            RowVersion: batchReview.RowVersion,
          }),
        }),
      );
      $('batchReview').textContent = batchSummary(result.Batch);
      $('queueBatch').disabled = true;
      batchReview = null;
      message('Reviewed batch queued. Physical results remain pending.');
    });
  $('record').onchange = () => {
    $('previewTrees').replaceChildren(
      el('div', 'Record changed. Run preview to see this record’s folders.', 'callout'),
    );
    controls();
  };
  $('save').onclick = () =>
    task(async () => {
      state.saved = JSON.parse(
        await api('asx_CreateDraft', { Request: JSON.stringify(payload()) }),
      );
      state.template = { asx_templateid: state.saved.TemplateId };
      await loadTemplates();
      state.editBase = { ...state.saved };
      $('revision').textContent = 'Saved draft ' + state.saved.RevisionId;
      await revisions();
      $('savedRevision').value = state.saved.RevisionId;
      message('Draft saved. Preview it, then publish with the Publisher role.');
    });
  $('preview').onclick = () =>
    task(async () => {
      const result = JSON.parse(
        await api('asx_PreviewTemplate', {
          Request: JSON.stringify({
            RevisionId: state.saved.RevisionId,
            RecordId: $('record').value,
          }),
        }),
      );
      renderPreview(result);
      message('Caller-authorized preview completed. No folders were created.');
    });
  $('publish').onclick = () =>
    task(async () => {
      const result = await api('asx_PublishTemplate', {
        RevisionId: state.saved.RevisionId,
        RowVersion: state.saved.RowVersion,
      });
      state.saved.Status = result;
      state.editBase = { ...state.saved };
      $('revision').textContent = 'Published ' + state.saved.RevisionId;
      await revisions();
      $('savedRevision').value = state.saved.RevisionId;
      await loadRevision();
      message(
        'Revision published. Provisioning follows the separately configured runtime profile.',
      );
    });
  const security = { runtime: null, operations: [] };
  const etag = (row) => (row['@odata.etag'] || '').replace(/^W\/"|"$/g, '');
  const readinessText = {
    Ready: 'Ready',
    Pending: 'Pending: not registered',
    Missing: 'Missing steps: Save to repair',
    WrongWorker: 'Registered for a different worker: Save to repair',
    WrongMode: 'Registered synchronously: Save to repair',
    WrongState: 'Enabled state differs: Save to repair',
    Outdated: 'Outdated registration: Save to repair',
    WorkerCannotRead: 'Worker cannot read this table: grant organization-level Read',
  };
  async function workers() {
    const rows = [];
    let options =
      '?$select=systemuserid,fullname&$filter=applicationid ne null and isdisabled eq false&$orderby=fullname';
    do {
      const result = await xrm.WebApi.retrieveMultipleRecords('systemuser', options);
      rows.push(...result.entities);
      options = result.nextLink
        ? new URL(result.nextLink, xrm.Utility.getGlobalContext().getClientUrl()).search
        : null;
    } while (options);
    return rows;
  }
  function runtimeResult(result, workerRows) {
    security.runtime = result;
    if (workerRows) {
      $('runtimeWorker').replaceChildren(
        ...workerRows.map((w) => option(w.systemuserid, w.fullname)),
      );
    }
    $('runtimeWorker').value = result.WorkerId;
    setReadiness(result);
    $('runtimeSites').value = result.SharePointHosts.join(',');
    $('runtimeEnabled').checked = result.Enabled;
    $('runtimeRecordUpdates').checked = result.ProcessRecordUpdates === true;
    const readiness = result.Registration?.Readiness || [];
    $('runtimeReadiness').replaceChildren(
      ...readiness.map((r) => {
        const item = document.createElement('li');
        item.textContent =
          (r.Scope === 'team' ? 'Team access events' : r.Scope) +
          ': ' +
          (readinessText[r.Status] || r.Status);
        return item;
      }),
    );
    if (result.Registration?.Error) {
      const item = document.createElement('li');
      item.textContent = 'Registration check failed: ' + result.Registration.Error;
      $('runtimeReadiness').append(item);
    }
    $('runtimePending').textContent =
      'Tables pending registration: Save to register their event steps (large sets register in batches, so Save again until none are pending).';
    $('runtimePending').hidden = !readiness.some((r) => r.Status !== 'Ready');
    renderTemplateTree();
    $('runtimeStatus').textContent = result.Enabled
      ? 'Runtime enabled; flow activation and role assignments are separate.'
      : 'Runtime paused: events queue until it is enabled.';
  }
  async function runtimeCommand(payload) {
    return JSON.parse(await api('asx_RuntimeAdmin', { Request: JSON.stringify(payload) }));
  }
  $('loadRuntime').onclick = () =>
    task(async () => {
      if (!(await runtimeRows()).length)
        throw new Error(
          'No runtime profile is installed. Import the Documents solution, then load the runtime profile.',
        );
      runtimeResult(await runtimeCommand({ Command: 'Get' }), await workers());
    });
  $('saveRuntime').onclick = () =>
    task(async () => {
      if (!security.runtime) throw new Error('Load the current runtime profile first.');
      runtimeResult(
        await runtimeCommand({
          Command: 'Save',
          RowVersion: security.runtime.RowVersion,
          WorkerId: $('runtimeWorker').value,
          SharePointHosts: $('runtimeSites')
            .value.split(',')
            .map((s) => s.trim().toLowerCase())
            .filter(Boolean),
          Enabled: $('runtimeEnabled').checked,
          ProcessRecordUpdates: $('runtimeRecordUpdates').checked,
        }),
      );
      message('Runtime profile saved and event registration verified.');
    });
  $('unregisterRuntime').onclick = () => {
    $('unregisterConfirm').hidden = false;
  };
  $('cancelUnregister').onclick = () => {
    $('unregisterConfirm').hidden = true;
  };
  $('confirmUnregister').onclick = () =>
    task(async () => {
      $('unregisterConfirm').hidden = true;
      runtimeResult(await runtimeCommand({ Command: 'Unregister' }));
      message('All Documents event registrations were removed.');
    });
  const failedJobs = { next: null };
  function failedJobQuery() {
    const steps = security.runtime?.Registration?.StepIds || [];
    if (!steps.length)
      throw new Error('Load the runtime profile first; no event steps are registered.');
    return (
      '?$select=asyncoperationid,message,createdon,_regardingobjectid_value' +
      '&$filter=statuscode eq 31 and (' +
      steps.map((id) => '_owningextensionid_value eq ' + id).join(' or ') +
      ')&$orderby=createdon desc'
    );
  }
  function renderFailed(rows, append) {
    if (!append) $('failedJobs').replaceChildren();
    for (const row of rows) {
      const table = row['_regardingobjectid_value@Microsoft.Dynamics.CRM.lookuplogicalname'] || '';
      const item = document.createElement('li');
      const box = document.createElement('input');
      box.type = 'checkbox';
      box.dataset.table = table;
      box.dataset.record = row._regardingobjectid_value || '';
      item.append(box);
      const text = document.createElement('span');
      text.textContent =
        ' ' +
        table +
        ' ' +
        box.dataset.record +
        ' · ' +
        (row.message || 'No error text') +
        ' · ' +
        row.createdon;
      item.append(text);
      $('failedJobs').append(item);
    }
  }
  async function loadFailedPage(options, append) {
    const result = await xrm.WebApi.retrieveMultipleRecords('asyncoperation', options, 50);
    renderFailed(result.entities, append);
    failedJobs.next = result.nextLink
      ? new URL(result.nextLink, xrm.Utility.getGlobalContext().getClientUrl()).search
      : null;
    $('moreFailedJobs').hidden = !failedJobs.next;
  }
  $('loadFailedJobs').onclick = () => task(() => loadFailedPage(failedJobQuery(), false));
  $('moreFailedJobs').onclick = () => task(() => loadFailedPage(failedJobs.next, true));
  $('replanFailed').onclick = () =>
    task(async () => {
      const picked = [...$('failedJobs').children]
        .map((li) => li.children[0])
        .filter((box) => box.checked && box.dataset.record);
      if (!picked.length) throw new Error('Select at least one failed job with a record.');
      let queued = 0;
      const skipped = [];
      for (const box of picked) {
        const templates = state.templates.filter((t) => t.asx_table === box.dataset.table);
        if (!box.dataset.table) skipped.push('table unknown for record ' + box.dataset.record);
        else if (!templates.length)
          skipped.push(
            "table '" + box.dataset.table + "' has no template (record " + box.dataset.record + ')',
          );
        for (const template of templates) {
          await api('asx_ManageWork', {
            Request: JSON.stringify({
              Command: 'Queue',
              RequestId: crypto.randomUUID(),
              TemplateId: template.asx_templateid,
              RecordId: box.dataset.record,
            }),
          });
          queued++;
        }
      }
      const why = skipped.length
        ? ' Skipped ' + skipped.length + ' record(s): ' + skipped.join('; ') + '.'
        : '';
      if (!queued) throw new Error('Nothing was queued.' + why);
      message('Queued ' + queued + ' record plans for the selected failed jobs.' + why);
    });
  const blockedRecords = { next: null };
  function blockedWork(row) {
    try {
      return JSON.parse(row.asx_payload || '{}');
    } catch {
      return {};
    }
  }
  function renderBlocked(rows, append) {
    if (!append) $('blockedRecords').replaceChildren();
    for (const row of rows) {
      const work = blockedWork(row);
      const item = document.createElement('li');
      const text = document.createElement('span');
      text.textContent =
        (work.Table || 'No table') +
        ' ' +
        (work.RecordId || work.Key || '') +
        ' · ' +
        ((work.Notices || [])[0] || 'No notice') +
        ' · ' +
        row.modifiedon +
        ' ';
      item.append(text);
      const retry = document.createElement('button');
      retry.className = 'secondary';
      retry.textContent = 'Retry';
      retry.disabled = !work.Key;
      retry.onclick = () =>
        task(async () => {
          const result = JSON.parse(
            await api('asx_ManageWork', {
              Request: JSON.stringify({ Command: 'RetryOutbox', Key: work.Key }),
            }),
          );
          retry.disabled = true;
          message(
            result.Status === 'Pending'
              ? 'Record queued for planning again. It blocks again if the cause remains.'
              : 'Record is ' + result.Status + '; nothing to retry.',
          );
        });
      item.append(retry);
      $('blockedRecords').append(item);
    }
  }
  async function loadBlockedPage(options, append) {
    const result = await xrm.WebApi.retrieveMultipleRecords('asx_outbox', options, 50);
    renderBlocked(result.entities, append);
    blockedRecords.next = result.nextLink
      ? new URL(result.nextLink, xrm.Utility.getGlobalContext().getClientUrl()).search
      : null;
    $('moreBlockedRecords').hidden = !blockedRecords.next;
    if (!append && !result.entities.length) message('No blocked records.');
  }
  $('loadBlockedRecords').onclick = () =>
    task(() =>
      loadBlockedPage(
        "?$select=asx_payload,modifiedon&$filter=asx_status eq 'Blocked'&$orderby=modifiedon desc",
        false,
      ),
    );
  $('moreBlockedRecords').onclick = () => task(() => loadBlockedPage(blockedRecords.next, true));
  $('loadOperations').onclick = () =>
    task(async () => {
      const rows = await xrm.WebApi.retrieveMultipleRecords(
        'asx_operation',
        '?$select=asx_payload&$orderby=createdon desc&$top=50',
      );
      security.operations = rows.entities.map((r) => JSON.parse(r.asx_payload));
      $('operation').replaceChildren(option('', 'Select an operation'));
      security.operations.forEach((o) =>
        $('operation').append(option(o.Key, o.Status + ' · ' + o.Key)),
      );
      message('Latest operations loaded. This is a bounded activity view.');
    });
  $('operation').onchange = () =>
    task(async () => {
      if (!$('operation').value) return;
      const result = JSON.parse(
        await api('asx_ManageWork', {
          Request: JSON.stringify({ Command: 'Inspect', Key: $('operation').value }),
        }),
      );
      $('operationStatus').textContent =
        result.Status +
        '\n' +
        (result.Notices || []).join('\n') +
        (result.LeaseUntilUtc ? '\nClaim expiry: ' + result.LeaseUntilUtc : '');
      $('recoveryResponse').value = '';
      $('recoveryRun').value = result.RunId || '';
      $('recoveryToken').value = result.RunId ? result.Token : '';
    });
  async function manage(command) {
    if (!$('operation').value) throw new Error('Select an operation.');
    const result = JSON.parse(
      await api('asx_ManageWork', {
        Request: JSON.stringify({ Command: command, Key: $('operation').value }),
      }),
    );
    $('operationStatus').textContent = result.Status;
    message('Operator action recorded: ' + result.Status);
  }
  $('retryOperation').onclick = () => task(() => manage('Retry'));
  $('cancelOperation').onclick = () => task(() => manage('Cancel'));
  $('recoverOperation').onclick = () =>
    task(async () => {
      if (!$('operation').value)
        throw new Error('Select the expired operation before recording recovery evidence.');
      if (
        !$('recoveryRun').value ||
        !validGuid($('recoveryToken').value) ||
        !$('recoveryEvidence').value.trim()
      )
        throw new Error(
          'Provide the exact prior run, a valid claim token and verified termination evidence. Recovery is unavailable without all three.',
        );
      if (!$('operation').value) throw new Error('Select an operation.');
      const result = JSON.parse(
        await api('asx_RecoverWorker', {
          Request: JSON.stringify({
            Key: $('operation').value,
            RunId: $('recoveryRun').value,
            Token: $('recoveryToken').value,
            Evidence: $('recoveryEvidence').value,
            ResponseBody: $('recoveryResponse').value || null,
          }),
        }),
      );
      message('Recovery audit recorded: ' + result.Status);
    });
  $('replanRecord').onclick = () =>
    task(async () => {
      if (!state.root || !$('record').value) throw new Error('Select a table and record first.');
      if (!state.template) throw new Error('Select a saved template first.');
      const result = JSON.parse(
        await api('asx_ManageWork', {
          Request: JSON.stringify({
            Command: 'Replan',
            TemplateId: state.template.asx_templateid,
            RecordId: $('record').value,
            RequestId: crypto.randomUUID(),
          }),
        }),
      );
      if (result.Status === 'Inactive') {
        message('No work queued. The template is deactivated or outside its scheduled dates.');
        return;
      }
      message(
        'Published template replan queued: ' +
          result.Key +
          '. Editor changes are not published by this action.',
      );
    });
  const validGuid = (value) =>
    /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value || '');
  async function runtimeRows() {
    const rows = await xrm.WebApi.retrieveMultipleRecords(
      'asx_runtime',
      "?$select=asx_runtimeid&$filter=asx_name eq 'Default'",
    );
    if (rows.nextLink || rows.entities.length > 1) throw new Error('Runtime setup is ambiguous.');
    return rows.entities;
  }
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
  document
    .querySelectorAll('[data-view]')
    .forEach((tab) => (tab.onclick = () => showView(tab.dataset.view)));
  window.addEventListener('hashchange', () => showView(location.hash.slice(1)));
  showView(location.hash.slice(1));
  render();
  if (!xrm?.WebApi || !xrm?.Utility) {
    message(
      'Open this admin page inside its Dataverse app to connect. No simulated records or changes are active.',
      true,
    );
    $('table').replaceChildren(option('', 'Dataverse connection required'));
    return;
  }
  task(async () => {
    state.tables = (
      await all(
        'EntityDefinitions?$select=LogicalName,DisplayName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,IsDocumentManagementEnabled&$filter=IsDocumentManagementEnabled eq true',
      )
    )
      .filter((t) => t.PrimaryNameAttribute)
      .sort((a, b) => display(a).localeCompare(display(b)));
    $('table').replaceChildren(option('', 'Select a business table'));
    state.tables.forEach((t) => $('table').append(option(t.LogicalName, display(t))));
    $('table').disabled = false;
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
    await loadTemplates();
    try {
      setReadiness(await runtimeCommand({ Command: 'Get' }));
    } catch {
      state.readiness = null; // Not permitted for this user: show no badge.
    }
    renderEnablePicker();
    renderTemplateTree();
    render();
    message('Connected. Select a document-enabled table to start a new draft.');
  });
})();
