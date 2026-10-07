'use strict';
// Monitor and Settings. Monitor's lists load when the page opens; each row comes
// from asx_ManageWork ListProblems with names and the actions the server allows. Settings loads
// the automation profile at once. Text is only ever set with textContent.
(() => {
  const ui = window.AsxdUi;
  const xrm = window.parent?.Xrm || window.Xrm;
  const $ = (id) => document.getElementById(id);
  const { el, button } = ui;
  const work = (request) => ui.api('asx_ManageWork', request);
  const catalog = (request) => ui.api('asx_CatalogAdmin', request);
  const runtimeApi = (request) => ui.api('asx_RuntimeAdmin', request);
  const EMPTY = '00000000-0000-0000-0000-000000000000';
  const OPERATOR = 'prvCreateasx_operatorcommand';
  const LISTS = [
    {
      id: 'TemplateRuns',
      title: 'Template re-runs',
      tile: 'Re-runs',
      empty: 'No re-runs in progress.',
    },
    {
      id: 'NotCaptured',
      title: 'Changes not captured',
      tile: 'Changes not captured',
      empty: 'No missed changes.',
    },
    {
      id: 'BlockedRecords',
      title: 'Blocked records',
      tile: 'Blocked records',
      empty: 'Nothing blocked.',
    },
    {
      id: 'WaitingRecords',
      title: 'Waiting for record data',
      tile: 'Waiting for record data',
      empty: 'No records waiting.',
    },
    { id: 'BlockedJobs', title: 'Blocked jobs', tile: 'Blocked jobs', empty: 'No blocked jobs.' },
    {
      id: 'RetryingJobs',
      title: 'Retrying automatically',
      tile: 'Retrying',
      empty: 'Nothing waiting to retry.',
    },
  ];
  const COLUMNS = {
    TemplateRuns: ['Template', 'Progress', 'State', 'Started', 'Actions'],
    NotCaptured: ['Select', 'Record', 'Table', 'Problem', 'Since', 'Actions'],
    BlockedRecords: ['Record', 'Table · Template', 'Problem', 'Since', 'Actions'],
    WaitingRecords: ['Record', 'Table · Template', 'What it needs', 'Since', 'Actions'],
    BlockedJobs: ['What', 'Kind', 'Problem and fix', 'Since', 'Actions'],
    RetryingJobs: ['What', 'Kind', 'Last error', 'Attempt', 'Next attempt', 'Actions'],
  };
  const RUN_STATE = {
    Running: 'Running',
    Waiting: 'Waiting for other work',
    Retrying: 'Retrying after a temporary error',
    Paused: 'Paused',
    Blocked: 'Needs attention',
    Done: 'Done',
    Cancelled: 'Cancelled',
  };
  const LABEL = {
    Retry: 'Retry',
    Cancel: 'Cancel job',
    Pause: 'Pause',
    Resume: 'Resume',
    CancelRun: 'Cancel re-run',
    UseLibrary: 'Use the library that was created',
    CreateAgain: 'Create it again',
    CheckAgain: 'Check again',
    Rerun: 'Re-run',
    Dismiss: 'Dismiss',
    Check: 'Check',
    OpenRecord: 'Open record',
  };
  // Retrying automatically says "Retry now": the job retries anyway; this runs it at once.
  const labelFor = (list, action) =>
    list === 'RetryingJobs' && action === 'Retry' ? 'Retry now' : LABEL[action];
  const BUSY = {
    Retry: 'Retrying…',
    Cancel: 'Cancelling…',
    Pause: 'Pausing…',
    Resume: 'Resuming…',
    CancelRun: 'Cancelling…',
    UseLibrary: 'Using…',
    UseCandidate: 'Using…',
    CreateAgain: 'Creating…',
    CheckAgain: 'Checking…',
    Rerun: 'Re-running…',
    Dismiss: 'Dismissing…',
  };
  const DONE = {
    Retry: 'Queued again',
    Cancel: 'Cancelled',
    Pause: 'Paused',
    Resume: 'Running',
    CancelRun: 'Cancelled',
    UseLibrary: 'Using the library',
    UseCandidate: 'Using the library',
    CreateAgain: 'Creating again',
    CheckAgain: 'Checking SharePoint…',
    Rerun: 'Re-run queued',
    Dismiss: 'Dismissed',
  };
  const TRACKING = {
    Ready: ['Ready', false],
    Pending: ['Not tracked yet', true],
    Missing: ['Some steps missing', true],
    WrongWorker: ['Set up for a different user', true],
    WrongMode: ['Set up incorrectly', true],
    WrongState: ['Turned off', true],
    Outdated: ['Out of date', true],
    WorkerCannotRead: ["Run-as user can't read this table", false],
  };
  const OPERATION_PREFIXES = ['folderjob:', 'librarycreate:', 'catalogprobe:', 'policywork:'];
  const INSTALL_DOCS =
    'https://github.com/ascentix-software/Ascentix-Documents/blob/main/docs/customer-installation.md';
  const ADMIN_ONLY = 'Only a System Administrator can change these settings.';
  // templates and hasLibrary stay null until read, so the setup checklist never flashes open.
  const monitor = {
    summary: null,
    lists: new Map(),
    tables: [],
    templates: null,
    hasLibrary: null,
    record: null,
    recent: null,
    recentLoad: null,
  };
  // edited: the form has changes not saved yet, which a runtime change from elsewhere keeps.
  const settings = { workers: [], edited: false };
  const number = (n) => Number(n || 0).toLocaleString('en-US');
  const counted = (list) =>
    monitor.summary?.Capped?.includes(list)
      ? '5,000+'
      : number(monitor.summary?.[list] ?? monitor.lists.get(list)?.rows.length ?? 0);
  const plural = ui.plural;

  // Monitor -----------------------------------------------------------------------------------

  async function openMonitor() {
    renderLists();
    renderPill();
    ui.onRuntime(() => {
      renderPill();
      renderSetup();
    });
    $('monitor-settings-link').onclick = () => ui.navigate('settings');
    $('monitor-refresh').onclick = () =>
      ui.busy($('monitor-refresh'), 'Refreshing…', 'monitor', () => refresh(true));
    $('advanced').ontoggle = () => ($('advanced').open ? ensureRecent() : undefined);
    $('recent-more').onclick = () =>
      ui.busy($('recent-more'), 'Loading…', 'advanced', () => loadRecent(true));
    $('operation-lookup').onclick = () =>
      ui.busy($('operation-lookup'), 'Looking up…', 'advanced', () =>
        lookUp($('operation-id').value),
      );
    wireCheck();
    await Promise.all([refresh(false), loadCheckPickers()]);
    const link = ui.deeplink();
    if (link) await focusLink(link);
  }

  async function refresh(announce) {
    const [summary] = await Promise.all([
      loadSummary(),
      ...LISTS.map((l) => loadList(l.id, false)),
    ]);
    $('monitor-checked').textContent = 'Checked ';
    $('monitor-checked').append(ui.time(summary?.CountedUtc || new Date().toISOString()));
    if (!announce || !summary) return;
    // Counted as the tiles count: a list Dataverse stopped counting says "5,000+".
    const say = (list, one, many, after = '') =>
      summary[list]
        ? counted(list) +
          ' ' +
          (summary[list] === 1 && !summary.Capped?.includes(list) ? one : many) +
          after
        : null;
    const parts = [
      say('TemplateRuns', 're-run', 're-runs', ' in progress'),
      say('NotCaptured', 'change', 'changes', ' not captured'),
      say('BlockedRecords', 'blocked record', 'blocked records'),
      say('WaitingRecords', 'record', 'records', ' waiting for data'),
      say('BlockedJobs', 'blocked job', 'blocked jobs'),
      say('RetryingJobs', 'job', 'jobs', ' retrying'),
    ].filter(Boolean);
    ui.feedback(
      'monitor',
      'Refreshed. ' + (parts.length ? parts.join(', ') + '.' : 'Nothing needs attention.'),
    );
  }

  async function loadSummary() {
    try {
      monitor.summary = (await work({ Command: 'Summary' })).Summary;
    } catch {
      monitor.summary = null;
    }
    renderTiles();
    renderSetup();
    return monitor.summary;
  }

  function renderSetup() {
    const runtime = ui.runtime();
    const steps = [
      [
        'Choose who runs automation (Settings)',
        'settings',
        !!runtime?.WorkerId && runtime.WorkerId !== EMPTY,
      ],
      ['Add a site and a library (Sites & access)', 'access', monitor.hasLibrary !== false],
      [
        'Enable a table and publish a template (Folder templates)',
        'templates',
        monitor.templates?.length !== 0,
      ],
      ['Turn automation on', null, !!runtime?.Enabled],
    ];
    const open = runtime && steps.some((s) => !s[2]);
    $('setup-checklist').hidden = !open;
    $('monitor-tiles').hidden = !!open;
    if (!open) return;
    // Redrawn on every count and runtime change: focus keys bring focus back to its step.
    ui.withFocus(() =>
      $('setup-checklist').replaceChildren(
        ...steps.map(([text, tab, done]) => {
          const item = el('li', null, done ? 'is-done' : '');
          if (done)
            item.append(
              el('span', '✓ ', 'check'),
              el('span', text),
              el('span', ' (done)', 'sr-only'),
            );
          else if (tab) {
            const step = button(text, () => ui.navigate(tab), 'link');
            step.dataset.focusKey = 'setup:' + tab;
            item.append(step);
          } else item.append(el('span', text));
          return item;
        }),
      ),
    );
  }

  function renderTiles() {
    // Redrawn on every tick: focus keys bring focus back to the same tile.
    ui.withFocus(() =>
      $('monitor-tiles').replaceChildren(
        ...LISTS.map((list) => {
          const tile = button(list.tile + ' ', () => $('h-' + list.id).focus(), 'tile');
          tile.dataset.focusKey = 'tile:' + list.id;
          tile.append(el('strong', counted(list.id)));
          return tile;
        }),
      ),
    );
  }

  // Automation as a pill in the Monitor header: running or paused, from the runtime Get or, for
  // roles that cannot make it, the Default runtime row. The switch itself is in Settings.
  async function renderPill() {
    const pill = $('monitor-automation');
    const state = await ui.automation();
    pill.hidden = !state;
    if (!state) return;
    pill.textContent = state.Enabled ? 'Automation running' : 'Automation paused';
    pill.dataset.tone = state.Enabled ? 'ok' : 'warning';
  }

  // While Monitor is the open page, the 60-second tick re-reads the counts, the re-runs and the
  // blocked jobs. Progress updates silently. A re-run that reaches Done or Needs attention, and
  // a library setup that leaves "Checking SharePoint…", are announced once in the Monitor
  // feedback line.
  async function watch() {
    const before = new Map(
      ['TemplateRuns', 'BlockedJobs'].flatMap((id) =>
        (monitor.lists.get(id)?.rows || []).map((r) => [r.Key, r]),
      ),
    );
    await Promise.all([loadSummary(), reread('TemplateRuns'), reread('BlockedJobs')]);
    // One line holds every announcement of the tick, so none replaces another.
    const said = [];
    for (const row of monitor.lists.get('TemplateRuns')?.rows || []) {
      const was = before.get(row.Key)?.Run?.State;
      const now = row.Run?.State;
      if (was && was !== now && (now === 'Done' || now === 'Blocked'))
        said.push('Re-run of ' + row.Run.TemplateName + ': ' + RUN_STATE[now] + '.');
    }
    for (const row of monitor.lists.get('BlockedJobs')?.rows || []) {
      const was = before.get(row.Key)?.Recovery?.State;
      if (was === 'Checking' && row.Recovery && row.Recovery.State !== 'Checking')
        said.push(ui.recoverySentence(row.Title, row.Recovery));
    }
    if (said.length) ui.feedback('monitor', said.join(' '));
  }
  // The watch's re-read of a list's first page, with no skeleton. The fresh page replaces the
  // first page, so a row that was retried, cancelled or resolved and has left the list is gone;
  // rows that "Show 50 more" appended stay. A failed read leaves the list as it was until the
  // next tick.
  async function reread(id) {
    const state = monitor.lists.get(id);
    if (!state) return;
    try {
      const page = await work({ Command: 'ListProblems', List: id, Page: null });
      const fresh = new Set(page.Problems.map((r) => r.Key));
      const appended = state.rows.filter((r) => state.extra.has(r.Key) && !fresh.has(r.Key));
      state.extra = new Set(appended.map((r) => r.Key));
      state.rows = page.Problems.concat(appended);
      if (!appended.length) state.next = page.Next;
      redraw(id);
    } catch {
      // Quiet: the next tick tries again.
    }
  }
  // Redraws a list unless the admin is in the middle of something there: an open in-page
  // confirmation, open Details, or focus on a control a redraw cannot give focus back to (Copy,
  // a checkbox, the table filter, the Since sort). The next tick redraws it once that is over.
  function redraw(id) {
    const section = listOf(id);
    const active = document.activeElement;
    const held =
      !!section?.querySelector('.confirm[role=group]') ||
      [...(section?.querySelectorAll('details') || [])].some((d) => d.open) ||
      (!!section?.contains(active) && active !== $('h-' + id) && !active.dataset?.focusKey);
    if (!held) ui.withFocus(() => renderList(id));
  }

  function renderLists() {
    $('monitor-lists').replaceChildren(
      ...LISTS.map((list) => {
        const section = el('section', null, 'problem-list');
        section.id = 'list-' + list.id;
        section.setAttribute('data-focus-scope', '');
        const heading = el('h3', list.title + ' · Loading…');
        heading.id = 'h-' + list.id;
        heading.tabIndex = -1;
        heading.setAttribute('data-focus-heading', '');
        const body = el('div');
        body.id = 'body-' + list.id;
        section.append(heading, body);
        // extra: keys of rows that "Show 50 more" appended, which the tick's re-read keeps.
        monitor.lists.set(list.id, {
          rows: [],
          next: null,
          extra: new Set(),
          filter: '',
          order: 'desc',
          meta: list,
        });
        return section;
      }),
    );
  }

  async function loadList(id, append) {
    const state = monitor.lists.get(id);
    const body = $('body-' + id);
    body.setAttribute('aria-busy', 'true');
    if (!append) body.replaceChildren(...[0, 1, 2].map(() => el('div', null, 'skeleton')));
    try {
      const page = await work({
        Command: 'ListProblems',
        List: id,
        Page: append ? state.next : null,
      });
      if (append) {
        const known = new Set(state.rows.map((r) => r.Key));
        const added = page.Problems.filter((r) => !known.has(r.Key));
        added.forEach((r) => state.extra.add(r.Key));
        state.rows = state.rows.concat(added);
      } else {
        state.rows = page.Problems;
        state.extra = new Set();
      }
      state.next = page.Next;
      ui.withFocus(() => renderList(id));
    } catch (error) {
      if (append) throw error;
      const failed = el(
        'p',
        "Couldn't load " + state.meta.title + ': ' + (error.message || String(error)),
        'error',
      );
      failed.setAttribute('role', 'alert');
      body.replaceChildren(
        failed,
        button('Try again', () => loadList(id, false)),
      );
      $('h-' + id).textContent = state.meta.title;
    } finally {
      body.removeAttribute('aria-busy');
    }
  }

  function renderList(id) {
    const state = monitor.lists.get(id);
    $('h-' + id).textContent = state.meta.title + ' · ' + counted(id);
    const body = $('body-' + id);
    if (!state.rows.length) {
      body.replaceChildren(el('p', state.meta.empty, 'empty'));
      return;
    }
    const tables = [...new Set(state.rows.map((r) => r.Record?.TableLabel).filter(Boolean))];
    const parts = [];
    if (tables.length > 1) {
      const filter = el('select');
      filter.id = 'filter-' + id;
      filter.append(option('', 'All tables'), ...tables.map((t) => option(t, t)));
      filter.value = state.filter;
      filter.onchange = () => {
        state.filter = filter.value;
        ui.withFocus(() => renderList(id));
      };
      const label = el('label', 'Table');
      label.append(filter);
      parts.push(label);
    }
    if (id === 'NotCaptured') {
      // Its own action row, so its busy state leaves the rows' buttons alone.
      const tools = el('div', null, 'row');
      tools.setAttribute('data-actions', '');
      tools.append(rerunSelected());
      parts.push(tools);
    }
    const table = el('table');
    const caption = el('caption', state.meta.title, 'sr-only');
    const head = el('thead');
    const header = el('tr');
    for (const name of COLUMNS[id]) {
      const th = el('th', name === 'Select' ? null : name);
      th.setAttribute('scope', 'col');
      if (name === 'Select') th.append(el('span', 'Select', 'sr-only'));
      if (name === 'Since') {
        th.setAttribute('aria-sort', state.order === 'desc' ? 'descending' : 'ascending');
        th.replaceChildren(
          button(
            'Since',
            () => {
              state.order = state.order === 'desc' ? 'asc' : 'desc';
              ui.withFocus(() => renderList(id));
            },
            'link',
          ),
        );
      }
      header.append(th);
    }
    head.append(header);
    const rows = el('tbody');
    rows.id = 'rows-' + id;
    const shown = state.rows
      .filter((r) => !state.filter || r.Record?.TableLabel === state.filter)
      .sort(
        (a, b) =>
          (state.order === 'desc' ? -1 : 1) *
          String(a.SinceUtc || '').localeCompare(String(b.SinceUtc || '')),
      );
    for (const row of shown) rows.append(renderRow(id, row));
    table.append(caption, head, rows);
    parts.push(table);
    if (state.next) {
      const more = button('Show 50 more', () =>
        ui.busy(more, 'Loading…', 'monitor', () => loadList(id, true)),
      );
      more.id = 'more-' + id;
      more.dataset.focusKey = 'more:' + id;
      const row = el('div', null, 'row');
      row.setAttribute('data-actions', '');
      row.append(more);
      parts.push(row);
    }
    body.replaceChildren(...parts);
  }

  const option = (value, text) => {
    const o = el('option', text);
    o.value = value;
    return o;
  };
  const cell = (...content) => {
    const td = el('td');
    td.append(...content.filter((c) => c != null));
    return td;
  };
  // The record's name opens its form in a new window; its ID labels the row's checkbox.
  function recordCell(row, nameId) {
    const name = row.Record?.Name;
    const shown = !row.Record
      ? el('span', row.Title)
      : !name
        ? el('span', 'Record not available to you')
        : button(name, () => openRecord(row), 'link');
    shown.id = nameId;
    if (name) shown.dataset.focusKey = 'row:' + row.Key + ':open';
    return cell(shown);
  }
  const openRecord = (row) =>
    xrm.Navigation.openForm({
      entityName: row.Record.Table,
      entityId: row.Record.Id,
      openInNewWindow: true,
    });

  function renderRow(list, row) {
    const tr = el('tr');
    tr.setAttribute('data-focus-row', '');
    tr.dataset.key = row.Key;
    const status = el('span', null, 'row-status');
    const actions = el('td', null, 'actions');
    actions.setAttribute('data-actions', '');
    const nameId = 'name-' + list + '-' + tr.dataset.key.replace(/[^\w-]/g, '_');
    switch (list) {
      case 'TemplateRuns':
        tr.append(...runCells(row, status));
        break;
      case 'NotCaptured': {
        // A team row or a row without a record cannot be re-run (the server offers no Rerun).
        let box = null;
        if (row.Actions?.includes('Rerun')) {
          box = el('input');
          box.type = 'checkbox';
          box.dataset.key = row.Key;
          box.setAttribute('aria-labelledby', nameId);
          box.onchange = () => updateSelected();
        }
        tr.append(
          cell(box),
          recordCell(row, nameId),
          cell(el('span', row.Record?.TableLabel || '')),
          cell(el('span', row.Problem), status),
          cell(ui.time(row.SinceUtc)),
        );
        break;
      }
      case 'BlockedRecords':
      case 'WaitingRecords': {
        const what = el('span', row.Problem);
        const more = row.More?.length
          ? button(
              '+' + row.More.length + ' more',
              () => more.replaceWith(...row.More.map((m) => el('p', m))),
              'link',
            )
          : null;
        tr.append(
          recordCell(row, nameId),
          cell(el('span', [row.Record?.TableLabel, row.TemplateName].filter(Boolean).join(' · '))),
          cell(what, more, status),
          cell(ui.time(row.SinceUtc)),
        );
        break;
      }
      case 'RetryingJobs':
        tr.append(
          cell(el('span', row.Title)),
          cell(el('span', row.KindLabel)),
          cell(el('span', row.Problem), status),
          cell(el('span', String(row.Attempt))),
          cell(ui.time(row.NextAttemptUtc)),
        );
        break;
      default:
        tr.append(
          cell(el('span', row.Title)),
          cell(el('span', row.KindLabel)),
          problemCell(row, status),
          cell(ui.time(row.SinceUtc)),
        );
    }
    for (const action of row.Actions || []) {
      // Lists with a Record column open the record from its name; job lists get a button.
      if (
        action === 'OpenRecord' &&
        ['NotCaptured', 'BlockedRecords', 'WaitingRecords'].includes(list)
      )
        continue;
      if (action === 'UseCandidate') continue;
      if (action === 'Cancel' && row.Kind === 'LibrarySetup') {
        rowButton(actions, list, row, 'Cancel', 'Cancel setup', 'danger', tr, status);
        continue;
      }
      const style =
        action === 'Cancel' || action === 'CancelRun'
          ? 'danger'
          : action === 'UseLibrary' || action === 'CreateAgain'
            ? 'primary'
            : 'secondary';
      rowButton(actions, list, row, action, labelFor(list, action), style, tr, status);
    }
    actions.append(ui.details(row.Key, 'Details', row.Title));
    tr.append(actions);
    return tr;
  }

  // Adds one action button to its container, then marks it disabled with its reason when the
  // caller lacks the privilege (the reason is placed after the button, so it must be attached).
  function rowButton(into, list, row, action, text, style, tr, status, candidate = null) {
    const control = button(
      text,
      () => act(list, row, action, control, tr, status, candidate),
      style,
    );
    control.dataset.action = action;
    control.dataset.focusKey =
      'row:' + list + ':' + row.Key + ':' + action + (candidate ? ':' + candidate.ListId : '');
    control.setAttribute('aria-label', text + ' for ' + row.Title);
    into.append(control);
    const missing =
      action === 'Check' || action === 'OpenRecord'
        ? null
        : row.Kind === 'LibrarySetup'
          ? ui.needs('prvCreateasx_site')
          : ui.needs(OPERATOR);
    if (missing)
      ui.disable(control, 'reason-' + control.dataset.focusKey.replace(/[^\w-]/g, '_'), missing);
    return control;
  }

  function problemCell(row, status) {
    const td = el('td');
    const recovery = row.Recovery;
    if (recovery && row.Kind === 'LibrarySetup') {
      td.append(el('p', ui.recoverySentence(row.Title, recovery)));
      if (recovery.State === 'Ambiguous') td.append(candidates(row));
      td.append(status);
      return td;
    }
    td.append(el('p', row.Problem));
    if (row.Fix) td.append(el('p', row.Fix, 'muted'));
    td.append(status);
    return td;
  }

  // An ambiguous finding lists each candidate with the checks it failed and its own Use this one.
  // `list` is the area whose feedback line reports the choice: the Blocked jobs list or Advanced.
  function candidates(row, list = 'BlockedJobs') {
    const items = el('ul', null, 'candidates');
    for (const c of row.Recovery.Candidates) {
      const failed = ui.candidateChecks(c);
      const item = el('li');
      item.append(
        el('span', c.Title + ' · ' + c.Url + ' · created '),
        ui.time(c.CreatedUtc),
        el('span', failed ? ' · ' + failed : ''),
      );
      if (row.Actions.includes('UseCandidate') && c.IsLibrary && c.CatalogEntry !== 'Conflict')
        rowButton(item, list, row, 'UseCandidate', 'Use this one', 'secondary', null, null, c);
      items.append(item);
    }
    return items;
  }

  function runCells(row, status) {
    const run = row.Run;
    const total = (run.TotalEstimated ? 'about ' : '') + number(run.Total);
    const progressText = number(run.Planned) + ' of ' + total;
    const bar = el('progress');
    bar.max = Math.max(1, run.Total);
    bar.value = Math.min(run.Planned, bar.max);
    bar.setAttribute('aria-valuetext', progressText + ', ' + RUN_STATE[run.State].toLowerCase());
    bar.setAttribute('aria-label', 'Progress of ' + run.TemplateName);
    const finish = el('div', null, 'muted');
    if (['Running', 'Waiting', 'Retrying'].includes(run.State)) {
      if (run.EstimatedFinishUtc)
        finish.append(el('span', 'Ends around '), ui.time(run.EstimatedFinishUtc));
      else finish.textContent = 'Estimating…';
    }
    status.textContent = RUN_STATE[run.State];
    if (run.State === 'Retrying' && run.NextAttemptUtc)
      status.append(el('span', ' · next try '), ui.time(run.NextAttemptUtc));
    const state = el('td');
    state.append(status);
    if (run.State === 'Blocked') state.append(el('p', row.Problem));
    return [
      cell(el('span', run.TemplateName + ' · ' + run.TableLabel + ' · v' + run.Version)),
      cell(bar, el('div', progressText), finish),
      state,
      cell(ui.time(run.StartedUtc), el('span', run.StartedBy ? ' · ' + run.StartedBy : '')),
    ];
  }

  function confirmFor(row, action, candidate) {
    if (action === 'Cancel' && row.Kind === 'LibrarySetup')
      return {
        text:
          'Cancel the setup of ' +
          row.Title +
          '? Nothing in SharePoint is deleted. If SharePoint already created the library, add it as an existing library.',
        confirm: 'Cancel setup',
        keep: 'Keep setup',
        danger: true,
      };
    if (action === 'Cancel')
      return {
        text:
          'Cancel the ' +
          row.KindLabel.toLowerCase() +
          ' for ' +
          row.Title +
          '? Nothing in SharePoint is undone or deleted.',
        confirm: 'Cancel job',
        keep: 'Keep job',
        danger: true,
      };
    if (action === 'CancelRun')
      return {
        text:
          'Cancel the re-run of ' +
          row.Run.TemplateName +
          '? Records not yet planned are skipped. Folder work already queued for planned records continues, and nothing in SharePoint is undone.',
        confirm: 'Cancel re-run',
        keep: 'Keep re-run',
        danger: true,
      };
    if (action === 'UseCandidate')
      return {
        text:
          'Use ' +
          candidate.Title +
          ' at ' +
          candidate.Url +
          ' for this setup? Documents stops its permission inheritance if needed and manages its team access.',
        confirm: 'Use this library',
        keep: 'Keep looking',
      };
    return null;
  }

  function send(row, action, candidate) {
    const key = row.Key;
    const outbox = row.Kind === 'RecordPlan' || row.Kind === 'TemplateRun';
    switch (action) {
      case 'Retry':
        return work({ Command: outbox ? 'RetryOutbox' : 'Retry', Key: key });
      case 'Cancel':
        return row.Kind === 'LibrarySetup'
          ? catalog({ Command: 'CancelSetup', Key: key })
          : work({ Command: 'Cancel', Key: key });
      case 'Pause':
        return work({ Command: 'PauseTemplateRun', Key: key });
      case 'Resume':
        return work({ Command: 'ResumeTemplateRun', Key: key });
      case 'CancelRun':
        return work({ Command: 'CancelTemplateRun', Key: key });
      case 'UseLibrary':
      case 'UseCandidate':
        return catalog({
          Command: 'ResolveSetup',
          Key: key,
          Choice: 'UseLibrary',
          ListId: (candidate || row.Recovery.Candidates[0]).ListId,
          RowVersion: row.RowVersion,
        });
      case 'CreateAgain':
        return catalog({
          Command: 'ResolveSetup',
          Key: key,
          Choice: 'CreateAgain',
          RowVersion: row.RowVersion,
        });
      case 'CheckAgain':
        return catalog({ Command: 'RecheckSetup', Key: key });
      case 'Rerun':
        return row.Kind === 'CaptureJob'
          ? work({
              Command: 'RerunRecord',
              Table: row.Record.Table,
              RecordId: row.Record.Id,
              RequestId: crypto.randomUUID(),
            })
          : work({
              Command: 'Replan',
              TemplateId: row.TemplateId,
              RecordId: row.Record.Id,
              RequestId: crypto.randomUUID(),
            });
      case 'Dismiss':
        return work({ Command: 'DismissCaptureJob', JobId: key });
      default:
        throw new Error('Unknown action.');
    }
  }

  function said(row, action, result) {
    const name = row.Record?.Name || row.Title;
    switch (action) {
      case 'Retry':
        return result.Status === 'Pending'
          ? 'Queued again. It blocks again if the cause remains.'
          : 'Nothing to retry: it is ' + result.Status + '.';
      case 'Cancel':
        return [
          row.Kind === 'LibrarySetup'
            ? 'Setup cancelled. Nothing in SharePoint was deleted.'
            : 'Job cancelled. Nothing in SharePoint was undone or deleted.',
          ...(result.Notices || []),
        ].join(' ');
      case 'Pause':
        return 'Re-run paused.';
      case 'Resume':
        return 'Re-run resumed.';
      case 'CancelRun':
        return 'Re-run cancelled. Folder work already queued continues.';
      case 'UseLibrary':
      case 'UseCandidate':
        return 'Using the existing library. Setup continues.';
      case 'CreateAgain':
        return 'Creating the library again.';
      case 'CheckAgain':
        return 'Checking SharePoint again.';
      case 'Rerun':
        return result.Status === 'Inactive'
          ? row.Kind === 'CaptureJob'
            ? 'Not queued: no template for this table is on.'
            : 'Not queued: the template is off or outside its scheduled dates.'
          : ['Re-run queued for ' + name + '.', ...(result.Notices || [])].join(' ');
      case 'Dismiss':
        return 'Dismissed.';
      default:
        return '';
    }
  }

  // A confirmation for a table row renders in a full-width row under it; elsewhere inside the anchor.
  function confirmHost(anchor) {
    if (anchor?.tagName === 'TR') {
      const holder = el('tr', null, 'confirm-row');
      const host = el('td');
      host.setAttribute('colspan', String(anchor.children.length));
      holder.append(host);
      anchor.after(holder);
      return { host, remove: () => holder.remove() };
    }
    const host = el('div');
    anchor.append(host);
    return { host, remove: () => host.remove() };
  }

  async function act(list, row, action, control, tr, status, candidate) {
    if (ui.blocked(control)) return;
    if (action === 'OpenRecord') return openRecord(row);
    if (action === 'Check')
      return focusLink({
        tab: 'monitor',
        record: row.Record.Table + ':' + row.Record.Id,
        template: row.TemplateId,
      });
    const question = confirmFor(row, action, candidate);
    if (question) {
      const place = confirmHost(tr || control.closest('li') || control.parentNode);
      const ok = await ui.confirmInline(control, { ...question, host: place.host });
      place.remove();
      if (!ok) return;
    }
    const area = list === 'advanced' ? 'advanced' : 'monitor';
    let result = null;
    await ui.busy(control, BUSY[action], area, async () => {
      result = await send(row, action, candidate);
    });
    if (!result) return;
    ui.feedback(area, said(row, action, result));
    if (action === 'Dismiss') {
      // Gone from the list's rows too, so a later re-render does not bring it back.
      const state = monitor.lists.get(list);
      state.rows = state.rows.filter((r) => r.Key !== row.Key);
      ui.withFocus(() => tr.remove());
      return;
    }
    if (!tr) return;
    if (result.Run)
      tr.replaceChildren(
        ...renderRow(list, { ...row, Run: result.Run, Status: result.Status, Actions: [] })
          .children,
      );
    const line = tr.querySelector('.row-status');
    if (line) line.textContent = DONE[action];
    // The row stays and its actions disable; opening or checking the record still works.
    [...tr.querySelectorAll('button')]
      .filter(
        (b) =>
          b.dataset.action && b.dataset.action !== 'OpenRecord' && b.dataset.action !== 'Check',
      )
      .forEach((b) => (b.disabled = true));
  }

  function rerunSelected() {
    const run = button('Re-run selected (0)', () => {
      const picked = [...listOf('NotCaptured').querySelectorAll('input[type=checkbox]')].filter(
        (b) => b.checked,
      );
      if (!picked.length) return undefined;
      return ui.busy(run, 'Re-running…', 'monitor', async () => {
        const rows = monitor.lists.get('NotCaptured').rows;
        let queued = 0;
        const failed = [];
        // Each row on its own: one refusal is reported and the rest are still queued.
        for (const box of picked) {
          const row = rows.find((r) => r.Key === box.dataset.key);
          const name = row?.Record?.Name || row?.Title || 'A record';
          try {
            const result = await send(row, 'Rerun');
            if (result.Status === 'Inactive')
              failed.push(name + ': no template for this table is on.');
            else queued++;
          } catch (error) {
            failed.push(name + ': ' + String(error.message || error).replace(/\.?$/, '.'));
          }
        }
        if (!failed.length)
          ui.feedback('monitor', 'Re-run queued for ' + plural(queued, 'record', 'records') + '.');
        else
          ui.feedback(
            'monitor',
            'Queued ' +
              number(queued) +
              '; ' +
              number(failed.length) +
              " couldn't be re-run: " +
              failed.join(' '),
            'error',
          );
      });
    });
    monitor.rerunSelected = run;
    return run;
  }
  function updateSelected() {
    const count = [...listOf('NotCaptured').querySelectorAll('input[type=checkbox]')].filter(
      (b) => b.checked,
    ).length;
    monitor.rerunSelected.textContent = 'Re-run selected (' + count + ')';
  }

  // Check a record.
  function wireCheck() {
    const refreshButtons = () => {
      const reason = monitor.record && $('check-template').value ? null : 'Choose a record first';
      ui.disable($('check-run'), 'check-reason', reason);
      ui.disable($('check-rerun'), 'check-reason', reason);
    };
    $('check-table').onchange = () => {
      monitor.record = null;
      $('check-record-name').textContent = '';
      fillTemplates();
      refreshButtons();
    };
    $('check-template').onchange = refreshButtons;
    $('check-choose').onclick = async () => {
      const table = $('check-table').value;
      if (!table) return;
      const picked = await xrm.Utility.lookupObjects({
        entityTypes: [table],
        defaultEntityType: table,
        allowMultiSelect: false,
      });
      if (!picked?.length) return;
      monitor.record = {
        table,
        id: picked[0].id.replace(/[{}]/g, '').toLowerCase(),
        name: picked[0].name,
      };
      $('check-record-name').textContent = monitor.record.name;
      $('check-choose').focus();
      refreshButtons();
    };
    $('check-run').onclick = () => {
      if (ui.blocked($('check-run'))) return undefined;
      return ui.busy($('check-run'), 'Checking…', 'check', async () => {
        const result = await work({
          Command: 'InspectRecord',
          TemplateId: $('check-template').value,
          RecordId: monitor.record.id,
        });
        $('check-result').replaceChildren(...readable(result));
      });
    };
    $('check-rerun').onclick = () => {
      if (ui.blocked($('check-rerun'))) return undefined;
      return ui.busy($('check-rerun'), 'Re-running…', 'check', async () => {
        const result = await work({
          Command: 'Replan',
          TemplateId: $('check-template').value,
          RecordId: monitor.record.id,
          RequestId: crypto.randomUUID(),
        });
        ui.feedback(
          'check',
          result.Status === 'Inactive'
            ? 'Not queued: the template is off or outside its scheduled dates.'
            : 'Re-run queued for ' + monitor.record.name + '.',
        );
      });
    };
    refreshButtons();
  }
  async function loadCheckPickers() {
    try {
      const [tables, templates, libraries] = await Promise.all([
        everything('asx_runtimetable', '?$select=asx_logicalname&$orderby=asx_logicalname'),
        everything(
          'asx_template',
          '?$select=asx_templateid,asx_name,asx_table,_asx_publishedrevisionid_value&$filter=_asx_publishedrevisionid_value ne null&$orderby=asx_name',
        ),
        xrm.WebApi.retrieveMultipleRecords(
          'asx_library',
          '?$select=asx_libraryid&$filter=asx_approved eq true&$top=1',
        ),
      ]);
      monitor.tables = tables.map((t) => t.asx_logicalname);
      await loadLabels(monitor.tables);
      monitor.templates = templates;
      monitor.hasLibrary = libraries.entities.length > 0;
    } catch (error) {
      // The pickers stay empty and the checklist treats the unread steps as done.
      ui.feedback(
        'check',
        "Couldn't load tables and templates: " + (error.message || String(error)),
        'error',
      );
    }
    $('check-table').replaceChildren(
      option('', 'Choose a table'),
      ...monitor.tables.map((t) => option(t, label(t))),
    );
    fillTemplates();
    renderSetup();
  }
  // Table display names, read once per table from Dataverse metadata; the logical name until then.
  const labels = new Map();
  const label = (table) => labels.get(table) || table;
  async function loadLabels(tables) {
    await Promise.all(
      tables
        .filter((t) => t !== 'team' && !labels.has(t))
        .map(async (t) => {
          try {
            labels.set(t, (await xrm.Utility.getEntityMetadata(t, [])).DisplayName || t);
          } catch {
            labels.set(t, t);
          }
        }),
    );
  }
  function fillTemplates() {
    const table = $('check-table').value;
    const choices = (monitor.templates || []).filter((t) => t.asx_table === table);
    $('check-template').replaceChildren(
      ...choices.map((t) => option(t.asx_templateid, t.asx_name)),
    );
    $('check-template').value = choices[0]?.asx_templateid || '';
  }
  // A readable result (no pre, no internal prefixes): a status sentence, then each destination.
  function readable(result) {
    const sentence =
      {
        Planned: 'Folders are created.',
        Applied: 'Folders are created.',
        Waiting: 'Waiting for record data.',
        Blocked: 'Blocked.',
        NotPlanned: 'This record has not been planned yet.',
        DecommissionReview: 'The record was deleted; its documents are kept.',
        TemplateDeleted: 'The template was deleted.',
      }[result.Status] || result.Status;
    const states = (result.Record?.OperationStates || []).map((s, i) =>
      el('li', 'Destination ' + (i + 1) + ': ' + s),
    );
    const list = el('ul');
    list.append(...states);
    return [el('p', sentence), ...(result.Notices || []).map((n) => el('p', n)), list];
  }

  // Advanced: recent operations and look up by ID. The first page loads once, when
  // Advanced opens or a link opens it; a failed read is reported and tried again next time.
  function ensureRecent() {
    if (!monitor.recentLoad)
      monitor.recentLoad = loadRecent(false).catch((error) => {
        monitor.recentLoad = null;
        ui.feedback(
          'advanced',
          "Couldn't load recent operations: " + (error.message || String(error)),
          'error',
        );
      });
    return monitor.recentLoad;
  }
  async function loadRecent(append) {
    const page = await work({
      Command: 'ListProblems',
      List: 'RecentOperations',
      Page: append ? monitor.recent?.next : null,
    });
    monitor.recent = { next: page.Next };
    if (!append) $('recent-rows').replaceChildren();
    for (const row of page.Problems) {
      const tr = el('tr');
      const open = button(
        row.Title,
        () => ui.busy(open, 'Looking up…', 'advanced', () => lookUp(row.Key)),
        'link',
      );
      tr.append(
        cell(open),
        cell(el('span', row.KindLabel)),
        cell(el('span', row.Status)),
        cell(ui.time(row.SinceUtc)),
      );
      $('recent-rows').append(tr);
    }
    $('recent-more').hidden = !page.Next;
  }
  async function lookUp(raw) {
    const key = String(raw || '').trim();
    if (!OPERATION_PREFIXES.some((p) => key.startsWith(p)))
      throw new Error(
        'Operation IDs start with folderjob:, librarycreate:, catalogprobe: or policywork:',
      );
    const result = await work({ Command: 'Inspect', Key: key });
    const fields = $('operation-fields');
    fields.replaceChildren();
    for (const [term, value] of [
      ['Operation ID', key],
      ['Status', result.Status],
      ['Problem', (result.Notices || []).join(' ') || 'None'],
    ]) {
      fields.append(el('dt', term), el('dd', value));
    }
    fields.append(el('dd'));
    fields.lastElementChild.append(ui.details(key, 'Details', 'this operation'));
    const actions = $('operation-actions');
    actions.replaceChildren();
    const setup = key.startsWith('librarycreate:');
    const row = {
      Key: key,
      Kind: setup ? 'LibrarySetup' : 'FolderJob',
      KindLabel: 'job',
      Title: key,
      Recovery: result.Recovery,
      RowVersion: result.RowVersion,
      Actions: [],
    };
    if (setup && result.Recovery) {
      // A library setup with a SharePoint finding offers the same recovery choices as its
      // Blocked jobs row: Use the library, Use this one, Create it again,
      // Check again, Cancel setup.
      fields.append(el('dd', ui.recoverySentence(key, result.Recovery)));
      row.Actions = result.Recovery.Choices || [];
      for (const action of row.Actions) {
        if (action === 'UseCandidate') continue;
        const style =
          action === 'Cancel'
            ? 'danger'
            : action === 'UseLibrary' || action === 'CreateAgain'
              ? 'primary'
              : 'secondary';
        rowButton(
          actions,
          'advanced',
          row,
          action,
          action === 'Cancel' ? 'Cancel setup' : LABEL[action],
          style,
          actions,
          el('span'),
        );
      }
      if (result.Recovery.State === 'Ambiguous') actions.append(candidates(row, 'advanced'));
    } else {
      for (const action of result.Status === 'Blocked' || result.Status === 'RetryWait'
        ? ['Retry', 'Cancel']
        : ['Cancel'])
        rowButton(
          actions,
          'advanced',
          row,
          action,
          LABEL[action],
          action === 'Cancel' ? 'danger' : 'secondary',
          actions,
          el('span'),
        );
    }
    $('operation-result').hidden = false;
  }

  async function focusLink(link) {
    if (link.operation) {
      $('advanced').open = true;
      await ensureRecent();
      $('operation-id').value = link.operation;
      await ui.busy($('operation-lookup'), 'Looking up…', 'advanced', () => lookUp(link.operation));
      $('operation-result-title').focus();
    }
    if (link.record) {
      const [table, id] = link.record.split(':');
      $('check-record').open = true;
      $('check-table').value = table;
      fillTemplates();
      if (link.template) $('check-template').value = link.template;
      const row = [...monitor.lists.values()]
        .flatMap((l) => l.rows)
        .find((r) => r.Record?.Id === id);
      monitor.record = { table, id, name: row?.Record?.Name || 'the chosen record' };
      $('check-record-name').textContent = monitor.record.name;
      $('check-template').onchange();
      $('check-run').focus();
    }
    if (link.run)
      rowsOf('TemplateRuns')
        ?.querySelector('[data-key="' + link.run + '"] button')
        ?.focus();
  }
  // A list's section and its rows (tbody), which renderLists and renderList draw.
  const listOf = (id) => document.getElementById('list-' + id);
  const rowsOf = (id) => document.getElementById('rows-' + id);

  // Automation switch (Settings): SetEnabled, never a whole-profile save. The switch keeps its
  // label and state spans while it works: its state text says what it is doing.
  function renderSwitch() {
    const runtime = ui.runtime();
    const control = $('automation-switch-settings');
    control.hidden = !runtime;
    if (!runtime) return;
    if (!control.dataset.busy) {
      control.setAttribute('aria-checked', String(!!runtime.Enabled));
      $('automation-state-settings').textContent = runtime.Enabled ? 'Running' : 'Paused';
    }
    ui.disable(
      control,
      'automation-reason-settings',
      runtime.CanChange === false
        ? 'Only a System Administrator can pause or resume automation.'
        : null,
    );
    control.onclick = toggleAutomation;
  }
  async function toggleAutomation() {
    const control = $('automation-switch-settings');
    if (ui.blocked(control) || control.dataset.busy || !ui.runtime()) return;
    const turnOn = !ui.runtime().Enabled;
    control.dataset.busy = '1';
    control.setAttribute('aria-busy', 'true');
    $('automation-state-settings').textContent = turnOn ? 'Turning on…' : 'Pausing…';
    ui.clearFeedback('settings');
    try {
      // setAutomation shares the result through setRuntime: the switch, the Monitor pill and the
      // form's row version follow it. The state text says the new state; the feedback line only
      // adds what still needs work.
      const result = await ui.setAutomation(turnOn);
      const problem = turnOn ? result.Registration?.Error : null;
      if (problem) ui.feedback('settings', 'Automation running. ' + problem, 'error');
    } catch (error) {
      ui.feedback('settings', error.message || String(error), 'error');
    } finally {
      delete control.dataset.busy;
      control.removeAttribute('aria-busy');
      renderSwitch();
    }
  }

  // Settings ----------------------------------------------------------------------------------

  // The page starts before the shell's runtime Get may be back: the form stays busy
  // until it is, then follows every runtime change (the switch, Save, Repair). A refused Get shows the missing-settings alert.
  async function openSettings() {
    $('automation-settings').setAttribute('aria-busy', 'true');
    $('add-host').onclick = () => {
      addHost('', true);
      settings.edited = true;
      const inputs = $('hosts-list').querySelectorAll('input');
      inputs[inputs.length - 1].focus();
    };
    $('runtimeWorker').onchange = () => (settings.edited = true);
    $('record-updates').onclick = () => {
      if (ui.blocked($('record-updates'))) return;
      settings.edited = true;
      $('record-updates').setAttribute(
        'aria-checked',
        String($('record-updates').getAttribute('aria-checked') !== 'true'),
      );
    };
    $('save-settings').onclick = () => ui.busy($('save-settings'), 'Saving…', 'settings', save);
    // Once Repair all ends, the button shows what is still pending, not its progress text.
    $('repair-all').onclick = async () => {
      if (ui.blocked($('repair-all'))) return;
      await ui.busy($('repair-all'), 'Repairing…', 'settings', repairAll);
      if (ui.runtime()) renderTracking(ui.runtime());
    };
    $('stop-tracking').onclick = async () => {
      if (ui.blocked($('stop-tracking'))) return;
      const ok = await ui.confirmInline($('stop-tracking'), {
        text: 'Stop tracking changes for every table? Documents stops capturing record and team changes until you repair change tracking. Use this before uninstalling.',
        confirm: 'Stop tracking changes',
        keep: 'Keep tracking',
        danger: true,
      });
      if (!ok) return;
      await ui.busy($('stop-tracking'), 'Stopping…', 'settings', async () => {
        ui.setRuntime(await runtimeApi({ Command: 'Unregister' }));
        ui.feedback('settings', 'Change tracking stopped for every table.');
      });
    };
    [settings.workers] = await Promise.all([workers(), loadConnections(), ui.runtimeReady()]);
    ui.onRuntime(showSettings);
    showSettings(ui.runtime());
    const link = ui.deeplink();
    if (link?.table)
      $('tracking-rows')
        .querySelector(
          '[data-table="' + link.table + '"] button, [data-table="' + link.table + '"] td',
        )
        ?.focus();
  }

  function showSettings(runtime) {
    $('settings-missing').hidden = !!runtime;
    $('automation-settings').querySelector('.stack').hidden = !runtime;
    if (!runtime) {
      $('save-settings').hidden = true;
      $('settings-readonly').hidden = true;
      $('automation-settings').removeAttribute('aria-busy');
      return;
    }
    renderSettings(runtime);
    // Table names come from metadata, read once per table; the logical name shows until then.
    const scopes = (runtime.Registration?.Readiness || []).map((r) => r.Scope);
    if (scopes.some((s) => s !== 'team' && !labels.has(s)))
      loadLabels(scopes).then(() => ui.runtime() && renderTracking(ui.runtime()));
  }

  // Every row of a query, page by page as the platform returns them.
  async function everything(table, options) {
    const rows = [];
    let query = options;
    do {
      const page = await xrm.WebApi.retrieveMultipleRecords(table, query);
      rows.push(...page.entities);
      query = page.nextLink
        ? new URL(page.nextLink, xrm.Utility.getGlobalContext().getClientUrl()).search
        : null;
    } while (query);
    return rows;
  }
  // The enabled application users; a caller who cannot read users gets an empty list.
  async function workers() {
    try {
      return await everything(
        'systemuser',
        '?$select=systemuserid,fullname&$filter=applicationid ne null and isdisabled eq false&$orderby=fullname',
      );
    } catch {
      return [];
    }
  }

  function renderSettings(runtime) {
    const editable = runtime.CanChange !== false;
    $('settings-readonly').hidden = editable;
    $('save-settings').hidden = !editable;
    // Fields the admin is editing keep their values when the runtime changes elsewhere; Save
    // reads the current row version and Enabled from the shared runtime, not from the form.
    if (!settings.edited) {
      const worker = $('runtimeWorker');
      worker.replaceChildren(...settings.workers.map((w) => option(w.systemuserid, w.fullname)));
      // A disabled or deleted run-as user stays selectable so saving keeps it.
      if (
        runtime.WorkerId &&
        runtime.WorkerId !== EMPTY &&
        !settings.workers.some((w) => w.systemuserid === runtime.WorkerId)
      )
        worker.append(option(runtime.WorkerId, 'Current run-as user (disabled or not found)'));
      if (!runtime.WorkerId || runtime.WorkerId === EMPTY)
        worker.prepend(option('', 'Choose an application user'));
      worker.value = runtime.WorkerId && runtime.WorkerId !== EMPTY ? runtime.WorkerId : '';
      $('hosts-list').replaceChildren();
      for (const host of runtime.SharePointHosts?.length ? runtime.SharePointHosts : [''])
        addHost(host, editable);
      $('record-updates').setAttribute('aria-checked', String(!!runtime.ProcessRecordUpdates));
    }
    $('runtimeWorker').disabled = !editable;
    $('add-host').hidden = !editable;
    $('hosts-list')
      .querySelectorAll('input, button')
      .forEach((n) => (n.disabled = !editable));
    ui.disable($('record-updates'), 'record-updates-reason', editable ? null : ADMIN_ONLY);
    ui.disable($('stop-tracking'), 'stop-tracking-reason', editable ? null : ADMIN_ONLY);
    renderSwitch();
    renderTracking(runtime);
    $('automation-settings').removeAttribute('aria-busy');
  }

  function addHost(value, editable) {
    const item = el('li', null, 'host row');
    const field = el('input');
    field.type = 'text';
    field.setAttribute('inputmode', 'url');
    field.setAttribute('spellcheck', 'false');
    field.value = value;
    field.oninput = () => (settings.edited = true);
    const index = $('hosts-list').querySelectorAll('li').length + 1;
    field.setAttribute('aria-label', 'SharePoint host name ' + index);
    const remove = button('Remove', () => {
      settings.edited = true;
      item.remove();
      $('add-host').focus();
    });
    remove.setAttribute('aria-label', 'Remove host ' + (value || index));
    field.disabled = remove.disabled = !editable;
    item.append(field, remove);
    $('hosts-list').append(item);
  }

  async function save() {
    const fields = [...$('hosts-list').querySelectorAll('input')];
    let invalid = null;
    for (const field of fields) {
      const host = field.value.trim().toLowerCase();
      const bad =
        host && !/^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+$/.test(host);
      const errorId = field.getAttribute('aria-label').replace(/\W+/g, '-') + '-error';
      document.getElementById(errorId)?.remove();
      field.removeAttribute('aria-invalid');
      field.removeAttribute('aria-describedby');
      if (!bad) continue;
      const message = el('p', 'Enter a host name like contoso.sharepoint.com', 'error');
      message.id = errorId;
      message.setAttribute('role', 'alert');
      field.after(message);
      field.setAttribute('aria-invalid', 'true');
      field.setAttribute('aria-describedby', errorId);
      invalid = invalid || field;
    }
    if (invalid) {
      invalid.focus();
      return;
    }
    const current = ui.runtime();
    const result = await runtimeApi({
      Command: 'Save',
      RowVersion: current.RowVersion,
      WorkerId: $('runtimeWorker').value,
      SharePointHosts: fields.map((f) => f.value.trim().toLowerCase()).filter(Boolean),
      Enabled: current.Enabled,
      ProcessRecordUpdates: $('record-updates').getAttribute('aria-checked') === 'true',
    });
    settings.edited = false;
    ui.setRuntime(result);
    ui.feedback('settings', 'Settings saved.');
  }

  function renderTracking(runtime) {
    const registration = runtime.Registration || { Readiness: [] };
    $('tracking-error').hidden = !registration.Error;
    $('tracking-error').textContent = registration.Error || '';
    const editable = runtime.CanChange !== false;
    $('tracking-rows').replaceChildren(
      ...(registration.Readiness || []).map((r) => {
        const [text, repairable] = TRACKING[r.Status] || [r.Status, true];
        const tr = el('tr');
        tr.dataset.table = r.Scope;
        tr.setAttribute('data-focus-row', '');
        const action = el('td');
        action.setAttribute('data-actions', '');
        if (repairable) {
          const repair = button('Repair', () => {
            if (ui.blocked(repair)) return undefined;
            return ui.busy(repair, 'Repairing…', 'settings', async () => {
              const result = await runtimeApi({
                Command: 'Register',
                Table: r.Scope,
                RowVersion: ui.runtime().RowVersion,
              });
              ui.withFocus(() => ui.setRuntime(result));
            });
          });
          repair.dataset.focusKey = 'tracking:' + r.Scope;
          repair.setAttribute(
            'aria-label',
            'Repair ' + (r.Scope === 'team' ? 'team access events' : label(r.Scope)),
          );
          action.append(repair);
          if (!editable) ui.disable(repair, 'tracking-reason-' + r.Scope, ADMIN_ONLY);
        } else if (r.Status === 'WorkerCannotRead') {
          const how = el('a', 'How to grant access');
          how.setAttribute('href', INSTALL_DOCS + '#configure-and-enable');
          how.setAttribute('target', '_blank');
          how.setAttribute('rel', 'noopener');
          action.append(how);
        }
        tr.append(
          cell(el('span', r.Scope === 'team' ? 'Team access events' : label(r.Scope))),
          cell(el('span', text)),
          action,
        );
        return tr;
      }),
    );
    const pending =
      registration.Pending ??
      (registration.Readiness || []).filter((r) => TRACKING[r.Status]?.[1]).length;
    if (!$('repair-all').dataset.busy) {
      $('repair-all').hidden = pending < 2;
      $('repair-all').textContent = 'Repair all (' + pending + ')';
    }
    ui.disable($('repair-all'), 'repair-all-reason', editable ? null : ADMIN_ONLY);
  }

  // Register repeatedly until nothing is pending or a call makes no progress.
  async function repairAll() {
    let before = ui.runtime().Registration?.Pending ?? 0;
    const total = before;
    while (before > 0) {
      const result = await runtimeApi({
        Command: 'Register',
        Table: '',
        RowVersion: ui.runtime().RowVersion,
      });
      ui.setRuntime(result);
      const after = result.Registration?.Pending ?? 0;
      $('repair-all').textContent = 'Repaired ' + (total - after) + ' of ' + total;
      if (after >= before) {
        ui.feedback(
          'settings',
          'Repair made no progress. ' +
            plural(after, 'table still needs', 'tables still need') +
            ' repair.',
          'error',
        );
        return;
      }
      before = after;
    }
    ui.feedback('settings', 'Change tracking repaired for every table.');
  }

  async function loadConnections() {
    const environment = xrm.Utility.getGlobalContext().organizationSettings?.bapEnvironmentId;
    $('open-connections').setAttribute(
      'href',
      environment
        ? 'https://make.powerautomate.com/environments/' + environment + '/connections'
        : INSTALL_DOCS + '#prepare-and-import',
    );
    try {
      const rows = await xrm.WebApi.retrieveMultipleRecords(
        'connectionreference',
        "?$select=connectionreferencedisplayname,connectionid,connectorid&$filter=startswith(connectionreferencelogicalname,'asx_')",
      );
      $('connection-list').replaceChildren(
        ...rows.entities.map((r) =>
          el(
            'li',
            r.connectionreferencedisplayname +
              ' · ' +
              (r.connectionid ? 'Connected' : 'Not connected'),
          ),
        ),
      );
    } catch (error) {
      $('connection-list').replaceChildren(
        el('li', "Couldn't read the connections: " + (error.message || String(error)), 'error'),
      );
    }
  }

  ui.onTab('monitor', openMonitor);
  ui.onTab('settings', openSettings);
  ui.onLink('monitor', focusLink);
  // While Monitor is the open page and the page is visible, the 60-second tick re-reads its
  // counts, re-runs and blocked jobs and announces state changes (watch).
  document.addEventListener('DOMContentLoaded', () => {
    setInterval(() => {
      if (document.visibilityState !== 'visible') return undefined;
      if (ui.activeTab() !== 'monitor' || !xrm?.WebApi || !ui.can(OPERATOR)) return undefined;
      return watch();
    }, 60000);
  });
})();
