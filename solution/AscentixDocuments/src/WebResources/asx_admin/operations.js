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
  const ADMIN = 'prvWriteasx_runtime';
  // Every list Monitor reads. `title` names a list whose read failed; `empty` is what a filter
  // with no rows says.
  const LISTS = [
    { id: 'TemplateRuns', title: 'Template re-runs' },
    { id: 'NotCaptured', title: 'Changes not captured', empty: 'No missed changes.' },
    { id: 'BlockedRecords', title: 'Blocked records', empty: 'Nothing blocked.' },
    { id: 'WaitingRecords', title: 'Waiting for record data', empty: 'No records waiting.' },
    { id: 'BlockedJobs', title: 'Blocked jobs', empty: 'No blocked jobs.' },
    { id: 'RetryingJobs', title: 'Retrying automatically', empty: 'Nothing waiting to retry.' },
  ];
  // The five problem lists in chip order, with each chip's label and each row's type and tone.
  const PROBLEMS = {
    BlockedRecords: { chip: 'Blocked records', type: 'Blocked record', tone: 'danger' },
    WaitingRecords: { chip: 'Waiting for data', type: 'Waiting for data', tone: 'warning' },
    BlockedJobs: { chip: 'Blocked jobs', type: 'Blocked job', tone: 'danger' },
    RetryingJobs: { chip: 'Retrying', type: 'Retrying', tone: 'muted' },
    NotCaptured: { chip: 'Not captured', type: 'Not captured', tone: 'warning' },
  };
  const PROBLEM_LISTS = Object.keys(PROBLEMS);
  // Lists whose item is a record: its name opens the form, so Open record is not offered again.
  const RECORD_LISTS = ['NotCaptured', 'BlockedRecords', 'WaitingRecords'];
  // A row's one button is the first of these it offers; the rest go in its ⋯ menu. Cancel,
  // Cancel re-run and Dismiss are never the one button.
  const PRIMARY = [
    'UseLibrary',
    'CreateAgain',
    'Retry',
    'Check',
    'Rerun',
    'CheckAgain',
    'OpenRecord',
  ];
  const PAGE = 50;
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
    UseLibrary: 'Use the library',
    CreateAgain: 'Create it again',
    CheckAgain: 'Check again',
    Rerun: 'Re-run',
    Dismiss: 'Dismiss',
    Check: 'Check',
    OpenRecord: 'Open record',
  };
  // Retrying automatically says "Retry now": the job retries anyway; this runs it at once. A
  // library setup's Cancel cancels the setup.
  const labelFor = (list, action, row = null) =>
    list === 'RetryingJobs' && action === 'Retry'
      ? 'Retry now'
      : action === 'Cancel' && row?.Kind === 'LibrarySetup'
        ? 'Cancel setup'
        : LABEL[action];
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
  // templates and hasLibrary stay null until read, so the setup checklist never flashes open.
  // filter: the chosen chip ('all' or a problem list); order: Since newest first or oldest first;
  // limit: how many merged rows show; selected: the Not captured rows checked for Re-run selected.
  const monitor = {
    summary: null,
    lists: new Map(),
    filter: 'all',
    order: 'desc',
    limit: PAGE,
    selected: new Set(),
    setupOpen: false,
    tables: [],
    templates: null,
    hasLibrary: null,
    record: null,
    recent: null,
    recentLoad: null,
    rerunSelected: null,
  };
  // edited: the form has changes not saved yet, which a runtime change from elsewhere keeps.
  // tables: the document-enabled tables ＋ Add table offers; counts: templates per table (null
  // when unread); tablesStale: a Tables redraw waits for the open confirmation there.
  const settings = { workers: [], edited: false, tables: [], counts: null, tablesStale: false };
  const number = (n) => Number(n || 0).toLocaleString('en-US');
  const countOf = (list) => monitor.summary?.[list] ?? monitor.lists.get(list)?.rows.length ?? 0;
  const capped = (list) => !!monitor.summary?.Capped?.includes(list);
  const counted = (list) => (capped(list) ? '5,000+' : number(countOf(list)));
  const plural = ui.plural;

  // Monitor -----------------------------------------------------------------------------------

  async function openMonitor() {
    // extra: keys of rows that Show more appended, which the tick's re-read keeps.
    for (const list of LISTS)
      monitor.lists.set(list.id, {
        rows: [],
        next: null,
        extra: new Set(),
        loading: false,
        error: null,
        meta: list,
      });
    renderPill();
    ui.onRuntime(() => {
      renderPill();
      renderSetup();
    });
    $('monitor-settings-link').onclick = () => ui.navigate('settings');
    $('monitor-refresh').onclick = () =>
      ui.busy($('monitor-refresh'), 'Refreshing…', 'monitor', () => refresh(true));
    ui.menu($('monitor-tools'), $('monitor-tools-list'));
    $('tools-check').onclick = () => openCheck($('monitor-tools'));
    $('tools-lookup').onclick = () => openLookup($('monitor-tools'));
    $('check-close').onclick = () => panels.check?.close();
    $('lookup-close').onclick = () => panels.lookup?.close();
    $('problem-sort').onclick = () => {
      monitor.order = monitor.order === 'desc' ? 'asc' : 'desc';
      ui.withFocus(renderTable);
    };
    $('problem-more').onclick = showMore;
    monitor.rerunSelected = rerunSelected();
    $('rerun-selected-row').replaceChildren(monitor.rerunSelected);
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
    // Counted as the chips count: a list Dataverse stopped counting says "5,000+".
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
    renderFilters();
    renderMore();
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
    // A paused automation is a normal state, shown by the header pill: only the first three
    // steps hold the problem lists back.
    const open = !!runtime && steps.slice(0, 3).some((s) => !s[2]);
    monitor.setupOpen = open;
    $('setup-checklist').hidden = !open;
    for (const id of ['monitor-filters', 'monitor-errors', 'problem-table-card'])
      $(id).hidden = open;
    $('runs-strip').hidden = open || !activeRuns().length;
    $('rerun-selected-row').hidden = open || monitor.filter !== 'NotCaptured';
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
    if (said.length) announce(said.join(' '));
  }
  // An error in the Monitor line stays until the admin acts again: a tick's announcement goes
  // after it instead of replacing it.
  function announce(text) {
    const line = $('fb-monitor');
    if (line.getAttribute('role') === 'alert' && line.textContent)
      line.append(' ', el('span', text, 'feedback-note'));
    else ui.feedback('monitor', text);
  }
  // The watch's re-read of a list's first page, with no skeleton. The fresh page replaces the
  // first page, so a row that was retried, cancelled or resolved and has left the list is gone;
  // rows that Show more appended stay. A failed read leaves the list as it was until the next
  // tick.
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
  // Redraws the re-runs strip or the problems table unless the admin is in the middle of
  // something there: an open in-page confirmation, an open ⋯ menu, or focus on a control a
  // redraw cannot give focus back to (a checkbox). The next tick redraws it once that is over.
  function redraw(id) {
    const area = id === 'TemplateRuns' ? $('runs-strip') : $('problem-table-card');
    const active = document.activeElement;
    const held =
      !!area.querySelector('.confirm[role=group]') ||
      [...area.querySelectorAll('.menu')].some((menu) => !menu.hidden) ||
      (area.contains(active) && active !== $('problem-heading') && !active.dataset?.focusKey);
    if (!held) ui.withFocus(() => render(id));
  }
  function render(id) {
    if (id === 'TemplateRuns') return renderStrip();
    renderFilters();
    return renderTable();
  }

  async function loadList(id, append) {
    const state = monitor.lists.get(id);
    if (!append) {
      state.loading = true;
      render(id);
    }
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
      state.error = null;
    } catch (error) {
      if (append) throw error;
      state.error = error.message || String(error);
    } finally {
      if (!append) state.loading = false;
    }
    ui.withFocus(() => {
      render(id);
      renderErrors();
    });
  }
  // A list whose read failed says so above the table, with Try again; the others still show.
  function renderErrors() {
    $('monitor-errors').replaceChildren(
      ...LISTS.filter((l) => monitor.lists.get(l.id).error).map((l) => {
        const box = el('div', null, 'row load-error');
        box.setAttribute('data-actions', '');
        const failed = el(
          'p',
          "Couldn't load " + l.title + ': ' + monitor.lists.get(l.id).error,
          'error',
        );
        failed.setAttribute('role', 'alert');
        box.append(
          failed,
          button('Try again', () => loadList(l.id, false)),
        );
        return box;
      }),
    );
  }

  // Filter chips: All and the five problem lists, counted from Summary ("5,000+" for a list
  // Dataverse stopped counting). A list with nothing in it is plain text, not a button. When the
  // chosen list empties, the table goes back to All.
  function renderFilters() {
    const host = $('monitor-filters');
    if (!monitor.summary && PROBLEM_LISTS.some((id) => monitor.lists.get(id)?.loading)) {
      host.replaceChildren();
      return;
    }
    let refocus = false;
    if (monitor.filter !== 'all' && !countOf(monitor.filter)) {
      refocus = document.activeElement?.dataset?.focusKey === 'filter:' + monitor.filter;
      monitor.filter = 'all';
      monitor.limit = PAGE;
      redraw('problems');
    }
    const total = PROBLEM_LISTS.some(capped)
      ? '5,000+'
      : number(PROBLEM_LISTS.reduce((sum, id) => sum + countOf(id), 0));
    ui.withFocus(() =>
      host.replaceChildren(
        ...[
          ['all', 'All', total],
          ...PROBLEM_LISTS.map((id) => [id, PROBLEMS[id].chip, counted(id)]),
        ].map(([id, label, count]) => {
          const text = label + ' · ' + count;
          if (id !== 'all' && !countOf(id)) return el('span', text, 'filter-chip is-empty');
          const chip = button(text, () => choose(id), 'filter-chip');
          chip.setAttribute('aria-pressed', String(monitor.filter === id));
          chip.dataset.focusKey = 'filter:' + id;
          return chip;
        }),
      ),
    );
    if (refocus) host.querySelector('[data-focus-key="filter:all"]')?.focus();
  }
  function choose(id) {
    monitor.filter = id;
    monitor.limit = PAGE;
    ui.withFocus(() => {
      renderFilters();
      renderTable();
    });
  }

  // The lists the table shows: the five under All, else the chosen one.
  const shownLists = () => (monitor.filter === 'all' ? PROBLEM_LISTS : [monitor.filter]);
  // Their loaded rows as one list, by Since (newest or oldest first), ties by key.
  function mergedRows() {
    const sign = monitor.order === 'desc' ? -1 : 1;
    const when = (row) => ui.ms(row.SinceUtc) || 0;
    return shownLists()
      .flatMap((list) => (monitor.lists.get(list)?.rows || []).map((row) => ({ list, row })))
      .sort(
        (a, b) =>
          sign * (when(a.row) - when(b.row)) || String(a.row.Key).localeCompare(String(b.row.Key)),
      );
  }

  function renderTable() {
    const select = monitor.filter === 'NotCaptured';
    const head = $('problem-table').querySelector('thead tr');
    const box = head.querySelector('.col-select');
    if (select && !box) {
      const th = el('th', null, 'col-select');
      th.setAttribute('scope', 'col');
      th.setAttribute('role', 'columnheader');
      th.append(el('span', 'Select', 'sr-only'));
      head.prepend(th);
    } else if (!select) box?.remove();
    const sort = head.querySelector('th[aria-sort]');
    sort.setAttribute('aria-sort', monitor.order === 'desc' ? 'descending' : 'ascending');
    $('problem-sort-arrow').textContent = monitor.order === 'desc' ? ' ▾' : ' ▴';
    const lists = shownLists();
    const loading = lists.some((id) => monitor.lists.get(id)?.loading);
    if (loading) $('problem-table-card').setAttribute('aria-busy', 'true');
    else $('problem-table-card').removeAttribute('aria-busy');
    const shown = mergedRows().slice(0, monitor.limit);
    const span = String(head.children.length);
    const note = (content) => {
      const tr = el('tr');
      const td = el('td');
      td.setAttribute('colspan', span);
      td.append(content);
      tr.append(td);
      return asRow(tr);
    };
    $('problem-rows').replaceChildren(
      ...(shown.length
        ? shown.map(({ list, row }) => renderRow(list, row))
        : loading
          ? [0, 1, 2].map(() => note(el('div', null, 'skeleton')))
          : lists.some((id) => monitor.lists.get(id)?.error)
            ? []
            : [
                note(
                  el(
                    'p',
                    monitor.filter === 'all'
                      ? 'Nothing needs attention.'
                      : LISTS.find((l) => l.id === monitor.filter).empty,
                    'empty',
                  ),
                ),
              ]),
    );
    $('rerun-selected-row').hidden = monitor.setupOpen || !select;
    updateSelected();
    renderMore();
  }

  // "Show N more": the rows of the shown lists not shown yet, counted from Summary (the loaded
  // rows when it is missing), 50 at a time; "Show 50 more" when a shown list is capped.
  function renderMore() {
    const more = $('problem-more');
    const lists = shownLists();
    const merged = mergedRows().length;
    const shown = Math.min(merged, monitor.limit);
    const has = merged > monitor.limit || lists.some((id) => monitor.lists.get(id)?.next);
    more.hidden = !has;
    more.parentNode.hidden = !has;
    if (!has || more.dataset.busy) return;
    const total = lists.reduce((sum, id) => sum + countOf(id), 0);
    const remaining = Math.max(total - shown, merged - shown);
    more.textContent =
      'Show ' + (lists.some(capped) || remaining <= 0 ? PAGE : Math.min(PAGE, remaining)) + ' more';
  }
  function showMore() {
    return ui
      .busy($('problem-more'), 'Loading…', 'monitor', async () => {
        monitor.limit += PAGE;
        await Promise.all(
          shownLists()
            .map((id) => [id, monitor.lists.get(id)])
            .filter(([, state]) => state.next && state.rows.length < monitor.limit)
            .map(([id]) => loadList(id, true)),
        );
      })
      .then(() => ui.withFocus(renderTable));
  }

  // Re-runs in progress, one card each above the chips: progress, who started it, and its
  // actions. Finished and cancelled runs are left out; the strip hides when there are none.
  const activeRuns = () =>
    (monitor.lists.get('TemplateRuns')?.rows || []).filter(
      (r) => r.Run && r.Run.State !== 'Done' && r.Run.State !== 'Cancelled',
    );
  function renderStrip() {
    const runs = activeRuns();
    $('runs-strip').hidden = monitor.setupOpen || !runs.length;
    $('runs-strip').replaceChildren(...runs.map(runCard));
  }
  function runCard(row) {
    const run = row.Run;
    const card = el('section', null, 'run-card');
    card.dataset.key = row.Key;
    card.setAttribute('data-focus-row', '');
    const status = el('span', null, 'row-status');
    const total = (run.TotalEstimated ? 'about ' : '') + number(run.Total);
    const progressText = number(run.Planned) + ' of ' + total;
    const bar = el('progress');
    bar.max = Math.max(1, run.Total);
    bar.value = Math.min(run.Planned, bar.max);
    bar.setAttribute('aria-valuetext', progressText + ', ' + RUN_STATE[run.State].toLowerCase());
    bar.setAttribute('aria-label', 'Progress of ' + run.TemplateName);
    const text = el('p', progressText, 'progress-text');
    if (['Running', 'Waiting', 'Retrying'].includes(run.State)) {
      if (run.EstimatedFinishUtc) text.append(' · ends around ', ui.time(run.EstimatedFinishUtc));
      else text.append(' · Estimating…');
    }
    if (run.State !== 'Running') text.append(' · ' + RUN_STATE[run.State]);
    if (run.State === 'Retrying' && run.NextAttemptUtc)
      text.append(' · next try ', ui.time(run.NextAttemptUtc));
    const head = el('div', null, 'run-head');
    const started = el('p', 'Started ', 'sub');
    started.append(ui.time(run.StartedUtc), run.StartedBy ? ' by ' + run.StartedBy : '');
    head.append(el('h2', 'Re-run of ' + run.TemplateName + ' v' + run.Version), started);
    const progress = el('div', null, 'run-progress');
    progress.append(bar, text);
    if (run.State === 'Blocked' && row.Problem) progress.append(el('p', row.Problem));
    progress.append(status);
    const actions = el('div', null, 'row run-actions');
    actions.setAttribute('data-actions', '');
    for (const action of row.Actions || [])
      rowButton(
        actions,
        'TemplateRuns',
        row,
        action,
        LABEL[action],
        action === 'CancelRun' ? 'danger' : 'secondary',
        card,
        status,
      );
    card.append(head, progress, actions);
    return card;
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

  // A problem table row with explicit roles: phone widths lay rows out as cards, which some
  // browsers would no longer read as a table.
  function asRow(tr) {
    tr.setAttribute('role', 'row');
    for (const td of tr.children) td.setAttribute('role', 'cell');
    return tr;
  }
  // A cell named for its column, so narrow screens can lay a row out as a card.
  const classed = (td, name) => {
    td.className = name;
    return td;
  };
  // One table row: type, item, problem, since, then the primary action and the ⋯ menu. Under the
  // Not captured filter a row the server can re-run starts with its checkbox.
  function renderRow(list, row) {
    const tr = el('tr');
    tr.setAttribute('data-focus-row', '');
    tr.dataset.key = row.Key;
    const status = el('span', null, 'row-status');
    const nameId = 'name-' + list + '-' + tr.dataset.key.replace(/[^\w-]/g, '_');
    if (monitor.filter === 'NotCaptured') {
      // A team row or a row without a record cannot be re-run (the server offers no Rerun).
      let box = null;
      if (row.Actions?.includes('Rerun')) {
        box = el('input');
        box.type = 'checkbox';
        box.dataset.key = row.Key;
        box.checked = monitor.selected.has(row.Key);
        box.setAttribute('aria-labelledby', nameId);
        box.onchange = () => {
          if (box.checked) monitor.selected.add(row.Key);
          else monitor.selected.delete(row.Key);
          updateSelected();
        };
      }
      tr.append(classed(cell(box), 'select'));
    }
    const type = el('span', PROBLEMS[list].type, 'type');
    type.dataset.tone = PROBLEMS[list].tone;
    tr.append(
      classed(cell(type), 'kind'),
      itemCell(list, row, nameId),
      problemCell(row, status),
      classed(cell(ui.time(row.SinceUtc)), 'since'),
      actionsCell(list, row, tr, status),
    );
    return asRow(tr);
  }
  // A record links to its form with "Table · Template" under it; a job shows its title, with
  // its attempt and next try when it is retrying, else its site or kind.
  function itemCell(list, row, nameId) {
    let td;
    const sub = el('span', null, 'sub');
    if (RECORD_LISTS.includes(list)) {
      td = recordCell(row, nameId);
      sub.textContent = [row.Record?.TableLabel, row.TemplateName].filter(Boolean).join(' · ');
    } else {
      const title = el('strong', row.Title);
      title.id = nameId;
      td = cell(title);
      if (list === 'RetryingJobs') {
        sub.append('Attempt ' + row.Attempt);
        if (row.NextAttemptUtc) sub.append(' · next ', ui.time(row.NextAttemptUtc));
      } else sub.textContent = row.Site?.Name || row.KindLabel || '';
    }
    td.append(sub);
    td.className = 'item';
    return td;
  }

  // The actions a row's buttons offer: the server's, less the ones shown elsewhere (a record's
  // name opens it; an ambiguous setup's candidates each have their own Use this one).
  const offered = (list, row) =>
    (row.Actions || []).filter(
      (a) => a !== 'UseCandidate' && !(a === 'OpenRecord' && RECORD_LISTS.includes(list)),
    );
  const primaryAction = (list, row) =>
    PRIMARY.find((action) => offered(list, row).includes(action)) || null;
  function actionsCell(list, row, tr, status) {
    const td = el('td', null, 'actions');
    td.setAttribute('data-actions', '');
    const primary = primaryAction(list, row);
    if (primary)
      rowButton(
        td,
        list,
        row,
        primary,
        labelFor(list, primary, row),
        primary === 'UseLibrary' || primary === 'CreateAgain' ? 'primary' : 'secondary',
        tr,
        status,
      );
    td.append(rowMenu(list, row, tr, status));
    return td;
  }
  // The ⋯ menu: the row's other actions in the server's order, then Copy ID. An action chosen
  // here works through the ⋯ button: its confirmation returns focus there and it shows busy.
  function rowMenu(list, row, tr, status) {
    const wrap = el('span', null, 'menu-anchor');
    const key = 'row:' + list + ':' + row.Key;
    const trigger = button('⋯', null, 'secondary menu-button');
    trigger.dataset.focusKey = key + ':menu';
    trigger.setAttribute('aria-haspopup', 'menu');
    trigger.setAttribute('aria-expanded', 'false');
    trigger.setAttribute('aria-label', 'More actions for ' + row.Title);
    const menu = el('ul', null, 'menu');
    menu.id = 'menu-' + key.replace(/[^\w-]/g, '_');
    menu.setAttribute('role', 'menu');
    menu.hidden = true;
    trigger.setAttribute('aria-controls', menu.id);
    const item = () => {
      const li = el('li');
      li.setAttribute('role', 'none');
      menu.append(li);
      return li;
    };
    const primary = primaryAction(list, row);
    for (const action of offered(list, row).filter((a) => a !== primary)) {
      const danger = action === 'Cancel' || action === 'CancelRun';
      const choice = rowButton(
        item(),
        list,
        row,
        action,
        labelFor(list, action, row),
        danger ? 'danger' : null,
        tr,
        status,
        null,
        trigger,
      );
      choice.setAttribute('role', 'menuitem');
      choice.tabIndex = -1;
    }
    const copy = button(
      'Copy ID',
      async () => {
        trigger.focus();
        try {
          await navigator.clipboard.writeText(row.Key);
          ui.feedback('monitor', 'Copied');
        } catch {
          ui.feedback('monitor', 'Copy failed', 'error');
        }
      },
      null,
    );
    copy.setAttribute('role', 'menuitem');
    copy.tabIndex = -1;
    copy.setAttribute('aria-label', 'Copy ID for ' + row.Title);
    item().append(copy);
    ui.menu(trigger, menu);
    wrap.append(trigger, menu);
    return wrap;
  }

  // Adds one action button to its container, then marks it disabled with its reason when the
  // caller lacks the privilege (the reason is placed after the button, so it must be attached).
  // `via` is the control the action works through: the row's ⋯ for a menu item.
  function rowButton(
    into,
    list,
    row,
    action,
    text,
    style,
    tr,
    status,
    candidate = null,
    via = null,
  ) {
    const control = button(
      text,
      () => {
        if (ui.blocked(control)) return undefined;
        via?.focus();
        return act(list, row, action, via || control, tr, status, candidate);
      },
      style,
    );
    control.dataset.action = action;
    control.dataset.focusKey =
      'row:' + list + ':' + row.Key + ':' + action + (candidate ? ':' + candidate.ListId : '');
    control.setAttribute('aria-label', text + ' for ' + row.Title);
    into.append(control);
    // The role of the API the action goes to: a library setup's own choices go through the
    // catalog API; Retry, like every other action, through asx_ManageWork.
    const missing =
      action === 'Check' || action === 'OpenRecord'
        ? null
        : row.Kind === 'LibrarySetup' && action !== 'Retry'
          ? ui.needs('prvCreateasx_site')
          : ui.needs(OPERATOR);
    if (missing)
      ui.disable(control, 'reason-' + control.dataset.focusKey.replace(/[^\w-]/g, '_'), missing);
    return control;
  }

  function problemCell(row, status) {
    const td = el('td', null, 'problem');
    const recovery = row.Recovery;
    if (recovery && row.Kind === 'LibrarySetup') {
      td.append(el('p', ui.recoverySentence(row.Title, recovery)));
      if (recovery.State === 'Ambiguous') td.append(candidates(row));
      td.append(status);
      return td;
    }
    td.append(el('p', row.Problem));
    if (row.More?.length) {
      const more = button(
        '+' + row.More.length + ' more',
        () => more.replaceWith(...row.More.map((m) => el('p', m))),
        'link',
      );
      td.append(more);
    }
    if (row.Fix) td.append(el('p', row.Fix, 'muted'));
    td.append(status);
    return td;
  }

  // An ambiguous finding lists each candidate with the checks it failed and its own Use this one.
  // `list` is the area whose feedback line reports the choice: Monitor's or Look up an operation's.
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
            : 'Not queued: the template is off or outside its active dates.'
          : ['Re-run queued for ' + name + '.', ...(result.Notices || [])].join(' ');
      case 'Dismiss':
        return 'Dismissed.';
      default:
        return '';
    }
  }

  // A confirmation for a table row renders in a full-width row under it; elsewhere in a slot at
  // the end of the anchor (a re-run card, a Tables row lays the slot across its full width).
  function confirmHost(anchor) {
    if (anchor?.tagName === 'TR') {
      const holder = el('tr', null, 'confirm-row');
      const host = el('td');
      host.setAttribute('colspan', String(anchor.children.length));
      holder.append(host);
      if (anchor.getAttribute('role') === 'row') asRow(holder);
      anchor.after(holder);
      return { host, remove: () => holder.remove() };
    }
    const host = el('div', null, 'confirm-slot');
    anchor.append(host);
    return { host, remove: () => host.remove() };
  }

  async function act(list, row, action, control, tr, status, candidate) {
    if (ui.blocked(control)) return;
    if (action === 'OpenRecord') return openRecord(row);
    if (action === 'Check')
      return focusLink(
        {
          tab: 'monitor',
          record: row.Record.Table + ':' + row.Record.Id,
          template: row.TemplateId,
        },
        control,
      );
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
      // Gone from the list's rows too, so a later redraw does not bring it back.
      const state = monitor.lists.get(list);
      state.rows = state.rows.filter((r) => r.Key !== row.Key);
      monitor.selected.delete(row.Key);
      ui.withFocus(() => tr.remove());
      updateSelected();
      return;
    }
    if (!tr) return;
    if (result.Run) {
      const changed = { ...row, Run: result.Run, Status: result.Status, Actions: [] };
      tr.replaceChildren(
        ...(list === 'TemplateRuns' ? runCard(changed) : renderRow(list, changed)).children,
      );
    }
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

  // Re-run selected: the checked Not captured rows the server can re-run. The checks are kept
  // by key, so a redraw keeps them.
  const selectedRows = () =>
    (monitor.lists.get('NotCaptured')?.rows || []).filter(
      (r) => monitor.selected.has(r.Key) && r.Actions?.includes('Rerun'),
    );
  function rerunSelected() {
    const run = button('Re-run selected (0)', () => {
      const rows = selectedRows();
      if (!rows.length) return undefined;
      return ui.busy(run, 'Re-running…', 'monitor', async () => {
        let queued = 0;
        const failed = [];
        // Each row on its own: one refusal is reported and the rest are still queued.
        for (const row of rows) {
          const name = row.Record?.Name || row.Title || 'A record';
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
    run.dataset.focusKey = 'rerun-selected';
    return run;
  }
  function updateSelected() {
    const run = monitor.rerunSelected;
    if (run && !run.dataset.busy)
      run.textContent = 'Re-run selected (' + selectedRows().length + ')';
  }

  // The Tools panels. One is open at a time; Escape or Close returns focus to what opened it.
  const panels = { check: null, lookup: null };
  function openCheck(invoker) {
    panels.check = ui.sidePanel($('check-panel'), invoker);
  }
  function openLookup(invoker) {
    panels.lookup = ui.sidePanel($('lookup-panel'), invoker);
    return ensureRecent();
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
            ? 'Not queued: the template is off or outside its active dates.'
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

  // Look up an operation: recent operations and look up by ID. The first page loads once, when
  // the panel opens or a link opens it; a failed read is reported and tried again next time.
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

  // A link or a row's Check opens its panel: an operation in Look up an operation with its result,
  // a record in Check a record ready to check, a re-run at its first action. `invoker` gets focus
  // back when the panel closes.
  async function focusLink(link, invoker = $('monitor-tools')) {
    if (link.operation) {
      await openLookup(invoker);
      $('operation-id').value = link.operation;
      await ui.busy($('operation-lookup'), 'Looking up…', 'advanced', () => lookUp(link.operation));
      $('operation-result-title').focus();
    }
    if (link.record) {
      const [table, id] = link.record.split(':');
      openCheck(invoker);
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
      $('runs-strip')
        .querySelector('[data-key="' + link.run + '"] button')
        ?.focus();
  }

  // Settings ----------------------------------------------------------------------------------

  // The automation switch: SetEnabled at once, never part of Save. While it works its state text
  // says what it is doing. A caller who cannot change settings sees the state as text only.
  function renderSwitch() {
    const runtime = ui.runtime();
    const control = $('automation-switch-settings');
    if (!runtime) return;
    control.hidden = runtime.CanChange === false;
    if (!control.dataset.busy) {
      control.setAttribute('aria-checked', String(!!runtime.Enabled));
      $('automation-state-settings').textContent = runtime.Enabled
        ? 'Automation is running'
        : 'Automation is paused';
    }
    control.onclick = toggleAutomation;
  }
  async function toggleAutomation() {
    const control = $('automation-switch-settings');
    const runtime = ui.runtime();
    if (ui.blocked(control) || control.dataset.busy || !runtime || runtime.CanChange === false)
      return;
    const turnOn = !runtime.Enabled;
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

  // The page starts before the shell's runtime Get may be back: the form stays busy until it
  // is, then follows every runtime change (the switch, Save, Repair, a table added or removed).
  // A refused Get shows a System Administrator the missing-settings alert; anyone else sees
  // whether automation runs, as text.
  async function openSettings() {
    ui.problemPill($('settings-problems'));
    $('automation-settings').setAttribute('aria-busy', 'true');
    $('add-host').onclick = () => {
      addHost('');
      markEdited();
      const inputs = $('hosts-list').querySelectorAll('input');
      inputs[inputs.length - 1].focus();
    };
    $('runtimeWorker').onchange = markEdited;
    $('save-settings').onclick = () =>
      ui.busy($('save-settings'), 'Saving…', 'settings-save', save);
    $('settings-discard').onclick = () => {
      if (!ui.runtime()) return;
      settings.edited = false;
      ui.clearFeedback('settings-save');
      renderSettings(ui.runtime());
      markEdited();
      $('settings-title').focus();
    };
    $('add-table').onclick = () => {
      renderPicker();
      $('table-picker').hidden = !$('table-picker').hidden;
      if (!$('table-picker').hidden) $('enableTable').focus();
    };
    // Choosing a table only enables Add: arrow keys on a closed select fire change.
    $('enableTable').onchange = chooseTable;
    $('add-chosen-table').onclick = () => {
      const name = $('enableTable').value;
      if (ui.blocked($('add-chosen-table')) || !name) return undefined;
      return addTable(name);
    };
    // Once Repair all ends, the button shows what is still pending, not its progress text.
    $('repair-all').onclick = async () => {
      if (ui.blocked($('repair-all'))) return;
      await ui.busy($('repair-all'), 'Repairing…', 'settings', repairAll);
      if (ui.runtime()) renderTables(ui.runtime());
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
    [settings.workers] = await Promise.all([
      workers(),
      loadConnections(),
      loadTables(),
      countTemplates(),
      ui.runtimeReady(),
    ]);
    ui.onRuntime(showSettings);
    await showSettings(ui.runtime());
    const link = ui.deeplink();
    // A link with a table lands on its row; with an empty one (Manage tables), on the card.
    if (link && 'table' in link) focusTable(link.table);
  }

  function showSettings(runtime) {
    $('settings-missing').hidden = !!runtime || !ui.can(ADMIN);
    for (const id of ['automation-card', 'automation-settings', 'tables-card', 'danger-zone'])
      $(id).hidden = !runtime;
    if (!runtime) {
      $('settings-footer').hidden = true;
      $('automation-settings').removeAttribute('aria-busy');
      return ui.can(ADMIN) ? undefined : showAutomationText();
    }
    renderSettings(runtime);
    // Table names come from metadata, read once per table; the logical name shows until then.
    const scopes = (runtime.Registration?.Readiness || []).map((r) => r.Scope);
    if (scopes.some((s) => s !== 'team' && !labels.has(s)))
      loadLabels(scopes).then(() => ui.runtime() && renderTables(ui.runtime()));
  }

  // Only a System Administrator can read the profile. Anyone else sees whether automation runs,
  // as text, from the Default runtime row; with no row to read the card stays hidden.
  async function showAutomationText() {
    const state = await ui.automation();
    if (ui.runtime()) return;
    $('automation-switch-settings').hidden = true;
    $('automation-card').hidden = !state;
    if (state)
      $('automation-state-settings').textContent = state.Enabled
        ? 'Automation is running'
        : 'Automation is paused';
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
  // The tables document management is enabled for, which ＋ Add table offers; their display
  // names label the Tables card too.
  async function loadTables() {
    const context = xrm.Utility.getGlobalContext();
    let next =
      context.getClientUrl() +
      '/api/data/v9.2/EntityDefinitions?$select=LogicalName,DisplayName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,IsDocumentManagementEnabled&$filter=IsDocumentManagementEnabled eq true';
    try {
      const rows = [];
      while (next) {
        const response = await fetch(next, {
          credentials: 'same-origin',
          headers: { Accept: 'application/json', 'OData-Version': '4.0' },
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error?.message || 'Dataverse request failed.');
        rows.push(...data.value);
        next = data['@odata.nextLink'] || null;
      }
      settings.tables = rows
        .filter((t) => t.PrimaryNameAttribute)
        .map((t) => ({
          name: t.LogicalName,
          label: t.DisplayName?.UserLocalizedLabel?.Label || t.LogicalName,
        }))
        .sort((a, b) => a.label.localeCompare(b.label));
      for (const t of settings.tables) labels.set(t.name, t.label);
    } catch (error) {
      settings.tables = [];
      ui.feedback('settings', "Couldn't load tables: " + (error.message || String(error)), 'error');
    }
  }
  // How many templates each table has, from one read of every template's table. Null when the
  // read fails: the rows then show no count.
  async function countTemplates() {
    try {
      const counts = new Map();
      for (const row of await everything('asx_template', '?$select=asx_table'))
        counts.set(row.asx_table, (counts.get(row.asx_table) || 0) + 1);
      settings.counts = counts;
    } catch {
      settings.counts = null;
    }
  }

  const hasWorker = (runtime) => !!runtime.WorkerId && runtime.WorkerId !== EMPTY;
  function workerName(runtime) {
    if (!hasWorker(runtime)) return 'Choose an application user';
    return (
      settings.workers.find((w) => w.systemuserid === runtime.WorkerId)?.fullname ||
      'Current run-as user (disabled or not found)'
    );
  }

  function renderSettings(runtime) {
    const editable = runtime.CanChange !== false;
    // Fields the admin is editing keep their values when the runtime changes elsewhere; Save
    // reads the current row version and Enabled from the shared runtime, not from the form.
    if (!settings.edited) {
      const worker = $('runtimeWorker');
      worker.replaceChildren(...settings.workers.map((w) => option(w.systemuserid, w.fullname)));
      // A disabled or deleted run-as user stays selectable so saving keeps it.
      if (hasWorker(runtime) && !settings.workers.some((w) => w.systemuserid === runtime.WorkerId))
        worker.append(option(runtime.WorkerId, 'Current run-as user (disabled or not found)'));
      if (!hasWorker(runtime)) worker.prepend(option('', 'Choose an application user'));
      worker.value = hasWorker(runtime) ? runtime.WorkerId : '';
      const hosts = runtime.SharePointHosts || [];
      $('hosts-list').replaceChildren();
      if (editable) for (const host of hosts.length ? hosts : ['']) addHost(host, host);
      else $('hosts-list').append(...hosts.map((host) => el('li', host, 'host')));
    }
    // Read-only: the values as text, and no control that changes anything.
    $('runtimeWorker').hidden = $('runtimeWorker').disabled = !editable;
    $('worker-text').hidden = editable;
    $('worker-text').textContent = workerName(runtime);
    for (const id of ['add-host', 'add-table', 'stop-tracking']) $(id).hidden = !editable;
    if (!editable) $('table-picker').hidden = true;
    renderSwitch();
    renderTables(runtime);
    markEdited();
    $('automation-settings').removeAttribute('aria-busy');
  }

  // A host row. `original` is the saved host the row was drawn for ('' for the empty row of a
  // profile with none); a row the admin added has none.
  function addHost(value, original = null) {
    const item = el('li', null, 'host row');
    const field = el('input');
    field.type = 'text';
    field.setAttribute('inputmode', 'url');
    field.setAttribute('spellcheck', 'false');
    field.value = value;
    if (original != null) field.dataset.original = original;
    field.oninput = markEdited;
    const index = $('hosts-list').querySelectorAll('li').length + 1;
    field.setAttribute('aria-label', 'SharePoint host name ' + index);
    const remove = button(
      'Remove',
      () => {
        item.remove();
        markEdited();
        $('add-host').focus();
      },
      'link',
    );
    remove.setAttribute('aria-label', 'Remove host ' + (value || index));
    item.append(field, remove);
    $('hosts-list').append(item);
  }

  // Changes not saved yet: the run-as user and each host row that differs from
  // the saved host it was drawn for, is new, or is gone. Changed and new rows are tinted.
  function unsavedCount(runtime) {
    if (!runtime || runtime.CanChange === false) return 0;
    let count = 0;
    if ($('runtimeWorker').value !== (hasWorker(runtime) ? runtime.WorkerId : '')) count++;
    const kept = new Set();
    for (const field of $('hosts-list').querySelectorAll('input')) {
      const value = field.value.trim().toLowerCase();
      const known = 'original' in field.dataset;
      if (known) kept.add(field.dataset.original);
      const edited = known ? value !== field.dataset.original : value !== '';
      field.classList.toggle('is-edited', edited);
      if (edited) count++;
    }
    return count + (runtime.SharePointHosts || []).filter((h) => !kept.has(h)).length;
  }
  // The save bar shows while there are unsaved changes, with their count.
  function markEdited() {
    const runtime = ui.runtime();
    const count = unsavedCount(runtime);
    settings.edited = count > 0;
    $('settings-footer').hidden = !settings.edited;
    $('settings-unsaved').textContent =
      count === 1 ? '1 unsaved change' : count + ' unsaved changes';
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
    });
    settings.edited = false;
    ui.setRuntime(result);
    ui.feedback('settings', 'Settings saved.');
    $('settings-title').focus();
  }

  // The enabled tables: the profile's tables and every table change tracking reports on.
  const enabledTables = (runtime) =>
    new Set(
      [
        ...(runtime?.Tables || []),
        ...(runtime?.Registration?.Readiness || []).map((r) => r.Scope),
      ].filter((t) => t && t !== 'team'),
    );

  function renderPicker() {
    const enabled = enabledTables(ui.runtime());
    $('enableTable').replaceChildren(
      option('', 'Choose a table'),
      ...(settings.tables || [])
        .filter((t) => !enabled.has(t.name))
        .map((t) => option(t.name, t.label)),
    );
    $('enableTable').value = '';
    chooseTable();
  }
  // Add is available once a table is chosen.
  function chooseTable() {
    if ($('enableTable').value) $('add-chosen-table').removeAttribute('aria-disabled');
    else $('add-chosen-table').setAttribute('aria-disabled', 'true');
  }

  // One row per enabled table, by name: its template count, its change tracking, and Repair and
  // Remove for an administrator. Team access events are not a table: every registration keeps
  // them, so they have no row, and Repair all appears whenever they need repairing.
  function renderTables(runtime) {
    const rows = $('tables-rows');
    // Never under an open confirmation (as Monitor's lists): removeTable redraws when it closes.
    settings.tablesStale = !!rows.querySelector('.confirm[role=group]');
    if (settings.tablesStale) return;
    const registration = runtime.Registration || { Readiness: [] };
    $('tracking-error').hidden = !registration.Error;
    $('tracking-error').textContent = registration.Error || '';
    const editable = runtime.CanChange !== false;
    const status = new Map((registration.Readiness || []).map((r) => [r.Scope, r.Status]));
    const scopes = [...new Set([...(runtime.Tables || []), ...status.keys()])]
      .filter((scope) => scope && scope !== 'team')
      .sort((a, b) => label(a).localeCompare(label(b)));
    rows.replaceChildren(...scopes.map((s) => tableRow(s, status.get(s), editable)));
    const pending =
      registration.Pending ??
      (registration.Readiness || []).filter((r) => TRACKING[r.Status]?.[1]).length;
    const team = status.get('team');
    // As a table's row: a state not listed is offered for repair.
    const teamNeedsRepair = !!team && (TRACKING[team]?.[1] ?? true);
    // The count is of the rows shown; the team access events it also repairs have none.
    const shown = Math.max(0, pending - (teamNeedsRepair ? 1 : 0));
    if (!$('repair-all').dataset.busy) {
      $('repair-all').hidden = (shown < 2 && !teamNeedsRepair) || !editable;
      $('repair-all').textContent = shown ? 'Repair all (' + shown + ')' : 'Repair all';
    }
  }

  function tableRow(scope, state, editable) {
    const name = label(scope);
    const row = el('div', null, 'card-row tables-grid');
    row.dataset.table = scope;
    row.setAttribute('data-focus-row', '');
    row.setAttribute('role', 'listitem');
    const about = el('div');
    about.append(
      el('span', name, 'table-name'),
      el(
        'span',
        settings.counts ? plural(settings.counts.get(scope) || 0, 'template', 'templates') : '',
        'sub',
      ),
    );
    const tracking = el('div', null, 'tracking');
    const [text, repairable] = state ? TRACKING[state] || [state, true] : [null, false];
    if (state) tracking.append(ui.status(state === 'Ready' ? 'ok' : 'attention', text));
    if (state === 'WorkerCannotRead') {
      const how = el('a', 'How to grant access');
      how.setAttribute('href', INSTALL_DOCS + '#configure-and-enable');
      how.setAttribute('target', '_blank');
      how.setAttribute('rel', 'noopener');
      tracking.append(how);
    }
    const actions = el('div', null, 'row table-actions');
    actions.setAttribute('data-actions', '');
    if (editable && repairable) {
      const repair = button('Repair', () => {
        if (ui.blocked(repair)) return undefined;
        return ui.busy(repair, 'Repairing…', 'settings', async () => {
          const result = await runtimeApi({
            Command: 'Register',
            Table: scope,
            RowVersion: ui.runtime().RowVersion,
          });
          ui.withFocus(() => ui.setRuntime(result));
        });
      });
      repair.dataset.focusKey = 'tracking:' + scope;
      repair.setAttribute('aria-label', 'Repair ' + name);
      actions.append(repair);
    }
    if (editable) {
      const remove = button('Remove', () => removeTable(scope, remove), 'link muted');
      remove.dataset.focusKey = 'remove:' + scope;
      remove.setAttribute('aria-label', 'Remove ' + name);
      actions.append(remove);
    }
    row.append(about, tracking, actions);
    return row;
  }

  // A table's row takes focus: its first action, else the row itself.
  function focusTable(table) {
    const row = table && $('tables-rows').querySelector('[data-table="' + table + '"]');
    if (!row) return $('tables-title').focus();
    const first = row.querySelector('button');
    if (first) return first.focus();
    row.tabIndex = -1;
    return row.focus();
  }

  // A refused add clears the choice, so Add waits for the next one.
  async function addTable(name) {
    const picker = $('enableTable');
    let added = false;
    picker.disabled = true;
    await ui.busy($('add-chosen-table'), 'Add', 'settings', async () => {
      const result = await runtimeApi({ Command: 'AddTable', Table: name });
      await countTemplates();
      $('table-picker').hidden = true;
      ui.setRuntime(result);
      ui.feedback('settings', label(name) + ' added.');
      added = true;
    });
    picker.disabled = false;
    if (added) return focusTable(name);
    picker.value = '';
    return chooseTable();
  }

  async function removeTable(table, control) {
    if (ui.blocked(control)) return;
    const place = confirmHost(control.closest('[data-table]'));
    let ok = false;
    try {
      ok = await ui.confirmInline(control, {
        host: place.host,
        text:
          'Stop creating folders for ' +
          label(table) +
          '? Queued folder work for this table is cancelled. Templates are kept, and nothing in SharePoint is deleted.',
        confirm: 'Remove table',
        keep: 'Keep table',
        danger: true,
      });
    } finally {
      place.remove();
      if (settings.tablesStale && ui.runtime()) renderTables(ui.runtime());
    }
    if (!ok) return;
    await ui.busy(control, 'Remove', 'settings', async () => {
      const result = await runtimeApi({ Command: 'RemoveTable', Table: table });
      await countTemplates();
      ui.withFocus(() => ui.setRuntime(result));
      ui.feedback('settings', label(table) + ' removed. Its templates are kept.');
    });
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

  // The connection references the flows use, by name. The package also carries two references
  // that no flow uses, kept from 0.1.0.3; they are not listed.
  const CONNECTIONS = { asx_documentsdataverse: 'Dataverse', asx_documentshttp: 'SharePoint' };
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
        '?$select=connectionreferencelogicalname,connectionid&$filter=' +
          Object.keys(CONNECTIONS)
            .map((name) => "connectionreferencelogicalname eq '" + name + "'")
            .join(' or '),
      );
      const order = Object.keys(CONNECTIONS);
      const index = (r) => order.indexOf(r.connectionreferencelogicalname);
      $('connection-list').replaceChildren(
        ...rows.entities
          .filter((r) => index(r) >= 0)
          .sort((a, b) => index(a) - index(b))
          .map((r) => {
            const item = el('li', null, 'card-row');
            item.append(
              el('span', CONNECTIONS[r.connectionreferencelogicalname]),
              ui.status(
                r.connectionid ? 'ok' : 'attention',
                r.connectionid ? 'Connected' : 'Not connected',
              ),
            );
            return item;
          }),
      );
    } catch (error) {
      // A caller who cannot read the connections does not see the card at all.
      if (!ui.can(ADMIN)) {
        $('connections-card').hidden = true;
        return;
      }
      $('connection-list').replaceChildren(
        el(
          'li',
          "Couldn't read the connections: " + (error.message || String(error)),
          'card-row error',
        ),
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
