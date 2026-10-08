'use strict';
// Form script: keeps a record's Documents hidden until the record has a usable document location.
// When the Documents tab opens on a record without one, Dynamics creates its own folder and
// location, and there is no setting to stop it. Documents creates the location once the SharePoint
// folder exists, so this hides Documents until then.
//
// Register AscentixDocuments.Form.onLoad as an On Load handler with "Pass execution context as
// first parameter" checked. Optional parameters name more tabs or sections that show documents
// ("tab" or "tab.section"); the Related › Documents item (navSPDocuments) is always handled.
// Only formContext APIs are used: no browser dialogs.
(() => {
  const NAV = 'navSPDocuments';
  const NOTICE = 'asx-documents-pending';
  const PENDING =
    'Folders are being created in SharePoint. The Documents tab will be visible when complete.';
  const LATE =
    "Folders aren't ready yet. Reopen this record later, or ask an administrator to check Monitor in Ascentix Documents.";
  // Re-check every 5 seconds for up to 2 minutes.
  const INTERVAL = 5000;
  const CHECKS = 24;
  // The platform's own Documents-tab query: an active, usable location with a site collection.
  const FILTER = [
    'statecode eq 0',
    'statuscode eq 1',
    'servicetype eq 0',
    'locationtype eq 0',
    'sitecollectionid ne null',
  ].join(' and ');

  // One state per form: its names, the record it checks, and its timer.
  const forms = new WeakMap();
  const polling = new Set();

  const guid = (value) =>
    String(value || '')
      .replace(/[{}]/g, '')
      .toLowerCase();

  // Parameters arrive as separate strings, or as one comma-separated string.
  function names(params) {
    return params
      .filter((p) => typeof p === 'string')
      .flatMap((p) => p.split(','))
      .map((p) => p.trim())
      .filter(Boolean);
  }

  function parts(formContext, list) {
    const ui = formContext.ui;
    const found = [];
    const nav = ui.navigation && ui.navigation.items && ui.navigation.items.get(NAV);
    if (nav) found.push(nav);
    for (const name of list) {
      const [tabName, sectionName] = name.split('.');
      const tab = ui.tabs && ui.tabs.get(tabName);
      if (!tab) continue;
      if (!sectionName) {
        found.push(tab);
        continue;
      }
      const section = tab.sections && tab.sections.get(sectionName);
      if (section) found.push(section);
    }
    return found;
  }

  // A form that has closed, or a control it no longer offers, is ignored.
  function show(state, visible) {
    let found = [];
    try {
      found = parts(state.formContext, state.names);
    } catch {
      return;
    }
    for (const part of found) {
      try {
        part.setVisible(visible);
      } catch {
        // Ignored, as above.
      }
    }
  }

  function notify(state, message) {
    try {
      const ui = state.formContext.ui;
      if (message) ui.setFormNotification(message, 'INFO', NOTICE);
      else ui.clearFormNotification(NOTICE);
    } catch {
      // Ignored, as above.
    }
  }

  function stop(state) {
    if (state.timer !== null) clearTimeout(state.timer);
    state.timer = null;
    state.run++;
    polling.delete(state);
  }

  function currentId(state) {
    try {
      return guid(state.formContext.data.entity.getId());
    } catch {
      return null;
    }
  }

  async function located(id) {
    const query =
      '?$select=sharepointdocumentlocationid&$filter=_regardingobjectid_value eq ' +
      id +
      ' and ' +
      FILTER +
      '&$top=1';
    const result = await Xrm.WebApi.retrieveMultipleRecords('sharepointdocumentlocation', query);
    return Boolean(result && result.entities && result.entities.length);
  }

  // Shows or hides Documents for the record now on the form. Read refused or failed: show
  // everything with no notice, so nobody is stranded.
  async function check(state) {
    stop(state);
    const run = state.run;
    const id = currentId(state);
    state.id = id;
    if (!id) {
      // A new record: hidden until it is saved; post-save checks again.
      show(state, false);
      notify(state, null);
      return;
    }
    let attempt = 0;
    const poll = async () => {
      state.timer = null;
      let found;
      try {
        found = await located(id);
      } catch {
        found = null;
      }
      if (run !== state.run) return;
      if (found !== false) {
        stop(state);
        show(state, true);
        notify(state, null);
        return;
      }
      show(state, false);
      if (attempt >= CHECKS) {
        stop(state);
        notify(state, LATE);
        return;
      }
      if (attempt === 0) notify(state, PENDING);
      attempt++;
      polling.add(state);
      state.timer = setTimeout(tick, INTERVAL);
    };
    // The form closed or moved to another record: stop. A reload or post-save checks again.
    const tick = () => {
      if (run !== state.run) return;
      if (currentId(state) !== id) {
        stop(state);
        return;
      }
      poll();
    };
    await poll();
  }

  function onLoad(executionContext, ...params) {
    const formContext =
      executionContext && typeof executionContext.getFormContext === 'function'
        ? executionContext.getFormContext()
        : null;
    // Without the execution context there is nothing to hide; Documents stays as the form has it.
    if (!formContext) return Promise.resolve();
    let state = forms.get(formContext);
    if (!state) {
      state = { formContext, names: [], id: null, timer: null, run: 0 };
      state.postSave = () => check(state);
      forms.set(formContext, state);
    }
    state.names = names(params);
    const entity = formContext.data.entity;
    if (typeof entity.addOnPostSave === 'function') {
      entity.removeOnPostSave(state.postSave);
      entity.addOnPostSave(state.postSave);
    }
    return check(state);
  }

  // A page that unloads stops every poll it started.
  if (typeof window.addEventListener === 'function') {
    window.addEventListener('pagehide', () => [...polling].forEach(stop));
  }

  const root = window.AscentixDocuments || (window.AscentixDocuments = {});
  root.Form = { ...root.Form, onLoad };
})();
