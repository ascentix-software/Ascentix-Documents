// The TEST environment checker (spec 7.4). The controller pastes it into the claude-in-chrome
// javascript_tool on the Documents app's top page; it finds the admin frame and returns JSON
// about the shown tab. It reads the page only: no innerHTML assignment, no changes. Its
// outerHTML.slice(0, 80) reads are diagnostics in a test file that verify-admin.cjs does not scan.
(() => {
  const frames = [...document.querySelectorAll('iframe')]
    .map((f) => { try { return f.contentDocument; } catch { return null; } })
    .filter(Boolean);
  const d = frames.find((x) => x.getElementById('tab-monitor')) || document;
  const w = d.defaultView;
  const shown = (e) => e.getClientRects().length > 0 && w.getComputedStyle(e).visibility !== 'hidden';
  const panel = d.querySelector('[role=tabpanel]:not([hidden])');
  const clone = panel.cloneNode(true);
  clone.querySelectorAll('.details, details.advanced').forEach((x) => x.remove());
  const lum = (c) => {
    const v = (c.match(/\d+(\.\d+)?/g) || []).slice(0, 3).map((n) => n / 255)
      .map((n) => (n <= 0.03928 ? n / 12.92 : ((n + 0.055) / 1.055) ** 2.4));
    return 0.2126 * v[0] + 0.7152 * v[1] + 0.0722 * v[2];
  };
  const ratio = (a, b) => { const x = lum(a), y = lum(b); return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05); };
  const bg = (e) => { for (let n = e.parentElement; n; n = n.parentElement) {
    const c = w.getComputedStyle(n).backgroundColor; if (!/rgba\(.*,\s*0\)$|transparent/.test(c)) return c; } return 'rgb(255,255,255)'; };
  const controls = [...panel.querySelectorAll('input,select,textarea,button,[role=switch],[role=tab]')].filter(shown);
  const name = (e) => e.getAttribute('aria-label') || e.getAttribute('aria-labelledby') || (e.labels && e.labels.length) || e.textContent.trim();
  return JSON.stringify({
    activeTab: d.querySelector('[role=tab][aria-selected=true]')?.id,
    focused: d.activeElement?.id || d.activeElement?.tagName,
    tabs: [...d.querySelectorAll('[role=tab]')].map((t) => t.textContent.trim()),
    banner: !!d.getElementById('status'),
    chip: d.getElementById('automationChip')?.textContent.trim(),
    helpTexts: [...d.querySelectorAll('.help')].filter(shown).length,
    recoveryInputs: ['recoveryRun', 'recoveryToken', 'recoveryEvidence', 'recoveryResponse'].filter((id) => d.getElementById(id)),
    guidOrKey: (clone.innerText.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}|\b(folderjob|librarycreate|catalogprobe|policywork|templaterun):\S+|\d{4}-\d\d-\d\dT\d\d:\d\d/gi) || []).slice(0, 5),
    unnamed: controls.filter((e) => !name(e)).map((e) => e.outerHTML.slice(0, 80)),
    brokenRefs: [...d.querySelectorAll('[aria-describedby],[aria-labelledby]')].flatMap((e) =>
      ((e.getAttribute('aria-describedby') || '') + ' ' + (e.getAttribute('aria-labelledby') || '')).split(/\s+/).filter((id) => id && !d.getElementById(id))),
    smallText: [...panel.querySelectorAll('*')].filter((e) => shown(e) && [...e.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim()) && parseFloat(w.getComputedStyle(e).fontSize) < 12).length,
    weakBorders: controls.filter((e) => /^(INPUT|SELECT|TEXTAREA)$/.test(e.tagName) && ratio(w.getComputedStyle(e).borderTopColor, bg(e)) < 3).length,
    smallTargets: controls.filter((e) => { const r = e.getBoundingClientRect(); return r.width < 24 || r.height < 24; }).map((e) => (e.textContent || e.getAttribute('aria-label') || '').trim().slice(0, 30)),
  });
})();
