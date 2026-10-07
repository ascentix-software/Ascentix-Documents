'use strict';
// Every page in light and dark: axe-core (WCAG 2.2 AA), keyboard flows, computed contrast, text
// size and target size, and screenshots for the review. Mocked Dataverse; Edge.
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { chromium } = require(process.env.ASXD_PLAYWRIGHT_MODULE || 'playwright');
const root = path.resolve(__dirname, '../../client/admin');
const axe = require.resolve(process.env.ASXD_AXE_MODULE || 'axe-core');
const shots = path.resolve(__dirname, '../artifacts/browser');
const BUILD = /const BUILD = '([^']+)'/.exec(
  fs.readFileSync(path.join(root, 'shell.js'), 'utf8'),
)[1];
const types = {
  '.html': 'text/html',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.svg': 'image/svg+xml',
};
fs.mkdirSync(shots, { recursive: true });

// Waits until the shown page has loaded: a page is visible, nothing in it is aria-busy, and the
// mocked Dataverse has had no call in flight for a moment. (Only the shown page counts: the
// Settings form is aria-busy until Settings opens, also while another page is shown.)
const settle = (page) =>
  page.waitForFunction(() => {
    const panel = document.querySelector('section.page:not([hidden])');
    return !!panel && !panel.querySelector('[aria-busy=true]') && window.__mockIdle?.();
  });

async function open(context, tab, extra = '') {
  const page = await context.newPage();
  page.errors = [];
  page.on('pageerror', (e) => page.errors.push(e.message));
  await page.route('https://asxd.test/**', (route) => {
    const file = path.join(root, new URL(route.request().url()).pathname.replace(/^\//, ''));
    route.fulfill({
      status: 200,
      contentType: types[path.extname(file)] || 'text/plain',
      body: fs.readFileSync(file),
    });
  });
  await page.addInitScript({ path: path.join(__dirname, 'mock-xrm.js') });
  await page.addInitScript(() => sessionStorage.setItem('asxd.launched', '1'));
  if (extra) await page.addInitScript(extra);
  await page.goto('https://asxd.test/index.html?data=' + tab + '-' + BUILD);
  await settle(page);
  return page;
}

// The computed checks of spec 7.1: control borders at least 3:1 against their background, no
// visible text under 12px, and every interactive target at least 24×24.
const computed = () => {
  const lum = (c) => {
    const v = (c.match(/\d+(\.\d+)?/g) || [])
      .slice(0, 3)
      .map((n) => n / 255)
      .map((n) => (n <= 0.03928 ? n / 12.92 : ((n + 0.055) / 1.055) ** 2.4));
    return 0.2126 * v[0] + 0.7152 * v[1] + 0.0722 * v[2];
  };
  const ratio = (a, b) => {
    const x = lum(a),
      y = lum(b);
    return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05);
  };
  const bg = (e) => {
    for (let n = e.parentElement; n; n = n.parentElement) {
      const c = getComputedStyle(n).backgroundColor;
      if (!/rgba\(.*,\s*0\)$|transparent/.test(c)) return c;
    }
    return 'rgb(255,255,255)';
  };
  const shown = (e) => e.getClientRects().length > 0 && getComputedStyle(e).visibility !== 'hidden';
  const panel = document.querySelector('section.page:not([hidden])');
  const controls = [
    ...document.querySelectorAll('input,select,textarea,button,[role=switch],[role=tab],summary,a'),
  ].filter(shown);
  return {
    weakBorders: controls
      .filter(
        (e) =>
          /^(INPUT|SELECT|TEXTAREA)$/.test(e.tagName) &&
          ratio(getComputedStyle(e).borderTopColor, bg(e)) < 3,
      )
      .map((e) => e.id || e.outerHTML.slice(0, 60)),
    smallText: [...panel.querySelectorAll('*')]
      .filter(
        (e) =>
          shown(e) &&
          [...e.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim()) &&
          parseFloat(getComputedStyle(e).fontSize) < 12,
      )
      .map((e) => e.textContent.trim().slice(0, 40)),
    smallTargets: controls
      .filter((e) => {
        const r = e.getBoundingClientRect();
        return r.width < 24 || r.height < 24;
      })
      .map((e) => (e.textContent || e.getAttribute('aria-label') || e.id).trim().slice(0, 40)),
  };
};

// The page opens on the first template's overview; Edit template opens it in the editor.
async function openTemplate(page) {
  await page.locator('#overview-title', { hasText: 'Account onboarding' }).waitFor();
  await page.locator('#overview-edit').click();
  await page.locator('#step-1 .destination-card').waitFor();
  await settle(page);
}
// axe-core (WCAG 2.2 AA), page errors and the computed checks for the page as it is shown.
async function check(page, label) {
  if (!(await page.evaluate(() => !!window.axe))) await page.addScriptTag({ path: axe });
  const result = await page.evaluate(async () =>
    (
      await window.axe.run(document, {
        runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21aa', 'wcag22aa'] },
      })
    ).violations.map((v) => v.id + ': ' + v.nodes.map((n) => n.target.join(' ')).join(', ')),
  );
  assert.deepEqual(result, [], label);
  assert.deepEqual(page.errors, [], label + ' page errors');
  assert.deepEqual(
    await page.evaluate(computed),
    { weakBorders: [], smallText: [], smallTargets: [] },
    label,
  );
  await page.screenshot({
    path: path.join(shots, 'a11y-' + label.replace(/ /g, '-') + '.png'),
    fullPage: true,
  });
}

(async () => {
  const browser = await chromium.launch({ headless: true, channel: 'msedge' });
  try {
    for (const scheme of ['light', 'dark'])
      for (const tab of ['templates', 'access', 'monitor', 'settings']) {
        const context = await browser.newContext({
          viewport: { width: 1440, height: 1000 },
          colorScheme: scheme,
        });
        const page = await open(context, tab);
        // Folder templates: the overview, with its Schedule panel open, then the editor.
        if (tab === 'templates') {
          await page.locator('#overview-title', { hasText: 'Account onboarding' }).waitFor();
          await settle(page);
          await check(page, 'templates-overview ' + scheme);
          await page.getByRole('button', { name: 'Edit schedule' }).click();
          await page.getByRole('dialog', { name: 'Schedule' }).waitFor();
          await check(page, 'templates-schedule ' + scheme);
          await page.keyboard.press('Escape');
          await openTemplate(page);
        }
        // Sites & access: the access drawer of the first library is checked with the table.
        if (tab === 'access') {
          await page
            .locator('#ad-libraries')
            .getByRole('button', { name: 'General', exact: true })
            .click();
          await page.getByRole('dialog', { name: 'General' }).waitFor();
          await settle(page);
        }
        await check(page, tab + ' ' + scheme);
        await context.close();
      }
    // Keyboard flows (spec 5.6).
    const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
    // Tab from the top of Settings reaches the automation switch; Space toggles it.
    const settings = await open(context, 'settings');
    const settingsFocus = () => settings.evaluate(() => document.activeElement?.id);
    for (let i = 0; i < 20 && (await settingsFocus()) !== 'automation-switch-settings'; i++)
      await settings.keyboard.press('Tab');
    assert.equal(await settingsFocus(), 'automation-switch-settings');
    const before = await settings
      .locator('#automation-switch-settings')
      .getAttribute('aria-checked');
    await settings.keyboard.press('Space');
    await settings.waitForFunction(
      (was) =>
        document.getElementById('automation-switch-settings').getAttribute('aria-checked') !== was,
      before,
    );
    const templates = await open(context, 'templates');
    await templates.locator('#overview-title', { hasText: 'Account onboarding' }).waitFor();
    const active = () => templates.evaluate(() => document.activeElement.id);
    await templates.locator('#overview-menu').focus();
    await templates.keyboard.press('ArrowDown');
    assert.equal(await active(), 'menu-rerun');
    await templates.keyboard.press('Escape');
    assert.equal(await active(), 'overview-menu');
    // Tab out of the open menu closes it.
    await templates.keyboard.press('ArrowDown');
    await templates.keyboard.press('Tab');
    assert.equal(await templates.locator('#overview-menu-list').isHidden(), true);
    assert.equal(await templates.locator('#overview-menu').getAttribute('aria-expanded'), 'false');
    // All versions opens a side panel at its heading; Escape returns to All versions.
    await templates.getByRole('button', { name: 'All versions' }).focus();
    await templates.keyboard.press('Enter');
    assert.equal(await active(), 'history-title');
    await templates.keyboard.press('Escape');
    assert.equal(await templates.locator('#history-panel').isHidden(), true);
    assert.equal(
      await templates.evaluate(() => document.activeElement.textContent),
      'All versions',
    );
    // Edit template opens the editor at its heading; Close returns to Edit template.
    await openTemplate(templates);
    assert.equal(await active(), 'editor-title');
    await templates.locator('#editor-close').click();
    await templates.locator('#template-overview').waitFor();
    await templates.waitForFunction(() => document.activeElement?.id === 'overview-edit');
    await openTemplate(templates);
    // The steps are tabs: arrow keys move focus, Enter shows the step.
    await templates.locator('#step-tab-1').focus();
    await templates.keyboard.press('ArrowRight');
    assert.equal(await active(), 'step-tab-2');
    await templates.keyboard.press('Enter');
    await templates.locator('#destinations').getByRole('button', { name: 'General' }).focus();
    await templates.keyboard.press('Enter');
    assert.match(
      await templates.evaluate(() => document.activeElement.getAttribute('data-focus-key')),
      /^node:/,
    );
    await templates.getByLabel('When should this folder appear?').selectOption('conditional');
    assert.equal(
      await templates.evaluate(() => document.activeElement.getAttribute('aria-label')),
      'Field, condition 1',
    );
    const monitor = await open(context, 'monitor');
    const focused = (attr) =>
      monitor.evaluate((name) => document.activeElement?.getAttribute(name), attr);
    // A chip filters the table and keeps focus.
    await monitor.getByRole('button', { name: 'Blocked jobs · 1' }).click();
    assert.equal(await focused('data-focus-key'), 'filter:BlockedJobs');
    assert.equal(await monitor.locator('#problem-rows tr').count(), 1);
    // The row's ⋯ opens with the keyboard; Cancel job confirms; Escape keeps and returns to ⋯.
    await monitor.getByRole('button', { name: /^More actions for Contoso Ltd/ }).focus();
    await monitor.keyboard.press('ArrowDown');
    assert.equal(await monitor.evaluate(() => document.activeElement.textContent), 'Cancel job');
    await monitor.keyboard.press('Enter');
    assert.match(
      await monitor.evaluate(() => document.activeElement.textContent),
      /^Cancel the folder job for /,
    );
    await monitor.keyboard.press('Escape');
    assert.match(await focused('aria-label'), /^More actions for Contoso Ltd/);
    // Tools opens Check a record as a side panel; Escape closes it and returns to Tools.
    await monitor.locator('#monitor-tools').focus();
    await monitor.keyboard.press('ArrowDown');
    await monitor.keyboard.press('Enter');
    assert.equal(await monitor.evaluate(() => document.activeElement.id), 'check-title');
    assert.equal(await monitor.locator('#check-panel').getAttribute('role'), 'dialog');
    await monitor.keyboard.press('Escape');
    assert.equal(await monitor.locator('#check-panel').isHidden(), true);
    assert.equal(await monitor.evaluate(() => document.activeElement.id), 'monitor-tools');
    for (const state of ['Checking', 'Found', 'NotFound', 'Ambiguous']) {
      const shown = await open(context, 'monitor', 'window.__recovery = ' + JSON.stringify(state));
      await shown.screenshot({
        path: path.join(shots, 'monitor-recovery-' + state.toLowerCase() + '.png'),
        fullPage: true,
      });
    }
    console.log(
      'PASS accessibility in Edge: axe WCAG 2.2 AA on four pages in light and dark (Folder templates as the overview, its Schedule panel and the editor; Sites & access with its access drawer open), computed borders, text and targets, keyboard flows (the overview ⋯ menu, All versions, Edit template and Close, Monitor chips, a row menu and its confirmation, the Tools panel), screenshots. Mocked Dataverse.',
    );
  } finally {
    await browser.close();
  }
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
