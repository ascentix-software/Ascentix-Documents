'use strict';
// Every tab in light and dark: axe-core (WCAG 2.2 AA), keyboard flows, computed contrast, text
// size and target size, and screenshots for the review. Mocked Dataverse; Edge.
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { chromium } = require(process.env.ASXD_PLAYWRIGHT_MODULE || 'playwright');
const root = path.resolve(__dirname, '../../client/admin');
const axe = require.resolve(process.env.ASXD_AXE_MODULE || 'axe-core');
const shots = path.resolve(__dirname, '../artifacts/browser');
const types = {
  '.html': 'text/html',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.svg': 'image/svg+xml',
};
fs.mkdirSync(shots, { recursive: true });

// Waits until the shown tab has loaded: a panel is visible, nothing in it is aria-busy, and the
// mocked Dataverse has had no call in flight for a moment. (Only the shown panel counts: the
// Settings form is aria-busy until Settings opens, also while another tab is shown.)
const settle = (page) =>
  page.waitForFunction(() => {
    const panel = document.querySelector('[role=tabpanel]:not([hidden])');
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
  await page.goto('https://asxd.test/index.html?data=' + tab + '-ui20261006nav1');
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
  const panel = document.querySelector('[role=tabpanel]:not([hidden])');
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

// Opens the first table in the rail, then its template, and waits for it to load.
async function openTemplate(page) {
  await page.locator('#templateTree summary').first().click();
  await page.locator('#templateTree').getByRole('button', { name: 'Account onboarding' }).click();
  await page.locator('#destinations').getByRole('button', { name: 'General' }).waitFor();
  await settle(page);
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
        // The rail's tables are collapsed <details>; open the first, then the template.
        if (tab === 'templates') await openTemplate(page);
        await page.addScriptTag({ path: axe });
        const result = await page.evaluate(async () =>
          (
            await window.axe.run(document, {
              runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21aa', 'wcag22aa'] },
            })
          ).violations.map((v) => v.id + ': ' + v.nodes.map((n) => n.target.join(' ')).join(', ')),
        );
        assert.deepEqual(result, [], tab + ' ' + scheme);
        assert.deepEqual(page.errors, [], tab + ' ' + scheme + ' page errors');
        assert.deepEqual(
          await page.evaluate(computed),
          { weakBorders: [], smallText: [], smallTargets: [] },
          tab + ' ' + scheme,
        );
        await page.screenshot({
          path: path.join(shots, 'a11y-' + tab + '-' + scheme + '.png'),
          fullPage: true,
        });
        await context.close();
      }
    // Keyboard flows (spec 5.6).
    const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
    const page = await open(context, 'templates');
    const focused = () =>
      page.evaluate(
        () =>
          document.activeElement.id ||
          document.activeElement.getAttribute('data-focus-key') ||
          document.activeElement.textContent.trim(),
      );
    await page.locator('#tab-templates').focus();
    await page.keyboard.press('ArrowRight');
    assert.equal(await focused(), 'tab-access');
    await page.keyboard.press('End');
    assert.equal(await focused(), 'tab-settings');
    await page.keyboard.press('Enter');
    await page.waitForURL(/data=settings-/);
    await page.waitForFunction(() => document.activeElement?.id === 'tab-settings');
    const templates = await open(context, 'templates');
    await openTemplate(templates);
    await templates.locator('#destinations').getByRole('button', { name: 'General' }).focus();
    await templates.keyboard.press('Enter');
    assert.match(
      await templates.evaluate(() => document.activeElement.getAttribute('data-focus-key')),
      /^node:/,
    );
    await templates.locator('#template-menu').focus();
    await templates.keyboard.press('ArrowDown');
    assert.equal(await templates.evaluate(() => document.activeElement.id), 'menu-history');
    await templates.keyboard.press('Escape');
    assert.equal(await templates.evaluate(() => document.activeElement.id), 'template-menu');
    await templates.getByLabel('When should this folder appear?').selectOption('conditional');
    assert.equal(
      await templates.evaluate(() => document.activeElement.getAttribute('aria-label')),
      'Field, condition 1',
    );
    const monitor = await open(context, 'monitor');
    await monitor
      .locator('#monitor-tiles')
      .getByRole('button', { name: /^Blocked jobs/ })
      .click();
    assert.equal(await monitor.evaluate(() => document.activeElement.id), 'h-BlockedJobs');
    await monitor.getByRole('button', { name: /^Cancel job for/ }).click();
    assert.match(
      await monitor.evaluate(() => document.activeElement.textContent),
      /^Cancel the folder job for /,
    );
    await monitor.keyboard.press('Escape');
    assert.match(
      await monitor.evaluate(() => document.activeElement.getAttribute('aria-label')),
      /^Cancel job for/,
    );
    for (const state of ['Checking', 'Found', 'NotFound', 'Ambiguous']) {
      const shown = await open(context, 'monitor', 'window.__recovery = ' + JSON.stringify(state));
      await shown.screenshot({
        path: path.join(shots, 'monitor-recovery-' + state.toLowerCase() + '.png'),
        fullPage: true,
      });
    }
    console.log(
      'PASS accessibility in Edge: axe WCAG 2.2 AA on four tabs in light and dark, computed borders, text and targets, keyboard flows, screenshots. Mocked Dataverse.',
    );
  } finally {
    await browser.close();
  }
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
