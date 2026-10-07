'use strict';
// Sites & access in headless Edge on the real page (index.html, CSS and every script) with the
// mocked Dataverse of mock-xrm.js: a deleted team, staging and Apply, existing libraries,
// library creation, Add site, the Remove library confirmation, at 1440/800/400 in light and dark.
const fs = require('fs'),
  path = require('path'),
  assert = require('assert/strict');
const { chromium } = require(process.env.ASXD_PLAYWRIGHT_MODULE || 'playwright');
const root = path.resolve(__dirname, '../../client/admin');
const evidence =
  process.env.ASXD_BROWSER_EVIDENCE_DIR || path.resolve(__dirname, '../../artifacts/browser');
const types = {
  '.html': 'text/html',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.svg': 'image/svg+xml',
};
fs.mkdirSync(evidence, { recursive: true });

// As test-admin-a11y-browser.cjs: the page from client/admin, mock-xrm.js before it loads.
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

(async () => {
  const browser = await chromium.launch({ headless: true, channel: 'msedge' });
  try {
    const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
    // A team deleted in Dataverse still has access to the library; Apply removes it.
    const page = await open(
      context,
      'access',
      `(() => {
        const deleted = { TeamId: '00000040-0000-4000-8000-000000000000', Access: 'Read' };
        window.__mock.policy = {
          Status: 'Applied',
          RowVersion: '1',
          Policy: { Desired: [deleted], Applied: [deleted] },
          Teams: [{ TeamId: deleted.TeamId, Name: 'AcceptC Team 1', Deleted: true }],
        };
      })()`,
    );
    const requests = (command) =>
      page.evaluate((c) => window.__mock.requests.filter((r) => r.Command === c), command);
    await page.getByRole('heading', { name: 'General', exact: true }).waitFor();
    // The deleted team is shown as one, with what happens next, and Apply is enabled for it.
    const deletedRow = page.locator('#ad-teams tr').first();
    await deletedRow.getByText('Deleted team: AcceptC Team 1').waitFor();
    await deletedRow
      .getByText('Documents removes its access the next time access is applied.', { exact: true })
      .waitFor();
    assert.equal(await deletedRow.locator('select').isDisabled(), true);
    assert.equal(await page.locator('#ad-apply').isDisabled(), false);
    for (const width of [1440, 400]) {
      await page.setViewportSize({ width, height: 1000 });
      assert.equal(
        await page.evaluate(() => document.documentElement.scrollWidth > innerWidth),
        false,
        'Overflow with a deleted team ' + width,
      );
      await page.screenshot({
        path: path.join(evidence, `deleted-team-${width}.png`),
        fullPage: true,
      });
    }
    await page.setViewportSize({ width: 1440, height: 1000 });
    const operations = await page.evaluate(() => window.__mock.ids.operations);
    await page.locator('#ad-add-team').click();
    await page.locator('#ad-team-choice').selectOption(operations);
    await page.locator('#ad-stage-team').click();
    assert.equal((await requests('ApplyPolicy')).length, 0);
    await page.locator('#ad-apply').click();
    await page.getByText('Access submitted.', { exact: true }).waitFor();
    assert.equal(await page.locator('#ad-apply').isDisabled(), true);
    // Applied: the deleted team's row is gone and nothing failed.
    assert.equal(await page.getByText('Deleted team: AcceptC Team 1').count(), 0);
    assert.equal(
      await page.getByText('Status updates are unavailable', { exact: false }).count(),
      0,
    );
    await page.evaluate(() => {
      const desired = window.__mock.policy.Policy.Desired;
      window.__mock.policy = {
        Status: 'Applied',
        RowVersion: '3',
        Policy: { Desired: desired, Applied: desired },
      };
    });
    await page
      .getByText('Access and team membership confirmed.', { exact: true })
      .waitFor({ timeout: 12000 });
    await page.locator('#ad-existing').click();
    await page
      .getByRole('button', { name: 'Add Archive', exact: true })
      .waitFor({ timeout: 12000 });
    await page.getByRole('button', { name: 'Add Archive', exact: true }).click();
    await page.getByText('Checking Archive and setting up its navigation.').waitFor();
    assert.equal((await requests('AddLibrary'))[0].NativeParentId, undefined);
    await page.locator('#ad-create').click();
    await page.locator('#ad-library-name').fill('Projects');
    await page.locator('#ad-initial-team').selectOption(operations);
    await page.locator('#ad-provision').click();
    await page.getByText('Creating Projects.', { exact: true }).waitFor();
    assert.equal((await requests('CreateLibrary'))[0].Entries.length, 1);
    await page.getByText('Waiting for confirmation', { exact: true }).waitFor({ timeout: 12000 });
    assert.equal(await page.getByText('Open Administration', { exact: false }).count(), 0);
    assert.equal(
      await page.getByRole('progressbar', { name: 'Projects setup progress' }).count(),
      1,
    );
    for (const width of [1440, 800, 400])
      for (const scheme of ['light', 'dark']) {
        await page.setViewportSize({ width, height: 1000 });
        await page.emulateMedia({ colorScheme: scheme });
        assert.equal(
          await page.evaluate(() => document.documentElement.scrollWidth > innerWidth),
          false,
          'Overflow ' + width + ' ' + scheme,
        );
        const alignment = await page.locator('.ad-progress-card').evaluateAll((cards) =>
          cards.map((card) => {
            const bar = card.querySelector('progress'),
              r = bar.getBoundingClientRect(),
              marks = [...card.querySelectorAll('.ad-step-mark')].map((m) => {
                const b = m.getBoundingClientRect();
                return b.left + b.width / 2;
              });
            return {
              start: Math.abs(r.left - marks[0]),
              end: Math.abs(r.right - marks[marks.length - 1]),
              fill: Math.abs(r.left + (r.width * bar.value) / bar.max - marks[bar.value]),
            };
          }),
        );
        for (const a of alignment)
          assert(
            a.start < 1 && a.end < 1 && a.fill < 1,
            'Stage/bar alignment ' + JSON.stringify(a),
          );
        await page.screenshot({
          path: path.join(evidence, `runtime-${width}-${scheme}.png`),
          fullPage: true,
        });
      }
    await page.emulateMedia({ colorScheme: 'light' });
    await page.setViewportSize({ width: 1440, height: 1000 });
    assert.notEqual(
      await page
        .locator('.is-loading .is-current .ad-step-mark')
        .first()
        .evaluate((el) => getComputedStyle(el).animationName),
      'none',
    );
    await page.emulateMedia({ reducedMotion: 'reduce' });
    assert.equal(
      await page
        .locator('.is-loading .is-current .ad-step-mark')
        .first()
        .evaluate((el) => getComputedStyle(el).animationName),
      'none',
    );
    // Remove library: the confirmation renders below the library's ⋯ button, takes focus, and
    // Escape returns focus to the ⋯ button.
    const libraryMenu = page.locator('#ad-library-menu');
    await libraryMenu.click();
    await page.getByRole('menuitem', { name: 'Remove library' }).click();
    const box = await page.locator('#ad-library-detail .confirm').boundingBox();
    const button = await libraryMenu.boundingBox();
    assert(box.y > button.y + button.height, 'The confirmation is below the ⋯ button');
    assert.match(
      await page.evaluate(
        () => document.activeElement.className + ' ' + document.activeElement.textContent,
      ),
      /^confirm-text Remove General from Documents\?/,
    );
    await page.screenshot({ path: path.join(evidence, 'remove-library-confirm.png') });
    await page.keyboard.press('Escape');
    assert.equal(await page.evaluate(() => document.activeElement.id), 'ad-library-menu');
    assert.equal(await page.locator('#ad-library-detail .confirm').count(), 0);
    assert.equal((await requests('RemoveLibrary')).length, 0);
    // Add site: type in the combobox, pick with the keyboard, then add and check the site.
    await page.locator('#ad-add-site').click();
    const combo = page.getByRole('combobox', { name: 'SharePoint site' });
    await combo.fill('Deliv');
    await page.getByRole('option', { name: 'Delivery' }).waitFor();
    await combo.press('ArrowDown');
    assert.equal(await combo.getAttribute('aria-activedescendant'), 'ad-native-0');
    await combo.press('Enter');
    assert.equal(await combo.inputValue(), 'Delivery');
    await page.locator('#ad-validate').click();
    const siteProgress = page.getByRole('progressbar', { name: 'Delivery setup progress' });
    await siteProgress.waitFor();
    assert.equal(await siteProgress.getAttribute('max'), '3');
    await page.evaluate(() => (window.__mock.siteStatus = 'Captured'));
    await page.getByText('Finishing setup', { exact: true }).waitFor({ timeout: 12000 });
    assert.equal(await siteProgress.getAttribute('value'), '2');
    await page.evaluate(() => (window.__mock.siteStatus = 'Approved'));
    await page.getByText('Delivery is ready.', { exact: true }).waitFor({ timeout: 12000 });
    assert.equal(await siteProgress.getAttribute('value'), '3');
    assert.deepEqual(page.errors, []);
    console.log(
      'PASS Sites & access in headless Edge on the real page: a deleted Dataverse team shown and removed by Apply, team staging/apply/poll, existing-library discovery/add, library creation, the Remove library confirmation under its ⋯ button with focus and Escape, Add site by keyboard; 1440/800/400 light and dark; no page errors or horizontal overflow. Mocked Dataverse (mock-xrm.js).',
    );
  } finally {
    await browser.close();
  }
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
