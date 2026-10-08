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
// At 1920px a page's content is --content-max (1280px) wide and centered in its area, the space
// beside any rail. The contents of its header and footer bars line up with it. A box is the
// element's content box, or its border box for an element placed by its margins.
const CONTENT_MAX = 1280;
async function assertCentered(page, label, content, area, bars = []) {
  const m = await page.evaluate(
    ([content, area, bars]) => {
      const box = ([sel, kind = 'content']) => {
        const e = document.querySelector(sel);
        const r = e.getBoundingClientRect();
        const s = getComputedStyle(e);
        const inset = (side) =>
          kind === 'border'
            ? 0
            : parseFloat(s['border' + side + 'Width']) + parseFloat(s['padding' + side]);
        return { left: r.left + inset('Left'), right: r.right - inset('Right') };
      };
      return { content: box(content), area: box([area, 'border']), bars: bars.map(box) };
    },
    [content, area, bars],
  );
  const width = m.content.right - m.content.left;
  assert(width >= CONTENT_MAX - 0.5 && width <= CONTENT_MAX + 0.5, label + ' width ' + width);
  const before = m.content.left - m.area.left,
    after = m.area.right - m.content.right;
  assert(Math.abs(before - after) <= 1, label + ' centered ' + before + '/' + after);
  m.bars.forEach((bar, i) =>
    assert(
      Math.abs(bar.left - m.content.left) <= 1 && Math.abs(bar.right - m.content.right) <= 1,
      label + ' ' + bars[i][0] + ' lines up ' + JSON.stringify([bar, m.content]),
    ),
  );
}
// axe-core (WCAG 2.2 AA), page errors and the computed checks for the page as it is shown.
async function check(page, label) {
  // A panel that is still sliding in is part transparent; contrast is read once it has landed.
  await page.evaluate(() =>
    Promise.all(
      document
        .getAnimations()
        .filter((a) => a.effect?.getComputedTiming().iterations !== Infinity)
        .map((a) => a.finished),
    ),
  );
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
        // Folder templates: the overview, with its Status panel open, then the editor.
        if (tab === 'templates') {
          await page.locator('#overview-title', { hasText: 'Account onboarding' }).waitFor();
          await settle(page);
          await check(page, 'templates-overview ' + scheme);
          // A hovered primary button keeps its text readable (at least 4.5:1 against its fill),
          // and a disabled one keeps its disabled fill on hover.
          await page.locator('#overview-edit').hover();
          const hovered = await page.evaluate(() => {
            const lum = (c) => {
              const v = (c.match(/\d+(\.\d+)?/g) || [])
                .slice(0, 3)
                .map((n) => n / 255)
                .map((n) => (n <= 0.03928 ? n / 12.92 : ((n + 0.055) / 1.055) ** 2.4));
              return 0.2126 * v[0] + 0.7152 * v[1] + 0.0722 * v[2];
            };
            const edit = document.getElementById('overview-edit');
            const style = getComputedStyle(edit);
            const [a, b] = [lum(style.color), lum(style.backgroundColor)];
            const result = {
              hover: edit.matches(':hover'),
              color: style.color,
              background: style.backgroundColor,
              ratio: (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05),
            };
            edit.setAttribute('aria-disabled', 'true');
            const probe = document.createElement('div');
            probe.style.background = 'var(--disabled-bg)';
            document.body.append(probe);
            result.disabled = getComputedStyle(edit).backgroundColor;
            result.disabledBg = getComputedStyle(probe).backgroundColor;
            probe.remove();
            edit.removeAttribute('aria-disabled');
            return result;
          });
          const hoverLabel = 'hovered primary ' + scheme + ' ' + JSON.stringify(hovered);
          assert(hovered.hover, hoverLabel + ': not hovered');
          assert(hovered.ratio >= 4.5, hoverLabel + ': contrast');
          assert.equal(hovered.disabled, hovered.disabledBg, hoverLabel + ': disabled look');
          await page.mouse.move(0, 0);
          await page.getByRole('button', { name: 'Edit status' }).click();
          await page.getByRole('dialog', { name: 'Status' }).waitFor();
          await page.getByRole('switch', { name: 'Template is on' }).waitFor();
          await check(page, 'templates-status ' + scheme);
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
        // Monitor: with the Check a record panel open from Tools.
        if (tab === 'monitor') {
          await page.locator('#monitor-tools').click();
          await page.getByRole('menuitem', { name: 'Check a record' }).click();
          await page.getByRole('dialog', { name: 'Check a record' }).waitFor();
          await settle(page);
        }
        // Settings: with a change, so the save bar shows.
        if (tab === 'settings') {
          await page.locator('#hosts-list input').first().fill('fabrikam.sharepoint.com');
          await page.locator('#settings-footer').waitFor();
          await settle(page);
        }
        await check(page, tab + ' ' + scheme);
        // Folder templates step 2: a folder with a condition, a test record and the ＋ Field
        // popover open.
        if (tab === 'templates') {
          await page.locator('#step-tab-2').click();
          await page
            .locator('#folder-tree')
            .getByRole('button', { name: 'General', exact: true })
            .click();
          await page.getByRole('radio', { name: 'Only when…' }).click();
          await page.locator('#add-test-record').click();
          await page.locator('#test-records .test-record .state', { hasText: 'Created' }).waitFor();
          await page.locator('#add-field').click();
          await page.locator('#field-popover').waitFor();
          await settle(page);
          await check(page, 'templates-folders ' + scheme);
          // Step 3: a new folder in the change list, the test record's result, the Publishing
          // card with its re-run box.
          await page.keyboard.press('Escape');
          await page.locator('#add-subfolder').click();
          await page.locator('#folder-name').fill('Projects');
          await page.locator('#step-tab-3').click();
          await page.locator('#change-list .change-row').first().waitFor();
          await page.locator('#consequences li').nth(1).waitFor();
          await page.locator('#previewTrees .preview-card').waitFor();
          await settle(page);
          await check(page, 'templates-review ' + scheme);
        }
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
    await templates
      .locator('#folder-tree')
      .getByRole('button', { name: 'General', exact: true })
      .focus();
    await templates.keyboard.press('Enter');
    assert.match(
      await templates.evaluate(() => document.activeElement.getAttribute('data-focus-key')),
      /^node:/,
    );
    await templates.getByRole('radio', { name: 'Only when…' }).click();
    assert.equal(
      await templates.evaluate(() => document.activeElement.getAttribute('aria-label')),
      'Field, condition 1',
    );
    // ＋ Field with the keyboard: Tab from Name, Enter, type, ArrowDown, Enter inserts the field
    // at the caret and returns to Name.
    await templates.locator('#folder-name').focus();
    await templates.keyboard.press('End');
    await templates.keyboard.press('Tab');
    assert.equal(await active(), 'add-field');
    await templates.keyboard.press('Enter');
    assert.equal(await active(), 'field-search');
    await templates.keyboard.type('number');
    await templates.keyboard.press('ArrowDown');
    await templates.keyboard.press('Enter');
    assert.equal(await active(), 'folder-name');
    assert.equal(await templates.locator('#field-popover').isHidden(), true);
    assert.equal(
      await templates.locator('#folder-name').inputValue(),
      'General{root.accountnumber}',
    );
    assert.equal(
      await templates.locator('#folder-shows-as .token').textContent(),
      'Account Number',
    );
    // Escape closes only the popover.
    await templates.keyboard.press('Tab');
    await templates.keyboard.press('Enter');
    await templates.keyboard.press('Escape');
    assert.equal(await active(), 'folder-name');
    assert.equal(await templates.locator('#field-popover').isHidden(), true);
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
    // The access drawer: Escape in its confirmation keeps it; Escape again closes it and focus
    // returns to the library's row.
    const access = await open(context, 'access');
    await access
      .locator('#ad-libraries')
      .getByRole('button', { name: 'General', exact: true })
      .click();
    await access.getByRole('dialog', { name: 'General' }).waitFor();
    await access.locator('#ad-library-menu').click();
    await access.getByRole('menuitem', { name: 'Remove library' }).click();
    await access.keyboard.press('Escape');
    assert.equal(await access.locator('#ad-drawer').isVisible(), true, 'Escape kept the drawer');
    await access.keyboard.press('Escape');
    assert.equal(await access.locator('#ad-drawer').isHidden(), true);
    assert.match(await access.evaluate(() => document.activeElement.dataset.focusKey), /^library:/);
    // Appearance: with the browser in light, choosing Dark from a page header's ◐ menu turns the
    // page dark at once (the body takes the dark canvas token), and the page passes axe in that
    // state. The choice is kept for the next page; Match browser returns to the browser's light.
    const canvas = /--canvas:\s*light-dark\((#\w+),\s*(#\w+)\)/.exec(
      fs.readFileSync(path.join(root, 'admin.css'), 'utf8'),
    );
    const rgb = (hex) => {
      const n = hex.length === 4 ? [...hex.slice(1)].map((c) => c + c) : hex.slice(1).match(/../g);
      return 'rgb(' + n.map((h) => parseInt(h, 16)).join(', ') + ')';
    };
    const [lightCanvas, darkCanvas] = [rgb(canvas[1]), rgb(canvas[2])];
    const lit = await browser.newContext({
      viewport: { width: 1440, height: 1000 },
      colorScheme: 'light',
    });
    const background = (p) => p.evaluate(() => getComputedStyle(document.body).backgroundColor);
    for (const tab of ['templates', 'access', 'monitor', 'settings']) {
      const page = await open(lit, tab);
      if (tab === 'templates')
        await page.locator('#overview-title', { hasText: 'Account onboarding' }).waitFor();
      const header = page.locator('#' + tab + ' .page-header');
      const control = header.getByRole('button', { name: 'Appearance' });
      if (tab === 'templates') {
        assert.equal(await background(page), lightCanvas, 'The browser is in light');
        await control.focus();
        await page.keyboard.press('ArrowDown');
        assert.equal(
          await page.evaluate(() => document.activeElement.getAttribute('aria-checked')),
          'true',
          'The menu opens on the checked item',
        );
        await header.getByRole('menuitemradio', { name: 'Dark' }).click();
        assert.equal(await control.getAttribute('aria-expanded'), 'false');
      }
      assert.equal(await background(page), darkCanvas, tab + ': Dark applies');
      assert.equal(
        await header
          .getByRole('menuitemradio', { name: 'Dark', includeHidden: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await settle(page);
      await check(page, tab + ' chosen-dark');
      if (tab === 'settings') {
        await control.click();
        await header.getByRole('menuitemradio', { name: 'Match browser' }).click();
        assert.equal(await background(page), lightCanvas, 'Match browser follows the browser');
      }
      await page.close();
    }
    await lit.close();
    // Below 1000px the page area stacks: the templates list is a select under its heading and
    // ＋ New, step 2's panel sits under the tree, and the access drawer takes the full width.
    const narrow = await browser.newContext({ viewport: { width: 900, height: 800 } });
    const small = await open(narrow, 'templates');
    await small.locator('#overview-title', { hasText: 'Account onboarding' }).waitFor();
    assert.equal(await small.locator('#template-picker').isVisible(), true);
    assert.equal(await small.locator('#new-template').isVisible(), true);
    assert.equal(await small.locator('#template-groups').isHidden(), true);
    assert.equal(await small.locator('#template-search').isHidden(), true);
    await openTemplate(small);
    await small.locator('#step-tab-2').click();
    await small
      .locator('#folder-tree')
      .getByRole('button', { name: 'General', exact: true })
      .click();
    const tree = await small.locator('#folder-tree').boundingBox();
    const side = await small.locator('#folder-panel').boundingBox();
    assert(side.y >= tree.y + tree.height, 'The folder panel is under the tree');
    const sites = await open(narrow, 'access');
    await sites
      .locator('#ad-libraries')
      .getByRole('button', { name: 'General', exact: true })
      .click();
    await sites.getByRole('dialog', { name: 'General' }).waitFor();
    await sites.waitForTimeout(300);
    assert.equal((await sites.locator('#ad-drawer').boundingBox()).width, 900);
    // The page under the full-width drawer is inert: Tab stays in the drawer.
    let inside = 0;
    for (let i = 0; i < 30; i++) {
      await sites.keyboard.press('Tab');
      const where = await sites.evaluate(() =>
        document.activeElement === document.body
          ? 'body'
          : document.getElementById('ad-drawer').contains(document.activeElement)
            ? 'drawer'
            : document.activeElement.id || document.activeElement.tagName,
      );
      assert.ok(
        where === 'drawer' || where === 'body',
        'Tab reached ' + where + ' under the drawer',
      );
      if (where === 'drawer') inside++;
    }
    assert.ok(inside > 0, 'Tab reaches the drawer');
    // An install with no templates yet: ＋ New starts the first one.
    const empty = await open(narrow, 'templates', 'window.__mock.empty = true');
    await empty.locator('#new-template').click();
    await empty.locator('#template-editor').waitFor();
    assert.equal(await empty.locator('#templateName').isVisible(), true);
    // The header's meta line wraps under the title instead of hiding.
    const narrowSettings = await open(narrow, 'settings');
    await narrowSettings.locator('#settings-title').waitFor();
    assert.equal(await narrowSettings.locator('#settings-meta').isVisible(), true);
    const heading = await narrowSettings.locator('#settings-title').boundingBox();
    const line = await narrowSettings.locator('#settings-meta').boundingBox();
    assert.ok(line.y >= heading.y + heading.height - 1, 'The meta line is under the title');
    await narrow.close();
    // Every page at 1920, 1440, 1000, 800 and 400 in light and dark: nothing scrolls sideways,
    // and a Monitor row's ⋯ menu opens without being cut off. At 1920 each page's content is
    // centered at its maximum width.
    for (const scheme of ['light', 'dark'])
      for (const width of [1920, 1440, 1000, 800, 400]) {
        const sized = await browser.newContext({
          viewport: { width, height: 900 },
          colorScheme: scheme,
        });
        for (const tab of ['templates', 'access', 'monitor', 'settings']) {
          const page = await open(sized, tab);
          const wide = () => page.evaluate(() => document.documentElement.scrollWidth > innerWidth);
          assert.equal(await wide(), false, tab + ' overflow ' + width + ' ' + scheme);
          const big = width === 1920;
          // The ◐ control sits at the right end of the header, at the content's right edge.
          if (tab !== 'templates') await page.locator('#' + tab + ' .page-header').waitFor();
          else await page.locator('#overview-title', { hasText: 'Account onboarding' }).waitFor();
          const edge = await page.evaluate((tab) => {
            const header = document.querySelector('#' + tab + ' .page-header');
            const control = header.querySelector('[aria-label=Appearance]');
            const h = header.getBoundingClientRect();
            const c = control.getBoundingClientRect();
            const right = h.right - parseFloat(getComputedStyle(header).paddingRight);
            return {
              right,
              control: c.right,
              top: c.top - h.top,
              width: c.width,
              inner: innerWidth,
            };
          }, tab);
          const edgeLabel = tab + ' appearance ' + width + ' ' + JSON.stringify(edge);
          assert(edge.width >= 24, edgeLabel + ': shown');
          assert(Math.abs(edge.control - edge.right) <= 1, edgeLabel + ': right edge');
          assert(edge.control <= edge.inner, edgeLabel + ': in view');
          const editorBars = [['.editor-header'], ['#editor-footer']];
          if (tab === 'templates') {
            if (big)
              await assertCentered(page, 'Overview', ['#overview-cards'], '.templates-main', [
                ['#overview-header'],
              ]);
            await openTemplate(page);
            if (big)
              await assertCentered(page, 'Step 1', ['#step-1'], '.templates-main', editorBars);
            for (const n of [2, 3]) {
              if (big)
                await page.screenshot({
                  path: path.join(shots, `responsive-templates-step-${n - 1}-1920-${scheme}.png`),
                  fullPage: true,
                });
              await page.locator('#step-tab-' + n).click();
              await settle(page);
              assert.equal(await wide(), false, 'step ' + n + ' overflow ' + width + ' ' + scheme);
            }
            if (big) {
              await assertCentered(page, 'Step 3', ['#step-3'], '.templates-main', editorBars);
              await page.locator('#step-tab-2').click();
              await settle(page);
              // The tree's content and the folder panel's border span the same 1280px as the
              // other steps: the tree's text starts at the header's, and the panel ends where
              // the footer's Next button does.
              const edges = await page.evaluate(() => {
                const start = (sel) => {
                  const e = document.querySelector(sel);
                  return (
                    e.getBoundingClientRect().left + parseFloat(getComputedStyle(e).paddingLeft)
                  );
                };
                const right = (sel) => document.querySelector(sel).getBoundingClientRect().right;
                const area = document.querySelector('.templates-main').getBoundingClientRect();
                return {
                  tree: start('.tree-side'),
                  header: start('.editor-header'),
                  panel: right('.folder-panel'),
                  next: right('#step-next'),
                  area: { left: area.left, right: area.right },
                };
              });
              const label = 'Step 2 ' + JSON.stringify(edges);
              assert(Math.abs(edges.tree - edges.header) <= 1, label + ' tree lines up');
              assert(Math.abs(edges.panel - edges.next) <= 1, label + ' panel lines up');
              assert(Math.abs(edges.panel - edges.tree - CONTENT_MAX) <= 1, label + ' width');
              assert(
                Math.abs(edges.tree - edges.area.left - (edges.area.right - edges.panel)) <= 1,
                label + ' centered',
              );
              await page.locator('#step-tab-3').click();
              await settle(page);
            }
          }
          if (big && tab === 'access')
            await assertCentered(page, 'Sites & access', ['.ad-main'], '.ad-main', [
              ['#ad-site-header'],
            ]);
          if (big && tab === 'monitor')
            await assertCentered(page, 'Monitor', ['#problem-table-card', 'border'], '#monitor', [
              ['#monitor > .page-header'],
            ]);
          // Look up an operation holds the recent operations table: from 1000px up its panel is
          // at least 700px wide and every column at least 80px, below it the panel takes the
          // full width; nothing in it scrolls sideways and no column is narrower than its header.
          if (tab === 'monitor') {
            await page.locator('#monitor-tools').click();
            await page.getByRole('menuitem', { name: 'Look up an operation' }).click();
            await page.locator('#recent-rows tr').first().waitFor();
            await settle(page);
            const lookup = await page.evaluate(() => {
              const panel = document.getElementById('lookup-panel');
              const range = document.createRange();
              const content = (e) => {
                const style = getComputedStyle(e);
                return (
                  e.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight)
                );
              };
              return {
                width: panel.getBoundingClientRect().width,
                overflow: panel.scrollWidth - panel.clientWidth,
                columns: [...panel.querySelectorAll('thead th')].map((th) => {
                  range.selectNodeContents(th);
                  return {
                    name: th.textContent.trim(),
                    width: th.getBoundingClientRect().width,
                    content: content(th),
                    text: range.getBoundingClientRect().width,
                  };
                }),
                spilled: [...panel.querySelectorAll('#recent-rows td')].filter(
                  (td) => td.scrollWidth > td.clientWidth,
                ).length,
              };
            });
            const label =
              'Look up an operation ' + width + ' ' + scheme + ' ' + JSON.stringify(lookup);
            if (width >= 1000) assert(lookup.width >= 700, label + ': panel width');
            else assert(Math.abs(lookup.width - width) <= 1, label + ': full width');
            assert(lookup.overflow <= 0, label + ': sideways scroll');
            assert.equal(lookup.spilled, 0, label + ': a cell spills');
            for (const column of lookup.columns) {
              assert(column.content + 0.5 >= column.text, label + ': ' + column.name + ' header');
              if (width >= 1000) assert(column.width >= 80, label + ': ' + column.name + ' width');
            }
            if (width === 1440)
              await page.screenshot({
                path: path.join(shots, `responsive-lookup-1440-${scheme}.png`),
              });
            await page.locator('#lookup-close').click();
          }
          if (big && tab === 'settings') {
            await assertCentered(page, 'Settings', ['#settings .page-body'], '#settings', [
              ['#settings > .page-header'],
              ['#settings .section-card', 'border'],
            ]);
            // The save bar's buttons line up with the cards.
            await page.locator('#hosts-list input').first().fill('fabrikam.sharepoint.com');
            await page.locator('#settings-footer').waitFor();
            await settle(page);
            await assertCentered(page, 'Settings', ['#settings .page-body'], '#settings', [
              ['#settings-footer'],
            ]);
            assert.equal(await wide(), false, 'settings save bar overflow ' + scheme);
          }
          if (tab === 'monitor') {
            // The rows laid out as cards still read as a table.
            if (width === 400) await check(page, 'monitor-400 ' + scheme);
            await page
              .getByRole('button', { name: /^More actions for/ })
              .last()
              .click();
            const menu = page.locator('#problem-rows .menu:not([hidden])');
            const box = await menu.boundingBox();
            const shown = await menu.evaluate((m) => {
              const r = m.getBoundingClientRect();
              const hit = document.elementFromPoint(r.left + r.width / 2, r.bottom - 4);
              return m.contains(hit);
            });
            assert(shown && box.x >= 0 && box.x + box.width <= width, 'Menu shown ' + width);
            assert.equal(await wide(), false, 'monitor menu overflow ' + width + ' ' + scheme);
          }
          await page.screenshot({
            path: path.join(shots, 'responsive-' + tab + '-' + width + '-' + scheme + '.png'),
            fullPage: true,
          });
          assert.deepEqual(page.errors, [], tab + ' ' + width + ' page errors');
          await page.close();
        }
        await sized.close();
      }
    for (const state of ['Checking', 'Found', 'NotFound', 'Ambiguous']) {
      const shown = await open(context, 'monitor', 'window.__recovery = ' + JSON.stringify(state));
      await shown.screenshot({
        path: path.join(shots, 'monitor-recovery-' + state.toLowerCase() + '.png'),
        fullPage: true,
      });
    }
    console.log(
      'PASS accessibility in Edge: axe WCAG 2.2 AA on four pages in light and dark (Folder templates as the overview, its Status panel, the editor, step 2 with a condition, a test record and the ＋ Field popover, and step 3 Review and publish; Sites & access with its access drawer open; Monitor with Check a record open; Settings with its save bar), computed borders, text and targets, a hovered primary button readable and a disabled one unchanged, keyboard flows (the overview ⋯ menu, All versions, Edit template and Close, the steps, the step 2 tree, Only when… and ＋ Field, Monitor chips, a row menu and its confirmation, the Tools panel, Escape in the access drawer confirmation and then the drawer), the Appearance menu choosing Dark while the browser is light on every page (dark canvas, axe clean) and Match browser going back, below 1000px (the templates select, the folder panel under the tree, the full-width drawer with Tab kept inside it, the header meta line under the title), every page and step at 1920/1440/1000/800/400 in light and dark without sideways scrolling and with a whole Monitor row menu and the ◐ control at the right edge of the header, Look up an operation at least 700px wide from 1000px with no column squeezed or narrower than its header, the content of each page centered at 1280px at 1920 with its header and footer lined up, screenshots. Mocked Dataverse.',
    );
  } finally {
    await browser.close();
  }
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
