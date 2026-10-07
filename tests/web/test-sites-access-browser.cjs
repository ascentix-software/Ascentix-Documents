const fs = require('fs'),
  path = require('path'),
  assert = require('assert/strict');
const { chromium } = require(process.env.ASXD_PLAYWRIGHT_MODULE || 'playwright');
const root = path.resolve(__dirname, '../../client/admin');
const evidence =
  process.env.ASXD_BROWSER_EVIDENCE_DIR || path.resolve(__dirname, '../../artifacts/browser');
fs.mkdirSync(evidence, { recursive: true });
(async () => {
  const browser = await chromium.launch({ headless: true, channel: 'msedge' });
  try {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.route('**/*', (route) => route.abort());
    await page.setContent(
      fs
        .readFileSync(path.join(root, 'index.html'), 'utf8')
        .replace(/<script[^>]*src=[^>]*><\/script>/g, '')
        .replace(/<link[^>]+>/g, ''),
    );
    await page.addStyleTag({ content: fs.readFileSync(path.join(root, 'admin.css'), 'utf8') });
    await page.evaluate(() => {
      const id = (n) => String(n).padStart(8, '0') + '-0000-0000-0000-000000000000';
      // setContent has no secure origin; provide only the UUID capability normally supplied by HTTPS Dataverse.
      crypto.randomUUID = () => id(Math.floor(Math.random() * 10000000));
      const site = {
        asx_siteid: id(1),
        asx_name: 'Delivery',
        asx_approved: true,
        asx_url: 'https://example.sharepoint.com/sites/delivery',
        _asx_nativeid_value: id(2),
      };
      const lib = {
        asx_libraryid: id(3),
        asx_name: 'General',
        _asx_siteid_value: id(1),
        asx_approved: true,
        asx_policyapplied: true,
      };
      window.testRequests = [];
      // A team deleted in Dataverse still has access to the library; Apply removes it.
      const deleted = { TeamId: id(40), Access: 'Read' };
      window.testPolicy = {
        Status: 'Applied',
        RowVersion: '1',
        Policy: { Desired: [deleted], Applied: [deleted] },
        Teams: [{ TeamId: id(40), Name: 'AcceptC Team 1', Deleted: true }],
      };
      window.Xrm = {
        Utility: { getGlobalContext: () => ({ getClientUrl: () => 'https://example.test' }) },
        Navigation: { openUrl: () => {} },
        WebApi: {
          retrieveMultipleRecords: async (table) => ({
            entities:
              table === 'asx_site'
                ? [site]
                : table === 'asx_library'
                  ? [lib]
                  : table === 'team'
                    ? [{ teamid: id(4), name: 'Operations' }]
                    : table === 'sharepointsite'
                      ? [{ sharepointsiteid: id(2), name: 'Delivery' }]
                      : [],
          }),
          retrieveRecord: async (table) =>
            table === 'asx_site' ? site : table === 'asx_library' ? lib : { name: 'Operations' },
          online: {
            execute: async (req) => {
              const c = JSON.parse(req.Request);
              window.testRequests.push(c);
              let result = {
                Status: 'Pending',
                Key:
                  c.Command === 'CreateLibrary'
                    ? 'librarycreate:test'
                    : c.Command === 'AddSite'
                      ? 'siteprobe:test'
                      : 'catalogprobe:test',
              };
              if (c.Command === 'GetPolicy') result = window.testPolicy;
              if (c.Command === 'ApplyPolicy')
                // The server leaves deleted teams out.
                result = window.testPolicy = {
                  Status: 'Queued',
                  RowVersion: '2',
                  Policy: {
                    Desired: c.Entries.filter((e) => e.TeamId !== id(40)),
                    Applied: [],
                    OperationKey: 'policywork:test',
                  },
                };
              if (c.Command === 'Inspect' && c.Key === 'siteprobe:test')
                result = { Status: window.testSiteStatus || 'Inspecting' };
              else if (c.Command === 'Inspect' && c.Key === 'librarycreate:test')
                result = { Status: window.testSetupStatus || 'ExternalUnknown' };
              else if (c.Command === 'Inspect')
                result = {
                  Status: 'Discovered',
                  Key: c.Key,
                  RowVersion: '3',
                  Observation: {
                    SiteId: id(1),
                    WebUrl: site.asx_url,
                    Libraries: [{ Id: id(8), Title: 'Archive' }],
                  },
                };
              return { ok: true, json: async () => ({ Result: JSON.stringify(result) }) };
            },
          },
        },
      };
      window.fetch = async () => ({ ok: true, json: async () => ({ value: [] }) });
    });
    await page.addScriptTag({ content: fs.readFileSync(path.join(root, 'admin.js'), 'utf8') });
    await page.addScriptTag({
      content: fs.readFileSync(path.join(root, 'sites-access.js'), 'utf8'),
    });
    await page.locator('[data-view="access"]').click();
    await page.getByRole('heading', { name: 'General', exact: true }).waitFor();
    // The deleted team is shown as one, with what happens next, and Apply is enabled for it.
    const deletedRow = page.locator('#ad-teams tr').first();
    await deletedRow.getByText('Deleted team: AcceptC Team 1').waitFor();
    await deletedRow
      .getByText(
        'This team was deleted in Dataverse. Documents removes its access the next time access is applied.',
        { exact: true },
      )
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
    await page.locator('#ad-add-team').click();
    await page.locator('#ad-team-choice').selectOption('00000004-0000-0000-0000-000000000000');
    await page.locator('#ad-stage-team').click();
    assert.equal(
      await page.evaluate(() => testRequests.filter((x) => x.Command === 'ApplyPolicy').length),
      0,
    );
    await page.locator('#ad-apply').click();
    assert.equal(await page.locator('#ad-apply').isDisabled(), true);
    // Applied: the deleted team's row is gone and nothing failed.
    assert.equal(await page.getByText('Deleted team: AcceptC Team 1').count(), 0);
    assert.equal(
      await page.getByText('Status updates are unavailable', { exact: false }).count(),
      0,
    );
    await page.evaluate(() => {
      testPolicy = {
        Status: 'Applied',
        RowVersion: '3',
        Policy: { Desired: testPolicy.Policy.Desired, Applied: testPolicy.Policy.Desired },
      };
    });
    await page.getByText('Access and team membership confirmed.', { exact: true }).waitFor();
    await page.locator('#ad-existing').click();
    try {
      await page
        .getByRole('button', { name: 'Add Archive', exact: true })
        .waitFor({ timeout: 12000 });
    } catch (e) {
      console.log(
        await page.evaluate(() => ({
          requests: testRequests,
          message: document.getElementById('ad-message').textContent,
          progress: document.getElementById('ad-provision-progress').textContent,
          form: document.getElementById('ad-existing-form').outerHTML,
        })),
      );
      throw e;
    }
    await page.getByRole('button', { name: 'Add Archive', exact: true }).click();
    assert.equal(
      await page.evaluate(
        () => testRequests.find((x) => x.Command === 'AddLibrary').NativeParentId,
      ),
      undefined,
    );
    await page.locator('#ad-create').click();
    await page.locator('#ad-library-name').fill('Projects');
    await page.locator('#ad-initial-team').selectOption('00000004-0000-0000-0000-000000000000');
    await page.locator('#ad-provision').click();
    assert.equal(
      await page.evaluate(
        () => testRequests.find((x) => x.Command === 'CreateLibrary').Entries.length,
      ),
      1,
    );
    await page.getByText('Waiting for confirmation', { exact: true }).waitFor({ timeout: 12000 });
    assert.equal(
      await page
        .getByText(
          'The original request must be reconciled in Administration before setup can continue.',
          { exact: true },
        )
        .count(),
      0,
    );
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
        for (const a of alignment) {
          assert(
            a.start < 1 && a.end < 1 && a.fill < 1,
            'Stage/bar alignment ' + JSON.stringify(a),
          );
        }
        await page.screenshot({
          path: path.join(evidence, `runtime-${width}-${scheme}.png`),
          fullPage: true,
        });
      }
    assert.notEqual(
      await page
        .locator('.is-loading .is-current .ad-step-mark')
        .evaluate((el) => getComputedStyle(el).animationName),
      'none',
    );
    await page.emulateMedia({ reducedMotion: 'reduce' });
    assert.equal(
      await page
        .locator('.is-loading .is-current .ad-step-mark')
        .evaluate((el) => getComputedStyle(el).animationName),
      'none',
    );
    await page.locator('#ad-add-site').click();
    await page.locator('#ad-native').selectOption('00000002-0000-0000-0000-000000000000');
    await page.locator('#ad-validate').click();
    const siteProgress = page.getByRole('progressbar', { name: 'Delivery setup progress' });
    await siteProgress.waitFor();
    assert.equal(await siteProgress.getAttribute('max'), '3');
    await page.evaluate(() => {
      testSiteStatus = 'Captured';
    });
    await page.getByText('Finishing setup', { exact: true }).waitFor({ timeout: 12000 });
    assert.equal(await siteProgress.getAttribute('value'), '2');
    await page.evaluate(() => {
      testSiteStatus = 'Approved';
    });
    await page.getByText('Delivery is ready.', { exact: true }).waitFor({ timeout: 12000 });
    assert.equal(await siteProgress.getAttribute('value'), '3');
    assert.equal(await page.locator('#ad-team-search,#ad-initial-team-search').count(), 0);
    assert.deepEqual(errors, []);
    console.log(
      'PASS actual admin HTML, CSS and both scripts in headless Edge: a deleted Dataverse team shown and removed by Apply, team staging/apply/poll, existing-library discovery/add, library creation; 1440/800/400 light and dark; no page errors or horizontal overflow. All Dataverse APIs mocked; no live calls.',
    );
  } finally {
    await browser.close();
  }
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
