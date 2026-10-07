const fs = require('fs'),
  path = require('path'),
  assert = require('assert/strict');
const { chromium } = require(process.env.ASXD_PLAYWRIGHT_MODULE || 'playwright');
(async () => {
  const browser = await chromium.launch({ headless: true, channel: 'msedge' });
  try {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } }),
      errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    const root = path.resolve(__dirname, '../../client/admin');
    await page.setContent(
      fs
        .readFileSync(path.join(root, 'index.html'), 'utf8')
        .replace(/<script[^>]*src=[^>]*><\/script>/g, '')
        .replace(/<link[^>]+>/g, ''),
    );
    await page.addStyleTag({ content: fs.readFileSync(path.join(root, 'admin.css'), 'utf8') });
    await page.evaluate(() => {
      const templates = [
        { asx_templateid: 'template-1', asx_table: 'account', asx_name: 'Account onboarding' },
        { asx_templateid: 'template-2', asx_table: 'account', asx_name: 'Contract documents' },
      ];
      const meta = {
        LogicalName: 'account',
        DisplayName: { UserLocalizedLabel: { Label: 'Account' } },
        EntitySetName: 'accounts',
        PrimaryNameAttribute: 'name',
        PrimaryIdAttribute: 'accountid',
      };
      window.Xrm = {
        Utility: { getGlobalContext: () => ({ getClientUrl: () => 'https://example.test' }) },
        WebApi: {
          retrieveRecord: async (t, id) => templates.find((v) => v.asx_templateid === id),
          retrieveMultipleRecords: async (t) => ({
            entities:
              t === 'asx_template'
                ? templates
                : t === 'asx_revision'
                  ? templates.map((v, i) => ({
                      asx_revisionid: 'revision-' + (i + 1),
                      asx_version: 1,
                      asx_status: 'Draft',
                      _asx_templateid_value: v.asx_templateid,
                      modifiedon: '2026-10-03T09:00:00Z',
                      '_modifiedby_value@OData.Community.Display.V1.FormattedValue': 'Dana Reyes',
                    }))
                  : t === 'asx_library'
                    ? [
                        {
                          asx_libraryid: 'library-1',
                          asx_name: 'General',
                          _asx_siteid_value: 'site-1',
                        },
                      ]
                    : t === 'asx_site'
                      ? [{ asx_siteid: 'site-1', asx_name: 'Delivery' }]
                      : t === 'asx_runtimetable'
                        ? [{ asx_logicalname: 'account' }]
                        : [],
          }),
          online: {
            execute: async () => ({
              ok: true,
              json: async () => ({
                Result: JSON.stringify({
                  RevisionId: 'revision-1',
                  RowVersion: '1',
                  Status: 'Draft',
                  Version: 1,
                  Draft: {
                    Table: 'account',
                    Sources: [
                      {
                        Alias: 'root',
                        Table: 'account',
                        Lookup: null,
                        Columns: [{ Name: 'name', Kind: 'Text' }],
                      },
                    ],
                    Destinations: [
                      {
                        Key: 'general',
                        Name: 'Business documents',
                        LibraryId: 'library-1',
                        Folders: [
                          { Key: 'root', Parent: null, Name: '{root.name}', Condition: null },
                        ],
                      },
                    ],
                  },
                }),
              }),
            }),
          },
        },
      };
      window.fetch = async (url) => ({
        ok: true,
        json: async () => ({
          value: url.includes('LookupAttributeMetadata')
            ? []
            : url.includes('/Attributes?')
              ? [{ LogicalName: 'name', AttributeType: 'String' }]
              : url.includes('/accounts?')
                ? [{ accountid: 'record-1', name: 'Example account' }]
                : [meta],
        }),
      });
    });
    for (const name of ['shell.js', 'admin.js'])
      await page.addScriptTag({ content: fs.readFileSync(path.join(root, name), 'utf8') });
    await page.evaluate(() => {
      try {
        sessionStorage.setItem('asxd.launched', '1');
      } catch {
        // about:blank may refuse storage; the shell then treats the load as already launched.
      }
      document.dispatchEvent(new Event('DOMContentLoaded'));
    });
    // The page opens on the first template's overview; a row in the list opens another.
    await page.locator('#overview-title', { hasText: 'Account onboarding' }).waitFor();
    await page
      .locator('#template-groups .list-row')
      .filter({ hasText: 'Contract documents' })
      .click();
    await page.locator('#overview-title', { hasText: 'Contract documents' }).waitFor();
    assert.equal(await page.locator('#overview-pill').textContent(), 'Draft v1');
    for (const width of [1440, 1000, 800, 400])
      for (const scheme of ['light', 'dark']) {
        await page.setViewportSize({ width, height: 1000 });
        await page.emulateMedia({ colorScheme: scheme });
        assert.equal(
          await page.evaluate(() => document.documentElement.scrollWidth > innerWidth),
          false,
          'Overview overflow ' + width,
        );
        await page.screenshot({
          path: path.resolve(
            __dirname,
            `../artifacts/browser/templates-overview-${width}-${scheme}.png`,
          ),
          fullPage: true,
        });
      }
    await page.setViewportSize({ width: 1440, height: 1000 });
    await page.getByRole('button', { name: 'Continue Draft v1' }).click();
    await page.locator('#editor-pill').filter({ hasText: 'Draft v1' }).waitFor();
    assert.equal(await page.locator('#editor-name').textContent(), 'Contract documents');
    assert.equal(await page.locator('#templateName').isHidden(), true);
    // The steps are tabs: arrow keys move focus, Enter shows the step.
    await page.locator('#step-tab-1').focus();
    await page.keyboard.press('ArrowRight');
    assert.equal(await page.evaluate(() => document.activeElement.id), 'step-tab-2');
    assert.equal(await page.locator('#step-tab-2').getAttribute('aria-selected'), 'false');
    await page.keyboard.press('Enter');
    await page.locator('#step-2').waitFor();
    assert.equal(await page.locator('#step-1').isHidden(), true);
    assert.equal(await page.locator('#step-next').textContent(), 'Next: Review');
    await page.locator('#step-back').click();
    await page.locator('#step-1 .destination-card').waitFor();
    for (const width of [1440, 1000, 800, 400])
      for (const scheme of ['light', 'dark']) {
        await page.setViewportSize({ width, height: 1000 });
        await page.emulateMedia({ colorScheme: scheme });
        assert.equal(
          await page.evaluate(() => document.documentElement.scrollWidth > innerWidth),
          false,
          'Overflow ' + width,
        );
        await page.screenshot({
          path: path.resolve(__dirname, `../artifacts/browser/templates-${width}-${scheme}.png`),
          fullPage: true,
        });
      }
    // Adding a destination saves by itself after a pause; Close then returns to the overview
    // without asking.
    await page.getByRole('button', { name: '＋ Add destination' }).click();
    await page.locator('#step-1 .destination-row').waitFor();
    await page.locator('#save-status', { hasText: /^Saved / }).waitFor({ timeout: 5000 });
    await page.locator('#editor-close').click();
    await page.locator('#overview-title', { hasText: 'Contract documents' }).waitFor();
    assert.equal(await page.locator('#leavePrompt').textContent(), '');
    await page.locator('#new-template').click();
    await page.locator('#editor-pill').filter({ hasText: 'Draft v1' }).waitFor();
    assert.equal(await page.locator('#templateName').isDisabled(), false);
    assert.equal(
      await page.locator('#templateName').inputValue(),
      '',
      'A new template starts unnamed',
    );
    assert.match(await page.locator('#destinations').textContent(), /No folders yet/);
    assert.deepEqual(errors, []);
    console.log(
      'PASS template navigation in Edge: the templates list and overview, Continue Draft into the editor, the steps by keyboard, an added destination saved by itself, Close, ＋ New, light/dark and 1440/1000/800/400 layouts of the overview and the editor. APIs mocked.',
    );
  } finally {
    await browser.close();
  }
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
