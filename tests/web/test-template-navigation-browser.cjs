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
                  ? [{ asx_revisionid: 'revision-1', asx_version: 1, asx_status: 'Draft' }]
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
    await page.locator('#templateTree summary').click();
    await page
      .locator('#templateTree')
      .getByRole('button', { name: 'Contract documents', exact: true })
      .click();
    await page.locator('#version-chip').filter({ hasText: 'Draft v1' }).waitFor();
    assert.equal(await page.locator('#templateName').inputValue(), 'Contract documents');
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
    await page.getByRole('button', { name: '＋ New template' }).first().click();
    await page.locator('#version-chip').filter({ hasText: 'Draft v1' }).waitFor();
    assert.equal(await page.locator('#templateName').isDisabled(), false);
    assert.match(await page.locator('#destinations').textContent(), /No folders yet/);
    assert.deepEqual(errors, []);
    console.log(
      'PASS template navigation in Edge: two templates per table, selection, new template, light/dark and 1440/1000/800/400 layouts. APIs mocked.',
    );
  } finally {
    await browser.close();
  }
})().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
