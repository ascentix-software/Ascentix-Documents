// A mocked Dataverse for the admin page's browser tests. Added with page.addInitScript, so it
// runs before the page's scripts and on every navigation: it defines window.Xrm with canned
// answers for every call the four tabs make, and answers the metadata and privilege URLs the
// page fetches. Rows carry no IDs in visible fields, as the real APIs.
//
// Tests adjust it through window.__mock (set by a later init script): policy, siteStatus,
// setupStatus, recovery and empty (an install with no templates yet); window.__recovery names a library setup's recovery state for
// Monitor. Every request is recorded in window.__mock.requests. window.__mockIdle() is true
// once no call has been in flight for a moment, so a test can wait for a tab to finish loading.
(() => {
  const id = (n) => String(n).padStart(8, '0') + '-0000-4000-8000-000000000000';
  const ago = (minutes) => new Date(Date.now() - minutes * 60000).toISOString();
  const IDS = {
    worker: id(1),
    template: id(2),
    revision: id(3),
    library: id(4),
    site: id(5),
    native: id(6),
    operations: id(7),
    finance: id(8),
    record: id(9),
    connection: id(10),
    candidate: id(11),
    other: id(12),
  };
  const site = {
    asx_siteid: IDS.site,
    asx_name: 'Delivery',
    asx_approved: true,
    asx_url: 'https://contoso.sharepoint.com/sites/delivery',
    _asx_nativeid_value: IDS.native,
    statecode: 0,
  };
  const library = {
    asx_libraryid: IDS.library,
    asx_name: 'General',
    _asx_siteid_value: IDS.site,
    asx_approved: true,
    asx_policyapplied: true,
    asx_entryurl: 'https://contoso.sharepoint.com/sites/delivery/General',
    statecode: 0,
  };
  const template = {
    asx_templateid: IDS.template,
    asx_name: 'Account onboarding',
    asx_table: 'account',
    asx_disabled: false,
    asx_startsutc: null,
    asx_endsutc: null,
    _asx_publishedrevisionid_value: IDS.revision,
  };
  const teams = [
    { teamid: IDS.operations, name: 'Operations', teamtype: 0 },
    {
      teamid: IDS.finance,
      name: 'Finance',
      teamtype: 2,
      membershiptype: 0,
      azureactivedirectoryobjectid: id(20),
    },
  ];
  const runtime = {
    WorkerId: IDS.worker,
    Enabled: true,
    CanChange: true,
    RowVersion: '7',
    SharePointHosts: ['contoso.sharepoint.com'],
    ProcessRecordUpdates: false,
    Registration: {
      Readiness: [
        { Scope: 'account', Status: 'Ready' },
        { Scope: 'team', Status: 'Ready' },
      ],
      Pending: 0,
      Error: null,
    },
  };
  const draft = {
    TemplateId: IDS.template,
    Name: 'Account onboarding',
    Table: 'account',
    Sources: [
      { Alias: 'root', Table: 'account', Lookup: null, Columns: [{ Name: 'name', Kind: 'Text' }] },
    ],
    Destinations: [
      {
        Key: 'general',
        Name: 'Business documents',
        LibraryId: IDS.library,
        Folders: [
          { Key: 'root', Parent: null, Name: '{root.name}', Condition: null },
          { Key: 'general_docs', Parent: 'root', Name: 'General', Condition: null },
        ],
      },
    ],
  };
  const label = (text) => ({ UserLocalizedLabel: { Label: text } });
  const accountInfo = {
    LogicalName: 'account',
    DisplayName: label('Account'),
    DisplayCollectionName: label('Accounts'),
    EntitySetName: 'accounts',
    PrimaryIdAttribute: 'accountid',
    PrimaryNameAttribute: 'name',
    IsDocumentManagementEnabled: true,
  };
  const attributes = [
    { LogicalName: 'name', AttributeType: 'String', DisplayName: label('Account Name') },
    { LogicalName: 'accountnumber', AttributeType: 'String', DisplayName: label('Account Number') },
    { LogicalName: 'revenue', AttributeType: 'Money', DisplayName: label('Annual Revenue') },
    { LogicalName: 'statecode', AttributeType: 'State', DisplayName: label('Status') },
  ];
  const states = [
    {
      LogicalName: 'statecode',
      OptionSet: {
        Options: [
          { Value: 0, Label: label('Active') },
          { Value: 1, Label: label('Inactive') },
        ],
      },
    },
  ];

  // A library setup whose create answer was lost, in the recovery state a test names.
  function recoveryRow(state) {
    const candidate = {
      ListId: IDS.candidate,
      Title: 'Project documents',
      Url: '/sites/delivery/Project documents',
      CreatedUtc: ago(30),
      IsLibrary: true,
      TitleMatches: true,
      UrlMatches: true,
      CreatedAfterRequest: true,
      CatalogEntry: 'None',
    };
    const recovery = {
      Checking: { State: 'Checking', Candidates: [], Choices: ['Cancel'] },
      Found: {
        State: 'Found',
        Candidates: [candidate],
        Choices: ['UseLibrary', 'CheckAgain', 'Cancel'],
      },
      NotFound: {
        State: 'NotFound',
        Candidates: [],
        Choices: ['CreateAgain', 'CheckAgain', 'Cancel'],
      },
      Ambiguous: {
        State: 'Ambiguous',
        Reason: 'SharePoint has two libraries that could be this one.',
        Candidates: [
          candidate,
          {
            ...candidate,
            ListId: IDS.other,
            Title: 'Project documents 2',
            Url: '/sites/delivery/Project documents 2',
            TitleMatches: false,
            UrlMatches: false,
          },
        ],
        Choices: ['UseCandidate', 'CheckAgain', 'Cancel'],
      },
    }[state];
    return {
      Key: 'librarycreate:' + IDS.candidate,
      Kind: 'LibrarySetup',
      KindLabel: 'Library setup',
      Title: 'Project documents',
      Problem:
        state === 'Checking' ? 'Checking SharePoint…' : 'The create may have reached SharePoint.',
      SinceUtc: ago(45),
      Recovery: recovery,
      RowVersion: '12',
      Actions: recovery.Choices,
    };
  }
  // A record row of Monitor's problem table: the record, its table and template.
  const recordRow = (key, extra) => ({
    Key: key,
    Kind: 'RecordPlan',
    KindLabel: 'Record',
    Title: 'Fabrikam Logistics · Account onboarding',
    Record: { Table: 'account', TableLabel: 'Account', Id: IDS.record, Name: 'Fabrikam Logistics' },
    TemplateId: IDS.template,
    TemplateName: 'Account onboarding',
    ...extra,
  });
  function problems(list) {
    if (list === 'TemplateRuns')
      return [
        {
          Key: 'templaterun:' + IDS.template,
          Kind: 'TemplateRun',
          KindLabel: 'Template re-run',
          Title: 'Account onboarding',
          SinceUtc: ago(20),
          Run: {
            TemplateId: IDS.template,
            TemplateName: 'Account onboarding',
            TableLabel: 'Account',
            Version: 1,
            State: 'Running',
            Planned: 500,
            Total: 1240,
            TotalEstimated: false,
            StartedUtc: ago(20),
            StartedBy: 'Alex Admin',
            EstimatedFinishUtc: new Date(Date.now() + 40 * 60000).toISOString(),
          },
          Actions: ['Pause', 'CancelRun'],
        },
      ];
    if (list === 'BlockedRecords')
      return [
        recordRow('request:' + IDS.record, {
          Status: 'Blocked',
          Problem: 'Library General is not ready for folder templates.',
          SinceUtc: ago(70),
          Actions: ['Retry', 'Cancel', 'OpenRecord', 'Check'],
        }),
      ];
    if (list === 'WaitingRecords')
      return [
        recordRow('request:' + IDS.other, {
          Status: 'Waiting',
          Problem: 'Needs a value in Account Number.',
          SinceUtc: ago(60),
          Actions: ['Rerun', 'Check', 'OpenRecord'],
        }),
      ];
    if (list === 'RetryingJobs')
      return [
        {
          Key: 'folderjob:' + IDS.candidate,
          Kind: 'FolderJob',
          KindLabel: 'Folder job',
          Title: 'Folder job · Northwind Traders',
          Problem: 'SharePoint is limiting requests.',
          Status: 'RetryWait',
          Attempt: 3,
          NextAttemptUtc: new Date(Date.now() + 10 * 60000).toISOString(),
          SinceUtc: ago(50),
          Actions: ['Retry', 'Cancel'],
        },
      ];
    if (list === 'BlockedJobs')
      return [
        {
          Key: 'folderjob:' + IDS.record,
          Kind: 'FolderJob',
          KindLabel: 'Folder job',
          Title: 'Contoso Ltd · General',
          Problem: 'SharePoint refused the folder: the app has no access to this site.',
          Fix: 'Give the Documents app access to the site, then Retry.',
          SinceUtc: ago(90),
          RowVersion: '5',
          Actions: ['Retry', 'Cancel'],
        },
        ...(window.__recovery ? [recoveryRow(window.__recovery)] : []),
      ];
    return [];
  }
  const summary = () => ({
    TemplateRuns: 1,
    NotCaptured: 0,
    BlockedRecords: 1,
    WaitingRecords: 1,
    BlockedJobs: window.__recovery ? 2 : 1,
    RetryingJobs: 1,
    Capped: [],
    CountedUtc: new Date().toISOString(),
  });

  const mock = (window.__mock = {
    requests: [],
    policy: {
      Status: 'Applied',
      RowVersion: '1',
      Policy: {
        Desired: [{ TeamId: IDS.operations, Access: 'Read' }],
        Applied: [{ TeamId: IDS.operations, Access: 'Read' }],
      },
      Teams: [{ TeamId: IDS.operations, Name: 'Operations', Deleted: false }],
    },
    siteStatus: null,
    setupStatus: null,
    ids: IDS,
  });

  // In-flight calls, for __mockIdle.
  let pending = 0;
  let last = Date.now();
  const tracked = async (work) => {
    pending++;
    last = Date.now();
    try {
      await new Promise((resolve) => setTimeout(resolve, 5));
      return await work();
    } finally {
      pending--;
      last = Date.now();
    }
  };
  window.__mockIdle = () => pending === 0 && Date.now() - last > 300;
  const answer = (result) => ({ ok: true, json: async () => ({ Result: JSON.stringify(result) }) });

  function rows(table, options = '') {
    switch (table) {
      case 'asx_runtimetable':
        return [{ asx_runtimetableid: id(30), asx_logicalname: 'account' }];
      case 'asx_template':
        return mock.empty ? [] : [template];
      case 'asx_revision':
        return mock.empty
          ? []
          : [
              {
                asx_revisionid: IDS.revision,
                asx_version: 1,
                asx_status: 'Published',
                _asx_templateid_value: IDS.template,
                modifiedon: ago(60 * 24 * 4),
                '_modifiedby_value@OData.Community.Display.V1.FormattedValue': 'Dana Reyes',
              },
            ];
      // The published revision's destination is on General, so one template uses it.
      case 'asx_destination':
        return [{ _asx_libraryid_value: IDS.library, _asx_revisionid_value: IDS.revision }];
      case 'asx_library':
        return /asx_listid eq/.test(options) ? [] : [library];
      case 'asx_site':
        return [site];
      case 'team':
        return teams;
      case 'systemuser':
        return [{ systemuserid: IDS.worker, fullname: 'Documents worker' }];
      case 'sharepointsite':
        return [
          {
            sharepointsiteid: IDS.native,
            name: 'Delivery',
            absoluteurl: 'https://contoso.sharepoint.com/sites/delivery',
            ...validation(),
          },
        ];
      case 'connectionreference':
        return [
          {
            connectionreferencedisplayname: 'Ascentix Documents SharePoint',
            connectionid: 'connected',
            connectorid: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline',
          },
        ];
      default:
        return [];
    }
  }
  // Dynamics' validation status of the SharePoint site: Valid unless a test sets another.
  const statuses = { 1: 'Not Started', 2: 'In Progress', 3: 'Invalid', 4: 'Valid' };
  const validation = () => ({
    validationstatus: mock.validationStatus ?? 4,
    'validationstatus@OData.Community.Display.V1.FormattedValue':
      statuses[mock.validationStatus ?? 4],
  });
  function record(table, key) {
    switch (table) {
      case 'sharepointsite':
        return { sharepointsiteid: key, ...validation() };
      case 'asx_template':
        return template;
      case 'asx_revision':
        // A saved draft is the next version after the published v1.
        return { asx_revisionid: key, asx_version: key === IDS.revision ? 1 : 2 };
      case 'asx_library':
        return library;
      case 'asx_site':
        return site;
      case 'team':
        return teams.find((t) => t.teamid === key) || teams[0];
      default:
        return {};
    }
  }
  function catalog(command) {
    switch (command.Command) {
      case 'Inspect':
        if (command.Key === 'siteprobe:test')
          return { Key: command.Key, Status: mock.siteStatus || 'Inspecting' };
        if (command.Key === 'librarycreate:test')
          return { Key: command.Key, Status: mock.setupStatus || 'ExternalUnknown' };
        if (command.Key === 'libraryprobe:test')
          return { Key: command.Key, Status: 'Approved', CatalogId: IDS.library };
        return {
          Status: 'Discovered',
          Key: command.Key,
          RowVersion: '3',
          Observation: {
            SiteId: IDS.site,
            WebUrl: site.asx_url,
            Libraries: [{ Id: id(40), Title: 'Archive' }],
          },
        };
      case 'AddSite':
        return { Status: 'Pending', Key: 'siteprobe:test' };
      case 'CreateLibrary':
        return { Status: 'Pending', Key: 'librarycreate:test' };
      case 'AddLibrary':
        return { Status: 'Pending', Key: 'libraryprobe:test' };
      case 'RemoveLibrary':
      case 'RemoveSite':
        return { Status: 'Removed', CatalogId: command.CatalogId, Notices: [] };
      default:
        return { Status: 'Pending', Key: 'catalogprobe:test' };
    }
  }
  function securityAdmin(command) {
    if (command.Command === 'ApplyPolicy')
      // The server leaves deleted teams out.
      return (mock.policy = {
        Status: 'Queued',
        RowVersion: '2',
        Policy: {
          Desired: command.Entries.filter(
            (e) => !(mock.policy.Teams || []).some((t) => t.Deleted && t.TeamId === e.TeamId),
          ),
          Applied: [],
          OperationKey: 'policywork:test',
        },
        Teams: mock.policy.Teams,
      });
    return mock.policy;
  }
  function execute(request) {
    const name = request.getMetadata().operationName;
    const command = request.Request ? JSON.parse(request.Request) : { ...request };
    delete command.getMetadata;
    mock.requests.push({ api: name, ...command });
    switch (name) {
      case 'asx_RuntimeAdmin':
        if (command.Command === 'SetEnabled') runtime.Enabled = command.Enabled;
        return answer(runtime);
      case 'asx_ManageWork':
        if (command.Command === 'Summary') return answer({ Status: 'Counted', Summary: summary() });
        if (command.Command === 'ListProblems')
          return answer({ Status: 'Page', Problems: problems(command.List), Next: null });
        if (command.Command === 'CountRecords')
          return answer({ Status: 'Counted', Run: { Total: 1240, TotalEstimated: false } });
        return answer({ Status: 'Pending' });
      case 'asx_LoadDraft':
        return answer({
          RevisionId: IDS.revision,
          RowVersion: '3',
          Status: 'Published',
          Version: 1,
          Draft: draft,
        });
      case 'asx_CreateDraft':
        return answer({
          TemplateId: IDS.template,
          RevisionId: id(50),
          RowVersion: '4',
          Status: 'Draft',
        });
      case 'asx_PreviewTemplate':
        return answer({
          Folders: [
            { Section: 'general', Node: 'root', Name: 'Contoso Ltd', RelativePath: 'Contoso Ltd' },
            {
              Section: 'general',
              Node: 'general_docs',
              Name: 'General',
              RelativePath: 'Contoso Ltd/General',
            },
          ],
          Notices: [],
        });
      case 'asx_PublishTemplate':
        return answer({ Status: 'Published', Notices: [] });
      case 'asx_CatalogAdmin':
        return answer(catalog(command));
      case 'asx_SecurityAdmin':
        return answer(securityAdmin(command));
      default:
        return answer({ Status: 'Pending' });
    }
  }

  window.Xrm = {
    Utility: {
      getGlobalContext: () => ({
        getClientUrl: () => location.origin,
        userSettings: { userId: '{' + IDS.worker + '}' },
        organizationSettings: {},
      }),
      lookupObjects: async () => [
        { id: '{' + IDS.record.toUpperCase() + '}', name: 'Contoso Ltd', entityType: 'account' },
      ],
      getEntityMetadata: async () => ({ DisplayName: 'Account' }),
    },
    Navigation: {
      navigateTo: async (page) => {
        location.href = '/' + page.webresourceName.replace(/^asx_admin\//, '');
      },
      openForm: async () => {},
      openUrl: () => {},
    },
    WebApi: {
      retrieveMultipleRecords: (table, options) =>
        tracked(async () => ({ entities: rows(table, options) })),
      retrieveRecord: (table, key) => tracked(async () => record(table, key)),
      updateRecord: (table, key) => tracked(async () => ({ id: key })),
      deleteRecord: (table, key) => tracked(async () => ({ id: key })),
      online: { execute: (request) => tracked(async () => execute(request)) },
    },
  };

  // The metadata and privilege URLs the page reads with fetch.
  const json = (value) => new Response(JSON.stringify(value), { status: 200 });
  const real = window.fetch.bind(window);
  window.fetch = (url, init) => {
    const text = String(url);
    if (!text.includes('/api/data/')) return real(url, init);
    return tracked(async () => {
      if (text.includes('RetrieveUserPrivilegeByPrivilegeName'))
        return json({ RolePrivileges: [{ Depth: 'Global' }] });
      if (text.includes('DateTimeAttributeMetadata') || text.includes('MultiSelectPicklist'))
        return json({ value: [] });
      if (text.includes('LookupAttributeMetadata')) return json({ value: [] });
      if (text.includes('StateAttributeMetadata')) return json({ value: states });
      if (text.includes('AttributeMetadata')) return json({ value: [] });
      if (text.includes('/Attributes')) return json({ value: attributes });
      if (text.includes("EntityDefinitions(LogicalName='")) return json(accountInfo);
      if (text.includes('EntityDefinitions')) return json({ value: [accountInfo] });
      return json({ value: [] });
    });
  };
})();
