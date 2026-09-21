'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
let calls = 0;
const directory = process.argv[2] || path.join(root, 'solution/AscentixDocuments/src/Workflows');
const files = fs.readdirSync(directory).filter((name) => name.endsWith('.json'));
assert.equal(files.length, 2, 'Two solution worker flows');
for (const name of files) {
  const flow = JSON.parse(fs.readFileSync(path.join(directory, name), 'utf8'));
  const serialized = JSON.stringify(flow);
  assert(!serialized.includes('shared_sharepointonline'));
  assert(!serialized.includes('graph.microsoft.com'));
  assert(!serialized.includes('AscentixSharePointDemo'));
  const refs = flow.properties.connectionReferences;
  assert.equal(
    refs.shared_webcontents.connection.connectionReferenceLogicalName,
    'asx_documentshttp',
  );
  assert.equal(
    refs.shared_commondataserviceforapps.connection.connectionReferenceLogicalName,
    'asx_documentsdataverse',
  );
  assert(
    Object.values(refs).every((r) => !r.connection.name && !r.connection.id),
    'Solution references must not embed environment connections',
  );
  assert.deepEqual(
    flow.properties.definition.actions.Initialize_work.inputs.variables.filter(
      (v) => v.name === 'WorkerAction',
    ),
    [{ name: 'WorkerAction', type: 'string', value: 'asx_DocumentWorker' }],
  );
  const names = new Set();
  function actions(block) {
    for (const [key, action] of Object.entries(block)) {
      assert(!names.has(key), 'Duplicate action name: ' + key);
      names.add(key);
      if (action.type === 'SetVariable') assert.notEqual(action.inputs.name, 'WorkerAction');
      if (action.type === 'OpenApiConnection') {
        calls++;
        assert.equal(action.inputs.retryPolicy.type, 'none');
        if (action.inputs.host.connectionName === 'shared_webcontents') {
          assert.equal(action.inputs.host.operationId, 'InvokeHttp');
          assert(['GET', 'POST'].includes(action.inputs.parameters['request/method']));
          assert(action.inputs.parameters['request/headers']);
          assert(
            Object.keys(action.inputs.parameters).every((k) => k.startsWith('request/')),
            'InvokeHttp uses request-prefixed serialized connector parameters',
          );
          assert.equal(
            action.inputs.parameters['request/url'],
            "@concat(variables('Work')?['SiteUrl'],'/',variables('Work')?['Http']?['RelativeUri'])",
          );
          assert.deepEqual(Object.keys(action.runAfter).length, 1);
          assert(Object.keys(action.runAfter)[0].startsWith('Await_'));
        } else {
          assert.equal(action.inputs.host.operationId, 'PerformUnboundAction');
          assert.equal(action.inputs.parameters.actionName, "@variables('WorkerAction')");
          assert.deepEqual(Object.keys(action.inputs.parameters).sort(), ['actionName', 'item']);
          assert.deepEqual(Object.keys(action.inputs.parameters.item), ['Request']);
          assert(action.inputs.parameters.item.Request.startsWith('@string('));
        }
      }
      if (action.type === 'Until')
        assert.equal(
          action.operationOptions,
          'FailWhenLimitsReached',
          'Reaching a loop limit must not admit an unpermitted HTTP request',
        );
      if (key.startsWith('Await_')) {
        const kind = key.includes('read') ? 'read' : 'create';
        assert.deepEqual(action.runAfter, { ['Reset_' + kind + '_permit']: ['Succeeded'] });
        assert.deepEqual(block['Reset_' + kind + '_permit'].inputs.value, {});
      }
      if (action.type === 'Foreach')
        assert.equal(action.runtimeConfiguration.concurrency.repetitions, 1);
      if (action.type === 'Until') {
        if (key.startsWith('Await_')) {
          assert.equal(action.limit.count, 120);
          assert.equal(action.limit.timeout, 'PT2H');
          assert(action.expression.includes('Permit'));
        } else {
          assert.equal(action.limit.count, 80);
          assert.equal(action.limit.timeout, 'PT3H');
        }
      }
      if (action.actions) actions(action.actions);
      for (const item of Object.values(action.cases || {})) actions(item.actions);
      if (action.default?.actions) actions(action.default.actions);
      if (action.else?.actions) actions(action.else.actions);
    }
  }
  actions(flow.properties.definition.actions);
  for (const trigger of Object.values(flow.properties.definition.triggers))
    assert.equal(trigger.runtimeConfiguration.concurrency.runs, 1);
  for (const required of [
    'Claim_operation',
    'Prepare_create',
    'Read_HTTP',
    'Create_HTTP',
    'Observe_read',
    'Record_create_response',
    'Complete_operation',
    'Record_failed_operation',
    'Begin_read',
    'Begin_create',
    'Await_read_slot',
    'Await_create_slot',
  ])
    assert(names.has(required));
  // Never grant recovery automatically from a recurrence or manual worker run.
  assert(!serialized.includes('asx_RecoverWorker'));
}
console.log(
  'PASS: two flow definitions; ' +
    calls +
    ' connector actions have None retries, serialized loops, guarded API commands and approved transport shape. Live designer/import validation is NOT RUN.',
);
