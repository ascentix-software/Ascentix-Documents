'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');

// Every Dataverse worker call retries temporary failures (throttling, 5xx, timeouts, version
// conflicts) in the flow before its failure branch runs. Worker commands are transactional, so a
// failed call changed nothing and is safe to repeat.
const DATAVERSE_RETRY = {
  type: 'exponential',
  count: 4,
  interval: 'PT10S',
  minimumInterval: 'PT5S',
  maximumInterval: 'PT1H',
};

// Flow concurrency stays at 1 today. No flow may ever exceed 5 parallel loop repetitions or
// trigger runs: the shared writer budget admits two writers and the claim protocol serializes
// work per site, so more parallel runs only add Dataverse load and contention.
const MAX_CONCURRENCY = 5;

// Failure branches pass the failed action's status, error code and message to the worker so it
// can wait after a temporary failure instead of blocking.
// A SharePoint write runs only when BeginHttp granted it. BeginHttp answers Stopped when the
// destination was removed or suspended after the write was prepared; the drive then ends
// cleanly with no write and no further worker call for that operation.
const PERMITTED = "@equals(variables('HttpAdmission')?['Status'],'Permit')";
const PERMIT_OR_STOPPED =
  "@or(equals(variables('HttpAdmission')?['Status'],'Permit'),equals(variables('HttpAdmission')?['Status'],'Stopped'))";

const FAILURE_BRANCHES = [
  'Record_blocked_outbox',
  'Record_failed_claim',
  'Record_failed_operation',
];

/**
 * Checks one flow definition and returns the number of connector actions. Throws on a violation.
 */
function checkFlow(flow) {
  let calls = 0;
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
    flow.properties.definition.actions.Initialize_worker_action.inputs.variables.filter(
      (v) => v.name === 'WorkerAction',
    ),
    [{ name: 'WorkerAction', type: 'string', value: 'asx_DocumentWorker' }],
  );
  const names = new Set();
  // scope: the enclosing If (key, action, parent block) and whether this block is its else.
  function actions(block, scope) {
    for (const [key, action] of Object.entries(block)) {
      assert(!names.has(key), 'Duplicate action name: ' + key);
      names.add(key);
      if (action.type === 'SetVariable') assert.notEqual(action.inputs.name, 'WorkerAction');
      if (action.type === 'OpenApiConnection') {
        calls++;
        if (action.inputs.host.connectionName === 'shared_webcontents') {
          // SharePoint retries are decided by the server (Retry-After, unknown-outcome
          // quarantine), never by the connector: a repeated POST could write twice.
          assert.deepEqual(
            action.inputs.retryPolicy,
            { type: 'none' },
            'SharePoint HTTP action must not retry: ' + key,
          );
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
          if (action.inputs.parameters['request/method'] === 'GET') {
            assert.deepEqual(Object.keys(action.runAfter).length, 1);
            assert(Object.keys(action.runAfter)[0].startsWith('Await_'));
          } else {
            // A write is reachable only in the true branch of the Permit check, which runs
            // right after its admission loop.
            assert(
              scope && !scope.otherwise && scope.action.type === 'If',
              'SharePoint write must sit inside the Permit check: ' + key,
            );
            assert.equal(
              scope.action.expression,
              PERMITTED,
              'SharePoint write must run only on Permit: ' + key,
            );
            assert.deepEqual(action.runAfter, {}, 'SharePoint write runs first in its branch');
            const after = Object.keys(scope.action.runAfter);
            assert.equal(after.length, 1);
            assert(after[0].startsWith('Await_'), 'The Permit check follows its admission loop');
            // A stopped permit ends the drive: the else branch only hands the result to the
            // loop and makes no worker or SharePoint call.
            const otherwise = Object.values(scope.action.else?.actions || {});
            assert.equal(otherwise.length, 1, 'Stopped permit ends the drive with one step');
            assert.equal(otherwise[0].type, 'SetVariable');
            assert.equal(otherwise[0].inputs.name, 'Work');
            assert.equal(otherwise[0].inputs.value, "@variables('HttpAdmission')");
          }
        } else {
          assert.deepEqual(
            action.inputs.retryPolicy,
            DATAVERSE_RETRY,
            'Dataverse action must retry temporary failures: ' + key,
          );
          assert.equal(action.inputs.host.operationId, 'PerformUnboundAction');
          assert.equal(action.inputs.parameters.actionName, "@variables('WorkerAction')");
          assert.deepEqual(Object.keys(action.inputs.parameters).sort(), ['actionName', 'item']);
          assert.deepEqual(Object.keys(action.inputs.parameters.item), ['Request']);
          assert(action.inputs.parameters.item.Request.startsWith('@string('));
        }
      }
      if (FAILURE_BRANCHES.includes(key))
        for (const field of ["'StatusCode'", "'ErrorCode'", "'Error'"])
          assert(
            action.inputs.parameters.item.Request.includes(field),
            key + ' must pass the failure ' + field,
          );
      // A drive that failed with no failed Dataverse action reports DriveFailed (not temporary),
      // except when a slot wait hit its loop limit, which stays temporary.
      if (key === 'Record_failed_operation')
        for (const part of [
          "'DriveFailed'",
          "actions('Await_read_slot')",
          "actions('Await_create_slot')",
        ])
          assert(
            action.inputs.parameters.item.Request.includes(part),
            key + ' must classify a drive failure without a failed action: ' + part,
          );
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
      if (action.type === 'Foreach') {
        const repetitions = action.runtimeConfiguration?.concurrency?.repetitions;
        assert(
          Number.isInteger(repetitions) && repetitions >= 1 && repetitions <= MAX_CONCURRENCY,
          'Loop concurrency must be set and at most ' + MAX_CONCURRENCY + ': ' + key,
        );
      }
      if (action.type === 'Until') {
        if (key.startsWith('Await_')) {
          assert.equal(action.limit.count, 120);
          assert.equal(action.limit.timeout, 'PT2H');
          assert.equal(
            action.expression,
            key.includes('read') ? PERMITTED : PERMIT_OR_STOPPED,
            'Admission loop exit: ' + key,
          );
        } else {
          assert.equal(action.limit.count, 80);
          assert.equal(action.limit.timeout, 'PT3H');
        }
      }
      const inner = action.type === 'If' ? { key, action, otherwise: false } : scope;
      if (action.actions) actions(action.actions, inner);
      for (const item of Object.values(action.cases || {})) actions(item.actions, null);
      if (action.default?.actions) actions(action.default.actions, null);
      if (action.else?.actions) actions(action.else.actions, { key, action, otherwise: true });
    }
  }
  actions(flow.properties.definition.actions, null);
  for (const [key, trigger] of Object.entries(flow.properties.definition.triggers)) {
    const runs = trigger.runtimeConfiguration?.concurrency?.runs;
    assert(
      Number.isInteger(runs) && runs >= 1 && runs <= MAX_CONCURRENCY,
      'Trigger concurrency must be set and at most ' + MAX_CONCURRENCY + ': ' + key,
    );
  }
  for (const required of [
    'Claim_operation',
    'Prepare_create',
    'Read_HTTP',
    'Create_HTTP',
    'Observe_read',
    'Record_create_response',
    'Complete_operation',
    'Record_failed_operation',
    'Record_failed_claim',
    'Begin_read',
    'Begin_create',
    'Await_read_slot',
    'Await_create_slot',
  ])
    assert(names.has(required));
  // Never grant recovery automatically from a recurrence or manual worker run.
  assert(!serialized.includes('asx_RecoverWorker'));
  return calls;
}

function find(block, predicate) {
  for (const [key, action] of Object.entries(block)) {
    if (predicate(key, action)) return action;
    const nested = [
      action.actions,
      ...Object.values(action.cases || {}).map((c) => c.actions),
      action.default?.actions,
      action.else?.actions,
    ].filter(Boolean);
    for (const inner of nested) {
      const hit = find(inner, predicate);
      if (hit) return hit;
    }
  }
  return null;
}

/**
 * Proves the checks reject bad definitions: each temporary in-memory fixture breaks one rule.
 */
function selfTest(flow) {
  checkFlow(flow);
  const broken = {
    'loop concurrency above 5': (f) => {
      find(
        f.properties.definition.actions,
        (_, a) => a.type === 'Foreach',
      ).runtimeConfiguration.concurrency.repetitions = MAX_CONCURRENCY + 1;
    },
    'trigger concurrency above 5': (f) => {
      Object.values(f.properties.definition.triggers)[0].runtimeConfiguration.concurrency.runs =
        MAX_CONCURRENCY + 1;
    },
    'Dataverse action without retries': (f) => {
      find(
        f.properties.definition.actions,
        (_, a) => a.inputs?.host?.connectionName === 'shared_commondataserviceforapps',
      ).inputs.retryPolicy = { type: 'none' };
    },
    'SharePoint HTTP action with retries': (f) => {
      find(
        f.properties.definition.actions,
        (_, a) => a.inputs?.host?.connectionName === 'shared_webcontents',
      ).inputs.retryPolicy = { ...DATAVERSE_RETRY };
    },
    'SharePoint write without the Permit check': (f) => {
      const guard = find(f.properties.definition.actions, (k) => k === 'Write_when_permitted');
      guard.expression = PERMIT_OR_STOPPED;
    },
    'SharePoint write in the stopped branch': (f) => {
      const guard = find(f.properties.definition.actions, (k) => k === 'Write_when_permitted');
      [guard.actions, guard.else.actions] = [guard.else.actions, guard.actions];
    },
    'SharePoint write right after its admission loop': (f) => {
      const create = find(f.properties.definition.actions, (_, a) => a.cases?.Create).cases.Create
        .actions;
      const guard = create.Write_when_permitted;
      delete create.Write_when_permitted;
      Object.assign(create, guard.actions);
      create.Create_HTTP.runAfter = { Await_create_slot: ['Succeeded'] };
    },
    'stopped permit that does not end the drive': (f) => {
      const guard = find(f.properties.definition.actions, (k) => k === 'Write_when_permitted');
      guard.else.actions.Save_stopped_state.inputs.value = "@json(body('Begin_create')?['Result'])";
    },
    'failure branch without its cause': (f) => {
      const branch = find(f.properties.definition.actions, (k) => k === 'Record_failed_claim');
      branch.inputs.parameters.item.Request =
        "@string(setProperty(setProperty(json('{}'),'Command','FailUnclaimed'),'Key',items('For_each_operation')))";
    },
  };
  for (const [name, breakIt] of Object.entries(broken)) {
    const copy = JSON.parse(JSON.stringify(flow));
    breakIt(copy);
    assert.throws(() => checkFlow(copy), assert.AssertionError, 'Check must reject: ' + name);
  }
  return Object.keys(broken).length;
}

module.exports = { checkFlow, selfTest, DATAVERSE_RETRY, MAX_CONCURRENCY };

if (require.main === module) {
  const directory = process.argv[2] || path.join(root, 'solution/AscentixDocuments/src/Workflows');
  const files = fs.readdirSync(directory).filter((name) => name.endsWith('.json'));
  assert.equal(files.length, 2, 'Two solution worker flows');
  let calls = 0;
  let rejected = 0;
  for (const name of files) {
    const flow = JSON.parse(fs.readFileSync(path.join(directory, name), 'utf8'));
    calls += checkFlow(flow);
    rejected += selfTest(flow);
  }
  console.log(
    'PASS: two flow definitions; ' +
      calls +
      ' connector actions: Dataverse calls retry exponentially, SharePoint HTTP calls never retry, loops and triggers stay at or below ' +
      MAX_CONCURRENCY +
      ', failure branches pass their cause, guarded API commands and approved transport shape. ' +
      rejected +
      ' broken in-memory fixtures were rejected. Live designer/import validation is NOT RUN.',
  );
}
