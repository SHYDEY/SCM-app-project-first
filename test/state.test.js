import test from 'node:test';
import assert from 'node:assert/strict';
import { createDemoState, createEmptyState, getWorkflowSteps, getWorkflowScreens, getDashboardMetrics } from '../src/state.js';

test('demo state exposes the complete seven-step workflow', () => {
  const state = createDemoState();
  assert.equal(state.mode, 'demo');
  assert.equal(getWorkflowSteps().length, 7);
  assert.equal(state.workflowSteps.at(-1).id, 'export');
});

test('empty state starts without order outputs', () => {
  const state = createEmptyState();
  const metrics = getDashboardMetrics(state);
  assert.equal(state.mode, 'empty');
  assert.equal(metrics.totalOrderAmount, 0);
  assert.equal(metrics.exceptions, 0);
});

test('demo demand step is ready for the prototype flow', () => {
  const state = createDemoState();
  const demand = state.workflowSteps.find((step) => step.id === 'demand');
  assert.equal(demand.status, 'in-progress');
});

test('workflow screen registry contains all prototype screens', () => {
  const screenIds = ['dashboard', 'upload', 'demand', 'inventory', 'machine', 'option', 'review', 'report'];
  assert.deepEqual(getWorkflowScreens().map((screen) => screen.id), screenIds);
});

test('state serialization preserves workspace mode and order month', () => {
  const state = createDemoState();
  const parsed = JSON.parse(JSON.stringify(state));
  assert.equal(parsed.mode, 'demo');
  assert.equal(parsed.orderMonth, state.orderMonth);
});
