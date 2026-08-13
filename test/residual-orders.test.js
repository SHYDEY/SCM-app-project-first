import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

test('previous-month residual orders are modeled as a separate demand input', async () => {
  const builder = await readFile(new URL('../build_dummy_2025.mjs', import.meta.url), 'utf8');
  const app = await readFile(new URL('../src/app.js', import.meta.url), 'utf8');
  const state = await readFile(new URL('../src/state.js', import.meta.url), 'utf8');
  assert.match(builder, /전월도 잔여주문/);
  assert.match(builder, /적용 수요/);
  assert.match(builder, /전월도 잔여주문\(기기\)/);
  assert.match(app, /전월도 잔여주문/);
  assert.match(state, /residualOrder/);
});
