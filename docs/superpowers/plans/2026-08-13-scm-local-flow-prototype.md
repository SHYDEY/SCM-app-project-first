# SCM Local Flow Prototype Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a local browser prototype that lets one user inspect the complete SCM order workflow from dashboard through export without implementing production-grade calculations or backend integrations.

**Architecture:** Use a dependency-light static SPA served by a small Node HTTP server. Keep workflow state and demo data in a pure JavaScript module so the UI can later replace local state with Supabase storage and real calculation services. The first phase will expose the entire process as navigable screens with explicit “prototype” boundaries.

**Tech Stack:** Vanilla HTML/CSS/ES modules, Node.js built-in `http` server, Node test runner.

## Global Constraints

- Initial user model is single-user with no login or role permissions.
- Dashboard is the primary navigation surface.
- Excel upload is the primary input concept; this phase presents upload/template controls but does not parse production Excel files.
- Direct screen editing is a secondary prototype interaction.
- Both demo-data and empty-workspace modes are required.
- localStorage is the local persistence boundary; Supabase is deferred.
- Outputs are prototype downloads only; no real order execution occurs.
- The app must run locally with a single documented command.

---

### Task 1: Establish the prototype contract with tests

**Files:**
- Create: `src/state.js`
- Create: `test/state.test.js`

**Interfaces:**
- Produces `createDemoState()`, `createEmptyState()`, `getWorkflowSteps()`, and `getDashboardMetrics(state)`.
- State fields: `mode`, `workspaceName`, `orderMonth`, `activeStep`, `workflowSteps`, `metrics`, `records`, `lastSavedAt`.

- [ ] **Step 1: Write failing tests**

```js
import test from 'node:test';
import assert from 'node:assert/strict';
import { createDemoState, createEmptyState, getWorkflowSteps, getDashboardMetrics } from '../src/state.js';

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
```

- [ ] **Step 2: Run the test and verify it fails**

Run: `node --test test/state.test.js`  
Expected: FAIL because `src/state.js` does not exist yet.

- [ ] **Step 3: Implement the minimal state model**

Implement demo rows for demand, inventory, machine orders, option orders, and exceptions. Derive dashboard metrics from those rows without introducing Excel or backend dependencies.

- [ ] **Step 4: Run the test and verify it passes**

Run: `node --test test/state.test.js`  
Expected: PASS with 2 passing tests.

### Task 2: Build the local app shell and navigation

**Files:**
- Create: `index.html`
- Create: `src/app.js`
- Create: `src/styles.css`
- Create: `server.mjs`
- Create: `package.json`

**Interfaces:**
- `src/app.js` consumes the state functions from Task 1.
- The shell exposes navigation IDs: `dashboard`, `upload`, `demand`, `inventory`, `machine`, `option`, `review`, `report`.
- `server.mjs` serves the workspace root at `http://localhost:3000`.

- [ ] **Step 1: Write failing UI smoke test**

```js
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

test('app shell contains the dashboard and all workflow navigation labels', async () => {
  const html = await readFile(new URL('../index.html', import.meta.url), 'utf8');
  for (const label of ['대시보드', '데이터 업로드', '수요 확인', '재고·Open PO', '기기 발주', '옵션 발주', '예외·검토', '보고·출력']) {
    assert.match(html, new RegExp(label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
  }
});
```

- [ ] **Step 2: Run the test and verify it fails**

Run: `node --test test/app-shell.test.js`  
Expected: FAIL because the app shell has not been created.

- [ ] **Step 3: Implement the shell**

Create the responsive app layout: left navigation, top bar with workspace mode, content viewport, and a fixed prototype notice. Navigation changes the current view without a page reload.

- [ ] **Step 4: Run the test and verify it passes**

Run: `node --test test/app-shell.test.js`  
Expected: PASS.

### Task 3: Implement dashboard and workflow overview

**Files:**
- Modify: `src/app.js`
- Modify: `src/styles.css`

**Interfaces:**
- Dashboard renders `getDashboardMetrics(state)`.
- Each workflow card can navigate to its corresponding screen.
- Mode controls call `createDemoState()` and `createEmptyState()`.

- [ ] **Step 1: Add a failing state-transition test**

```js
test('workflow step state can move from not-started to in-progress', () => {
  const state = createDemoState();
  const demand = state.workflowSteps.find((step) => step.id === 'demand');
  assert.equal(demand.status, 'in-progress');
});
```

- [ ] **Step 2: Run the test and verify the expected failure**

Run: `node --test test/state.test.js`  
Expected: FAIL if the demo workflow statuses are not defined.

- [ ] **Step 3: Implement dashboard cards and status rail**

Show total order amount, machine amount, option amount, exception count, and workflow progress. Include clear labels that the values are demo or empty-state values.

- [ ] **Step 4: Run all tests**

Run: `node --test`  
Expected: PASS.

### Task 4: Add stage screens for the end-to-end flow

**Files:**
- Modify: `src/app.js`
- Modify: `src/styles.css`

**Interfaces:**
- Every screen receives the same state and renders a compact table or checklist.
- Screen actions update only prototype state: `markStepComplete(id)`, `selectRow(id)`, and `setWorkspaceMode(mode)`.

- [ ] **Step 1: Add failing coverage for required screen headings**

```js
test('workflow screen registry contains all prototype screens', () => {
  const screenIds = ['dashboard', 'upload', 'demand', 'inventory', 'machine', 'option', 'review', 'report'];
  assert.deepEqual(getWorkflowScreens().map((screen) => screen.id), screenIds);
});
```

- [ ] **Step 2: Run test and verify it fails**

Run: `node --test test/state.test.js`  
Expected: FAIL until `getWorkflowScreens()` exists.

- [ ] **Step 3: Implement the stage screens**

Implement these prototype screens:

- Upload: drag/drop-style panel, template download button, upload activity list, validation checklist.
- Demand: source comparison table with OL, SFDC, Trend, Bulk-deal, and meeting adjustment columns.
- Inventory: stock and Open PO table with “가용재고 계산 예정” callout.
- Machine order: sample order table and calculation explanation panel.
- Option order: BOM/attachment/Common/mandatory option table and warning badges.
- Review: exception queue and adjustment/approval placeholders.
- Report: executive summary, comparison table, and export actions.

All nonimplemented calculations must be labeled “2차 구현 예정” instead of appearing complete.

- [ ] **Step 4: Run tests and start the server**

Run: `node --test`  
Run: `node server.mjs`  
Expected: tests pass and the server listens on port 3000.

### Task 5: Add local persistence and prototype exports

**Files:**
- Modify: `src/app.js`
- Modify: `src/styles.css`
- Modify: `test/state.test.js`

**Interfaces:**
- `saveState(state)` writes to localStorage when available.
- `loadState()` restores a saved state or returns `null`.
- `downloadJson(state)` creates a JSON download.
- `downloadTemplate()` creates a prototype CSV/Excel-compatible template download.

- [ ] **Step 1: Write failing persistence tests**

```js
test('state serialization preserves workspace mode and order month', () => {
  const state = createDemoState();
  const parsed = JSON.parse(JSON.stringify(state));
  assert.equal(parsed.mode, 'demo');
  assert.equal(parsed.orderMonth, state.orderMonth);
});
```

- [ ] **Step 2: Run the test and verify it fails only if serialization contract is missing**

Run: `node --test test/state.test.js`  
Expected: FAIL until the state contains the required fields.

- [ ] **Step 3: Implement persistence and downloads**

Autosave after mode switches and prototype row edits. Add “JSON 백업”, “샘플 입력 템플릿”, and “작업공간 초기화” actions. Never overwrite the current state without a confirmation dialog for reset.

- [ ] **Step 4: Run all tests and manually exercise the flow**

Run: `node --test`  
Then open `http://localhost:3000` and check mode switch, navigation, reset confirmation, and downloads.

### Task 6: Verify the prototype against the PRD and hand off for approval

**Files:**
- Modify: `README.md`
- Create: `docs/verification/2026-08-13-scm-local-prototype.md`

- [ ] **Step 1: Document local startup and prototype boundaries**

Document the exact commands:

```powershell
npm start
```

Document that Excel parsing, real calculations, Supabase, authentication, and actual order execution are intentionally deferred.

- [ ] **Step 2: Run verification commands**

Run: `node --test`  
Run: `npm start` and load `http://localhost:3000`  
Check every navigation item and both data modes.

- [ ] **Step 3: Stop for user approval**

Do not start the second implementation phase until the user approves the visible prototype and explicitly asks to continue.
