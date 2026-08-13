import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

test('app shell contains the dashboard and workflow navigation labels', async () => {
  const html = await readFile(new URL('../index.html', import.meta.url), 'utf8');
  for (const label of ['대시보드', '데이터 업로드', '수요 확인', '재고·Open PO', '기기 발주', '옵션 발주', '예외·검토', '보고·출력']) {
    assert.match(html, new RegExp(label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
  }
});

test('design system uses the Fujifilm BI visual tokens', async () => {
  const css = await readFile(new URL('../src/styles.css', import.meta.url), 'utf8');
  assert.match(css, /--color-primary:\s*#018463/);
  assert.match(css, /--container-width:\s*1160px/);
  assert.match(css, /border-radius:\s*2px/);
});

test('dashboard uses a horizontal workflow progress rail', async () => {
  const app = await readFile(new URL('../src/app.js', import.meta.url), 'utf8');
  assert.match(app, /progress-rail/);
  assert.match(app, /progress-node/);
  assert.match(app, /progress-connector/);
});

test('data summary cards expose a detail interaction and the global menu is sticky', async () => {
  const app = await readFile(new URL('../src/app.js', import.meta.url), 'utf8');
  const css = await readFile(new URL('../src/styles.css', import.meta.url), 'utf8');
  assert.match(app, /data-detail/);
  assert.match(app, /detail-panel/);
  assert.match(css, /\.sidebar\s*\{[^}]*position:\s*sticky/);
  assert.match(css, /\.topbar\s*\{[^}]*position:\s*sticky/);
});

test('interactive workflow areas use green hover outlines and Korean text can wrap safely', async () => {
  const css = await readFile(new URL('../src/styles.css', import.meta.url), 'utf8');
  assert.match(css, /\.nav-item:hover[^}]*border-color:\s*var\(--color-primary\)/);
  assert.match(css, /\.progress-node:hover[^}]*border-color:\s*var\(--color-primary\)/);
  assert.match(css, /\.primary-button:hover[^}]*border-color:\s*var\(--color-primary\)/);
  assert.match(css, /word-break:\s*keep-all/);
  assert.match(css, /overflow-wrap:\s*anywhere/);
  assert.match(css, /td,\s*th,\s*\.progress-label,\s*\.progress-meta\s*\{[^}]*white-space:\s*normal/);
});
