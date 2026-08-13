import { createDemoState, createEmptyState, getDashboardMetrics, getWorkflowScreens } from './state.js';

const STORAGE_KEY = 'scm-local-flow-prototype';
let state = loadState() || createDemoState();

const money = (value) => `₩${new Intl.NumberFormat('ko-KR').format(value)}`;
const number = (value) => new Intl.NumberFormat('ko-KR').format(value || 0);
const escapeHtml = (value) => String(value ?? '').replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char]));

function saveState() {
  state.lastSavedAt = new Date().toISOString();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function loadState() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (!saved?.records) return saved;
    saved.records.inventory?.forEach((row, index) => { row.residualOrder ??= [3, 4, 2, 3][index] || 0; });
    saved.records.machine?.forEach((row, index) => { row.residualOrder ??= [3, 4, 2][index] || 0; });
    saved.records.option?.forEach((row, index) => { row.residualOrder ??= [2, 3, 2][index] || 0; });
    return saved;
  } catch { return null; }
}

function setState(next) {
  state = next;
  saveState();
  render();
}

function badge(value, tone='neutral') {
  return `<span class="badge badge-${tone}">${escapeHtml(value)}</span>`;
}

function statusTone(value) {
  if (['정상','확정','완료','가상데이터'].includes(value)) return 'success';
  if (['검토 필요','Flex 미달','부족 예상'].includes(value)) return 'danger';
  if (['MOQ 과잉 확인','조치 중','검토'].includes(value)) return 'warning';
  return 'neutral';
}

function table(headers, rows, emptyText='데이터가 없습니다.') {
  if (!rows.length) return `<div class="empty-state"><div class="empty-icon">∅</div><strong>${emptyText}</strong><span>데이터 업로드 또는 가상데이터로 시작해보세요.</span></div>`;
  return `<div class="table-wrap"><table><thead><tr>${headers.map((h) => `<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.join('')}</tbody></table></div>`;
}

function detailPanel(key) {
  const details = {
    total: ['총 발주 예정 금액', '기기와 옵션의 최종 발주 예정 금액을 합산한 데모 수치입니다.', [['기기 발주', '₩52,500,000'], ['옵션 발주', '₩22,500,000'], ['전월 대비', '+8.4%']]],
    machine: ['기기 발주 상세', '기종별 수요·가용재고·MOQ를 기준으로 계산된 현재 데모 결과입니다.', [['Apeos C7070', '15대 · 정상'], ['ApeosPort 5570', '15대 · Flex 미달'], ['Apeos C3060', '20대 · MOQ 과잉 확인']]],
    option: ['옵션 발주 상세', 'BOM·장착율·Common품을 반영한 현재 데모 결과입니다.', [['CT202480', '0개 · 재고 충분'], ['EL300123', '10개 · MOQ 과잉 확인'], ['OF300456', '20개 · Common품']]],
    exceptions: ['검토 필요 상세', '계산 결과 중 사용자의 확인이나 조정이 필요한 항목입니다.', [['Flex 미달', 'ApeosPort 5570'], ['MOQ 과잉', 'Apeos C3060 / EL300123'], ['데이터 누락', 'SFDC 수요 8건']]],
    ol: ['OL 제출 수량 상세', '영업에서 제출한 출고 Outlook 기준의 데모 수치입니다.', [['확정 OL', '76대'], ['조정 전 수량', '98대'], ['조정 사유', '전월 제출 OL 및 수급회의 반영']]],
    sfdc: ['SFDC 파이프라인 상세', '중요 파이프라인을 참고 수요로 표시한 데모 수치입니다.', [['검토 파이프라인', '28대'], ['반영 예정', '20대'], ['상태', '수급회의 검토']]],
    bulk: ['Bulk-deal 상세', 'OL 외 대형 프로젝트 수요로 별도 확인이 필요한 데모 수치입니다.', [['프로젝트 수요', '16대'], ['필요월도', '2026-10'], ['상태', '확정']]],
    finalDemand: ['최종 확정 수요 상세', 'OL·SFDC·Bulk-deal·Trend를 검토한 후 계산에 사용하는 데모 수치입니다.', [['최종 확정', '86대'], ['확정 기준', '수급회의'], ['상태', '확정 전 프로토타입']]],
  };
  const [title, description, rows] = details[key] || details.total;
  return `<div class="detail-panel" role="region" aria-label="${title}"><div class="detail-heading"><div><span class="eyebrow">DETAIL VIEW</span><h4>${title}</h4><p>${description}</p></div><button class="detail-close" data-detail-close aria-label="상세 닫기">×</button></div><div class="detail-items">${rows.map(([label, value]) => `<div><span>${label}</span><strong>${value}</strong></div>`).join('')}</div></div>`;
}

function layout(content) {
  const screens = getWorkflowScreens();
  const metrics = getDashboardMetrics(state);
  return `<div class="app-shell">
    <aside class="sidebar">
      <div class="brand"><div class="brand-mark">S</div><div><strong>SCM Flow</strong><small>수급·발주 관리</small></div></div>
      <div class="workspace-mini"><span class="eyebrow">CURRENT WORKSPACE</span><strong>${escapeHtml(state.workspaceName)}</strong><span>${state.orderMonth} · ${state.mode === 'demo' ? badge('DEMO','info') : badge('빈 작업공간')}</span></div>
      <nav>${screens.map((screen) => `<button class="nav-item ${state.activeScreen === screen.id ? 'active' : ''}" data-screen="${screen.id}"><span class="nav-icon">${screen.icon}</span>${screen.label}${screen.id === 'review' && metrics.exceptions ? `<em>${metrics.exceptions}</em>` : ''}</button>`).join('')}</nav>
      <div class="sidebar-bottom"><button class="text-button" data-action="demo">↻ 데모 데이터로 시작</button><button class="text-button" data-action="empty">＋ 빈 작업공간으로 시작</button><div class="save-status">자동 저장 ${state.lastSavedAt ? '완료' : '대기 중'}</div></div>
    </aside>
    <main class="main"><header class="topbar"><div><span class="eyebrow">MONTHLY ORDER CYCLE</span><h1>${pageTitle()}</h1></div><div class="top-actions"><span class="prototype-pill">PROTOTYPE</span><button class="icon-button" data-action="reset" title="작업공간 초기화">↺</button><button class="avatar">담당</button></div></header><div class="content">${content}</div><div class="prototype-banner"><span>●</span> 1차 프로토타입 · 전체 플로우 확인용입니다. Excel 파싱·실제 계산·Supabase 연동은 2차 구현에서 진행합니다.</div></main>
  </div>`;
}

function pageTitle() {
  return getWorkflowScreens().find((screen) => screen.id === state.activeScreen)?.label || '대시보드';
}

function render() {
  document.querySelector('#app').innerHTML = layout(renderScreen());
  bindEvents();
}

function renderScreen() {
  const views = { dashboard: renderDashboard, upload: renderUpload, demand: renderDemand, inventory: renderInventoryWithResidual, machine: renderMachineWithResidual, option: renderOptionWithResidual, review: renderReview, report: renderReport };
  return (views[state.activeScreen] || renderDashboard)();
}

function renderDashboard() {
  const m = getDashboardMetrics(state);
  const steps = state.workflowSteps;
  return `<section class="hero"><div><span class="eyebrow">${state.orderMonth} ORDER CYCLE</span><h2>이번 달 발주 현황을<br/><span>한눈에 확인하세요.</span></h2><p>수요 확정부터 System 출력까지, 월간 발주 프로세스의 현재 상태를 관리합니다.</p></div><div class="hero-actions"><button class="primary-button" data-screen="upload">데이터 업로드 시작 <span>→</span></button><button class="secondary-button" data-action="report">보고 자료 미리보기</button></div></section>
  <div class="kpi-grid"><button class="kpi-card accent data-card" data-detail="total"><span>총 발주 예정 금액</span><strong>${money(m.totalOrderAmount)}</strong><small>기기 + 옵션 · 데모 기준</small><b class="card-hint">상세 보기 ↗</b></button><button class="kpi-card data-card" data-detail="machine"><span>기기 발주</span><strong>${number(m.machineQuantity)} <small>대</small></strong><small>${money(m.machineAmount)}</small><b class="card-hint">상세 보기 ↗</b></button><button class="kpi-card data-card" data-detail="option"><span>옵션 발주</span><strong>${number(m.optionQuantity)} <small>개</small></strong><small>${money(m.optionAmount)}</small><b class="card-hint">상세 보기 ↗</b></button><button class="kpi-card warning data-card" data-detail="exceptions"><span>검토 필요</span><strong>${number(m.exceptions)} <small>건</small></strong><small>Flex ${m.flexExceeded}건 · 필수 옵션 ${m.mandatoryMissing}건</small><b class="card-hint">상세 보기 ↗</b></button></div><div class="detail-slot"></div>
  <div class="section-heading"><div><span class="eyebrow">WORKFLOW PROGRESS</span><h3>월간 발주 진행 현황</h3></div><span class="muted">${steps.filter((s) => s.status === 'completed').length} / ${steps.length} 완료</span></div><div class="progress-rail">${steps.map((step, index) => `<button class="progress-node ${step.status}" data-screen="${step.id === 'collect' ? 'upload' : step.id === 'export' ? 'report' : step.id}"><span class="progress-node-top"><span class="progress-icon">${step.status === 'completed' ? '✓' : step.status === 'in-progress' ? '•' : step.short}</span>${index < steps.length - 1 ? `<span class="progress-connector ${step.status === 'completed' ? 'complete' : ''}"></span>` : ''}</span><span class="progress-label">${step.label}</span><span class="progress-meta">${step.status === 'completed' ? '완료' : step.status === 'in-progress' ? '진행 중' : '대기'}</span></button>`).join('')}</div>
  <div class="dashboard-lower"><div class="panel"><div class="panel-heading"><div><span class="eyebrow">EXCEPTIONS</span><h3>검토가 필요한 항목</h3></div><button class="link-button" data-screen="review">전체 보기 →</button></div>${state.records.exceptions.slice(0, 3).map((row) => `<div class="exception-row"><span class="exception-dot ${row.severity}"></span><div><strong>${row.type}</strong><span>${row.target}</span></div>${badge(row.status, statusTone(row.status))}</div>`).join('') || '<div class="mini-empty">현재 검토 항목이 없습니다.</div>'}</div><div class="panel"><div class="panel-heading"><div><span class="eyebrow">LATEST UPLOAD</span><h3>최근 데이터 업로드</h3></div><button class="link-button" data-screen="upload">업로드 관리 →</button></div>${state.uploadHistory.map((row) => `<div class="upload-row"><div class="file-icon">×</div><div><strong>${row.name}</strong><span>${row.rows} rows · ${row.time}</span></div>${badge(row.status, 'info')}</div>`).join('') || '<div class="mini-empty">아직 업로드한 파일이 없습니다.</div>'}</div></div>`;
}

function renderUpload() {
  return `<div class="page-intro"><div><span class="eyebrow">STEP 01 · DATA INTAKE</span><h2>데이터 업로드</h2><p>통합 Excel 또는 업무별 파일을 업로드해 발주 사이클을 시작합니다.</p></div><button class="primary-button" data-action="template">템플릿 다운로드 ↓</button></div><div class="upload-layout"><div class="upload-drop"><div class="upload-symbol">↑</div><h3>Excel 파일을 여기에 놓으세요</h3><p>통합 파일 또는 업무별 파일을 지원합니다.</p><button class="secondary-button" data-action="fake-upload">파일 선택</button><small>지원 예정: .xlsx, .xls · 최대 10MB</small></div><div class="panel upload-check"><div class="panel-heading"><div><span class="eyebrow">IMPORT GUIDE</span><h3>업로드 구성</h3></div>${badge('프로토타입','info')}</div>${['수요', '재고', 'Open PO', '기준정보', 'BOM', '과거 실적', '발주 이력'].map((item, i) => `<div class="check-row"><span class="check-number">0${i + 1}</span><strong>${item}</strong><span class="muted">${i < 3 ? '핵심 입력' : '선택 입력'}</span></div>`).join('')}</div></div><div class="section-heading"><div><span class="eyebrow">UPLOAD HISTORY</span><h3>업로드 이력</h3></div></div>${table(['파일명','상태','행 수','업로드 시각'], state.uploadHistory.map((row) => `<tr><td class="strong">${row.name}</td><td>${badge(row.status, 'info')}</td><td>${number(row.rows)}</td><td>${row.time}</td></tr>`), '아직 업로드한 파일이 없습니다.')}`;
}

function renderDemand() {
  const rows = state.records.demand.map((row) => `<tr><td>${badge(row.source, row.source === 'OL' ? 'info' : 'neutral')}</td><td class="strong">${row.item}</td><td>${row.needMonth}</td><td class="number">${number(row.quantity)}</td><td class="number strong">${number(row.final)}</td><td>${badge(row.status, statusTone(row.status))}</td><td class="muted">${row.note}</td></tr>`);
  return `<div class="page-intro"><div><span class="eyebrow">STEP 02 · DEMAND CONFIRMATION</span><h2>수요 확인</h2><p>OL·SFDC·Bulk-deal·Trend를 비교하고 수급회의 기준으로 최종 수요를 확정합니다.</p></div><button class="primary-button" data-action="fake-complete">수요 검토 완료</button></div><div class="source-strip"><button class="data-card" data-detail="ol"><span>OL 제출 수량</span><strong>98 대</strong><small>상세 보기 ↗</small></button><button class="data-card" data-detail="sfdc"><span>SFDC 파이프라인</span><strong>28 대</strong><small>상세 보기 ↗</small></button><button class="data-card" data-detail="bulk"><span>Bulk-deal</span><strong>16 대</strong><small>상세 보기 ↗</small></button><button class="data-card" data-detail="finalDemand"><span>최종 확정 수요</span><strong>86 대</strong><small>상세 보기 ↗</small></button></div><div class="detail-slot"></div>${table(['출처','품목','필요월도','원천 수량','최종 확정','상태','비고'], rows)}`;
}

function renderInventory() {
  const rows = state.records.inventory.map((row) => `<tr><td class="strong">${row.item}</td><td>${badge(row.type, row.type === '기기' ? 'info' : 'neutral')}</td><td class="number">${number(row.stock)}</td><td class="number">${number(row.reserved)}</td><td class="number">${number(row.openPo)}</td><td>${row.eta}</td><td class="number strong">${number(row.available)}</td><td>${badge(row.flag, statusTone(row.flag))}</td></tr>`);
  return `<div class="page-intro"><div><span class="eyebrow">STEP 03 · STOCK & OPEN PO</span><h2>재고·Open PO</h2><p>전월말 재고와 미입고 발주를 확인해 발주 계산에 사용할 가용재고를 준비합니다.</p></div><button class="secondary-button" data-action="fake-complete">재고 확인 완료</button></div><div class="info-callout"><span class="callout-icon">i</span><div><strong>가용재고 계산 기준</strong><p>기말재고에서 예약·보류 수량을 차감하고, 필요월도 전 확정 입고 예정 Open PO를 반영합니다.</p></div><span class="prototype-label">2차 계산 예정</span></div>${table(['품목','구분','기말재고','예약/보류','Open PO','입고 예정','가용재고','상태'], rows)}`;
}

function renderMachine() {
  const rows = state.records.machine.map((row) => `<tr><td class="strong">${row.item}</td><td>${row.supplier}</td><td class="number">${number(row.demand)}</td><td class="number">${number(row.available)}</td><td class="number">${number(row.need)}</td><td class="number">${number(row.moq)}</td><td class="number strong">${number(row.order)}</td><td class="number">${money(row.amount)}</td><td>${badge(row.status, statusTone(row.status))}</td></tr>`);
  return `<div class="page-intro"><div><span class="eyebrow">STEP 04 · MACHINE ORDER</span><h2>기기 발주량</h2><p>수요와 가용재고를 비교해 기기별 발주 예정량을 확인합니다.</p></div><button class="primary-button" data-action="fake-complete">계산 결과 검토</button></div><div class="formula-card"><div><span class="eyebrow">CALCULATION LOGIC</span><strong>확정 수요 + 안전재고 - 가용재고 = 계산 필요량</strong><p>MOQ·Flex Rule을 적용한 최종 발주량은 2차 구현에서 자동 계산됩니다.</p></div><span class="logic-chip">MOQ · FLEX</span></div>${table(['기종','Supplier','확정 수요','가용재고','필요량','MOQ','발주 예정','금액','검토'], rows)}`;
}

function renderOption() {
  const rows = state.records.option.map((row) => `<tr><td class="strong">${row.item}</td><td>${row.model}</td><td class="number">${row.bom}</td><td class="number">${Math.round(row.attach * 100)}%</td><td class="number">${row.theory}</td><td class="number">${row.stock}</td><td class="number strong">${row.need}</td><td class="number">${row.order}</td><td>${row.common === 'Y' ? badge('Common','info') : '—'}</td><td>${row.mandatory === 'Y' ? badge('필수','warning') : '—'}</td><td>${badge(row.status, statusTone(row.status))}</td></tr>`);
  return `<div class="page-intro"><div><span class="eyebrow">STEP 05 · OPTION ORDER</span><h2>옵션 발주량</h2><p>BOM·장착율·평균 사용량을 기준으로 옵션 수요를 전개하고 Common품을 통합합니다.</p></div><button class="primary-button" data-action="fake-complete">옵션 결과 검토</button></div><div class="mini-kpis"><div><span>BOM 전개 품목</span><strong>12</strong></div><div><span>Common품</span><strong>4</strong></div><div><span>필수 옵션</span><strong>3</strong></div><div><span>경고</span><strong class="danger-text">1</strong></div></div>${table(['옵션','적용 기종','BOM','장착율','이론 수요','가용재고','필요량','발주 예정','구분','필수','검토'], rows)}`;
}

function renderReview() {
  const rows = state.records.exceptions.map((row) => `<tr><td>${badge(row.type, statusTone(row.status))}</td><td class="strong">${row.target}</td><td>${badge(row.severity === 'high' ? '높음' : row.severity === 'medium' ? '중간' : '낮음', row.severity === 'high' ? 'danger' : 'warning')}</td><td>${row.action}</td><td>${badge(row.status, statusTone(row.status))}</td><td><button class="table-action" data-action="resolve">조치 기록</button></td></tr>`);
  return `<div class="page-intro"><div><span class="eyebrow">STEP 06 · REVIEW & APPROVAL</span><h2>예외·검토</h2><p>자동 계산 결과의 예외와 수기 조정 사유를 검토합니다.</p></div><button class="primary-button" data-screen="report">보고 자료로 이동 →</button></div><div class="review-summary"><div><strong>${state.records.exceptions.length}</strong><span>전체 예외</span></div><div><strong class="danger-text">${state.records.exceptions.filter((r) => r.severity === 'high').length}</strong><span>높은 우선순위</span></div><div><strong>${state.records.exceptions.filter((r) => r.status === '조치 중').length}</strong><span>조치 중</span></div></div>${table(['예외 유형','대상','우선순위','권장 조치','상태','액션'], rows)}`;
}

function renderReport() {
  const m = getDashboardMetrics(state);
  return `<div class="page-intro"><div><span class="eyebrow">STEP 07 · REPORT & EXPORT</span><h2>보고·출력</h2><p>사장 보고용 요약과 System 업로드용 파일을 준비합니다.</p></div><div class="button-row"><button class="secondary-button" data-action="json">JSON 백업 ↓</button><button class="primary-button" data-action="export">System 업로드용 Excel ↓</button></div></div><div class="report-header"><div><span class="eyebrow">EXECUTIVE SUMMARY</span><h3>2026년 09월 발주 검토</h3><p>가상데이터 기준 미리보기 · 최종 확정 전</p></div>${badge('승인 전', 'warning')}</div><div class="report-grid"><div class="report-card"><span>총 발주 예정 금액</span><strong>${money(m.totalOrderAmount)}</strong><small>전월 대비 +8.4% <span class="up">↑</span></small></div><div class="report-card"><span>기기 발주 금액</span><strong>${money(m.machineAmount)}</strong><small>${number(m.machineQuantity)}대</small></div><div class="report-card"><span>옵션 발주 금액</span><strong>${money(m.optionAmount)}</strong><small>${number(m.optionQuantity)}개</small></div></div><div class="report-columns"><div class="panel"><div class="panel-heading"><div><span class="eyebrow">COMPARISON</span><h3>OL 제출 대비</h3></div></div>${table(['구분','전월 OL','최종 수요','발주 예정','차이'], [`<tr><td class="strong">기기</td><td>92</td><td>76</td><td>50</td><td class="number">-26</td></tr>`,`<tr><td class="strong">옵션</td><td>118</td><td>70</td><td>30</td><td class="number">-40</td></tr>`])}</div><div class="panel"><div class="panel-heading"><div><span class="eyebrow">DECISION NOTES</span><h3>주요 원인·리스크</h3></div></div><ul class="risk-list"><li><span class="exception-dot danger"></span>5570 기종 Flex 기준 미달 확인 필요</li><li><span class="exception-dot warning"></span>C3060 MOQ 적용으로 과잉 발주 가능</li><li><span class="exception-dot info"></span>SFDC 파이프라인 수급회의 반영 여부 확인</li></ul></div></div>`;
}

  function renderInventoryWithResidual() {
    const rows = state.records.inventory.map((row) => `<tr><td class="strong">${row.item}</td><td>${badge(row.type, row.type === '기기' ? 'info' : 'neutral')}</td><td class="number">${number(row.stock)}</td><td class="number">${number(row.reserved)}</td><td class="number">${number(row.residualOrder)}</td><td class="number">${number(row.openPo)}</td><td>${row.eta}</td><td class="number strong">${number(row.available)}</td><td>${badge(row.flag, statusTone(row.flag))}</td></tr>`);
    return `<div class="page-intro"><div><span class="eyebrow">STEP 03 · STOCK & OPEN PO</span><h2>재고·Open PO</h2><p>전월말 재고·Open PO와 전월도 잔여주문을 함께 확인해 발주 계산에 사용할 기준을 준비합니다.</p></div><button class="secondary-button" data-action="fake-complete">재고 확인 완료</button></div><div class="info-callout"><span class="callout-icon">i</span><div><strong>전월도 잔여주문 반영</strong><p>전월에 출고되지 못한 영업 주문은 재고가 아니라 이월 수요로 관리하고, 기기·옵션 발주 계산의 적용 수요에 더합니다.</p></div><span class="prototype-label">수요 반영</span></div>${table(['품목','구분','기말재고','예약/보류','전월도 잔여주문','Open PO','입고 예정','가용재고','상태'], rows)}`;
  }

  function renderMachineWithResidual() {
    const rows = state.records.machine.map((row) => `<tr><td class="strong">${row.item}</td><td>${row.supplier}</td><td class="number">${number(row.demand)}</td><td class="number">${number(row.residualOrder)}</td><td class="number">${number((row.demand || 0) + (row.residualOrder || 0))}</td><td class="number">${number(row.available)}</td><td class="number">${number(row.need)}</td><td class="number">${number(row.moq)}</td><td class="number strong">${number(row.order)}</td><td class="number">${money(row.amount)}</td><td>${badge(row.status, statusTone(row.status))}</td></tr>`);
    return `<div class="page-intro"><div><span class="eyebrow">STEP 04 · MACHINE ORDER</span><h2>기기 발주량</h2><p>당월 확정수요에 전월도 잔여주문을 더한 적용 수요와 가용재고를 비교해 발주량을 확인합니다.</p></div><button class="primary-button" data-action="fake-complete">계산 결과 검토</button></div><div class="formula-card"><div><span class="eyebrow">CALCULATION LOGIC</span><strong>확정 수요 + 전월도 잔여주문 + 안전재고 - 가용재고 = 계산 필요량</strong><p>잔여주문은 출고대기 상태로 남은 고객 수요이므로 재고가 아닌 추가 수요로 반영합니다.</p></div><span class="logic-chip">RESIDUAL · MOQ · FLEX</span></div>${table(['기종','Supplier','확정 수요','전월도 잔여주문','적용 수요','가용재고','필요량','MOQ','발주 예정','금액','검토'], rows)}`;
  }

  function renderOptionWithResidual() {
    const rows = state.records.option.map((row) => `<tr><td class="strong">${row.item}</td><td>${row.model}</td><td class="number">${row.bom}</td><td class="number">${Math.round(row.attach * 100)}%</td><td class="number">${row.theory}</td><td class="number">${number(row.residualOrder)}</td><td class="number">${number(row.stock)}</td><td class="number strong">${number(row.need)}</td><td class="number">${number(row.order)}</td><td>${row.common === 'Y' ? badge('Common','info') : '—'}</td><td>${row.mandatory === 'Y' ? badge('필수','warning') : '—'}</td><td>${badge(row.status, statusTone(row.status))}</td></tr>`);
    return `<div class="page-intro"><div><span class="eyebrow">STEP 05 · OPTION ORDER</span><h2>옵션 발주량</h2><p>기기 확정수요와 전월도 잔여주문을 합산한 적용 기기수요를 BOM·장착율에 전개합니다.</p></div><button class="primary-button" data-action="fake-complete">옵션 결과 검토</button></div><div class="mini-kpis"><div><span>BOM 전개 품목</span><strong>12</strong></div><div><span>Common품</span><strong>4</strong></div><div><span>필수 옵션</span><strong>3</strong></div><div><span>경고</span><strong class="danger-text">1</strong></div></div>${table(['옵션','적용 기종','BOM','장착율','이론 수요','전월도 잔여주문','가용재고','필요량','발주 예정','구분','필수','검토'], rows)}`;
  }

  function download(name, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = name; a.click(); URL.revokeObjectURL(url);
}

  function bindEvents() {
    document.querySelectorAll('[data-screen]').forEach((button) => button.addEventListener('click', () => { state.activeScreen = button.dataset.screen; saveState(); render(); }));
    document.querySelectorAll('[data-detail]').forEach((button) => button.addEventListener('click', () => {
      const slot = document.querySelector('.detail-slot');
      if (!slot) return;
      slot.innerHTML = detailPanel(button.dataset.detail);
      slot.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      slot.querySelector('[data-detail-close]')?.addEventListener('click', () => { slot.innerHTML = ''; });
    }));
    document.querySelectorAll('[data-action]').forEach((button) => button.addEventListener('click', () => {
    const action = button.dataset.action;
    if (action === 'demo') setState(createDemoState());
    if (action === 'empty') setState(createEmptyState());
    if (action === 'reset' && confirm('현재 작업공간을 초기화할까요?')) setState(createEmptyState());
    if (action === 'report') { state.activeScreen = 'report'; saveState(); render(); }
    if (action === 'fake-upload') { state.uploadHistory.unshift({ name: 'uploaded-workbook-preview.xlsx', status: '검증 예정', rows: 0, time: '방금 전' }); saveState(); render(); }
    if (action === 'fake-complete') { const current = state.workflowSteps.find((step) => step.id === state.activeScreen); if (current) current.status = 'completed'; saveState(); render(); }
    if (action === 'template') download('SCM_업로드_템플릿_프로토타입.csv', '시트,필수 컬럼,상태\n수요,기준월·품목코드·필요월도·수량,입력 예정\n재고,기준월·품목코드·기말재고,입력 예정\nOpen PO,PO번호·품목코드·미입고수량·입고예정월,입력 예정\n기준정보,품목코드·단가·MOQ·Lead Time,입력 예정\nBOM,기종코드·옵션코드·BOM수량·필수여부,입력 예정', 'text/csv;charset=utf-8');
    if (action === 'json') download(`${state.workspaceName}.json`, JSON.stringify(state, null, 2), 'application/json');
    if (action === 'export') download(`${state.workspaceName}_System업로드_프로토타입.csv`, '발주월,필요월도,품목코드,품목명,발주수량,단가,금액,승인상태\n2026-09,2026-09,DEMO-001,Apeos C7070,15,1500000,22500000,검토 예정', 'text/csv;charset=utf-8');
    if (action === 'resolve') alert('2차 구현에서 예외 조정 사유와 승인 이력을 저장할 수 있습니다.');
  }));
}

render();
