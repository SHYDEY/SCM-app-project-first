export const workflowSteps = [
  { id: 'collect', label: 'OL 취합', short: '01', status: 'completed' },
  { id: 'demand', label: '수요 확인', short: '02', status: 'in-progress' },
  { id: 'inventory', label: '재고·Open PO', short: '03', status: 'in-progress' },
  { id: 'machine', label: '기기 발주', short: '04', status: 'not-started' },
  { id: 'option', label: '옵션 발주', short: '05', status: 'not-started' },
  { id: 'review', label: '보고·검토', short: '06', status: 'not-started' },
  { id: 'export', label: 'System 출력', short: '07', status: 'not-started' },
];

export const screenRegistry = [
  { id: 'dashboard', label: '대시보드', icon: '⌂' },
  { id: 'upload', label: '데이터 업로드', icon: '↑' },
  { id: 'demand', label: '수요 확인', icon: '◎' },
  { id: 'inventory', label: '재고·Open PO', icon: '▣' },
  { id: 'machine', label: '기기 발주', icon: '▤' },
  { id: 'option', label: '옵션 발주', icon: '◇' },
  { id: 'review', label: '예외·검토', icon: '!' },
  { id: 'report', label: '보고·출력', icon: '↗' },
];

const demoRecords = {
  demand: [
    { source: 'OL', item: 'Apeos C7070', needMonth: '2026-09', quantity: 42, final: 40, status: '확정', note: '전월 제출 OL 조정' },
    { source: 'SFDC', item: 'ApeosPort 5570', needMonth: '2026-09', quantity: 28, final: 20, status: '검토', note: '중요 파이프라인' },
    { source: 'Bulk-deal', item: 'Apeos C3060', needMonth: '2026-10', quantity: 16, final: 16, status: '확정', note: '대형 프로젝트' },
    { source: 'Trend', item: 'Apeos C3070', needMonth: '2026-10', quantity: 12, final: 10, status: '참고', note: '최근 12개월 추세' },
  ],
  inventory: [
    { item: 'Apeos C7070', type: '기기', stock: 18, reserved: 4, openPo: 12, eta: '2026-09', available: 26, flag: '정상' },
    { item: 'ApeosPort 5570', type: '기기', stock: 6, reserved: 2, openPo: 0, eta: '-', available: 4, flag: '부족 예상' },
    { item: 'CT202480', type: '옵션', stock: 65, reserved: 12, openPo: 40, eta: '2026-09', available: 93, flag: '정상' },
    { item: 'EL300123', type: '옵션', stock: 22, reserved: 8, openPo: 0, eta: '-', available: 14, flag: 'MOQ 확인' },
  ],
  machine: [
    { item: 'Apeos C7070', supplier: '상해', demand: 40, available: 26, need: 14, moq: 5, order: 15, amount: 22500000, status: '정상' },
    { item: 'ApeosPort 5570', supplier: '도쿄', demand: 20, available: 4, need: 16, moq: 5, order: 15, amount: 18000000, status: 'Flex 미달' },
    { item: 'Apeos C3060', supplier: '심천', demand: 16, available: 3, need: 13, moq: 10, order: 20, amount: 12000000, status: 'MOQ 과잉 확인' },
  ],
  option: [
    { item: 'CT202480', model: 'C7070', bom: 1, attach: 0.82, theory: 33, stock: 93, need: 0, order: 0, mandatory: 'N', common: 'Y', status: '정상' },
    { item: 'EL300123', model: '5570', bom: 1, attach: 1, theory: 20, stock: 14, need: 6, order: 10, mandatory: 'Y', common: 'N', status: 'MOQ 과잉 확인' },
    { item: 'OF300456', model: 'C3060/3070', bom: 1, attach: 0.65, theory: 17, stock: 4, need: 13, order: 20, mandatory: 'N', common: 'Y', status: '정상' },
  ],
  exceptions: [
    { type: 'Flex 미달', target: 'ApeosPort 5570', severity: 'high', action: 'Supplier Flex 기준 확인 필요', status: '검토 필요' },
    { type: 'MOQ 과잉', target: 'Apeos C3060 / EL300123', severity: 'medium', action: '과잉 금액 확인 후 승인', status: '검토 필요' },
    { type: '데이터 누락', target: 'SFDC 수요 8건', severity: 'low', action: '수급회의 반영 여부 확인', status: '조치 중' },
  ],
};

demoRecords.inventory.forEach((row, index) => { row.residualOrder = [3, 4, 2, 3][index] || 0; });
demoRecords.machine.forEach((row, index) => { row.residualOrder = [3, 4, 2][index] || 0; });
demoRecords.option.forEach((row, index) => { row.residualOrder = [2, 3, 2][index] || 0; });

function baseState(mode) {
  return {
    mode,
    workspaceName: mode === 'demo' ? '2026년 09월 발주 검토 데모' : '새 발주 작업공간',
    orderMonth: '2026-09',
    activeScreen: 'dashboard',
    workflowSteps: workflowSteps.map((step) => ({ ...step })),
    records: mode === 'demo' ? structuredClone(demoRecords) : { demand: [], inventory: [], machine: [], option: [], exceptions: [] },
    uploadHistory: mode === 'demo' ? [{ name: 'demo-sample-workbook.xlsx', status: '가상데이터', rows: 68, time: '방금 전' }] : [],
    lastSavedAt: null,
  };
}

export function createDemoState() {
  return baseState('demo');
}

export function createEmptyState() {
  const state = baseState('empty');
  state.workflowSteps = workflowSteps.map((step) => ({ ...step, status: 'not-started' }));
  return state;
}

export function getWorkflowSteps() {
  return workflowSteps.map((step) => ({ ...step }));
}

export function getWorkflowScreens() {
  return screenRegistry.map((screen) => ({ ...screen }));
}

export function getDashboardMetrics(state) {
  const machineAmount = state.records.machine.reduce((sum, row) => sum + (row.amount || 0), 0);
  const optionAmount = state.records.option.reduce((sum, row) => sum + ((row.order || 0) * 750000), 0);
  return {
    totalOrderAmount: machineAmount + optionAmount,
    machineAmount,
    optionAmount,
    machineQuantity: state.records.machine.reduce((sum, row) => sum + (row.order || 0), 0),
    optionQuantity: state.records.option.reduce((sum, row) => sum + (row.order || 0), 0),
    exceptions: state.records.exceptions.filter((row) => row.status === '검토 필요').length,
    flexExceeded: state.records.machine.filter((row) => row.status === 'Flex 미달').length,
    mandatoryMissing: state.records.option.filter((row) => row.mandatory === 'Y' && row.need > 0).length,
  };
}
