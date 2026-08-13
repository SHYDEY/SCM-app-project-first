import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const inputPath = 'C:/Users/fujifilm/Documents/ChatGPT/SCM 수급 개선 프로젝트/SCM_월간_발주_업무용_템플릿.xlsx';
const outputDir = 'outputs/SCM_2025_더미데이터';
const outputPath = `${outputDir}/SCM_2025_1년치_더미데이터.xlsx`;
await fs.mkdir(outputDir, { recursive: true });

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const sheet = (name) => wb.worksheets.getItem(name);
const months = Array.from({ length: 12 }, (_, i) => new Date(2025, i, 1));

const machines = [
  ['기기', 'M-C7070', 'Apeos C7070', 'A3 컬러 복합기', '후지필름BI', 1500000, 'KRW', 5, 1, 45, 0.8, 1.2, '정상', new Date(2025, 11, 31)],
  ['기기', 'M-5570', 'ApeosPort 5570', 'A3 컬러 복합기', '후지필름BI', 1100000, 'KRW', 5, 1, 45, 0.8, 1.2, '정상', new Date(2025, 11, 31)],
  ['기기', 'M-C3060', 'Apeos C3060', 'A3 컬러 복합기', '후지필름BI', 850000, 'KRW', 10, 1, 40, 0.75, 1.15, '정상', new Date(2025, 11, 31)],
  ['기기', 'M-C3070', 'Apeos C3070', 'A3 컬러 복합기', '후지필름BI', 780000, 'KRW', 5, 1, 40, 0.75, 1.15, '정상', new Date(2025, 11, 31)],
  ['기기', 'M-C2570', 'Apeos C2570', 'A3 컬러 복합기', '후지필름BI', 650000, 'KRW', 5, 1, 35, 0.8, 1.2, '정상', new Date(2025, 11, 31)],
  ['기기', 'M-C2060', 'Apeos C2060', 'A3 컬러 복합기', '후지필름BI', 520000, 'KRW', 5, 1, 35, 0.8, 1.2, '정상', new Date(2025, 11, 31)],
  ['기기', 'M-5330', 'Apeos 5330', 'A4 복합기', '후지필름BI', 420000, 'KRW', 5, 1, 30, 0.8, 1.2, '정상', new Date(2025, 11, 31)],
  ['기기', 'M-3060', 'Apeos 3060', 'A4 복합기', '후지필름BI', 360000, 'KRW', 10, 1, 30, 0.75, 1.15, '정상', new Date(2025, 11, 31)],
];

const options = [
  ['옵션', 'OP-CT202480', 'CT202480 블랙 토너', '소모품', '후지필름BI', 120000, 'KRW', 10, 1, 30, 0.8, 1.2, '정상', new Date(2025, 11, 31)],
  ['옵션', 'OP-EL300123', 'EL300123 피니셔', '피니셔', '후지필름BI', 450000, 'KRW', 5, 1, 45, 0.8, 1.2, '정상', new Date(2025, 11, 31)],
  ['옵션', 'OP-OF300456', 'OF300456 오피스 피니셔', '피니셔', '후지필름BI', 380000, 'KRW', 5, 1, 45, 0.8, 1.2, '정상', new Date(2025, 11, 31)],
  ['옵션', 'OP-WC300789', 'WC300789 무선 키트', '통신 옵션', '후지필름BI', 180000, 'KRW', 10, 1, 35, 0.8, 1.2, '정상', new Date(2025, 11, 31)],
  ['옵션', 'OP-TR400111', 'TR400111 추가 트레이', '급지 옵션', '후지필름BI', 250000, 'KRW', 5, 1, 40, 0.8, 1.2, '정상', new Date(2025, 11, 31)],
  ['옵션', 'OP-TN5570', 'TN5570 블랙 토너', '소모품', '후지필름BI', 95000, 'KRW', 10, 1, 30, 0.8, 1.2, '정상', new Date(2025, 11, 31)],
  ['옵션', 'OP-ST2020', 'ST2020 스테이플', '소모품', '후지필름BI', 15000, 'KRW', 20, 1, 25, 0.8, 1.2, '정상', new Date(2025, 11, 31)],
  ['옵션', 'OP-FN3060', 'FN3060 피니셔', '피니셔', '후지필름BI', 320000, 'KRW', 5, 1, 45, 0.8, 1.2, '정상', new Date(2025, 11, 31)],
];

const allItems = [...machines, ...options];
const quantity = (base, monthIndex, itemIndex) => Math.max(2, Math.round(base * (1 + ((monthIndex % 4) - 1.5) * 0.08) + ((itemIndex * 3 + monthIndex) % 5)));
const residualHeaderFormat = { fill: '#2E74B5', font: { bold: true, color: '#FFFFFF', size: 10 }, horizontalAlignment: 'center', verticalAlignment: 'center', wrapText: true, borders: { preset: 'all', style: 'thin', color: '#D9E2F3' } };
const residualInputFormat = { fill: '#FFF4CC', font: { color: '#1F2937', size: 10 }, horizontalAlignment: 'right', borders: { preset: 'all', style: 'thin', color: '#D9E2F3' } };
const residualCalcFormat = { fill: '#EAF4EA', font: { color: '#1F2937', size: 10 }, horizontalAlignment: 'right', borders: { preset: 'all', style: 'thin', color: '#D9E2F3' } };

// 07 기준정보
const master = sheet('07_기준정보');
master.getRange('A5:N20').values = allItems;
master.getRange('F5:F20').format.numberFormat = '#,##0';
master.getRange('H5:J20').format.numberFormat = '#,##0';
master.getRange('K5:L20').format.numberFormat = '0.0%';
master.getRange('N5:N20').format.numberFormat = 'yyyy-mm-dd';

// 03 수요입력: 월별 기기 8종의 OL/SFDC/Bulk-deal/Trend를 최종 수요로 확정
const demandRows = [];
const sourceTypes = ['OL', 'SFDC', 'Bulk-deal', 'Trend'];
for (let m = 0; m < 12; m += 1) {
  for (let i = 0; i < machines.length; i += 1) {
    const finalQty = quantity(24 - i * 2, m, i);
    const confidence = i % 4 === 1 ? 0.75 : i % 3 === 0 ? 0.95 : 0.85;
    const source = sourceTypes[(m + i) % sourceTypes.length];
    const olQty = Math.max(0, finalQty + ((m + i) % 5) - 2);
    const adjustment = finalQty - olQty;
    demandRows.push([months[m], '기기', ['직판', '대리점', '렌탈'][i % 3], machines[i][1], machines[i][2], new Date(2025, Math.min(11, m + 1), 1), olQty, confidence, finalQty, m % 7 === 0 ? 'N' : '', source === 'Bulk-deal' ? 'Y' : 'N', adjustment, null, source, ['김영업', '이수급', '박기획'][i % 3], '확정', adjustment === 0 ? '변경 없음' : '수급회의 조정 반영']);
  }
}
const demand = sheet('03_수요입력');
demand.getRange('A5:Q204').clear({ contentsOnly: true });
demand.getRange(`A5:Q${4 + demandRows.length}`).values = demandRows;
demand.getRange(`A5:A${4 + demandRows.length}`).format.numberFormat = 'yyyy-mm';
demand.getRange(`F5:F${4 + demandRows.length}`).format.numberFormat = 'yyyy-mm';
demand.getRange(`G5:G${4 + demandRows.length}`).format.numberFormat = '#,##0';
demand.getRange(`H5:H${4 + demandRows.length}`).format.numberFormat = '0%';
demand.getRange(`I5:M${4 + demandRows.length}`).format.numberFormat = '#,##0';
demand.getRange('M5').formulas = [['=IF(A5="","",I5+L5)']];
demand.getRange(`M5:M${4 + demandRows.length}`).fillDown();

// 04 재고_OpenPO: 월별 기기/옵션 전월말 재고와 입고 예정
const inventoryRows = [];
for (let m = 0; m < 12; m += 1) {
  for (let i = 0; i < allItems.length; i += 1) {
    const item = allItems[i];
    const stock = item[0] === '기기' ? 18 + ((i * 4 + m * 2) % 18) : 8 + ((i * 5 + m * 3) % 15);
    const reserved = item[0] === '기기' ? 2 + ((m + i) % 4) : 4 + ((m + i) % 7);
    const qualityHold = (m + i) % 11 === 0 ? 2 : 0;
    const outbound = 1 + ((m * 2 + i) % 4);
    const openPo = 4 + ((m + i * 2) % 12);
    inventoryRows.push([months[m], item[0], item[1], item[2], item[4], m % 2 === 0 ? '서울DC' : '부산DC', stock, reserved, qualityHold, outbound, null, openPo, Math.round(openPo * 0.7), new Date(2025, Math.min(11, m + 1), 15), null, qualityHold ? '품질보류 재고 포함' : '정상']);
  }
}
const inventory = sheet('04_재고_OpenPO');
inventory.getRange('Q4').values = [['전월도 잔여주문']];
inventory.getRange('Q4').format = residualHeaderFormat;
inventory.getRange('A5:Q204').clear({ contentsOnly: true });
inventory.getRange(`A5:Q${4 + inventoryRows.length}`).values = inventoryRows.map((row, index) => {
  const itemIndex = index % allItems.length;
  const monthIndex = Math.floor(index / allItems.length);
  const item = allItems[itemIndex];
  const residual = monthIndex === 0 ? (item[0] === '기기' ? 3 + (itemIndex % 3) : 2 + (itemIndex % 4)) : ((monthIndex + itemIndex) % 7 === 0 ? 2 + (itemIndex % 3) : 0);
  return [...row, residual];
});
inventory.getRange(`A5:A${4 + inventoryRows.length}`).format.numberFormat = 'yyyy-mm';
inventory.getRange(`N5:N${4 + inventoryRows.length}`).format.numberFormat = 'yyyy-mm';
inventory.getRange(`G5:M${4 + inventoryRows.length}`).format.numberFormat = '#,##0';
inventory.getRange(`O5:O${4 + inventoryRows.length}`).format.numberFormat = '#,##0';
inventory.getRange(`Q5:Q${4 + inventoryRows.length}`).format.numberFormat = '#,##0';
inventory.getRange(`Q5:Q${4 + inventoryRows.length}`).format = residualInputFormat;
inventory.getRange('Q:Q').format.columnWidth = 16;
inventory.getRange('K5').formulas = [['=IF(C5="","",MAX(0,G5-H5-I5-J5))']];
inventory.getRange(`K5:K${4 + inventoryRows.length}`).fillDown();
inventory.getRange('O5').formulas = [['=IF(C5="","",K5+M5)']];
inventory.getRange(`O5:O${4 + inventoryRows.length}`).fillDown();

// 05 기기발주계산: 월별 기기 수요와 재고를 입력하고 템플릿 수식은 유지
const machineCalcRows = [];
for (let m = 0; m < 12; m += 1) {
  for (let i = 0; i < machines.length; i += 1) {
    const item = machines[i];
    const demandQty = quantity(24 - i * 2, m, i);
    const safety = 5 + (i % 4);
    const stock = 18 + ((i * 4 + m * 2) % 18);
    const reserved = 2 + ((m + i) % 4);
    const inbound = 3 + ((m + i) % 6);
    machineCalcRows.push([months[m], new Date(2025, Math.min(11, m + 1), 1), item[1], item[2], item[4], demandQty, safety, stock, reserved, inbound, null, null, item[7], null, null, null, null, null, (m + i) % 9 === 0 ? '예외 확인' : '검토 예정']);
  }
}
const mach = sheet('05_기기발주계산');
mach.getRange('T4:U4').values = [['전월도 잔여주문', '적용 수요']];
mach.getRange('T4:U4').format = residualHeaderFormat;
mach.getRange('A5:U204').clear({ contentsOnly: true });
mach.getRange(`A5:U${4 + machineCalcRows.length}`).values = machineCalcRows.map((row, index) => {
  const itemIndex = index % machines.length;
  const monthIndex = Math.floor(index / machines.length);
  const residual = monthIndex === 0 ? 3 + (itemIndex % 3) : ((monthIndex + itemIndex) % 7 === 0 ? 2 + (itemIndex % 3) : 0);
  return [...row, residual, null];
});
mach.getRange(`A5:B${4 + machineCalcRows.length}`).format.numberFormat = 'yyyy-mm';
mach.getRange(`F5:O${4 + machineCalcRows.length}`).format.numberFormat = '#,##0';
mach.getRange(`Q5:R${4 + machineCalcRows.length}`).format.numberFormat = '#,##0';
mach.getRange(`T5:U${4 + machineCalcRows.length}`).format.numberFormat = '#,##0';
mach.getRange(`T5:T${4 + machineCalcRows.length}`).format = residualInputFormat;
mach.getRange(`U5:U${4 + machineCalcRows.length}`).format = residualCalcFormat;
mach.getRange('T:T').format.columnWidth = 16;
mach.getRange('U:U').format.columnWidth = 14;
mach.getRange('U5').formulas = [['=IF(C5="","",F5+T5)']];
mach.getRange(`U5:U${4 + machineCalcRows.length}`).fillDown();
mach.getRange('K5').formulas = [['=IF(C5="","",MAX(0,H5-I5+J5))']];
mach.getRange(`K5:K${4 + machineCalcRows.length}`).fillDown();
mach.getRange('L5').formulas = [['=IF(C5="","",MAX(0,U5+G5-K5))']];
mach.getRange(`L5:L${4 + machineCalcRows.length}`).fillDown();
mach.getRange('N5').formulas = [['=IF(C5="","",U5*(1-0.2))']];
mach.getRange(`N5:N${4 + machineCalcRows.length}`).fillDown();
mach.getRange('O5').formulas = [['=IF(C5="","",U5*(1+0.2))']];
mach.getRange(`O5:O${4 + machineCalcRows.length}`).fillDown();
mach.getRange('P5').formulas = [['=IF(C5="","",IF(L5=0,"정상",IF(L5<N5,"Flex 미달",IF(L5>O5,"Flex 초과","정상"))))']];
mach.getRange(`P5:P${4 + machineCalcRows.length}`).fillDown();
mach.getRange('Q5').formulas = [['=IF(C5="","",IF(L5=0,0,ROUNDUP(L5/M5,0)*M5))']];
mach.getRange(`Q5:Q${4 + machineCalcRows.length}`).fillDown();
mach.getRange('R5').formulas = [["=IF(C5=\"\",\"\",Q5*SUMIF('07_기준정보'!$B$5:$B$204,C5,'07_기준정보'!$F$5:$F$204))"]];
mach.getRange(`R5:R${4 + machineCalcRows.length}`).fillDown();

// 06 옵션발주계산: BOM·장착율·평균사용량·필수 여부를 월별 입력
const bomByOption = [1, 1, 1, 1, 1, 1, 2, 1];
const attachByOption = [0.65, 0.35, 0.28, 0.55, 0.22, 0.75, 0.9, 0.3];
const machineForOption = ['M-C7070', 'M-5570', 'M-C3060', 'M-C3070', 'M-C2570', 'M-C2060', 'M-5330', 'M-3060'];
const optionRows = [];
for (let m = 0; m < 12; m += 1) {
  for (let i = 0; i < options.length; i += 1) {
    const option = options[i];
    const machineIndex = machines.findIndex((item) => item[1] === machineForOption[i]);
    const machineDemand = quantity(24 - machineIndex * 2, m, machineIndex);
    const stock = 8 + ((i * 5 + m * 3) % 15);
    const safety = 8 + (i % 5);
    optionRows.push([months[m], new Date(2025, Math.min(11, m + 1), 1), option[1], option[2], machineForOption[i], machineDemand, bomByOption[i], attachByOption[i], 1, i % 4 === 0 ? 3 : 0, null, stock, safety, null, option[7], null, ['OP-EL300123', 'OP-OF300456', 'OP-ST2020'].includes(option[1]) ? 'Y' : 'N', null, null, option[5], null]);
  }
}
const opt = sheet('06_옵션발주계산');
opt.getRange('V4:W4').values = [['전월도 잔여주문(기기)', '적용 기기수요']];
opt.getRange('V4:W4').format = residualHeaderFormat;
opt.getRange('A5:W204').clear({ contentsOnly: true });
opt.getRange(`A5:W${4 + optionRows.length}`).values = optionRows.map((row, index) => {
  const optionIndex = index % options.length;
  const monthIndex = Math.floor(index / options.length);
  const residual = monthIndex === 0 ? 2 + (optionIndex % 3) : ((monthIndex + optionIndex) % 6 === 0 ? 2 : 0);
  return [...row, residual, null];
});
opt.getRange(`A5:B${4 + optionRows.length}`).format.numberFormat = 'yyyy-mm';
opt.getRange(`F5:G${4 + optionRows.length}`).format.numberFormat = '#,##0';
opt.getRange(`H5:H${4 + optionRows.length}`).format.numberFormat = '0%';
opt.getRange(`I5:O${4 + optionRows.length}`).format.numberFormat = '#,##0.0';
opt.getRange(`S5:U${4 + optionRows.length}`).format.numberFormat = '#,##0';
opt.getRange(`V5:W${4 + optionRows.length}`).format.numberFormat = '#,##0';
opt.getRange(`V5:V${4 + optionRows.length}`).format = residualInputFormat;
opt.getRange(`W5:W${4 + optionRows.length}`).format = residualCalcFormat;
opt.getRange('V:V').format.columnWidth = 18;
opt.getRange('W:W').format.columnWidth = 14;
opt.getRange('W5').formulas = [['=IF(C5="","",F5+V5)']];
opt.getRange(`W5:W${4 + optionRows.length}`).fillDown();
opt.getRange('K5').formulas = [['=IF(C5="","",W5*G5*H5*I5+J5)']];
opt.getRange(`K5:K${4 + optionRows.length}`).fillDown();
opt.getRange('N5').formulas = [['=IF(C5="","",MAX(0,K5+M5-L5))']];
opt.getRange(`N5:N${4 + optionRows.length}`).fillDown();
opt.getRange('P5').formulas = [['=IF(C5="","",IF(N5=0,"정상",IF(N5>O5*10,"MOQ 과잉 확인","정상")))']];
opt.getRange(`P5:P${4 + optionRows.length}`).fillDown();
opt.getRange('R5').formulas = [['=IF(C5="","",IF(Q5="Y",IF(K5>0,"정상","누락"),"정상"))']];
opt.getRange(`R5:R${4 + optionRows.length}`).fillDown();
opt.getRange('S5').formulas = [['=IF(C5="","",IF(N5=0,0,ROUNDUP(N5/O5,0)*O5))']];
opt.getRange(`S5:S${4 + optionRows.length}`).fillDown();
opt.getRange('U5').formulas = [['=IF(C5="","",S5*T5)']];
opt.getRange(`U5:U${4 + optionRows.length}`).fillDown();

// 08 예외·승인: 실제 흐름에서 검토가 필요한 대표 케이스를 월별로 기록
const exceptionRows = [];
for (let m = 0; m < 12; m += 1) {
  exceptionRows.push([months[m], '기기 발주량', machines[(m + 1) % machines.length][1], 'Flex 하한 미달', 12 + m, 10 + m, m % 3 === 0 ? '승인 완료' : '조치 완료', 'SFDC 확도 반영 후 수급회의에서 보수적으로 조정', '이수급', '김팀장']);
  if (m % 2 === 0) exceptionRows.push([months[m], '옵션 발주량', options[2][1], 'MOQ 과잉 가능', 7 + m, 5 + m, m % 4 === 0 ? '승인 완료' : '조치 완료', 'Common품 재고와 다음달 수요를 함께 검토', '박기획', '김팀장']);
}
const exceptions = sheet('08_예외_승인');
exceptions.getRange('A5:J204').clear({ contentsOnly: true });
exceptions.getRange(`A5:J${4 + exceptionRows.length}`).values = exceptionRows;
exceptions.getRange(`A5:A${4 + exceptionRows.length}`).format.numberFormat = 'yyyy-mm';
exceptions.getRange(`E5:F${4 + exceptionRows.length}`).format.numberFormat = '#,##0';

// 대시보드와 보고자료 기준월을 2025년 12월로 설정
sheet('02_발주대시보드').getRange('B4').values = [[new Date(2025, 11, 1)]];
sheet('02_발주대시보드').getRange('B4').format.numberFormat = 'yyyy-mm';
sheet('02_발주대시보드').getRange('B5').values = [['검토 완료']];
sheet('09_보고자료').getRange('B4').values = [[new Date(2025, 11, 1)]];
sheet('09_보고자료').getRange('B4').format.numberFormat = 'yyyy-mm';
sheet('09_보고자료').getRange('B5').values = [['작성 완료']];

const previewSheets = ['03_수요입력', '04_재고_OpenPO', '05_기기발주계산', '06_옵션발주계산', '07_기준정보', '08_예외_승인', '09_보고자료'];
for (const name of previewSheets) {
  const blob = await wb.render({ sheetName: name, range: name === '09_보고자료' ? 'A1:F19' : 'A1:U18', scale: 1, format: 'png' });
  await fs.writeFile(`${outputDir}/${name}.png`, new Uint8Array(await blob.arrayBuffer()));
}

const check = await wb.inspect({ kind: 'table', sheetId: '03_수요입력', range: 'A4:Q10', include: 'values,formulas', tableMaxRows: 10, tableMaxCols: 20, maxChars: 6000 });
console.log(check.ndjson);
const calcCheck = await wb.inspect({ kind: 'table', sheetId: '05_기기발주계산', range: 'A4:S10', include: 'values,formulas', tableMaxRows: 10, tableMaxCols: 20, maxChars: 6000 });
console.log(calcCheck.ndjson);
const errors = await wb.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A', options: { useRegex: true, maxResults: 200 }, summary: '2025 dummy data formula error scan' });
console.log(errors.ndjson);

const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);
console.log(outputPath);
