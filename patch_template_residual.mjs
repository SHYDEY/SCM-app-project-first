import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const inputPath = 'C:/Users/fujifilm/Documents/ChatGPT/SCM 수급 개선 프로젝트/SCM_월간_발주_업무용_템플릿.xlsx';
const outputPath = 'outputs/SCM_월간_발주_업무용_템플릿_v2_잔여주문반영/SCM_월간_발주_업무용_템플릿_v2_잔여주문반영.xlsx';
await fs.mkdir('outputs/SCM_월간_발주_업무용_템플릿_v2_잔여주문반영', { recursive: true });
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const blue = '#2E74B5';
const white = '#FFFFFF';
const inputFill = '#FFF4CC';
const calcFill = '#EAF4EA';
const border = '#D9E2F3';
const headerFormat = { fill: blue, font: { bold: true, color: white, size: 10 }, horizontalAlignment: 'center', verticalAlignment: 'center', wrapText: true, borders: { preset: 'all', style: 'thin', color: border } };

const inventory = wb.worksheets.getItem('04_재고_OpenPO');
inventory.getRange('Q4').values = [['전월도 잔여주문']];
inventory.getRange('Q4').format = headerFormat;
inventory.getRange('Q5:Q204').format = { fill: inputFill, font: { color: '#1F2937', size: 10 }, horizontalAlignment: 'right', borders: { preset: 'all', style: 'thin', color: border } };
inventory.getRange('Q5:Q204').format.numberFormat = '#,##0';
inventory.getRange('Q:Q').format.columnWidth = 16;

const machine = wb.worksheets.getItem('05_기기발주계산');
machine.getRange('T4:U4').values = [['전월도 잔여주문', '적용 수요']];
machine.getRange('T4:U4').format = headerFormat;
machine.getRange('T5:T204').format = { fill: inputFill, font: { color: '#1F2937', size: 10 }, horizontalAlignment: 'right', borders: { preset: 'all', style: 'thin', color: border } };
machine.getRange('U5:U204').format = { fill: calcFill, font: { color: '#1F2937', size: 10 }, horizontalAlignment: 'right', borders: { preset: 'all', style: 'thin', color: border } };
machine.getRange('T5:T204').format.numberFormat = '#,##0';
machine.getRange('U5:U204').format.numberFormat = '#,##0';
machine.getRange('U5').formulas = [['=IF(C5="","",F5+T5)']];
machine.getRange('U5:U204').fillDown();
machine.getRange('T:T').format.columnWidth = 16;
machine.getRange('U:U').format.columnWidth = 14;

const option = wb.worksheets.getItem('06_옵션발주계산');
option.getRange('V4:W4').values = [['전월도 잔여주문(기기)', '적용 기기수요']];
option.getRange('V4:W4').format = headerFormat;
option.getRange('V5:V204').format = { fill: inputFill, font: { color: '#1F2937', size: 10 }, horizontalAlignment: 'right', borders: { preset: 'all', style: 'thin', color: border } };
option.getRange('W5:W204').format = { fill: calcFill, font: { color: '#1F2937', size: 10 }, horizontalAlignment: 'right', borders: { preset: 'all', style: 'thin', color: border } };
option.getRange('V5:W204').format.numberFormat = '#,##0';
option.getRange('W5').formulas = [['=IF(C5="","",F5+V5)']];
option.getRange('W5:W204').fillDown();
option.getRange('V:V').format.columnWidth = 18;
option.getRange('W:W').format.columnWidth = 14;

const previews = [
  ['04_재고_OpenPO', 'A1:Q16'],
  ['05_기기발주계산', 'A1:U16'],
  ['06_옵션발주계산', 'A1:W16'],
];
for (const [sheetName, range] of previews) {
  const blob = await wb.render({ sheetName, range, scale: 1, format: 'png' });
  const safe = sheetName.replace(/[\\/:*?"<>|]/g, '_');
  await fs.writeFile(`outputs/SCM_월간_발주_업무용_템플릿_v2_잔여주문반영/${safe}.png`, new Uint8Array(await blob.arrayBuffer()));
}
const errors = await wb.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A', options: { useRegex: true, maxResults: 100 }, summary: 'template residual order formula error scan' });
console.log(errors.ndjson);
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);
console.log(outputPath);
