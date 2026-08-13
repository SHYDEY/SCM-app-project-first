import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const inputPath = 'C:/Users/fujifilm/Documents/ChatGPT/SCM 수급 개선 프로젝트/SCM_월간_발주_업무용_템플릿.xlsx';
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
for (const sheet of wb.worksheets.items) {
  const used = sheet.getUsedRange();
  if (!used) continue;
  const result = await wb.inspect({ kind: 'table', sheetId: sheet.name, range: used.address, include: 'values,formulas', tableMaxRows: 6, tableMaxCols: 25, maxChars: 12000 });
  console.log(`SHEET=${sheet.name}\n${result.ndjson}`);
}
