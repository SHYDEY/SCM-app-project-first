import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const inputPath = 'C:/Users/fujifilm/Downloads/2차시_프로세스분석_워크숍시트.xlsx';
const blob = await FileBlob.load(inputPath);
const wb = await SpreadsheetFile.importXlsx(blob);
const summary = await wb.inspect({kind:'workbook,sheet,table', maxChars:12000, tableMaxRows:20, tableMaxCols:20, tableMaxCellChars:200});
console.log(summary.ndjson);
const sheets = await wb.inspect({kind:'sheet', include:'id,name'});
console.log('SHEETS\n'+sheets.ndjson);
for (const item of wb.worksheets.items) {
  const used = item.getUsedRange();
  if (!used) continue;
  const address = used.address;
  console.log(`RANGE ${item.name} ${address}`);
  const region = await wb.inspect({kind:'region', sheetId:item.name, range:address, maxChars:18000, tableMaxRows:100, tableMaxCols:30, tableMaxCellChars:300});
  console.log(region.ndjson);
  const preview = await wb.render({sheetName:item.name, autoCrop:'all', scale:1, format:'png'});
  const safe = item.name.replace(/[\\/:*?"<>|]/g,'_');
  await fs.writeFile(`inspect_${safe}.png`, new Uint8Array(await preview.arrayBuffer()));
}
