import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const p = 'outputs/SCM_월간_발주_업무용_템플릿/SCM_월간_발주_업무용_템플릿.xlsx';
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(p));
const sheets = await wb.inspect({kind:'sheet',include:'id,name'});
console.log(sheets.ndjson);
const calc = await wb.inspect({kind:'table',sheetId:'05_기기발주계산',range:'A4:S8',include:'values,formulas',tableMaxRows:5,tableMaxCols:20,maxChars:8000});
console.log(calc.ndjson);
const errors = await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A',options:{useRegex:true,maxResults:100},summary:'final formula error scan'});
console.log(errors.ndjson);
