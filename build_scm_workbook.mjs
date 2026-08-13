import fs from 'node:fs/promises';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const outDir = 'outputs/SCM_월간_발주_업무용_템플릿';
await fs.mkdir(outDir, { recursive: true });
const wb = Workbook.create();

const colors = {
  navy: '#1F4D78', blue: '#2E74B5', lightBlue: '#E8EEF5', lightGray: '#F2F4F7',
  input: '#FFF4CC', calc: '#EAF4EA', warn: '#FCE8E6', white: '#FFFFFF', text: '#1F2937',
  border: '#D9E2F3', muted: '#666666'
};

function title(sheet, text, subtitle='') {
  sheet.mergeCells('A1:J1');
  sheet.getRange('A1').values = [[text]];
  sheet.getRange('A1:J1').format = { fill: colors.navy, font: { bold: true, color: colors.white, size: 16 }, horizontalAlignment: 'left', verticalAlignment: 'center' };
  sheet.getRange('A1:J1').format.rowHeight = 28;
  if (subtitle) {
    sheet.mergeCells('A2:J2');
    sheet.getRange('A2').values = [[subtitle]];
    sheet.getRange('A2:J2').format = { font: { color: colors.muted, italic: true, size: 10 }, wrapText: true };
  }
}

function header(sheet, range, values) {
  sheet.getRange(range).values = [values];
  sheet.getRange(range).format = { fill: colors.blue, font: { bold: true, color: colors.white, size: 10 }, horizontalAlignment: 'center', verticalAlignment: 'center', wrapText: true, borders: { preset: 'all', style: 'thin', color: colors.border } };
}

function bodyStyle(sheet, range, fill=null) {
  sheet.getRange(range).format = { font: { color: colors.text, size: 10 }, verticalAlignment: 'center', wrapText: true, borders: { preset: 'all', style: 'thin', color: colors.border }, ...(fill ? { fill } : {}) };
}

function widths(sheet, pairs) {
  for (const [col, width] of pairs) sheet.getRange(`${col}:${col}`).format.columnWidth = width;
}

function addTable(sheet, address, name) {
  try { sheet.tables.add(address, true, name); } catch (e) { /* best effort */ }
}

// 01 사용안내
const guide = wb.worksheets.add('01_사용안내');
guide.showGridLines = false;
title(guide, 'SCM 월간 기기·옵션 발주 업무용 템플릿', '워크숍 결과 반영 | 노란색: 사용자가 입력 | 녹색: 자동 계산 | 빨간색: 확인 필요');
header(guide, 'A4:D4', ['순서','사용 시트','목적','사용 방법']);
guide.getRange('A5:D15').values = [
  [1,'03_수요입력','OL·SFDC·Bulk-deal·Trend·수급회의 결과 입력','월도별 수요를 입력하고 확정 여부를 선택'],
  [2,'04_재고_OpenPO','전월말 재고 및 미입고 발주 입력','기말재고·예약·보류·Open PO를 품목별 입력'],
  [3,'05_기기발주계산','기기 발주 필요량 계산','수요·재고·안전재고·MOQ·Flex 입력 후 수식 결과 확인'],
  [4,'06_옵션발주계산','BOM·장착율 기반 옵션 수량 계산','기종별 옵션 수요 입력 후 옵션코드별 합산 결과 확인'],
  [5,'08_예외_승인','예외·조정·승인 이력 관리','경고 발생 건의 조치·승인 상태 기록'],
  [6,'09_보고자료','사장 보고용 요약','기기·옵션 금액과 전월 비교, OL 대비 차이를 확인'],
  [7,'10_PoC_데이터요구','데이터 확보 현황 관리','워크숍에서 식별된 필수 데이터의 담당자·기한 입력'],
  [8,'07_기준정보','기종·옵션·BOM·공급 조건 관리','실제 운영 전 기준정보를 최신화'],
  [9,'11_용어','용어 정의 및 확인사항','확정 정의와 확인 담당자 기록'],
  [10,'02_발주대시보드','전체 진행 상태 확인','발주월과 승인 상태를 입력하면 주요 KPI 표시'],
  [11,'','주의','본 파일은 PoC용 템플릿이며 ERP/SFDC 실시간 연동 전에는 업로드 방식으로 사용'],
];
bodyStyle(guide, 'A5:D15', colors.white);
guide.getRange('A15:D15').format = { fill: colors.warn, font: { color: '#9B1C1C', bold: true, size: 10 }, wrapText: true, borders: { preset: 'all', style: 'thin', color: colors.border } };
header(guide, 'A18:C18', ['색상','의미','예시']);
guide.getRange('A19:C21').values = [['노란색', '사용자 입력 영역', '수요·재고·단가·MOQ·Flex·승인'], ['녹색', '수식 계산 결과', '가용재고·필요량·MOQ 반영량·금액'], ['빨간색', '예외·검토 필요', 'Flex 초과·MOQ 과잉·필수 옵션 누락']];
bodyStyle(guide, 'A19:C21');
guide.getRange('A19').format.fill = colors.input; guide.getRange('A20').format.fill = colors.calc; guide.getRange('A21').format.fill = colors.warn;
widths(guide, [['A',10],['B',22],['C',34],['D',70]]); guide.freezePanes.freezeRows(4);

// 02 dashboard
const dash = wb.worksheets.add('02_발주대시보드'); dash.showGridLines = false;
title(dash, '월간 발주 대시보드', '발주월과 승인 상태를 기준으로 자동 집계됩니다. 계산 시트 입력 후 확인하세요.');
dash.getRange('A4:B6').values = [['발주월',''],['승인상태','승인 요청'],['마지막 업데이트','']];
dash.getRange('B4').values = [[new Date(2026,7,1)]]; dash.getRange('B4').format.numberFormat = 'yyyy-mm';
dash.getRange('B6').formulas = [['=TODAY()']]; dash.getRange('B6').format.numberFormat = 'yyyy-mm-dd';
dash.getRange('A4:A6').format = { fill: colors.lightBlue, font: { bold: true, color: colors.navy }, borders: { preset: 'all', style: 'thin', color: colors.border } };
dash.getRange('B4:B6').format = { fill: colors.input, borders: { preset: 'all', style: 'thin', color: colors.border } };
dash.getRange('B5').dataValidation = { rule: { type: 'list', values: ['작성 중','수요 확정','산출 완료','승인 요청','승인 완료','발주 입력 완료','입고/지불 관리'] } };
header(dash, 'D4:I4', ['KPI','당월','전월','증감','증감률','설명']);
dash.getRange('D5:I12').values = [
  ['총 발주금액','','','','','기기+옵션'], ['기기 발주금액','','','','','05_기기발주계산'], ['옵션 발주금액','','','','','06_옵션발주계산'], ['기기 발주수량','','','','','최종 발주량'], ['옵션 발주수량','','','','','최종 발주량'], ['예외 건수','','','','','08_예외_승인'], ['Flex 초과','','','','','기기·옵션 계산 결과'], ['필수 옵션 누락','','','','','BOM 필수 여부 검증']
];
dash.getRange('E5').formulas = [["=SUM('05_기기발주계산'!$R$6:$R$205)+SUM('06_옵션발주계산'!$T$6:$T$205)"]];
dash.getRange('E6').formulas = [["=SUM('05_기기발주계산'!$R$6:$R$205)"]];
dash.getRange('E7').formulas = [["=SUM('06_옵션발주계산'!$T$6:$T$205)"]];
dash.getRange('E8').formulas = [["=SUM('05_기기발주계산'!$Q$6:$Q$205)"]];
dash.getRange('E9').formulas = [["=SUM('06_옵션발주계산'!$S$6:$S$205)"]];
dash.getRange('E10').formulas = [["=COUNTIF('08_예외_승인'!$G$6:$G$205,\"검토 필요\")"]];
dash.getRange('E11').formulas = [["=COUNTIF('05_기기발주계산'!$O$6:$O$205,\"Flex 초과\")+COUNTIF('06_옵션발주계산'!$Q$6:$Q$205,\"Flex 초과\")"]];
dash.getRange('E12').formulas = [["=COUNTIF('06_옵션발주계산'!$R$6:$R$205,\"누락\")"]];
dash.getRange('F5:F12').formulas = Array.from({length:8},(_,i)=>[`=0`]);
dash.getRange('G5:G12').formulas = Array.from({length:8},(_,i)=>[`=E${5+i}-F${5+i}`]);
dash.getRange('H5:H12').formulas = Array.from({length:8},(_,i)=>[`=IFERROR(G${5+i}/F${5+i},0)`]);
dash.getRange('I5:I12').values = [['단가 포함'],['기기 단가 포함'],['옵션 단가 포함'],['수량'],['수량'],['수동 입력'],['규칙 위반'],['BOM 필수품']];
bodyStyle(dash, 'D5:I12'); dash.getRange('D5:D12').format = { fill: colors.lightBlue, font:{bold:true,color:colors.navy}, borders:{preset:'all',style:'thin',color:colors.border} }; dash.getRange('E5:H12').format.fill = colors.calc;
dash.getRange('E5:G7').format.numberFormat = '#,##0'; dash.getRange('E8:G12').format.numberFormat = '#,##0'; dash.getRange('H5:H12').format.numberFormat = '0.0%';
header(dash, 'A16:F16', ['진행 점검','상태','담당','기한','결과','메모']);
dash.getRange('A17:F23').values = [['OL 취합','미착수','','','',''],['수요 확정','미착수','','','',''],['재고/Open PO','미착수','','','',''],['기기 발주량','미착수','','','',''],['옵션 발주량','미착수','','','',''],['보고·승인','미착수','','','',''],['System 입력·입고','미착수','','','','']];
bodyStyle(dash, 'A17:F23'); dash.getRange('B17:B23').dataValidation = { rule: { type:'list', values:['미착수','진행 중','완료','보류'] } };
widths(dash, [['A',20],['B',18],['C',20],['D',14],['E',18],['F',32],['D',22],['E',16],['F',16],['G',16],['H',14],['I',22]]); dash.freezePanes.freezeRows(4);

// 03 demand input
const demand = wb.worksheets.add('03_수요입력'); demand.showGridLines = false;
title(demand, '수요 입력 및 확정', 'OL·SFDC·Bulk-deal·과거 Trend·수급회의 결과를 한 행 단위로 입력합니다.');
header(demand, 'A4:Q4', ['월도','수요구분','법인/채널','기종/옵션코드','품목명','필요월도','수량','확도','확정수량','변경/취소','Bulk-deal','수급회의 조정','최종 확정','출처','담당자','확정상태','조정사유']);
const demandRows = Array.from({length:200},()=>['','','','','','',null,'',null,'','',null,null,'','','미확정','']); demand.getRange('A5:Q204').values = demandRows; bodyStyle(demand,'A5:Q204',colors.input); addTable(demand,'A4:Q204','DemandInput');
demand.getRange('I5').formulas = [['=IF(A5="","",MAX(0,G5+L5+IF(K5="Y",G5,0)-IF(J5="취소",G5,0)))']]; demand.getRange('I5:I204').fillDown(); demand.getRange('M5').formulas = [['=IF(P5="확정",I5,0)']]; demand.getRange('M5:M204').fillDown(); demand.getRange('F5:F204').format.numberFormat = 'yyyy-mm'; demand.getRange('G5:G204').format.numberFormat = '#,##0'; demand.getRange('I5:I204').format.numberFormat = '#,##0'; demand.getRange('L5:M204').format.numberFormat = '#,##0';
for (const c of ['B5:B204','H5:H204','J5:J204','K5:K204','N5:N204','P5:P204']) demand.getRange(c).dataValidation = {rule:{type:'list',values: c.startsWith('P')?['미확정','확정','보류']:c.startsWith('K')?['Y','N']:c.startsWith('J')?['정상','변경','취소']:c.startsWith('B')?['OL','SFDC','Bulk-deal','Trend','수급회의']:['높음','중간','낮음']}};
demand.getRange('I5:I204').format.fill = colors.calc; demand.getRange('M5:M204').format.fill = colors.calc;
widths(demand, [['A',12],['B',13],['C',18],['D',18],['E',22],['F',12],['G',12],['H',10],['I',13],['J',12],['K',12],['L',14],['M',13],['N',14],['O',14],['P',12],['Q',28]]); demand.freezePanes.freezeRows(4);

// 04 inventory & open PO
const inv = wb.worksheets.add('04_재고_OpenPO'); inv.showGridLines = false;
title(inv, '전월말 재고 및 Open PO', '재고 상태와 미입고 발주를 분리 입력하여 가용재고를 계산합니다.');
header(inv, 'A4:P4', ['기준월','품목구분','품목코드','품목명','Supplier','창고','기말재고','예약/할당','품질보류','출고예정','가용기말재고','Open PO 수량','확정입고수량','입고예정월','실제 가용재고','비고']);
inv.getRange('A5:P204').values = Array.from({length:200},()=>['','','','','','',null,null,null,null,null,null,null,'',null,'']); bodyStyle(inv,'A5:P204',colors.input); addTable(inv,'A4:P204','InventoryOpenPO');
inv.getRange('K5').formulas = [['=IF(C5="","",MAX(0,G5-H5-I5-J5))']]; inv.getRange('K5:K204').fillDown(); inv.getRange('O5').formulas = [['=IF(C5="","",K5+M5)']]; inv.getRange('O5:O204').fillDown(); inv.getRange('G5:O204').format.numberFormat = '#,##0'; inv.getRange('K5:K204').format.fill = colors.calc; inv.getRange('O5:O204').format.fill = colors.calc; inv.getRange('B5:B204').dataValidation = {rule:{type:'list',values:['기기','PRT','옵션','소모품','부품']}}; widths(inv,[['A',12],['B',12],['C',18],['D',22],['E',16],['F',16],['G',12],['H',12],['I',12],['J',12],['K',15],['L',12],['M',14],['N',14],['O',14],['P',26]]); inv.freezePanes.freezeRows(4);

// 07 master
const master = wb.worksheets.add('07_기준정보'); master.showGridLines = false;
title(master, '기준정보 관리', '실제 운영 전 기종·옵션·BOM·공급 조건을 등록합니다. 유효기간과 변경 이력을 남겨야 합니다.');
header(master,'A4:N4',['품목구분','품목코드','품목명','제품군','Supplier','단가','통화','MOQ','발주단위','Lead Time(일)','Flex 하한%','Flex 상한%','단종/대체','유효기간']);
master.getRange('A5:N204').values = Array.from({length:200},()=>['','','','','',null,'',null,null,null,null,null,'','']); bodyStyle(master,'A5:N204',colors.input); addTable(master,'A4:N204','MasterItems'); master.getRange('F5:F204').format.numberFormat = '#,##0'; master.getRange('K5:L204').format.numberFormat = '0.0%'; master.getRange('A5:A204').dataValidation={rule:{type:'list',values:['기기','PRT','옵션','소모품','부품']}}; widths(master,[['A',12],['B',18],['C',24],['D',16],['E',16],['F',12],['G',10],['H',10],['I',12],['J',14],['K',12],['L',12],['M',18],['N',16]]); master.freezePanes.freezeRows(4);

// 05 machine calculation
const mach = wb.worksheets.add('05_기기발주계산'); mach.showGridLines = false;
title(mach, '기기 발주량 계산', '노란색 입력 → 녹색 수식 → 빨간색 경고. 최종 발주량은 계산값을 검토 후 확정 입력할 수 있습니다.');
header(mach,'A4:S4',['발주월','필요월도','기종코드','기종명','Supplier','확정수요','안전재고','기말재고','예약/보류','확정입고','가용재고','계산 필요량','MOQ','Flex 하한','Flex 상한','Flex 판정','최종 발주량','발주금액','검토상태']);
mach.getRange('A5:S204').values = Array.from({length:200},()=>['','','','','',null,null,null,null,null,null,null,null,null,null,'',null,null,'미검토']); bodyStyle(mach,'A5:S204',colors.input); addTable(mach,'A4:S204','MachineOrderCalc');
mach.getRange('K5').formulas = [['=IF(C5="","",MAX(0,H5-I5+J5))']]; mach.getRange('K5:K204').fillDown(); mach.getRange('L5').formulas = [['=IF(C5="","",MAX(0,F5+G5-K5))']]; mach.getRange('L5:L204').fillDown(); mach.getRange('N5').formulas = [['=IF(C5="","",F5*(1-0.2))']]; mach.getRange('N5:N204').fillDown(); mach.getRange('O5').formulas = [['=IF(C5="","",F5*(1+0.2))']]; mach.getRange('O5:O204').fillDown(); mach.getRange('P5').formulas = [['=IF(C5="","",IF(L5=0,"정상",IF(L5<N5,"Flex 미달",IF(L5>O5,"Flex 초과","정상"))))']]; mach.getRange('P5:P204').fillDown(); mach.getRange('Q5').formulas = [['=IF(C5="","",IF(L5=0,0,ROUNDUP(L5/M5,0)*M5))']]; mach.getRange('Q5:Q204').fillDown(); mach.getRange('R5').formulas = [['=IF(C5="","",Q5*0)']]; mach.getRange('R5:R204').fillDown(); mach.getRange('K5:Q204').format.fill = colors.calc; mach.getRange('R5:R204').format.fill = colors.calc; mach.getRange('A5:B204').format.numberFormat = 'yyyy-mm'; mach.getRange('F5:O204').format.numberFormat = '#,##0'; mach.getRange('Q5:Q204').format.numberFormat = '#,##0'; mach.getRange('R5:R204').format.numberFormat = '#,##0'; mach.getRange('P5:P204').conditionalFormats.add('containsText',{text:'Flex 초과',format:{fill:colors.warn,font:{color:'#9B1C1C',bold:true}}}); mach.getRange('S5:S204').dataValidation={rule:{type:'list',values:['미검토','정상','예외 승인','확정']}}; widths(mach,[['A',12],['B',12],['C',18],['D',22],['E',16],['F',12],['G',12],['H',12],['I',12],['J',12],['K',12],['L',14],['M',10],['N',12],['O',12],['P',13],['Q',14],['R',14],['S',14]]); mach.freezePanes.freezeRows(4);

// Use unit price from 07_기준정보 to calculate machine order amount.
mach.getRange('R5').formulas = [["=IF(C5=\"\",\"\",Q5*SUMIF('07_기준정보'!$B$5:$B$204,C5,'07_기준정보'!$F$5:$F$204))"]];
mach.getRange('R5:R204').fillDown();
mach.getRange('R5:R204').format.fill = colors.calc;
mach.getRange('R5:R204').format.numberFormat = '#,##0';

// 06 option calculation
const opt = wb.worksheets.add('06_옵션발주계산'); opt.showGridLines = false;
title(opt, '옵션 발주량 계산', '기종별 BOM·장착율을 입력한 후 옵션코드별로 Common품을 합산합니다.');
header(opt,'A4:U4',['발주월','필요월도','옵션코드','옵션명','기종코드','기기확정수요','BOM수량','장착율','평균사용량','별도수요','이론수요','옵션가용재고','안전재고','계산 필요량','MOQ','Flex 판정','필수옵션','누락검증','최종 발주량','단가','발주금액']);
opt.getRange('A5:U204').values = Array.from({length:200},()=>['','','','','',null,null,null,null,null,null,null,null,null,null,'정상','','정상',null,null,null]); bodyStyle(opt,'A5:U204',colors.input); addTable(opt,'A4:U204','OptionOrderCalc');
opt.getRange('K5').formulas = [['=IF(C5="","",F5*G5*H5*I5+J5)']]; opt.getRange('K5:K204').fillDown(); opt.getRange('N5').formulas = [['=IF(C5="","",MAX(0,K5+M5-L5))']]; opt.getRange('N5:N204').fillDown(); opt.getRange('P5').formulas = [['=IF(C5="","",IF(N5=0,"정상",IF(N5>O5*10,"MOQ 과잉 확인","정상")))']]; opt.getRange('P5:P204').fillDown(); opt.getRange('R5').formulas = [['=IF(C5="","",IF(Q5="Y",IF(K5>0,"정상","누락"),"정상"))']]; opt.getRange('R5:R204').fillDown(); opt.getRange('S5').formulas = [['=IF(C5="","",IF(N5=0,0,ROUNDUP(N5/O5,0)*O5))']]; opt.getRange('S5:S204').fillDown(); opt.getRange('U5').formulas = [['=IF(C5="","",S5*T5)']]; opt.getRange('U5:U204').fillDown(); opt.getRange('K5:N204').format.fill = colors.calc; opt.getRange('P5:P204').format.fill = colors.calc; opt.getRange('R5:R204').format.fill = colors.calc; opt.getRange('S5:U204').format.fill = colors.calc; opt.getRange('A5:B204').format.numberFormat = 'yyyy-mm'; opt.getRange('H5:H204').format.numberFormat = '0.0%'; opt.getRange('F5:G204').format.numberFormat = '#,##0'; opt.getRange('I5:O204').format.numberFormat = '#,##0.0'; opt.getRange('S5:U204').format.numberFormat = '#,##0'; opt.getRange('Q5:Q204').dataValidation={rule:{type:'list',values:['Y','N']}}; opt.getRange('P5:P204').conditionalFormats.add('containsText',{text:'MOQ 과잉',format:{fill:colors.warn,font:{color:'#9B1C1C',bold:true}}}); opt.getRange('R5:R204').conditionalFormats.add('containsText',{text:'누락',format:{fill:colors.warn,font:{color:'#9B1C1C',bold:true}}}); widths(opt,[['A',12],['B',12],['C',18],['D',22],['E',18],['F',12],['G',10],['H',10],['I',13],['J',12],['K',12],['L',14],['M',12],['N',14],['O',10],['P',14],['Q',10],['R',13],['S',14],['T',12],['U',14]]); opt.freezePanes.freezeRows(4);

// 08 exception/approval
const ex = wb.worksheets.add('08_예외_승인'); ex.showGridLines = false;
title(ex, '예외·조정·승인 이력', '자동 계산 경고 또는 사람의 수량 조정은 반드시 사유·승인자·재시작 단계를 기록합니다.');
header(ex,'A4:J4',['발생월','단계','품목코드','예외유형','시스템 계산값','조정값','상태','조치/사유','담당자','승인자']);
ex.getRange('A5:J204').values = Array.from({length:200},()=>['','','','','',null,'미등록','','','']); bodyStyle(ex,'A5:J204',colors.input); addTable(ex,'A4:J204','ExceptionApproval'); ex.getRange('G5:G204').dataValidation={rule:{type:'list',values:['미등록','검토 필요','조치 중','승인 요청','승인 완료','반려','종결']}}; ex.getRange('B5:B204').dataValidation={rule:{type:'list',values:['0 OL 취합','① 수요','② 재고','③ 기기','④ 옵션','⑤ 보고','⑥ 입력','⑦ 입고/지불']}}; ex.getRange('G5:G204').conditionalFormats.add('containsText',{text:'검토 필요',format:{fill:colors.warn,font:{color:'#9B1C1C',bold:true}}}); widths(ex,[['A',12],['B',16],['C',18],['D',24],['E',14],['F',12],['G',14],['H',36],['I',16],['J',16]]); ex.freezePanes.freezeRows(4);

// 09 report
const rep = wb.worksheets.add('09_보고자료'); rep.showGridLines = false;
title(rep, '사장 보고용 발주 요약', '수치가 채워진 계산 시트에서 자동 집계됩니다. 변동 원인은 08_예외_승인 또는 메모란에 기록합니다.');
rep.getRange('A4:B5').values = [['보고월',new Date(2026,7,1)],['보고상태','작성 중']]; rep.getRange('B4').format.numberFormat='yyyy-mm'; rep.getRange('A4:A5').format={fill:colors.lightBlue,font:{bold:true,color:colors.navy},borders:{preset:'all',style:'thin',color:colors.border}}; rep.getRange('B4:B5').format={fill:colors.input,borders:{preset:'all',style:'thin',color:colors.border}}; rep.getRange('B5').dataValidation={rule:{type:'list',values:['작성 중','검토 중','승인 요청','승인 완료']}};
header(rep,'A8:F8',['구분','당월 수량','당월 금액','전월 수량','전월 금액','증감 원인']);
rep.getRange('A9:F11').values=[['기기',null,null,0,0,'수요·재고·MOQ·Flex 변동'],['옵션',null,null,0,0,'BOM·장착율·Common품·MOQ 변동'],['합계',null,null,0,0,'']];
rep.getRange('B9').formulas=[["=SUM('05_기기발주계산'!$Q$6:$Q$205)"]]; rep.getRange('C9').formulas=[["=SUM('05_기기발주계산'!$R$6:$R$205)"]]; rep.getRange('B10').formulas=[["=SUM('06_옵션발주계산'!$S$6:$S$205)"]]; rep.getRange('C10').formulas=[["=SUM('06_옵션발주계산'!$U$6:$U$205)"]]; rep.getRange('B11').formulas=[['=SUM(B9:B10)']]; rep.getRange('C11').formulas=[['=SUM(C9:C10)']]; rep.getRange('D11').formulas=[['=SUM(D9:D10)']]; rep.getRange('E11').formulas=[['=SUM(E9:E10)']]; bodyStyle(rep,'A9:F11'); rep.getRange('B9:E11').format.numberFormat='#,##0'; rep.getRange('B9:C11').format.fill=colors.calc; rep.getRange('A11:F11').format={fill:colors.lightBlue,font:{bold:true,color:colors.navy},borders:{preset:'all',style:'thin',color:colors.border}};
header(rep,'A15:F15',['비교 항목','전월 OL','최종 확정 수요','실제 발주 필요량','차이','차이 사유']); rep.getRange('A16:F19').values=[['기기',0,null,null,null,''],['옵션',0,null,null,null,''],['Flex 초과 건수',0,null,null,null,''],['필수 옵션 누락',0,null,null,null,'']]; rep.getRange('C16').formulas=[["=SUM('03_수요입력'!$M$5:$M$204)"]]; rep.getRange('D16').formulas=[["=SUM('05_기기발주계산'!$L$6:$L$205)"]]; rep.getRange('E16').formulas=[['=D16-C16']]; rep.getRange('C17').formulas=[["=SUM('06_옵션발주계산'!$K$6:$K$205)"]]; rep.getRange('D17').formulas=[["=SUM('06_옵션발주계산'!$N$6:$N$205)"]]; rep.getRange('E17').formulas=[['=D17-C17']]; rep.getRange('C18').formulas=[["=COUNTIF('05_기기발주계산'!$P$6:$P$205,\"Flex 초과\")+COUNTIF('06_옵션발주계산'!$P$6:$P$205,\"Flex 초과\")"]]; rep.getRange('C19').formulas=[["=COUNTIF('06_옵션발주계산'!$R$6:$R$205,\"누락\")"]]; bodyStyle(rep,'A16:F19'); rep.getRange('C16:E19').format.fill=colors.calc; widths(rep,[['A',24],['B',16],['C',18],['D',18],['E',16],['F',42]]); rep.freezePanes.freezeRows(8);

// 10 data requirements
const data = wb.worksheets.add('10_PoC_데이터요구'); data.showGridLines=false;
title(data,'PoC 데이터 요구사항','워크숍에서 식별된 데이터의 보유 여부·위치·담당자·기한을 관리합니다.');
header(data,'A4:J4',['데이터명','왜 필요한가','보유 여부','위치/시스템','형태','확보 담당자','기한','마스킹','우선순위','비고']);
data.getRange('A5:J20').values=[
 ['기종 마스터','기종 코드·제품군·MOQ·발주단위·단가','','','','','', '단가','높음',''],['옵션 마스터','옵션 코드·MOQ·발주단위·Common 여부·단가','','','','','', '단가','높음',''],['BOM (기종-옵션)','옵션 소요량 전개·필수 여부','','','','','', '', '높음',''],['전월말 기기 재고','기기 소요량 산출','','','','','', '', '높음',''],['전월말 옵션 재고','옵션 소요량 산출','','','','','', '', '높음',''],['기기 판매 실적 24개월','Trend·장착율 분모','','','','','', '고객사명','중간',''],['옵션 판매 실적 24개월','장착율 분자','','','','','', '고객사명','중간',''],['발주~입고 이력','Lead time 실측','','','','','', '거래처명','최우선','발주일·입고일 연결 필수'],['미입고 발주 현황','가용재고 계산','','','','','', '', '높음',''],['OL 제출 이력 12개월','Flex 검증·OL 대비 비교','','','','','', '', '높음',''],['SFDC 파이프라인','수요 반영','','','','','', '고객사명','중간',''],['Flexibility rule 문서','허용 범위 판정','','','','','', '', '높음',''],['전월 발주 실적·금액','전월 대비 비교','','','','','', '단가','높음',''],['사장 보고 샘플','출력 형식 참고','','','','','', '금액','중간',''],['지난달 발주 산출 엑셀','현행 계산 로직 파악','','','','','', '단가','높음',''],['승인·반려 이력','재시작 단계·변경 사유 관리','','','','','', '', '중간','']
]; bodyStyle(data,'A5:J20',colors.input); addTable(data,'A4:J20','PocDataRequirements'); data.getRange('C5:C20').dataValidation={rule:{type:'list',values:['미확인','보유','일부 보유','미보유']}}; data.getRange('I5:I20').dataValidation={rule:{type:'list',values:['최우선','높음','중간','낮음']}}; widths(data,[['A',24],['B',38],['C',14],['D',24],['E',14],['F',16],['G',14],['H',16],['I',12],['J',32]]); data.freezePanes.freezeRows(4);

// 11 glossary
const gloss = wb.worksheets.add('11_용어'); gloss.showGridLines=false;
title(gloss,'용어 및 정책 확인','워크숍 추정 정의는 실제 운영 규정 확인 후 확정 정의에 반영합니다.');
header(gloss,'A4:F4',['용어','현재 이해','확정 정의','확인 담당자','확인 상태','비고']);
gloss.getRange('A5:F20').values=[
 ['OL','영업에서 취합하여 공유하는 판매계획 수량','','','미확인','제품·부품 포함 여부'],['필요월도','재고가 실제로 필요한 고객 출고 예측월도','','','미확인','발주월과 분리'],['Flexibility rule','Supplier 제출 OL 대비 발주 변경 허용 범위','','','미확인','전월 ±20%, 전전월 ±30%는 추정값'],['PRT','A3/A4 Printer 계열 제품군','','','미확인','MOQ 별도 적용'],['장착율','기기 판매 시 옵션 동반 판매 비율','','','미확인','산출 기간 필요'],['Common품','복수 기종 공통 사용 옵션/부품/소모품','','','미확인','옵션코드 통합'],['옵션 필수품','특정 기기·옵션 출하 시 반드시 포함되는 옵션','','','미확인','누락 경고'],['평균 사용량','3·6·12개월 기간별 평균 사용량','','','미확인','특수 실적 제외 여부'],['MOQ','최소 발주 수량 단위','','','미확인','발주단위와 구분'],['BOM','기종별 구성 자재 명세','','','미확인','유효기간 필요'],['SFDC','Salesforce 영업 파이프라인','','','미확인','확도 반영 정책'],['수급회의','월 단위 수요 최종 검증 회의','','','미확인','결정사항 기록'],['전월말 재고','전월 말 기준 마감 재고','','','미확인','재고 상태별 구분'],['Open PO','발주했으나 미입고인 발주 잔량','','','미확인','필요월도 전 입고 가능량'],['FX-LIVE','산출 발주량을 최종 확정하여 Supplier에 제출하는 System','','','미확인','실제 시스템명 확인'],['PO Match','PO·입고·Invoice 매칭을 통한 지불 처리','','','미확인','차이 처리 기준']
]; bodyStyle(gloss,'A5:F20',colors.input); addTable(gloss,'A4:F20','Glossary'); gloss.getRange('E5:E20').dataValidation={rule:{type:'list',values:['미확인','확인 중','확정']}}; widths(gloss,[['A',22],['B',48],['C',40],['D',16],['E',14],['F',28]]); gloss.freezePanes.freezeRows(4);

// Apply global formatting and render previews
for (const s of wb.worksheets.items) {
  const used = s.getUsedRange();
  if (used) used.format.font.name = 'Aptos';
}

const previewSheets = ['01_사용안내','02_발주대시보드','05_기기발주계산','06_옵션발주계산','09_보고자료'];
for (const name of previewSheets) {
  const blob = await wb.render({sheetName:name, autoCrop:'all', scale:1, format:'png'});
  await fs.writeFile(`${outDir}/${name}.png`, new Uint8Array(await blob.arrayBuffer()));
}
const inspect = await wb.inspect({kind:'table', sheetId:'05_기기발주계산', range:'A4:S10', include:'values,formulas', tableMaxRows:10, tableMaxCols:20, maxChars:5000});
console.log(inspect.ndjson);
const errors = await wb.inspect({kind:'match', searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A', options:{useRegex:true,maxResults:100}, summary:'formula error scan'});
console.log(errors.ndjson);
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(`${outDir}/SCM_월간_발주_업무용_템플릿.xlsx`);
console.log(`${outDir}/SCM_월간_발주_업무용_템플릿.xlsx`);
