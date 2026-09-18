// Requires the Codex workspace @oai/artifact-tool runtime. See references/automation.md.
import fs from 'node:fs/promises';
import path from 'node:path';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
import {calculate,FIELDS} from './pnl.mjs';

const [inputFile,outputFile,mode] = process.argv.slice(2);
if (!inputFile || !outputFile || (mode && mode!=='--blank')) throw new Error('Usage: node scripts/build-workbook.mjs input.json output.xlsx [--blank]');
const input=JSON.parse(await fs.readFile(inputFile,'utf8'));
const control=calculate(input);
if(control.status==='invalid')throw new Error('Invalid input. Run pnl.mjs calculate and resolve structural issues first.');
if(input.sites.length!==2 || !input.scenarios.includes('actual') || !input.scenarios.includes('budget'))throw new Error('Workbook requires exactly two sites and actual + budget; core supports a wider perimeter.');
const blank=mode==='--blank';
const scopes=[...input.sites,'CENTER','ELIM'];
const data=structuredClone(input);
if(blank){for(const r of data.records){for(const f of FIELDS)r[f]=null;r.closed=false;r.source_ref='';}for(const a of data.allocations)for(const s of input.sites)a.bases[s]=null;}
const ordered=[];
for(const period of data.periods)for(const scenario of data.scenarios)for(const scope of scopes){const row=data.records.find(r=>r.period===period&&r.scenario===scenario&&r.scope===scope);if(!row)throw new Error('Workbook requires explicit rows for every expected period/scenario/scope, even with null amounts.');ordered.push(row);}
const w=Workbook.create();
const out=w.worksheets.add('P&L');
const raw=w.worksheets.add('Данные');
const allocation=w.worksheets.add('Распределение');
const calc=w.worksheets.add('Расчёт');
const check=w.worksheets.add('Проверки');
const number='#,##0.00;(#,##0.00);"—"';
const serial=p=>(Date.UTC(+p.slice(0,4),+p.slice(5)-1,1)-Date.UTC(1899,11,30))/86400000;
const f=(s,cell,value)=>{s.getRange(cell).formulas=[[value]];s.getRange(cell).format.font.color=s===out?'#17243B':value.includes('!')?'#16703B':'#17243B';};
const v=(s,cell,value)=>s.getRange(cell).values=[[value]];
function style(s,lastCol,lastRow,title){
  const area=s.getRange(`B2:${lastCol}${lastRow}`);
  area.format.font={name:'Arial',size:10,color:'#17243B'};
  area.format.rowHeight=24;
  area.format.columnWidth=15;
  area.format.verticalAlignment='center';
  s.getRange(`A1:A${lastRow}`).format.columnWidth=3;
  s.showGridLines=false;v(s,'B2',title);s.getRange('B2').format.font={name:'Arial',size:16,bold:true,color:'#17243B'};
  s.getRange(`B5:${lastCol}5`).format={fill:'#263A58',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},rowHeight:40,wrapText:true,horizontalAlignment:'center'};
  s.freezePanes.freezeRows(5);
}
const last=5+ordered.length;
style(raw,'S',last,'Исходные месячные суммы');
raw.getRange('B3').values=[[blank?'Заполните синие ячейки. Пустое значение не равно нулю.':'Источник: examples/demo.json — вымышленные данные; рубли, возмещаемый НДС исключён.']];
raw.getRange('B5:S5').values=[['ID','Месяц','Объект','Сценарий','Выручка','Себестоимость','Коммерческие','Управленческие','Прочие доходы','Прочие расходы','Амортизация в P&L','Фин. доходы','Фин. расходы','Налог','Закрыт: 1/0','Источник строки','Число сумм','Пул известен']];
const literal=value=>typeof value==='string'&&/^[=+@-]/.test(value)?"'"+value:value;
const rawRows=ordered.map(r=>[literal(r.id),serial(r.period),literal(r.scope),r.scenario,...FIELDS.map(k=>r[k]),r.closed?1:0,literal(r.source_ref)]);
raw.getRange(`B6:Q${last}`).values=rawRows;
raw.getRange(`B6:B${last}`).format.columnWidth=29;
raw.getRange(`Q6:Q${last}`).format.columnWidth=40;
raw.getRange(`C6:C${last}`).setNumberFormat('mmm-yy');
raw.getRange(`F6:O${last}`).setNumberFormat(number);
raw.getRange(`F6:Q${last}`).format.font.color='#185ABD';
raw.getRange(`F6:Q${last}`).format.fill='#F2F6FC';
raw.getRange(`P6:P${last}`).dataValidation={rule:{type:'list',values:['0','1']}};
for(let r=6;r<=last;r++){f(raw,`R${r}`,`=COUNT(F${r}:O${r})`);f(raw,`S${r}`,`=IF(ISNUMBER(I${r}),1,0)`);}

const allocRows=[];
const allocIndex=new Map();
for(const period of data.periods)for(const scenario of data.scenarios){const a=data.allocations.find(r=>r.period===period&&r.scenario===scenario);if(!a)throw new Error('Missing allocation group');allocRows.push([serial(period),scenario,null,...data.sites.map(s=>a.bases[s])]);allocIndex.set(`${period}/${scenario}`,5+allocRows.length);}
const alast=5+allocRows.length;
style(allocation,'L',alast,'Распределение расходов центра');
v(allocation,'B3','Весь пул — управленческие расходы CENTER. При нулевых базах расходы остаются в центре.');
allocation.getRange('B5:L5').values=[['Месяц','Сценарий','Пул, руб.',`База ${data.sites[0]}`,`База ${data.sites[1]}`,'Всего база',`Расход ${data.sites[0]}`,`Расход ${data.sites[1]}`,'Остаток центра','Разница, руб.','Состояние']];
allocation.getRange(`B6:F${alast}`).values=allocRows;
allocation.getRange(`B6:B${alast}`).setNumberFormat('mmm-yy');
allocation.getRange(`D6:K${alast}`).setNumberFormat(number);
allocation.getRange(`E6:F${alast}`).format.font.color='#185ABD';
allocation.getRange(`E6:F${alast}`).format.fill='#F2F6FC';
allocation.getRange(`L6:L${alast}`).format.columnWidth=24;
for(let r=6;r<=alast;r++){
 const condition=`'Данные'!$C$6:$C$${last},B${r},'Данные'!$E$6:$E$${last},C${r},'Данные'!$D$6:$D$${last},"CENTER"`;
 f(allocation,`D${r}`,`=IF(COUNTIFS(${condition},'Данные'!$S$6:$S$${last},1)=1,SUMIFS('Данные'!$I$6:$I$${last},${condition}),"н/д")`);
 f(allocation,`G${r}`,`=IF(AND(COUNT(E${r}:F${r})=2,MIN(E${r}:F${r})>=0),SUM(E${r}:F${r}),"н/д")`);
 f(allocation,`H${r}`,`=IF(COUNT(D${r},G${r})=2,IF(D${r}<0,"н/д",IF(G${r}=0,0,ROUND(D${r}*E${r}/G${r},2))),"н/д")`);
 f(allocation,`I${r}`,`=IF(COUNT(D${r},G${r},H${r})=3,IF(G${r}=0,0,D${r}-H${r}),"н/д")`);
 f(allocation,`J${r}`,`=IF(COUNT(D${r},H${r}:I${r})=3,D${r}-SUM(H${r}:I${r}),"н/д")`);
 f(allocation,`K${r}`,`=IF(COUNT(D${r},H${r}:J${r})=4,D${r}-SUM(H${r}:J${r}),"н/д")`);
 f(allocation,`L${r}`,`=IF(COUNT(D${r}:J${r})<>7,"Нет данных",IF(J${r}<>0,"Не распределено","Распределено"))`);
}
style(calc,'V',last,'Расчёт по объектам');
v(calc,'B3','CENTER — корпоративный центр; ELIM — подготовленные корректировки свода. Амортизация уже включена в затраты.');
calc.getRange('B5:P5').values=[['Месяц','Объект','Сценарий','Выручка','Себестоимость','Валовая прибыль','EBIT до распр.','Распределение','EBIT','Амортизация','EBITDA','До налога','Налог','Чистая прибыль','Чистая маржа']];
calc.getRange('Q5:V5').values=[['Коммерческие','Управленческие после распр.','Прочие доходы','Прочие расходы','Фин. доходы','Фин. расходы']];
calc.getRange(`Q6:V${last}`).setNumberFormat(number);
calc.getRange(`E6:O${last}`).setNumberFormat(number);calc.getRange(`P6:P${last}`).setNumberFormat('0.0%');calc.getRange(`B6:B${last}`).setNumberFormat('mmm-yy');
const safe=(cell)=>`IF(ISNUMBER('Данные'!${cell}),'Данные'!${cell},"н/д")`;
for(const [i,row]of ordered.entries()){
 const r=i+6,a=allocIndex.get(`${row.period}/${row.scenario}`);
 f(calc,`B${r}`,`='Данные'!C${r}`);f(calc,`C${r}`,`='Данные'!D${r}`);f(calc,`D${r}`,`='Данные'!E${r}`);
 f(calc,`E${r}`,`=${safe(`F${r}`)}`);f(calc,`F${r}`,`=${safe(`G${r}`)}`);
 f(calc,`G${r}`,`=IF(COUNT(E${r}:F${r})=2,E${r}-F${r},"н/д")`);
 f(calc,`H${r}`,`=IF(COUNT(G${r},'Данные'!H${r}:K${r})=5,G${r}-'Данные'!H${r}-'Данные'!I${r}+'Данные'!J${r}-'Данные'!K${r},"н/д")`);
 const allocationFormula=row.scope===data.sites[0]?`'Распределение'!H${a}`:row.scope===data.sites[1]?`'Распределение'!I${a}`:row.scope==='CENTER'?`IF(COUNT('Распределение'!H${a}:I${a})=2,-SUM('Распределение'!H${a}:I${a}),"н/д")`:'0';
 f(calc,`I${r}`,`=${allocationFormula}`);f(calc,`J${r}`,`=IF(COUNT(H${r}:I${r})=2,H${r}-I${r},"н/д")`);
 f(calc,`K${r}`,`=${safe(`L${r}`)}`);f(calc,`L${r}`,`=IF(COUNT(J${r}:K${r})=2,SUM(J${r}:K${r}),"н/д")`);
 f(calc,`M${r}`,`=IF(COUNT(J${r},'Данные'!M${r}:N${r})=3,J${r}+'Данные'!M${r}-'Данные'!N${r},"н/д")`);
 f(calc,`N${r}`,`=${safe(`O${r}`)}`);f(calc,`O${r}`,`=IF(COUNT(M${r}:N${r})=2,M${r}-N${r},"н/д")`);
 f(calc,`P${r}`,`=IF(COUNT(E${r},O${r})=2,IF(E${r}=0,"н/д",O${r}/E${r}),"н/д")`);
 for(const [dest,source]of [['Q','H'],['S','J'],['T','K'],['U','M'],['V','N']])f(calc,`${dest}${r}`,`=${safe(`${source}${r}`)}`);
 f(calc,`R${r}`,`=IF(COUNT('Данные'!I${r},I${r})=2,'Данные'!I${r}+I${r},"н/д")`);
}
const monthCount=data.periods.length;
const summaryEnd=6+monthCount;
style(out,'K',summaryEnd+35,blank?'Шаблон производственного P&L':'Производственный P&L — демонстрация');
v(out,'B3',blank?'Заполните листы «Данные» и «Распределение». Суммы в рублях.':'ВЫМЫШЛЕННЫЕ ДАННЫЕ. Суммы в рублях, возмещаемый НДС исключён.');
out.getRange('B5:K5').values=[['Месяц','Выручка факт','Выручка план','Δ выручки','EBIT факт','EBITDA факт','Прибыль факт','Прибыль план','Δ прибыли','Чистая маржа']];
out.getRange(`C6:J${summaryEnd}`).setNumberFormat(number);out.getRange(`K6:K${summaryEnd}`).setNumberFormat('0.0%');
function aggregate(field,period,scenario){const rows=ordered.flatMap((r,i)=>r.period===period&&r.scenario===scenario?[i+6]:[]);const refs=rows.map(r=>`'Расчёт'!${field}${r}`).join(',');return `IF(COUNT(${refs})=${scopes.length},SUM(${refs}),"н/д")`;}
for(const [i,period]of data.periods.entries()){
 const r=i+6;v(out,`B${r}`,serial(period));out.getRange(`B${r}`).setNumberFormat('mmm-yy');
 for(const [dest,field,scenario]of [['C','E','actual'],['D','E','budget'],['F','J','actual'],['G','L','actual'],['H','O','actual'],['I','O','budget']])f(out,`${dest}${r}`,`=${aggregate(field,period,scenario)}`);
 f(out,`E${r}`,`=IF(COUNT(C${r}:D${r})=2,C${r}-D${r},"н/д")`);f(out,`J${r}`,`=IF(COUNT(H${r}:I${r})=2,H${r}-I${r},"н/д")`);
 f(out,`K${r}`,`=IF(COUNT(C${r},H${r})=2,IF(C${r}=0,"н/д",H${r}/C${r}),"н/д")`);
}
v(out,`B${summaryEnd}`,'Итого');
for(const c of ['C','D','E','F','G','H','I','J'])f(out,`${c}${summaryEnd}`,`=IF(COUNT(${c}6:${c}${summaryEnd-1})=${monthCount},SUM(${c}6:${c}${summaryEnd-1}),"н/д")`);
f(out,`K${summaryEnd}`,`=IF(COUNT(C${summaryEnd},H${summaryEnd})=2,IF(C${summaryEnd}=0,"н/д",H${summaryEnd}/C${summaryEnd}),"н/д")`);
out.getRange(`B${summaryEnd}:K${summaryEnd}`).format.fill='#E7EDF5';out.getRange(`B${summaryEnd}:K${summaryEnd}`).format.font.bold=true;
const detailStart=summaryEnd+4;
out.getRange(`B${detailStart}:F${detailStart}`).values=[['Объект','EBIT до распр.','Распределение','EBIT после','Чистая прибыль']];
out.getRange(`B${detailStart}:F${detailStart}`).format={fill:'#263A58',font:{bold:true,color:'#FFFFFF'},rowHeight:40,wrapText:true};
for(const [i,scope]of scopes.entries()){
 const r=detailStart+i+1;v(out,`B${r}`,scope);
 for(const [dest,field]of [['C','H'],['D','I'],['E','J'],['F','O']]){
 const rows=ordered.flatMap((x,j)=>x.scenario==='actual'&&x.scope===scope?[j+6]:[]);const refs=rows.map(x=>`'Расчёт'!${field}${x}`).join(',');f(out,`${dest}${r}`,`=IF(COUNT(${refs})=${monthCount},SUM(${refs}),"н/д")`);}
}
out.getRange(`C${detailStart+1}:F${detailStart+4}`).setNumberFormat(number);
v(out,`B${detailStart+6}`,'Статус периода');
f(out,`C${detailStart+6}`,`=IF(COUNT('Данные'!F6:O${last})<>${ordered.length*10},"Нет всех сумм",IF(COUNT('Распределение'!K6:K${alast})<>${allocRows.length},"Нет распределения",IF(COUNTBLANK('Данные'!Q6:Q${last})>0,"Нет источников",IF(COUNTIFS('Данные'!P6:P${last},1)<>${ordered.length},"Предварительный",IF(SUM('Распределение'!J6:J${alast})<>0,"Есть остаток центра","Рассчитан")))))`);
v(out,`B${detailStart+8}`,'Синие ячейки — ввод. Идентификаторы и состав месяцев меняйте в JSON, затем пересоздайте книгу.');
v(out,`B${detailStart+9}`,'«Рассчитан» не подтверждает сверку с 1С. См. «Проверки». Факт и бюджет заданы отдельно.');
const statementStart=detailStart+12;
out.getRange(`B${statementStart}:F${statementStart}`).values=[['P&L за период','Факт','Бюджет','Отклонение','Отклонение, %']];
out.getRange(`B${statementStart}:F${statementStart}`).format={fill:'#263A58',font:{bold:true,color:'#FFFFFF'},rowHeight:40,wrapText:true};
const statement=[['Выручка','E'],['Себестоимость','F'],['Валовая прибыль','G'],['Коммерческие','Q'],['Управление','R'],['Прочие доходы','S'],['Прочие расходы','T'],['EBIT','J'],['EBITDA, справ.','L'],['Фин. доходы','U'],['Фин. расходы','V'],['До налога','M'],['Налог','N'],['Чистая прибыль','O']];
for(const [i,[label,field]] of statement.entries()){
 const r=statementStart+i+1;v(out,`B${r}`,label);
 for(const [c,scenario] of [['C','actual'],['D','budget']]){
  const rows=ordered.flatMap((x,j)=>x.scenario===scenario?[j+6]:[]);const refs=rows.map(x=>`'Расчёт'!${field}${x}`).join(',');
  f(out,`${c}${r}`,`=IF(COUNT(${refs})=${rows.length},SUM(${refs}),"н/д")`);
 }
 f(out,`E${r}`,`=IF(COUNT(C${r}:D${r})=2,C${r}-D${r},"н/д")`);
 f(out,`F${r}`,`=IF(COUNT(C${r}:D${r})=2,IF(D${r}=0,"н/д",E${r}/ABS(D${r})),"н/д")`);
 if(['G','J','O'].includes(field)){out.getRange(`B${r}:F${r}`).format.fill='#E7EDF5';out.getRange(`B${r}:F${r}`).format.font.bold=true;}
}
out.getRange(`C${statementStart+1}:E${statementStart+statement.length}`).setNumberFormat(number);
out.getRange(`F${statementStart+1}:F${statementStart+statement.length}`).setNumberFormat('0.0%');
out.getRange(`C${detailStart+6}:D${detailStart+6}`).format.columnWidth=20;
out.getRange(`E6:E${summaryEnd}`).conditionalFormats.add('cellIs',{operator:'lessThan',formula:0,format:{fill:'#FCE8E6',font:{color:'#A61B1B'}}});
out.getRange(`J6:J${summaryEnd}`).conditionalFormats.add('cellIs',{operator:'lessThan',formula:0,format:{fill:'#FCE8E6',font:{color:'#A61B1B'}}});

style(check,'E',12,'Контроль полноты и расчётов');
check.getRange('B5:E5').values=[['Проверка','Отклонение / число','Ожидается','Значение проверки']];
check.getRange('B6:B10').values=[['Недостающие суммы'],['Незакрытые строки'],['Незаполненные источники'],['Распределения без расчёта'],['Разница распределений']];
f(check,'C6',`=${ordered.length*10}-COUNT('Данные'!F6:O${last})`);
f(check,'C7',`=${ordered.length}-COUNTIFS('Данные'!P6:P${last},1)`);
f(check,'C8',`=COUNTBLANK('Данные'!Q6:Q${last})`);
f(check,'C9',`=${allocRows.length}-COUNT('Распределение'!K6:K${alast})`);
f(check,'C10',`=IF(COUNT('Распределение'!K6:K${alast})=${allocRows.length},SUM('Распределение'!K6:K${alast}),"н/д")`);
check.getRange('D6:D10').values=[[0],[0],[0],[0],[0]];
check.getRange('E6:E10').values=[['Полнота 10 денежных полей'],['Статус закрытия задан пользователем'],['Наличие ссылки, не проверка документа'],['Пул и базы известны'],['Пул = площадки + остаток']];
check.getRange('B6:B10').format.columnWidth=34;check.getRange('E6:E10').format.columnWidth=44;check.getRange('C6:D10').setNumberFormat('0.00');
check.getRange('C6:C10').conditionalFormats.add('cellIs',{operator:'notEqual',formula:0,format:{fill:'#FCE8E6',font:{bold:true,color:'#A61B1B'}}});
v(check,'B12','Структуру ключей и повторные выгрузки проверяет pnl.mjs перед созданием книги.');
v(check,'B13','После изменения состава исходников пересоздайте книгу; не добавляйте строки вручную.');

// All comparisons below read the recalculated formula results, not copied expected values.
w.recalculate();
for(const [s,range]of [[raw,`F6:P${last}`],[calc,`E6:V${last}`],[allocation,`D6:K${alast}`],[out,`C6:K${summaryEnd}`],[out,`C${statementStart+1}:F${statementStart+statement.length}`]])s.getRange(range).format.horizontalAlignment='right';
function equalNumber(sheet,address,expected){const actual=sheet.getRange(address).values[0][0];if(typeof expected==='number'?typeof actual!=='number'||Math.abs(actual-expected)>0.000001:actual!==expected)throw new Error(`${sheet.name}!${address}: ${actual} != ${expected}`);}
if(!blank){
 for(const [i,period]of input.periods.entries()){
 const g=control.groups.find(x=>x.period===period&&x.scenario==='actual').total;
 for(const [col,field]of [['C','revenue'],['F','ebit'],['G','ebitda'],['H','net'],['K','margin']])equalNumber(out,`${col}${i+6}`,g[field]===null?'н/д':g[field]);
 }
 if(control.groups[0].total.net!==null){
 const row=raw.getRange('F6').values[0][0];raw.getRange('F6').values=[[row+100]];w.recalculate();equalNumber(out,'H6',control.groups[0].total.net+100);raw.getRange('F6').values=[[row]];
 const cost=raw.getRange('G6').values[0][0];raw.getRange('G6').values=[[null]];w.recalculate();equalNumber(out,'H6','н/д');raw.getRange('G6').values=[[cost]];
 const bases=allocation.getRange('E6:F6').values;const pool=allocation.getRange('D6').values[0][0];allocation.getRange('E6:F6').values=[[0,0]];w.recalculate();equalNumber(out,'H6',control.groups[0].total.net);equalNumber(allocation,'J6',pool);allocation.getRange('E6:F6').values=bases;
 const firstBase=allocation.getRange('E6').values[0][0];allocation.getRange('E6').values=[[null]];w.recalculate();equalNumber(out,`C${detailStart+6}`,'Нет распределения');allocation.getRange('E6').values=[[firstBase]];
 }
}else equalNumber(out,'H6','н/д');
w.recalculate();
const errors=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:30},summary:'Final formula scan'});
console.log(errors.ndjson);
await fs.mkdir(path.dirname(outputFile),{recursive:true});
if(process.env.PNL_PREVIEW_DIR){
 await fs.mkdir(process.env.PNL_PREVIEW_DIR,{recursive:true});
 for(const [sheetName,range]of [['P&L',`B2:K${detailStart+6}`],['P&L',`B${statementStart}:F${statementStart+statement.length}`],['Данные','B2:K12'],['Данные','L5:S12'],['Распределение',`B2:L${alast}`],['Расчёт','B2:J12'],['Расчёт','K5:V12'],['Проверки','B2:E13']]){
 const preview=await w.render({sheetName,range,scale:1.3,format:'png'});await fs.writeFile(path.join(process.env.PNL_PREVIEW_DIR,`${sheetName}-${range.replace(':','-')}.png`),new Uint8Array(await preview.arrayBuffer()));
 }
}
const output=await SpreadsheetFile.exportXlsx(w);await output.save(outputFile);
console.log(JSON.stringify({file:outputFile,mode:blank?'blank':'demo',periods:input.periods.length,sourceRows:ordered.length,totalRow:summaryEnd,formulaChecks:'passed'}));
