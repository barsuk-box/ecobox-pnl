// Normalized monthly P&L. No network, database writes or external dependencies.
import fs from 'node:fs/promises';
import { pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';

export const FIELDS = ['revenue','cogs','selling','admin','other_income','other_expense','da','financial_income','financial_expense','tax'];
export const DERIVED = ['gross','ebit_before','allocation','ebit','ebitda','pretax','net','margin'];
const combine = (values, fn) => values.every(v => v !== null) ? fn(...values) : null;
const sum = values => combine(values, (...xs) => xs.reduce((a,b) => a+b,0));
const cents = n => Math.round(n*100);
const units = n => n === null ? null : n/100;
const safeMoney = n => typeof n === 'number' && Number.isFinite(n) && Math.abs(n) <= 1e9 && Math.abs(n*100-Math.round(n*100)) < 0.0001;
const key = r => JSON.stringify([r.period,r.scenario,r.scope]);
export function variance(actual, budget) {
  return {amount: actual === null || budget === null ? null : (cents(actual)-cents(budget))/100,
    percent: actual === null || budget === null || budget === 0 ? null : (actual-budget)/Math.abs(budget)};
}
export function inventoryBridge({wip_open,production_cost,wip_close,fg_open,fg_close}) {
  const inputs = [wip_open,production_cost,wip_close,fg_open,fg_close];
  if (!inputs.every(safeMoney)) throw new Error('Inventory bridge requires five known monetary inputs. Other movements must be handled separately.');
  const output = cents(wip_open)+cents(production_cost)-cents(wip_close);
  return {production:output/100,cogs:(cents(fg_open)+output-cents(fg_close))/100};
}
export function calculate(input) {
  const issues = [];
  const issue = (code, location, detail, severity='error') => issues.push({code,location,detail,severity});
  if (!input || input.schema_version !== 1 || input.currency !== 'RUB' || input.unit !== 1 || input.basis !== 'accrual') {
    throw new Error('Expected schema_version=1, currency=RUB, unit=1, basis=accrual. No automatic FX, scale or cash-to-accrual conversion.');
  }
  for (const field of ['periods','sites','scenarios','records','allocations']) if (!Array.isArray(input[field])) throw new Error(`Missing array: ${field}`);
  for (const field of ['periods','sites','scenarios']) {
    if (!input[field].length || input[field].some(x => typeof x !== 'string' || !x.trim()) || new Set(input[field]).size !== input[field].length) throw new Error(`Invalid or duplicate ${field}`);
  }
  if (input.sites.some(s => ['CENTER','ELIM','TOTAL'].includes(s))) throw new Error('Reserved site name');
  if (input.periods.some(p => !/^\d{4}-(0[1-9]|1[0-2])$/.test(p))) throw new Error('Invalid YYYY-MM period');
  if (input.scenarios.some(s => !['actual','budget','forecast'].includes(s))) throw new Error('Unsupported scenario');
  if (input.vat !== 'excluded') throw new Error('Input must already exclude recoverable VAT; see methodology.');
  const scopes = [...input.sites,'CENTER','ELIM'];
  if (input.periods.length*input.scenarios.length*scopes.length>1000 || input.records.length>1000) throw new Error('At most 1000 monthly snapshots per run');
  const expected = new Set(input.periods.flatMap(period => input.scenarios.flatMap(scenario => scopes.map(scope => key({period,scenario,scope})))));
  const byKey = new Map();
  const ids = new Set();
  const sourceRefs = new Set();
  for (const row of input.records) {
    if (!row || typeof row !== 'object' || Array.isArray(row)) throw new Error('Record must be an object');
    const k = key(row);
    const allowed = ['id','period','scenario','scope','source_ref','closed',...FIELDS];
    if (Object.keys(row).some(f => !allowed.includes(f))) issue('unknown_field', k, 'Unknown field; map it explicitly before calculation.');
    if (!expected.has(k)) { issue('outside_scope',k,'Unexpected period, scenario or scope'); continue; }
    if (byKey.has(k)) { issue('duplicate_key',k,'Monthly snapshot occurs twice'); continue; }
    byKey.set(k,row);
    for (const [field,set] of [['id',ids],['source_ref',sourceRefs]]) {
      if (typeof row[field] !== 'string' || !row[field].trim()) issue('missing_identity',k,field);
      else if (set.has(row[field])) issue('duplicate_identity',k,field);
      else set.add(row[field]);
    }
    if (typeof row.closed !== 'boolean') issue('missing_close_status',k,'closed must be true or false');
    else if (!row.closed) issue('open_period',k,'Period is not closed','warning');
    for (const field of FIELDS) {
      if (row[field] === null || row[field] === undefined) issue('missing_amount',k,field);
      else if (!safeMoney(row[field])) issue('invalid_amount',k,`${field}: numeric RUB with up to two decimal places required`);
    }
  }
  // Structural ambiguities block aggregation, rather than choosing a first match.
  const structural = issues.some(i => !['missing_amount','open_period'].includes(i.code));
  if (structural) return {status:'invalid',issues,groups:[],totals:[],variance:[]};
  const allocationMap = new Map();
  for (const row of input.allocations) {
    if (!row || typeof row !== 'object') throw new Error('Allocation must be an object');
    const k = JSON.stringify([row.period,row.scenario]);
    if (!input.periods.includes(row.period) || !input.scenarios.includes(row.scenario) || allocationMap.has(k)) {
      issue('invalid_allocation_key',k,'Unknown or duplicate allocation group'); continue;
    }
    allocationMap.set(k,row);
    if (!row.bases || Object.keys(row.bases).length !== input.sites.length || input.sites.some(s => !(s in row.bases))) issue('invalid_bases',k,'Exactly one driver per site required');
    else for (const [site,value] of Object.entries(row.bases)) if (typeof value !== 'number' || !Number.isFinite(value) || value<0 || value>1e9) issue('invalid_base',k,site);
  }
  if (issues.some(i => ['invalid_allocation_key','invalid_bases','invalid_base'].includes(i.code))) return {status:'invalid',issues,groups:[],totals:[],variance:[]};
  const groups = [];
  for (const period of input.periods) for (const scenario of input.scenarios) {
    const groupKey = JSON.stringify([period,scenario]);
    const raw = scopes.map(scope => {
      const k = key({period,scenario,scope});
      const row = byKey.get(k);
      if (!row) issue('missing_record',k,'Expected monthly snapshot is absent');
      return {scope,...Object.fromEntries(FIELDS.map(f => [f,row && safeMoney(row[f]) ? cents(row[f]) : null]))};
    });
    const center = raw.find(r => r.scope==='CENTER');
    const allocation = allocationMap.get(groupKey);
    if (!allocation) issue('missing_allocation',groupKey,'Explicit drivers required, including a zero pool');
    if (center.admin !== null && center.admin<0) issue('negative_pool',groupKey,'Negative corporate admin pool requires a separate policy');
    const drivers = allocation ? input.sites.map(s=>allocation.bases[s]) : [];
    const driverTotal = drivers.reduce((a,b)=>a+b,0);
    let distributed = 0;
    let cumulativeDriver = 0;
    const amounts = input.sites.map((s,index) => {
      if (!allocation || center.admin === null || center.admin<0) return null;
      let amount = 0;
      if (driverTotal>0) {
        // Cumulative rounding preserves the pool and cannot make a recipient negative.
        cumulativeDriver += drivers[index];
        amount = Math.round(center.admin*cumulativeDriver/driverTotal)-distributed;
        distributed += amount;
      }
      return amount;
    });
    const residual = combine([center.admin,...amounts],(pool,...xs)=>pool-xs.reduce((a,b)=>a+b,0));
    if (allocation && driverTotal===0 && center.admin>0) issue('unallocated_pool',groupKey,'Pool remains in CENTER','warning');
    const rows = raw.map(r => {
      const index = input.sites.indexOf(r.scope);
      const expense = index>=0 ? amounts[index] : r.scope==='CENTER' ? combine(amounts,(...xs)=>-xs.reduce((a,b)=>a+b,0)) : 0;
      const gross = combine([r.revenue,r.cogs],(a,b)=>a-b);
      const ebitBefore = combine([gross,r.selling,r.admin,r.other_income,r.other_expense],(g,s,a,i,e)=>g-s-a+i-e);
      const ebit = combine([ebitBefore,expense],(a,b)=>a-b);
      const ebitda = combine([ebit,r.da],(a,b)=>a+b);
      const pretax = combine([ebit,r.financial_income,r.financial_expense],(a,b,c)=>a+b-c);
      const net = combine([pretax,r.tax],(a,b)=>a-b);
      return {...r,gross,ebit_before:ebitBefore,allocation:expense,ebit,ebitda,pretax,net,margin:net===null || r.revenue===null || r.revenue===0 ? null : net/r.revenue};
    });
    const total = Object.fromEntries([...FIELDS,...DERIVED.filter(f=>f!=='margin')].map(f=>[f,sum(rows.map(r=>r[f]))]));
    total.margin = total.net===null || total.revenue===null || total.revenue===0 ? null : total.net/total.revenue;
    const convert = row=>Object.fromEntries(Object.entries(row).map(([f,v])=>[f,f==='scope'||f==='margin' ? v : units(v)]));
    groups.push({period,scenario,rows:rows.map(convert),total:convert(total),allocation_residual:units(residual)});
  }
  const totals = input.scenarios.map(scenario=>{
    const selected = groups.filter(g=>g.scenario===scenario);
    const total = Object.fromEntries([...FIELDS,...DERIVED.filter(f=>f!=='margin')].map(f=>[f,combine(selected.map(g=>g.total[f]),(...xs)=>xs.reduce((a,b)=>a+cents(b),0)/100)]));
    total.margin = total.net===null || total.revenue===null || total.revenue===0 ? null : total.net/total.revenue;
    return {scenario,...total};
  });
  const comparisons = [];
  if (input.scenarios.includes('actual') && input.scenarios.includes('budget')) for (const period of [...input.periods,'TOTAL']) {
    const get = scenario => period==='TOTAL' ? totals.find(g=>g.scenario===scenario) : groups.find(g=>g.period===period&&g.scenario===scenario).total;
    comparisons.push({period,net:variance(get('actual').net,get('budget').net),revenue:variance(get('actual').revenue,get('budget').revenue)});
  }
  return {status:issues.some(i=>i.severity==='error')?'incomplete':issues.length?'preliminary':'calculated',issues,groups,totals,variance:comparisons};
}

export function compareReports(before,after) {
  const flatten = report=>new Map((report.groups??[]).flatMap(g=>Object.entries(g.total).map(([field,value])=>[JSON.stringify([g.period,g.scenario,field]),value])));
  const a = flatten(before), b = flatten(after);
  return [...new Set([...a.keys(),...b.keys()])].sort().filter(k=>a.get(k)!==b.get(k)).map(k=>({key:JSON.parse(k),before:a.has(k)?a.get(k):null,after:b.has(k)?b.get(k):null}));
}
async function main() {
  const [command,source,destination] = process.argv.slice(2);
  if (command==='compare' && source && destination) {
    console.log(JSON.stringify(compareReports(JSON.parse(await fs.readFile(source,'utf8')),JSON.parse(await fs.readFile(destination,'utf8'))),null,2)); return;
  }
  if (command!=='calculate' || !source || !destination) throw new Error('Usage: node scripts/pnl.mjs calculate input.json report.json | compare before.json after.json');
  const bytes = await fs.readFile(source,'utf8');
  const report = calculate(JSON.parse(bytes));
  report.input_sha256 = createHash('sha256').update(bytes).digest('hex');
  await fs.writeFile(destination,JSON.stringify(report,null,2)+'\n',{flag:'wx'});
  console.log(`${report.status}: ${report.issues.length} issues. Output: ${destination}`);
  if (['invalid','incomplete'].includes(report.status)) process.exitCode=2;
}
if (process.argv[1] && import.meta.url===pathToFileURL(process.argv[1]).href) main().catch(e=>{console.error(e.message);process.exitCode=1;});
