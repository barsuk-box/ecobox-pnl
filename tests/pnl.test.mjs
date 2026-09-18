import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {spawnSync} from 'node:child_process';
import {fileURLToPath} from 'node:url';
import {calculate,variance,inventoryBridge,compareReports,FIELDS} from '../scripts/pnl.mjs';
const fixture=JSON.parse(await fs.readFile(new URL('../examples/demo.json',import.meta.url),'utf8'));
const expected=JSON.parse(await fs.readFile(new URL('../examples/expected.json',import.meta.url),'utf8'));
const clone=()=>structuredClone(fixture);
const actual=r=>r.groups.find(g=>g.period==='2026-01'&&g.scenario==='actual');
test('full demo matches independent monthly and total controls',()=>{
  const result=calculate(fixture); assert.equal(result.status,'calculated'); assert.deepEqual(result.issues,[]);
  for(const exp of expected.actual){ const g=result.groups.find(x=>x.period===exp.period&&x.scenario==='actual'); for(const [k,v] of Object.entries(exp))if(k!=='period')assert.equal(g.total[k],v,k); }
  for(const scenario of ['actual','budget']){const total=result.totals.find(x=>x.scenario===scenario);for(const [k,v]of Object.entries(expected[`${scenario}_total`]))assert.equal(total[k],v,k);}
  assert.equal(result.variance.at(-1).net.amount,55000);
  assert.equal(result.totals[0].margin,1060000/5100000);
});
test('inventory bridge uses opening and closing inventories',()=>assert.deepEqual(inventoryBridge({wip_open:100,production_cost:600,wip_close:150,fg_open:200,fg_close:250}),{production:550,cogs:500}));
test('pool allocation preserves consolidated result',()=>{
  const g=actual(calculate(fixture)); assert.deepEqual(g.rows.map(r=>r.allocation),[80000,40000,-120000,0]); assert.equal(g.allocation_residual,0);
});
test('zero bases retain pool in center',()=>{const d=clone();d.allocations[0].bases={A:0,B:0};const r=calculate(d);assert.equal(r.status,'preliminary');assert.equal(actual(r).allocation_residual,120000);assert.equal(actual(r).total.net,335000);});
test('missing COGS propagates to profit without hiding known revenue',()=>{const d=clone();d.records[0].cogs=null;const r=calculate(d);assert.equal(r.status,'incomplete');assert.equal(actual(r).total.net,null);assert.equal(actual(r).total.revenue,1500000);});
test('missing monthly row does not become zero',()=>{const d=clone();d.records.shift();const r=calculate(d);assert.equal(r.status,'incomplete');assert.equal(actual(r).total.revenue,null);});
test('duplicate snapshot blocks totals',()=>{const d=clone();d.records.push({...d.records[0],id:'different'});const r=calculate(d);assert.equal(r.status,'invalid');assert.equal(r.groups.length,0);});
test('duplicate source identity blocks totals',()=>{const d=clone();d.records[1].source_ref=d.records[0].source_ref;assert.equal(calculate(d).status,'invalid');});
test('unmapped fields are not silently discarded',()=>{const d=clone();d.records[0].unmapped_cost=100;assert.equal(calculate(d).status,'invalid');});
test('bank cash basis is rejected',()=>{const d=clone();d.basis='cash';assert.throws(()=>calculate(d),/basis=accrual/);});
test('amounts reject numeric strings, nonfinite values and sub-kopecks',()=>{for(const v of ['100',NaN,Infinity,1.234,true]){const d=clone();d.records[0].revenue=v;assert.equal(calculate(d).status,'invalid');}});
test('missing D&A affects EBITDA only',()=>{const d=clone();d.records[0].da=null;const r=actual(calculate(d));assert.equal(r.total.ebitda,null);assert.equal(r.total.ebit,425000);});
test('zero and negative budget denominators',()=>{assert.deepEqual(variance(25,0),{amount:25,percent:null});assert.deepEqual(variance(-5,-10),{amount:5,percent:0.5});});
test('negative or missing drivers block calculation',()=>{for(const value of [-1,null]){const d=clone();d.allocations[0].bases.A=value;assert.equal(calculate(d).status,'invalid');}});
test('missing allocation is explicit and propagates',()=>{const d=clone();d.allocations.shift();const r=calculate(d);assert.equal(r.status,'incomplete');assert.equal(actual(r).total.net,null);});
test('rounding preserves pennies and zero driver receives zero',()=>{const d=clone();d.records[2].admin=0.01;d.allocations[0].bases={A:1,B:0};const g=actual(calculate(d));assert.deepEqual(g.rows.map(r=>r.allocation),[0.01,0,-0.01,0]);});
test('internal transaction correction removes unrealized profit',()=>{
  const d=clone();for(const row of d.records)for(const f of FIELDS)row[f]=0;
  d.records[0].revenue=100;d.records[0].cogs=60;d.records[1].revenue=80;d.records[1].cogs=50;d.records[3].revenue=-100;d.records[3].cogs=-80;
  const g=actual(calculate(d));assert.equal(g.total.revenue,80);assert.equal(g.total.cogs,30);assert.equal(g.total.net,50);
});
test('repeated and reordered input produce the same financial results',()=>{const d=clone();const before=JSON.stringify(d);const a=calculate(d);assert.equal(JSON.stringify(d),before);assert.deepEqual(calculate(d),a);d.records.reverse();assert.deepEqual(calculate(d),a);assert.deepEqual(compareReports(a,a),[]);});
test('report comparison pinpoints changed month and measures',()=>{const d=clone();const a=calculate(d);d.records[0].revenue+=100;const b=calculate(d);const changes=compareReports(a,b);assert(changes.length>0);assert(changes.every(x=>x.key[0]==='2026-01'&&x.key[1]==='actual'));assert(changes.some(x=>x.key[2]==='net'&&x.after-x.before===100));});
test('unclosed month remains preliminary',()=>{const d=clone();d.records[0].closed=false;assert.equal(calculate(d).status,'preliminary');});
test('appending a complete month extends totals',()=>{const d=clone();d.periods.push('2026-04');d.records.push(...d.records.filter(r=>r.period==='2026-03').map(r=>({...r,period:'2026-04',id:'new-'+r.id,source_ref:'new-'+r.source_ref})));d.allocations.push(...d.allocations.filter(r=>r.period==='2026-03').map(r=>({...r,period:'2026-04'})));const r=calculate(d);assert.equal(r.status,'calculated');assert.equal(r.totals[0].net,1445000);});
test('cumulative rounding stays nonnegative for four small shares',()=>{
 const d=clone();d.sites.push('C','D');for(const a of d.allocations){a.bases={A:1,B:1,C:1,D:1};}
 for(const period of d.periods)for(const scenario of d.scenarios)for(const scope of ['C','D'])d.records.push({...d.records.find(r=>r.period===period&&r.scenario===scenario&&r.scope==='ELIM'),scope,id:`${period}-${scenario}-${scope}`,source_ref:`synthetic:${period}-${scenario}-${scope}`});
 d.records[2].admin=0.02;const rows=actual(calculate(d)).rows;assert.deepEqual(rows.slice(0,4).map(r=>r.allocation),[0.01,0,0.01,0]);assert.equal(rows[4].allocation,-0.02);
});
test('negative corporate pool remains incomplete',()=>{const d=clone();d.records[2].admin=-100;const r=calculate(d);assert.equal(r.status,'incomplete');assert.equal(actual(r).total.net,null);});
test('duplicate dimensions and invalid dates are rejected',()=>{const d=clone();d.periods.push('2026-01');assert.throws(()=>calculate(d),/duplicate/);d.periods=['2026-13'];assert.throws(()=>calculate(d),/period/);});
test('CLI preserves existing files and writes incomplete diagnostics with code 2',async()=>{
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'ecobox-pnl-test-'));
 const script=fileURLToPath(new URL('../scripts/pnl.mjs',import.meta.url));
 try{
  const input=path.join(dir,'input.json'), output=path.join(dir,'report.json');
  await fs.writeFile(input,JSON.stringify(fixture));
  let result=spawnSync(process.execPath,[script,'calculate',input,output],{encoding:'utf8'});assert.equal(result.status,0,result.stderr);
  const old=await fs.readFile(output,'utf8');assert.equal(JSON.parse(old).input_sha256.length,64);
  result=spawnSync(process.execPath,[script,'calculate',input,output],{encoding:'utf8'});assert.equal(result.status,1);assert.equal(await fs.readFile(output,'utf8'),old);
  const d=clone();d.records[0].cogs=null;await fs.writeFile(input,JSON.stringify(d));
  result=spawnSync(process.execPath,[script,'calculate',input,path.join(dir,'incomplete.json')],{encoding:'utf8'});assert.equal(result.status,2);assert.equal(JSON.parse(await fs.readFile(path.join(dir,'incomplete.json'),'utf8')).status,'incomplete');
 }finally{
  if(path.dirname(path.resolve(dir))!==path.resolve(os.tmpdir()) || !path.basename(dir).startsWith('ecobox-pnl-test-'))throw new Error('Unexpected cleanup target');
  await fs.rm(dir,{recursive:true,force:true});
 }
});
