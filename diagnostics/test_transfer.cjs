// Profile transfer against a simulated DAD143: 0x2300 writes are rejected ("local control") unless the
// previous write on the device was the TAC unlock 0x2300:0x03, as seen on hardware on 2026-10-07.
const vm=require('node:vm'),fs=require('node:fs'),assert=require('node:assert/strict');
const elements=new Map();
function el(id){if(!elements.has(id))elements.set(id,{value:'',textContent:'',innerHTML:'',style:{},classList:{add(){},remove(){}},addEventListener(){},checked:false,scrollHeight:0});return elements.get(id);}
let dev=0,mem={},unlocked={},calls=[],confirms=[],answers=[],tacBroken=false,zeroHangs=new Set(),hung=new Set(),gross={},zeroRejects=0;
const key=(i,s)=>`${i}:${s}`;
const context=vm.createContext({console,Date,Math,Number,String,Array,JSON,Error,Promise,performance,Blob,URL,
  setInterval:()=>1,clearInterval(){},setTimeout:f=>{f();return 1},
  document:{getElementById:el,querySelectorAll:()=>[],querySelector:()=>null,createElement:()=>({click(){}})},
  window:{addEventListener(){}},navigator:{sendBeacon(){}},alert(){},
  confirm:m=>{confirms.push(m);return answers.length?answers.shift():true},
  fetch:async(p,o={})=>{const q=JSON.parse(o.body||'{}');calls.push({dev,p,q});
    const ok=v=>({ok:true,status:200,json:async()=>({ok:true,value:v})}),bad=e=>({ok:false,status:502,json:async()=>({ok:false,error:e})});
    if(p==='/api/select-device'){dev=q.address;mem[dev]??={[key(0x2300,3)]:7,[key(0x2100,0x0A)]:5,[key(0x2100,0x0B)]:5000};return ok()}
    if(hung.has(dev))return bad('SDO Read: HTTP 500: Ecat: Timeout');
    if(p==='/api/dashboard'){const g=gross[dev];return {ok:true,status:200,json:async()=>({ok:true,gross:Array.isArray(g)?g.shift()??'0.0':(g??'0.0')})}}
    const m=mem[dev];
    // Gross weight noise shrinks with the filter level FL (0x2100:04).
    if(p==='/api/read'&&q.index===0x2900&&q.sub===1){const a={1:0.8,2:0.5,3:0.3,4:0.2,5:0.1,6:0.05,7:0.02,8:0.01}[m[key(0x2100,4)]]??1;m.t=(m.t||0)+1;const sp=m.spike&&m[key(0x2100,4)]===m.spike.fl&&!m.spike.done&&m.t%7===0?(m.spike.n=(m.spike.n||0)+1,m.spike.n>=3&&(m.spike.done=true),5):0;return ok(([a,-a,-a,a][m.t%4])+(m.drift||0)*m.t+sp)}
    if(p==='/api/read')return ok(m[key(q.index,q.sub)]??0);
    if(p==='/api/write'){const u=unlocked[dev];unlocked[dev]=false;
      if(q.index===0x2300&&q.sub===3){unlocked[dev]=!tacBroken;return ok()}
      if(q.index===0x2300&&q.sub===0x0A&&dev===zeroRejects&&u)return bad('SDO Write: HTTP 500: Ecat SDO: General error');
      if(q.index===0x2300&&q.sub===0x0A&&zeroHangs.has(dev)&&u){hung.add(dev);return bad('SDO Write: HTTP 500: Ecat: Timeout')}
      if((q.index===0x2300||(q.index===0x2004&&q.sub===2))&&!u)return bad('SDO Write: HTTP 500: Ecat SDO: Data cannot be transferred (local control)');
      m[key(q.index,q.sub)]=Number(q.value);if(q.index===0x2300&&q.sub===0x0A)m[key(0x2300,2)]=1313+dev%10;return ok()}
    return ok();
  }});
const run=s=>vm.runInContext(s,context);
run(fs.readFileSync('source/app.js','utf8').replace(/L\('START[^;]*;/,'').replace('init().catch(e=>M(e.message));',''));
run('__f9=f9');
const sub=[1,2,3,4,7,8,0x0A,0x0B,0x11];
const P=[[0x2100,4,'Filter setting'],[0x2100,9,'Filter mode'],[0x2100,0x0A,'No motion range'],[0x2100,0x0B,'No motion time'],[0x2300,0x0C,'Display step size'],[0x2900,1,'Gross weight'],
  ...[1,2,3,4,5].map(s=>[0x2004,s,'Save '+s])].map(([Index,SubIndex,EntryName])=>({Index,SubIndex,EntryName})).concat([
  ...sub.map(s=>({Index:0x2300,SubIndex:s,EntryName:{3:'TAC',4:'Gain',0x0A:'Zero'}[s]||`Setting ${s}`}))])
  .map(p=>({...p,ValueBytes:4,DataTypeName:'INTEGER32',ReadOp:true,WriteOp:true,IndexHex:'0x'+p.Index.toString(16),SubIndexHex:'0x'+p.SubIndex.toString(16)}));
context.__P=P;
// Profile exactly as saved by v1.9.4 (still contains calibration commands 0x04 and 0x0A).
const prof={formatVersion:1,parameters:[{index:0x2100,subIndex:4,name:'Filter setting',value:4,category:'cfg'},
  ...[[1,1654],[2,1219],[4,4000],[7,15000],[8,-10000],[0x0A,32588],[0x0B,1],[0x11,300]].map(([s,v])=>({index:0x2300,subIndex:s,name:{4:'Calibrate Gain',0x0A:'Calibrate Zero'}[s]||`Setting ${s}`,value:v,category:'cal'}))]};
context.__F=prof;
async function transfer(devs,ans){calls=[];confirms=[];answers=[...ans];
  run(`P=__P;F=__F;MC=true;C=true;DA=0;B=[];R=()=>{};V=()=>{};f9=async()=>{};DS=()=>${JSON.stringify(devs)}`);
  el('ic').checked=true;el('ik').checked=true;await run('f34()');return el('pd').textContent}
(async()=>{
  // 1. Normal transfer with EEPROM save.
  let out=await transfer([1013,1014],[true,true,false,true]);
  const w=calls.filter(c=>c.p==='/api/write');
  assert.ok(!w.some(c=>c.q.index===0x2300&&(c.q.sub===4||c.q.sub===0x0A)),'calibration commands must never be written');
  assert.match(confirms[0],/NICHT übertragen[\s\S]*0x2300:0x02[\s\S]*0x2300:0x04[\s\S]*0x2300:0x0A/);
  for(let i=0;i<w.length;i++)if(w[i].q.index===0x2300&&w[i].q.sub!==3)assert.deepEqual([w[i-1].q.index,w[i-1].q.sub],[0x2300,3],'TAC unlock directly before each calibration write');
  for(const d of [1013,1014]){assert.equal(mem[d][key(0x2300,7)],15000);assert.equal(mem[d][key(0x2300,0x11)],300);assert.equal(mem[d][key(0x2300,4)],undefined);assert.equal(mem[d][key(0x2004,2)],0)}
  assert.match(out,/1013: 6 OK · 0 Fehler · EEPROM gespeichert \(Setup, Kalibrierung\)/);
  // Each written group is saved with its own command (manual 9.12): setup WP 0x2004:03 for 0x2100, calibration CS 0x2004:02 last.
  const sv=w.filter(c=>c.dev===1013&&c.q.index===0x2004).map(c=>c.q.sub);assert.deepEqual(sv,[3,2]);
  assert.ok(!w.some(c=>c.q.index===0x2300&&c.q.sub===2),'Absolute zero belongs to the target scale');
  // 2. Unlock not accepted: calibration block stops after the first error, no EEPROM save.
  tacBroken=true;mem={};out=await transfer([1020],[true,true,true,true]);
  const c2=calls.filter(c=>c.p==='/api/write'&&c.q.index===0x2300&&c.q.sub!==3);
  assert.equal(c2.length,1,'only one failing calibration write per device');
  assert.match(out,/1020: 1 OK · 1 Fehler · 4 übersprungen · Nullpunkt übersprungen · NICHT gespeichert/);
  assert.match(out,/local control/);
  assert.ok(!calls.some(c=>c.p==='/api/write'&&c.q.index===0x2004));
  tacBroken=false;
  // 3. Without EEPROM confirmation nothing is stored persistently.
  mem={};out=await transfer([1021],[true,true,false,false]);
  assert.ok(!calls.some(c=>c.p==='/api/write'&&c.q.index===0x2004));assert.match(out,/nicht dauerhaft gespeichert/);
  // 4. Zero calibration on request: runs once per device with value 0, after the settings, with a TAC unlock directly before it, then EEPROM save.
  mem={};out=await transfer([1023,1027],[true,true,true,true]);
  assert.match(confirms[2],/Nullpunkt[\s\S]*entlastet/);
  const w4=calls.filter(c=>c.p==='/api/write');
  const z=w4.map((c,i)=>[c,i]).filter(([c])=>c.q.index===0x2300&&c.q.sub===0x0A);
  assert.equal(z.length,2);for(const[c,i]of z){assert.equal(c.q.value,'0');assert.deepEqual([w4[i-1].q.index,w4[i-1].q.sub],[0x2300,3]);assert.deepEqual(w4.slice(i+1,i+4).map(c=>[c.q.index,c.q.sub]),[[0x2004,3],[0x2300,3],[0x2004,2]])}
  assert.ok(!w4.some(c=>c.q.index===0x2300&&c.q.sub===4),'gain is never calibrated by the transfer');
  // The zero calibration rewrites Absolute zero (0x2300:0x02); that is not a mismatch (seen on hardware: 1219 -> 1313).
  assert.match(out,/1023: 6 OK · 0 Fehler · Nullpunkt kalibriert \(Absolute zero 0→1316\) · EEPROM gespeichert/);
  assert.equal(mem[1027][key(0x2300,2)],1320);
  // 5. Weight not at rest: the zero calibration command is not sent, settings are still saved.
  mem={};gross={1030:['0.0','0.4','0.1','0.0','0.3','0.0']};out=await transfer([1030],[true,true,true,true]);
  assert.ok(!calls.some(c=>c.p==='/api/write'&&c.q.sub===0x0A),'no zero command while the weight moves');
  assert.match(out,/1030: 6 OK · 0 Fehler · Nullpunkt NICHT kalibriert – Waage nicht ruhig \(Brutto schwankt: 0 0.4 0.1 0 0.3 0 [^)]*Spanne 0.4, Grenze 0.2, NR 5 d, NT 5000 ms\)\) · EEPROM gespeichert/);
  assert.ok(calls.filter(c=>c.p==='/api/dashboard').length>=15,'observes the device no-motion time NT (5000 ms)');
  assert.match(el('log').textContent,/NULLPUNKT 1030 \| Brutto schwankt/);
  // Steady within two steps of the last digit: zero command is sent.
  mem={};gross={1031:['0.1','0.2','0.1','0.0','0.2','0.1']};out=await transfer([1031],[true,true,true,true]);
  assert.ok(calls.some(c=>c.p==='/api/write'&&c.q.sub===0x0A));assert.match(out,/1031: 6 OK · 0 Fehler · Nullpunkt kalibriert/);
  // Gross not readable: no zero command.
  mem={};gross={1034:'—'};out=await transfer([1034],[true,true,true,true]);
  assert.ok(!calls.some(c=>c.p==='/api/write'&&c.q.sub===0x0A));assert.match(out,/Brutto nicht lesbar/);
  gross={};
  // 5a. Zero calibration rejected by the device (hardware 1023: "Ecat SDO: General error" after 33 ms): that device is not saved, the run continues.
  mem={};zeroRejects=1023;out=await transfer([1023,1027],[true,true,true,true]);zeroRejects=0;
  assert.match(out,/1023: 6 OK · 1 Fehler · Nullpunkt-Fehler · NICHT gespeichert/);assert.match(out,/General error/);
  assert.match(out,/1027: 6 OK · 0 Fehler · Nullpunkt kalibriert[^\n]*EEPROM gespeichert/);assert.doesNotMatch(out,/ÜBERTRAGUNG BEENDET/);
  // 5b. Zero calibration times out (hardware 1015): no further requests to that device, it is skipped, the run continues and it can be retried.
  mem={};zeroHangs=new Set([1015]);out=await transfer([1014,1015,1016],[true,true,true,true]);
  const zi=calls.findIndex(c=>c.dev===1015&&c.p==='/api/write'&&c.q.sub===0x0A);
  assert.ok(zi>0);assert.equal(calls.slice(zi+1).filter(c=>c.dev===1015&&c.p!=='/api/select-device').length,0,'no requests after the hang');
  assert.match(out,/1014: 6 OK · 0 Fehler · Nullpunkt kalibriert[^\n]*EEPROM gespeichert/);
  assert.match(out,/1015: 6 OK · 1 Fehler · Nullpunkt-Fehler · Gerät antwortet nicht mehr, übersprungen/);
  assert.match(out,/1016: 6 OK · 0 Fehler · Nullpunkt kalibriert[^\n]*EEPROM gespeichert/);
  assert.match(out,/Fehlgeschlagen\/übersprungen: 1015\n/);assert.equal(el('rt').disabled,false);assert.match(el('rt').textContent,/\(1\)/);
  // Manual retry after the scale was restarted: only 1015 is processed again.
  zeroHangs.clear();hung.clear();calls=[];confirms=[];answers=[true,true,true,true];await run('f36()');out=el('pd').textContent;
  assert.deepEqual([...new Set(calls.filter(c=>c.p==='/api/select-device').map(c=>c.q.address))],[1015]);
  assert.match(out,/1015: 6 OK · 0 Fehler · Nullpunkt kalibriert[^\n]*EEPROM gespeichert/);assert.equal(el('rt').disabled,true);
  // Three scales without answer to the zero calibration in one run: general problem, the rest is not touched.
  mem={};zeroHangs=new Set([1020,1021,1022]);out=await transfer([1020,1021,1022,1023],[true,true,true,true]);
  assert.ok(!calls.some(c=>c.dev===1023));assert.match(out,/ÜBERTRAGUNG BEENDET: 3 Waagen ohne Antwort[^\n]*Nicht bearbeitet: 1023/);
  assert.match(out,/Fehlgeschlagen\/übersprungen: 1020, 1021, 1022, 1023/);
  zeroHangs.clear();hung.clear();
  // 6. Cancel on the calibration confirmation writes nothing.
  mem={};await transfer([1022],[false]);assert.ok(!calls.some(c=>c.p==='/api/write'));
  // 7. No second dashboard poll or bus scan while one is still pending.
  calls=[];run('f9=__f9;C=true;MC=true;Z=null');await Promise.all([run('f9()'),run('f9()'),run('f25()'),run('f25()')]);
  assert.equal(calls.filter(c=>c.p==='/api/dashboard').length,1);assert.equal(calls.filter(c=>c.p.startsWith('/api/bus')).length,1);
  calls=[];await run('f9()');assert.equal(calls.length,1,'guard released after completion');
  // 8. New profiles do not contain calibration commands at all.
  run('P=__P');assert.deepEqual(run("Q(0,1).map(p=>p.SubIndex)"),[0x0C,1,7,8,0x0B,0x11]);
  // 9. Filter optimisation: weakest filter level within the target is proposed, saved with WP only after confirmation, otherwise restored.
  dev=1050;mem={1050:{[key(0x2100,4)]:4,[key(0x2100,9)]:0,[key(0x2300,0x0B)]:1,[key(0x2300,0x0C)]:1}};run('P=__P;C=true;DA=1050');
  el('fo_t').value='1';el('fo_s').value='3';el('fo_m').value='0';calls=[];confirms=[];answers=[true,true];await run('FO()');
  assert.match(el('fo').textContent,/Vorschlag: FL 6/);assert.equal(mem[1050][key(0x2100,4)],6);
  assert.deepEqual(calls.filter(c=>c.p==='/api/write'&&c.q.index===0x2004).map(c=>c.q.sub),[3],'only the setup group is saved');
  assert.match(el('log').textContent,/FILTER 1050 \| FM 0 FL 5 \| Unruhe 2.0 d/);assert.match(confirms[1],/FL 6 übernehmen/);
  mem[1050][key(0x2100,4)]=4;calls=[];answers=[true,false];await run('FO()');
  assert.equal(mem[1050][key(0x2100,4)],4,'restored');assert.ok(!calls.some(c=>c.p==='/api/write'&&c.q.index===0x2004));assert.match(el('fo').textContent,/Wiederhergestellt: FM 0 \/ FL 4/);
  // Nothing reaches the target: FL 8 proposed with a hint at mechanical causes.
  el('fo_t').value='1';mem[1050][key(0x2100,4)]=4;mem[1050][key(0x2300,0x0B)]=3;answers=[true];calls=[];confirms=[];await run('FO()');assert.match(el('fo').textContent,/Kein Filter erreicht 1 d[^\n]*keine Änderung/);
  assert.equal(confirms.filter(c=>/übernehmen/.test(c)).length,0,'no adoption offered without a suitable level');assert.equal(mem[1050][key(0x2100,4)],4);
  // Slow drift and a disturbance in one pass (seen on 1022) do not decide: detrended noise, the better pass counts.
  mem[1050][key(0x2300,0x0B)]=1;mem[1050][key(0x2100,4)]=4;mem[1050].drift=0.002;mem[1050].spike={fl:6};answers=[true,true];await run('FO()');
  assert.match(el('fo').textContent,/Vorschlag: FL 6/);assert.match(el('fo').textContent,/\n 6 [^\n]*!/);assert.equal(mem[1050][key(0x2100,4)],6);
  mem[1050].drift=0;mem[1050].spike=null;
  console.log('PASS: calibration commands excluded, TAC unlock per calibration write, stop after first error, read-back, EEPROM save only on clean devices, zero calibration only on request, zero only when the weight is at rest, skip failed devices and retry them, stop after three zero calibrations without answer, cancel writes nothing, no overlapping polls/scans, group-wise EEPROM save, filter optimisation');
})().catch(e=>{console.error(e);process.exitCode=1;});
