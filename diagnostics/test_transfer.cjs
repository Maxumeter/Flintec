// Profile transfer against a simulated DAD143: 0x2300 writes are rejected ("local control") unless the
// previous write on the device was the TAC unlock 0x2300:0x03, as seen on hardware on 2026-10-07.
const vm=require('node:vm'),fs=require('node:fs'),assert=require('node:assert/strict');
const elements=new Map();
function el(id){if(!elements.has(id))elements.set(id,{value:'',textContent:'',innerHTML:'',style:{},classList:{add(){},remove(){}},addEventListener(){},checked:false,scrollHeight:0});return elements.get(id);}
let dev=0,mem={},unlocked={},calls=[],confirms=[],answers=[],tacBroken=false,zeroHangs=0,hung=new Set();
const key=(i,s)=>`${i}:${s}`;
const context=vm.createContext({console,Date,Math,Number,String,Array,JSON,Error,Promise,performance,Blob,URL,
  setInterval:()=>1,clearInterval(){},setTimeout:()=>1,
  document:{getElementById:el,querySelectorAll:()=>[],querySelector:()=>null,createElement:()=>({click(){}})},
  window:{addEventListener(){}},navigator:{sendBeacon(){}},alert(){},
  confirm:m=>{confirms.push(m);return answers.length?answers.shift():true},
  fetch:async(p,o={})=>{const q=JSON.parse(o.body||'{}');calls.push({dev,p,q});
    const ok=v=>({ok:true,status:200,json:async()=>({ok:true,value:v})}),bad=e=>({ok:false,status:502,json:async()=>({ok:false,error:e})});
    if(p==='/api/select-device'){dev=q.address;mem[dev]??={[key(0x2300,3)]:7};return ok()}
    if(hung.has(dev))return bad('SDO Read: HTTP 500: Ecat: Timeout');
    const m=mem[dev];
    if(p==='/api/read')return ok(m[key(q.index,q.sub)]??0);
    if(p==='/api/write'){const u=unlocked[dev];unlocked[dev]=false;
      if(q.index===0x2300&&q.sub===3){unlocked[dev]=!tacBroken;return ok()}
      if(q.index===0x2300&&q.sub===0x0A&&dev===zeroHangs&&u){hung.add(dev);return bad('SDO Write: HTTP 500: Ecat: Timeout')}
      if((q.index===0x2300||(q.index===0x2004&&q.sub===2))&&!u)return bad('SDO Write: HTTP 500: Ecat SDO: Data cannot be transferred (local control)');
      m[key(q.index,q.sub)]=Number(q.value);if(q.index===0x2300&&q.sub===0x0A)m[key(0x2300,2)]=1313+dev%10;return ok()}
    return ok();
  }});
const run=s=>vm.runInContext(s,context);
run(fs.readFileSync('source/app.js','utf8').replace(/L\('START[^;]*;/,'').replace('init().catch(e=>M(e.message));',''));
run('__f9=f9');
const sub=[1,2,3,4,7,8,0x0A,0x0B,0x11];
const P=[{Index:0x2100,SubIndex:1,EntryName:'Filter'},{Index:0x2004,SubIndex:2,EntryName:'EEPROM'},
  ...sub.map(s=>({Index:0x2300,SubIndex:s,EntryName:{3:'TAC',4:'Gain',0x0A:'Zero'}[s]||`Setting ${s}`}))]
  .map(p=>({...p,ValueBytes:4,DataTypeName:'INTEGER32',ReadOp:true,WriteOp:true,IndexHex:'0x'+p.Index.toString(16),SubIndexHex:'0x'+p.SubIndex.toString(16)}));
context.__P=P;
// Profile exactly as saved by v1.9.4 (still contains calibration commands 0x04 and 0x0A).
const prof={formatVersion:1,parameters:[{index:0x2100,subIndex:1,name:'Filter',value:4,category:'cfg'},
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
  assert.match(confirms[0],/NICHT übertragen[\s\S]*0x2300:0x04[\s\S]*0x2300:0x0A/);
  for(let i=0;i<w.length;i++)if(w[i].q.index===0x2300&&w[i].q.sub!==3)assert.deepEqual([w[i-1].q.index,w[i-1].q.sub],[0x2300,3],'TAC unlock directly before each calibration write');
  for(const d of [1013,1014]){assert.equal(mem[d][key(0x2300,7)],15000);assert.equal(mem[d][key(0x2300,0x11)],300);assert.equal(mem[d][key(0x2300,4)],undefined);assert.equal(mem[d][key(0x2004,2)],0)}
  assert.match(out,/1013: 7 OK · 0 Fehler · EEPROM gespeichert/);
  // 2. Unlock not accepted: calibration block stops after the first error, no EEPROM save.
  tacBroken=true;mem={};out=await transfer([1020],[true,true,true,true]);
  const c2=calls.filter(c=>c.p==='/api/write'&&c.q.index===0x2300&&c.q.sub!==3);
  assert.equal(c2.length,1,'only one failing calibration write per device');
  assert.match(out,/1020: 1 OK · 1 Fehler · 5 übersprungen · Nullpunkt übersprungen · NICHT gespeichert/);
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
  assert.equal(z.length,2);for(const[c,i]of z){assert.equal(c.q.value,'0');assert.deepEqual([w4[i-1].q.index,w4[i-1].q.sub],[0x2300,3]);assert.deepEqual([w4[i+1].q.index,w4[i+1].q.sub,w4[i+2].q.index],[0x2300,3,0x2004])}
  assert.ok(!w4.some(c=>c.q.index===0x2300&&c.q.sub===4),'gain is never calibrated by the transfer');
  // The zero calibration rewrites Absolute zero (0x2300:0x02); that is not a mismatch (seen on hardware: 1219 -> 1313).
  assert.match(out,/1023: 7 OK · 0 Fehler · Nullpunkt kalibriert \(Absolute zero 1219→1316\) · EEPROM gespeichert/);
  assert.equal(mem[1027][key(0x2300,2)],1320);
  // 5. Zero calibration times out (hardware 1015): no further requests to that device, nothing saved, remaining devices untouched.
  mem={};zeroHangs=1015;out=await transfer([1014,1015,1016],[true,true,true,true]);
  const zi=calls.findIndex(c=>c.dev===1015&&c.p==='/api/write'&&c.q.sub===0x0A);
  assert.ok(zi>0);assert.equal(calls.slice(zi+1).filter(c=>c.dev===1015&&c.p!=='/api/select-device').length,0,'no requests after the hang');
  assert.ok(!calls.some(c=>c.dev===1016),'devices after a failed zero calibration are not touched');
  assert.match(out,/1014: 7 OK · 0 Fehler · Nullpunkt kalibriert[^\n]*EEPROM gespeichert/);
  assert.match(out,/1015: 7 OK · 1 Fehler · Nullpunkt-Fehler · Gerät antwortet nicht mehr/);
  assert.match(out,/ÜBERTRAGUNG BEENDET[^\n]*1015[^\n]*Nicht bearbeitet: 1016/);
  zeroHangs=0;hung.clear();
  // 6. Cancel on the calibration confirmation writes nothing.
  mem={};await transfer([1022],[false]);assert.ok(!calls.some(c=>c.p==='/api/write'));
  // 7. No second dashboard poll or bus scan while one is still pending.
  calls=[];run('f9=__f9;C=true;MC=true;Z=null');await Promise.all([run('f9()'),run('f9()'),run('f25()'),run('f25()')]);
  assert.equal(calls.filter(c=>c.p==='/api/dashboard').length,1);assert.equal(calls.filter(c=>c.p.startsWith('/api/bus')).length,1);
  calls=[];await run('f9()');assert.equal(calls.length,1,'guard released after completion');
  // 8. New profiles do not contain calibration commands at all.
  run('P=__P');assert.deepEqual(run("Q(0,1).map(p=>p.SubIndex)"),[1,2,7,8,0x0B,0x11]);
  console.log('PASS: calibration commands excluded, TAC unlock per calibration write, stop after first error, read-back, EEPROM save only on clean devices, zero calibration only on request, stop after a failed zero calibration, cancel writes nothing, no overlapping polls/scans');
})().catch(e=>{console.error(e);process.exitCode=1;});
