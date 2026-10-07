// Profile transfer against a simulated DAD143: 0x2300 writes are rejected ("local control") unless the
// previous write on the device was the TAC unlock 0x2300:0x03, as seen on hardware on 2026-10-07.
const vm=require('node:vm'),fs=require('node:fs'),assert=require('node:assert/strict');
const elements=new Map();
function el(id){if(!elements.has(id))elements.set(id,{value:'',textContent:'',innerHTML:'',style:{},classList:{add(){},remove(){}},addEventListener(){},checked:false,scrollHeight:0});return elements.get(id);}
let dev=0,mem={},unlocked={},calls=[],confirms=[],answers=[],tacBroken=false;
const key=(i,s)=>`${i}:${s}`;
const context=vm.createContext({console,Date,Math,Number,String,Array,JSON,Error,Promise,performance,Blob,URL,
  setInterval:()=>1,clearInterval(){},setTimeout:()=>1,
  document:{getElementById:el,querySelectorAll:()=>[],querySelector:()=>null,createElement:()=>({click(){}})},
  window:{addEventListener(){}},navigator:{sendBeacon(){}},alert(){},
  confirm:m=>{confirms.push(m);return answers.length?answers.shift():true},
  fetch:async(p,o={})=>{const q=JSON.parse(o.body||'{}');calls.push({dev,p,q});
    const ok=v=>({ok:true,status:200,json:async()=>({ok:true,value:v})}),bad=e=>({ok:false,status:502,json:async()=>({ok:false,error:e})});
    if(p==='/api/select-device'){dev=q.address;mem[dev]??={[key(0x2300,3)]:7};return ok()}
    const m=mem[dev];
    if(p==='/api/read')return ok(m[key(q.index,q.sub)]??0);
    if(p==='/api/write'){const u=unlocked[dev];unlocked[dev]=false;
      if(q.index===0x2300&&q.sub===3){unlocked[dev]=!tacBroken;return ok()}
      if((q.index===0x2300||(q.index===0x2004&&q.sub===2))&&!u)return bad('SDO Write: HTTP 500: Ecat SDO: Data cannot be transferred (local control)');
      m[key(q.index,q.sub)]=Number(q.value);return ok()}
    return ok();
  }});
const run=s=>vm.runInContext(s,context);
run(fs.readFileSync('source/app.js','utf8').replace(/L\('START[^;]*;/,'').replace('init().catch(e=>M(e.message));',''));
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
  let out=await transfer([1013,1014],[true,true,true]);
  const w=calls.filter(c=>c.p==='/api/write');
  assert.ok(!w.some(c=>c.q.index===0x2300&&(c.q.sub===4||c.q.sub===0x0A)),'calibration commands must never be written');
  assert.match(confirms[0],/NICHT übertragen[\s\S]*0x2300:0x04[\s\S]*0x2300:0x0A/);
  for(let i=0;i<w.length;i++)if(w[i].q.index===0x2300&&w[i].q.sub!==3)assert.deepEqual([w[i-1].q.index,w[i-1].q.sub],[0x2300,3],'TAC unlock directly before each calibration write');
  for(const d of [1013,1014]){assert.equal(mem[d][key(0x2300,7)],15000);assert.equal(mem[d][key(0x2300,0x11)],300);assert.equal(mem[d][key(0x2300,4)],undefined);assert.equal(mem[d][key(0x2004,2)],0)}
  assert.match(out,/1013: 7 OK · 0 Fehler · EEPROM gespeichert/);
  // 2. Unlock not accepted: calibration block stops after the first error, no EEPROM save.
  tacBroken=true;mem={};out=await transfer([1020],[true,true,true]);
  const c2=calls.filter(c=>c.p==='/api/write'&&c.q.index===0x2300&&c.q.sub!==3);
  assert.equal(c2.length,1,'only one failing calibration write per device');
  assert.match(out,/1020: 1 OK · 1 Fehler · 5 übersprungen · NICHT gespeichert/);
  assert.match(out,/local control/);
  assert.ok(!calls.some(c=>c.p==='/api/write'&&c.q.index===0x2004));
  tacBroken=false;
  // 3. Without EEPROM confirmation nothing is stored persistently.
  mem={};out=await transfer([1021],[true,true,false]);
  assert.ok(!calls.some(c=>c.p==='/api/write'&&c.q.index===0x2004));assert.match(out,/nicht dauerhaft gespeichert/);
  // 4. Cancel on the calibration confirmation writes nothing.
  mem={};await transfer([1022],[false]);assert.ok(!calls.some(c=>c.p==='/api/write'));
  // 5. New profiles do not contain calibration commands at all.
  run('P=__P');assert.deepEqual(run("Q(0,1).map(p=>p.SubIndex)"),[1,2,7,8,0x0B,0x11]);
  console.log('PASS: calibration commands excluded, TAC unlock per calibration write, stop after first error, read-back, EEPROM save only on clean devices, cancel writes nothing');
})().catch(e=>{console.error(e);process.exitCode=1;});
