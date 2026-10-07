const vm=require('node:vm'),fs=require('node:fs'),assert=require('node:assert/strict');
const elements=new Map();
function el(id){if(!elements.has(id))elements.set(id,{value:'',textContent:'',innerHTML:'',style:{},classList:{add(){},remove(){}},addEventListener(){},checked:false,scrollHeight:0});return elements.get(id);}
const calls=[];let failSelect=false,failWrite=false,failDashboard=false;
const context=vm.createContext({console,Date,Math,Number,String,Array,JSON,Error,Promise,performance,Blob,URL,
  setInterval:()=>1,clearInterval(){},setTimeout:()=>1,
  document:{getElementById:el,querySelectorAll:()=>[],querySelector:()=>null,createElement:()=>({click(){}})},
  window:{addEventListener(){}},navigator:{sendBeacon(){}},confirm:()=>false,alert(){},
  fetch:async(p,o={})=>{const q=JSON.parse(o.body||'{}');calls.push({p,q});
    const bad=(p==='/api/select-device'&&failSelect)||(p==='/api/write'&&failWrite)||(p==='/api/dashboard'&&failDashboard)||p==='/api/connect';
    return {ok:!bad,status:bad?500:200,json:async()=>bad?{ok:false,error:p==='/api/connect'?'SECRET_USER SECRET_PASS access_token=SECRET_TOKEN':'SDO Write: HTTP 500: Ecat: Timeout'}:{ok:true,value:7,raw:'07000000'}};
  }});
const run=s=>vm.runInContext(s,context);
let code=fs.readFileSync('source/app.js','utf8').replace('init().catch(e=>M(e.message));','');
run(code);run('R=()=>{};V=()=>{};f9=async()=>{};MC=true;Z={master:"TEST_MASTER"};P=[];');
(async()=>{
  el('master').value='TEST_MASTER';el('address').value='1014';
  calls.length=0;
  await run('f26(1014,true)');
  assert.equal(calls.filter(c=>c.p==='/api/select-device').length,1,'same address must reach backend');
  assert.equal(run('DA'),1014);assert.equal(run('C'),true);
  run('B=[{address:1014,isDAD143:true,name:"Device A",state:"OP"},{address:1015,isDAD143:true,name:"Device B",state:"OP"}]');
  el('address').value='1015';run('f21()');
  assert.match(el('busBody').innerHTML,/Device A[\s\S]*?<button disabled>Aktiv/);
  assert.match(el('busBody').innerHTML,/f26\(1015,true\)/);
  failSelect=true;el('gross').textContent='stale';await run('f26(1015,true)');
  assert.equal(run('DA'),0);assert.equal(run('C'),false);assert.equal(el('gross').textContent,'—');
  assert.match(el('log').textContent,/Ziel=1015/);failSelect=false;
  run('MC=true;C=true;DA=1014;P=[{Index:8960,SubIndex:3,ValueBytes:4,DataTypeName:"INTEGER32"},{Index:8960,SubIndex:4,ValueBytes:4,DataTypeName:"INTEGER32"}]');
  el('cs').textContent='Kalibriere Verstärkung';calls.length=0;
  await run('CWZ(CP(0x2300,4),1000)');
  assert.deepEqual(calls.map(c=>[c.p,c.q.index,c.q.sub,c.q.value]),[['/api/read',8960,3,undefined],['/api/write',8960,3,'7'],['/api/write',8960,4,'1000']]);
  assert.match(el('log').textContent,/0x2300:0x04.*Typ=INTEGER32.*Bytes=4.*Wert=1000.*Kalibriere Verstärkung/);
  failWrite=true;await assert.rejects(run('WP(CP(0x2300,4),1000,1)'),/Ecat: Timeout/);
  assert.match(el('log').textContent,/FEHLER HTTP 500.*ms.*Ecat: Timeout/);
  assert.match(el('log').textContent,/Schreibausgang ungewiss/);
  await assert.rejects(run('A("/api/connect",{method:"POST",body:JSON.stringify({username:"SECRET_USER",password:"SECRET_PASS"})})'));
  assert.doesNotMatch(el('log').textContent,/SECRET_USER|SECRET_PASS|SECRET_TOKEN/);
  failDashboard=true;
  for(let i=0;i<3;i++)await assert.rejects(run('A("/api/dashboard")'));
  assert.equal((el('log').textContent.match(/GET \/api\/dashboard.*FEHLER/g)||[]).length,1);
  failDashboard=false;await run('A("/api/dashboard")');assert.match(el('log').textContent,/Lesen wieder erfolgreich/);
  run('for(let i=0;i<3100;i++)L("test "+i)');assert.ok(el('log').textContent.split('\n').length<=3001);
  console.log('PASS: same-address selection, confirmed-address highlighting, failed selection clearing, unchanged calibration command sequence, SDO error context, credential redaction, dashboard dedup/recovery, bounded log');
})().catch(e=>{console.error(e);process.exitCode=1;});
