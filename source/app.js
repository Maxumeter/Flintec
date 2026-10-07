
let DA=0,MC=false,C=false,P=[],S=null,T=null,F=null,Z=null,B=[];
let HB=setInterval(()=>fetch('/api/heartbeat',{method:'POST',cache:'no-store'}).catch(()=>{}),2000);fetch('/api/heartbeat',{method:'POST',cache:'no-store'}).catch(()=>{});window.addEventListener('pagehide',()=>{try{navigator.sendBeacon('/api/window-close','')}catch{}});const $=x=>document.getElementById(x);function L(s) {
  const node = $('log');
  const lines = (node.textContent + `[${new Date().toISOString()}] ${s}\n`).split('\n');
  node.textContent = lines.slice(-3001).join('\n');
  node.scrollTop = node.scrollHeight;
}
function exportLog() {
  const data = `Flintec Control Center 1.9.5\n${$('log').textContent}`;
  const url = URL.createObjectURL(new Blob([data], {type:'text/plain;charset=utf-8'}));
  const a = document.createElement('a');
  a.href = url; a.download = `Flintec_Log_${new Date().toISOString().replace(/[:.]/g,'-')}.txt`;
  a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function M(s){let t=$('toast');t.textContent=s;t.style.display='block';setTimeout(()=>t.style.display='none',3500)}async function A(p, o = {}) {
  const id = ++A.sequence, started = performance.now();
  let q = {}; try { q = JSON.parse(o.body || '{}'); } catch {}
  const auth = p === '/api/connect';
  const hex = (n, width) => '0x' + Number(n).toString(16).toUpperCase().padStart(width,'0');
  const method = o.method || 'GET';
  const sdo = p === '/api/read' || p === '/api/write';
  const address = p === '/api/select-device' ? q.address : DA;
  let context = `#${id} ${method} ${p} | DAD=${address || '-'} | Master=${Z?.master || $('master').value}`;
  if (sdo) context += ` | ${hex(q.index,4)}:${hex(q.sub,2)} | Typ=${q.type} | Bytes=${q.length}`;
  if (p === '/api/write') context += ` | Wert=${q.value} | persistent=${!!q.confirmPersistent}`;
  if (sdo && $('cs')?.textContent) context += ` | Kalibrierstatus=${$('cs').textContent}`;
  const quiet = p === '/api/dashboard' || p === '/api/info' || p === '/api/parameters';
  const safe = message => {
    let text = String(message);
    if (auth) return 'Anmeldung fehlgeschlagen; Zugangsdaten werden nicht protokolliert.';
    return text.replace(/(Bearer\s+)\S+/gi,'$1[entfernt]').replace(/((?:password|access_token|refresh_token|authorization)["\s:=]+)[^\s,}]+/gi,'$1[entfernt]');
  };
  let status = 'Netzwerk';
  if (!quiet) L(context + ' | START');
  try {
    const r = await fetch(p, {headers:{'Content-Type':'application/json'}, ...o});
    status = r.status;
    let j; try { j = await r.json(); } catch { throw Error('Ungültige JSON-Antwort'); }
    if (!r.ok || j.ok === false) throw Error(safe(j.error || `HTTP ${r.status}`));
    if (!quiet) L(context + ` | OK HTTP ${status} | ${Math.round(performance.now()-started)} ms` + (p === '/api/read' ? ` | Wert=${j.value} | Roh=${j.raw ?? '-'}` : ''));
    if (p === '/api/dashboard' && A.dashboardError) { L(context + ' | Lesen wieder erfolgreich'); A.dashboardError = ''; }
    return j;
  } catch(e) {
    const error = safe(e.message);
    const detail = `${context} | FEHLER HTTP ${status} | ${Math.round(performance.now()-started)} ms | ${error}`;
    if (p !== '/api/dashboard' || A.dashboardError !== error) L(detail);
    if (p === '/api/dashboard') A.dashboardError = error;
    if (p === '/api/write') L(`#${id} Schreibausgang ungewiss: Fehler bestätigt nicht, ob das Gerät den Befehl ausgeführt hat. Nicht blind erneut ausführen.`);
    throw Error(`HTTP ${status}: ${error}`);
  }
}
A.sequence = 0;
function tab(id){document.querySelectorAll('.tab,.view').forEach(x=>x.classList.remove('active'));document.querySelector(`.tab[data-v="${id}"]`)?.classList.add('active');$(id)?.classList.add('active')}
async function init(){let i=await A('/api/info');$('ver').textContent=i.version;$('ip').value=i.config.ip||'192.168.179.10';$('master').value=i.config.master||'A1_PLC1_KEB1';$('address').value=i.config.address||'';P=(await A('/api/parameters')).parameters;R();document.querySelectorAll('.tab').forEach(b=>b.onclick=()=>tab(b.dataset.v));$('search').oninput=R;$('systemObjects').onchange=R;$('live').onchange=V;$('pf').onchange=e=>f31(e.target.files[0]);if(i.masterConnected)Z={ip:$('ip').value.trim(),master:$('master').value.trim(),address:i.selectedAddress||0};f5(i.masterConnected||false,i.deviceConnected||false,i.selectedAddress||0)}
function f5(master,device,address=0){MC=!!master;C=!!device;DA=C?Number(address):0;$('dot').className='dot'+(C?' on':MC?' master':'');let a=address||+$('address').value||0;$('statusText').textContent=C?`DAD143 ${a} verbunden`:MC?'EtherCAT Master verbunden':'Offline';$('connect').textContent=MC?'Master trennen':(($('address').value||'').trim()?'Verbinden':'Master verbinden');$('switchDevice').style.display=MC?'inline-block':'none';$('switchDevice').textContent=C?'Gerät wechseln':'Gerät auswählen';$('selectedInfo').textContent=C?`DAD143 · Adresse ${a}`:'Kein DAD143 ausgewählt';if(B.length)f21();V()}
$('address')?.addEventListener?.('input',()=>{if(!MC)$('connect').textContent=$('address').value?'Verbinden':'Master verbinden'});
async function f6(){if(!MC)return M('Zuerst EtherCAT Master verbinden.');tab('bus');if(!B.length)await f25()}
async function f7(){if(MC){await A('/api/disconnect',{method:'POST',body:'{}'}).catch(()=>0);B=[];f21();f22();f5(false,false);return}try{$('statusText').textContent='Verbinde EtherCAT Master…';let av=String($('address').value||'').trim(),q={ip:$('ip').value.trim(),master:$('master').value.trim(),address:av?+av:0,username:$('username').value,password:$('password').value};Z={ip:q.ip,master:q.master,address:q.address};let r=await Promise.race([A('/api/connect',{method:'POST',body:JSON.stringify(q)}),new Promise((_,rej)=>setTimeout(()=>rej(Error('Verbindungstimeout')),12000))]);$('password').value='';f5(true,!r.masterOnly,q.address);L(`EtherCAT Master ${q.master} verbunden${q.address?' · DAD143 '+q.address:''}`);if(r.masterOnly){tab('bus');await f25()}else await f26(q.address,true)}catch(e){f5(false,false);M(e.message);L('VERBINDUNG: '+e.message)}}
function V(){clearInterval(T);if(C&&$('live').checked)T=setInterval(f9,1000)}async function f9x(){if(!C)return;try{let d=await A('/api/dashboard');for(let k of ['gross','net','tare','deviceId','firmware','serial'])$(k).textContent=d[k]??'—'}catch(e){}}
function CR(p){return p.ReadPreOp||p.ReadSafeOp||p.ReadOp}function CW(p){return p.WritePreOp||p.WriteSafeOp||p.WriteOp}function esc(s){return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}function R(){let q=$('search').value.toLowerCase(),s=$('systemObjects').checked,a=P.filter(p=>(s||(p.Index>=0x2000&&p.Index<0x6000))&&(!q||`${p.IndexHex} ${p.ObjectName} ${p.EntryName}`.toLowerCase().includes(q)));$('tbody').innerHTML=a.map(p=>`<tr data-k="${p.Index}:${p.SubIndex}" onclick="SR(this)"><td>${p.IndexHex}</td><td>${p.SubIndexHex}</td><td>${esc(p.ObjectName)}</td><td>${esc(p.EntryName)}</td><td>${p.DataTypeName}</td><td>${p.AccessText}</td><td class="value" id="v-${p.Index}-${p.SubIndex}">${p.Value??''}</td><td class="set" id="s-${p.Index}-${p.SubIndex}" contenteditable="${CW(p)}"></td><td><button onclick="event.stopPropagation();S='${p.Index}:${p.SubIndex}';f17()" ${CR(p)?'':'disabled'}>Lesen</button><button class="primary" onclick="event.stopPropagation();S='${p.Index}:${p.SubIndex}';f19()" ${CW(p)?'':'disabled'}>Schreiben</button></td></tr>`).join('')}function SR(r){document.querySelectorAll('#tbody tr').forEach(x=>x.classList.remove('selected'));r.classList.add('selected');S=r.dataset.k}function SP(){if(!S)return M('Auswählen.');let[i,s]=S.split(':').map(Number);return P.find(p=>p.Index===i&&p.SubIndex===s)}async function RP(p,q=0){if(!C)throw Error('Zuerst DAD143 auswählen.');let r=await A('/api/read',{method:'POST',body:JSON.stringify({index:p.Index,sub:p.SubIndex,length:p.ValueBytes||Math.ceil(p.BitLength/8)||4,type:p.DataTypeName})}),e=$(`v-${p.Index}-${p.SubIndex}`);if(e)e.textContent=r.value;p.Value=r.value;if(!q)L(`${p.IndexHex}:${p.SubIndexHex} READ ${r.value}`);return r.value}async function f17(){let p=SP();if(p)try{await RP(p)}catch(e){M(e.message)}}async function WP(p,v,q=0){let ep=p.Index==0x2004;if(ep&&p.SubIndex==2||p.Index==0x2100&&p.SubIndex==0x12){let c=P.find(x=>x.Index==0x2300&&x.SubIndex==3),t=await RP(c,1);await WP(c,t,1)}await A('/api/write',{method:'POST',body:JSON.stringify({index:p.Index,sub:p.SubIndex,length:ep?4:(p.ValueBytes||Math.ceil(p.BitLength/8)||4),type:ep?'INTEGER32':p.DataTypeName,value:String(v),confirmPersistent:ep})});if(!q)L(`${p.IndexHex}:${p.SubIndexHex} WRITE ${v}`)}async function f19(){let p=SP();if(!p||!CW(p))return;let v=$(`s-${p.Index}-${p.SubIndex}`).textContent.trim();if(!v)return M('Wert setzen.');if(p.Index===0x2004&&!confirm('EEPROM schreiben?'))return;if(p.Index===0x2300&&p.SubIndex!==3&&!confirm('TAC freigeben?'))return;try{await WP(p,v);M('Geschrieben.');if(CR(p))await RP(p,1)}catch(e){M(e.message);L('WRITE: '+e.message)}}
function f20(x){let a=[];if(x.vendorId)a.push(`VID 0x${Number(x.vendorId).toString(16).toUpperCase()}`);if(x.productCode)a.push(`PID 0x${Number(x.productCode).toString(16).toUpperCase()}`);if(x.serial)a.push(`SN ${x.serial}`);return a.join(' · ')||'—'}function f21(){let a=$('onlyDad')?.checked?B.filter(x=>x.isDAD143):B;if(!a.length){$('busBody').innerHTML='<tr><td colspan="8" class="muted">Keine Geräte gefunden.</td></tr>';return}let cur=DA;$('busBody').innerHTML=a.map(x=>{let active=x.isDAD143&&x.address===cur;return `<tr class="scan ${x.isDAD143?'dad':''} ${active?'activeDevice':''}" ${x.isDAD143&&!active?`onclick="f26(${x.address},true)"`:''}><td class="${x.status?'good':'bad'}">${x.status?'✓':'!'}</td><td><span class="deviceTitle">${esc(x.name)}</span></td><td>${x.address}</td><td>${esc(x.state)}</td><td>${x.isDAD143?'<span class="badge dadbadge">DAD143</span>':'<span class="badge">EtherCAT</span>'}</td><td>${esc(f20(x))}</td><td>${esc(x.diagnostics||'—')}</td><td class="action">${x.isDAD143?(active?'<button disabled>Aktiv</button>':`<button class="primary" onclick="event.stopPropagation();f26(${x.address},true)">Wechseln</button>`):'—'}</td></tr>`}).join('')}
function f22(){let d=B.filter(x=>x.isDAD143);$('dv').innerHTML=d.length?d.map(x=>`<tr><td><input class="dev" type="checkbox" checked value="${x.address}"></td><td>${esc(x.name)}</td><td>${x.address}</td><td>${esc(x.state)}</td><td>${esc(x.diagnostics)}</td></tr>`).join(''):'<tr><td colspan="5" class="muted">Keine DAD143 gefunden.</td></tr>'}function SA(v){document.querySelectorAll('.dev').forEach(x=>x.checked=!!v)}function DS(){return[...document.querySelectorAll('.dev:checked')].map(x=>+x.value)}
async function f25x(){if(!MC)return M('Zuerst EtherCAT Master verbinden.');if(Z&&($('ip').value.trim()!=Z.ip||$('master').value.trim()!=Z.master)){let m='Verbindungsdaten geändert – Master trennen und neu verbinden.';$('busStatus').textContent=m;L('BUS: '+m);return M(m)}$('busStatus').textContent='Bus wird eingelesen…';try{let s=+$('ss').value||1001,e=+$('se').value||1128,r=await A(`/api/bus?start=${s}&end=${e}`);B=r.devices||[];$('busStatus').textContent=`${B.length} EtherCAT-Geräte · ${B.filter(x=>x.isDAD143).length} DAD143`;f21();f22();L(`BUS: ${B.length} Geräte, ${B.filter(x=>x.isDAD143).length} DAD143`)}catch(e){$('busStatus').textContent='Scanfehler';M(e.message);L('BUS: '+e.message)}}async function f26(a, stayBus = false) {
  if (f26.b) return;
  a = Number(a);
  if (!MC || !Number.isInteger(a) || a < 1 || a > 65535) return M('Masterverbindung und gültige Geräteadresse erforderlich.');
  f26.b = 1; clearInterval(T);
  try {
    L(`GERÄTEWECHSEL | vorher=${DA || '-'} | Ziel=${a}`);
    // Always confirm backend selection. The editable address is not session state.
    await A('/api/select-device', {method:'POST',body:JSON.stringify({address:a})});
    $('address').value = a;
    if (Z) Z.address = a;
    f5(true,true,a); clearInterval(T);
    for (const k of ['gross','net','tare','deviceId','firmware','serial']) $(k).textContent = '—';
    P.forEach(p => p.Value = ''); R();
    let n = 0, errors = 0;
    for (const p of P.filter(CR)) {
      if (!C) break;
      try { await RP(p,1); n++; } catch { errors++; }
      $('busStatus').textContent = `DAD143 ${a}: ${n} gelesen, ${errors} Fehler`;
    }
    L(`PARAMETERLESUNG | DAD=${a} | ${n} OK | ${errors} Fehler`);
    R(); await f9(); if (!stayBus) tab('overview');
  } catch(e) {
    f5(MC,false,0);
    P.forEach(p => p.Value = ''); R();
    for (const k of ['gross','net','tare','deviceId','firmware','serial']) $(k).textContent = '—';
    M(e.message); L(`DAD VERBINDUNG | Ziel=${a} | ${e.message}`);
  } finally { f26.b = 0; V(); }
}

const G=[0x2100,0x2200,0x2500,0x2600,0x2680,0x2700,0x2800,0x2d00];
// 0x2300 sub-indices that execute a calibration (zero 0x0A, gain 0x04) or hold the scale's own zero (0x02 Absolute zero). Never transfer them.
const XC=[0x02,0x04,0x0A],XN=/calib|gain|span|adjust|command|execute|absolute zero/i;
function KX(i,s,n){return i===0x2300&&(s===3||XC.includes(s)||XN.test(n||''))}
function K(p){if(p.Index===0x2300&&!KX(p.Index,p.SubIndex,p.EntryName)&&!/^Not used/.test(p.EntryName))return'cal';if(G.includes(p.Index))return'cfg';return''}
function Q(c,k){return P.filter(p=>CR(p)&&CW(p)&&((c&&K(p)==='cfg')||(k&&K(p)==='cal')))}function PG(n,t,s){$('pr').value=t?n/t*100:0;$('ps').textContent=s||`${n}/${t}`}async function f30(){if(!C)return M('DAD143 auswählen.');let a=Q($('pc').checked,$('pk').checked),o=[];if(!a.length)return M('Keine Auswahl.');for(let i=0;i<a.length;i++){let p=a[i];PG(i,a.length,`Lese ${i+1}/${a.length}`);try{o.push({index:p.Index,subIndex:p.SubIndex,name:p.EntryName,type:p.DataTypeName,value:await RP(p,1),category:K(p)})}catch{}}let f={formatVersion:1,device:'DAD143',created:new Date().toISOString(),parameters:o},u=URL.createObjectURL(new Blob([JSON.stringify(f,null,2)])),a1=document.createElement('a');a1.href=u;a1.download=`DAD143_${Date.now()}.flintecprofile`;a1.click();URL.revokeObjectURL(u);PG(a.length,a.length,`${o.length} gespeichert`)}async function f31(f){if(!f)return;try{let o=JSON.parse(await f.text());if(o.formatVersion!==1||!Array.isArray(o.parameters))throw Error('Ungültiges Profil');F=o;$('pi').textContent=`${f.name} · ${o.parameters.length} Parameter`;$('pd').textContent='Geladen.'}catch(e){F=null;M(e.message)}}
// Profiles saved by older versions still contain the calibration commands; filter them here as well.
function E(){return F?F.parameters.filter(e=>!KX(e.index,e.subIndex,e.name)&&((e.category==='cfg'&&$('ic').checked)||(e.category==='cal'&&$('ik').checked))):[]}
function EX(){return F?F.parameters.filter(e=>e.category==='cal'&&KX(e.index,e.subIndex,e.name)):[]}
// The zero calibration (0x2300:0x0A) hung 1014 and 1015 sporadically (no SDO answer, then SAFE-OP), presumably while the weight was not at rest.
// Send it only after the gross weight has been steady for about 2 s (6 readings, spread at most 2 steps of the last displayed digit).
async function ST(){let v=[];for(let i=0;i<6;i++){if(i)await new Promise(r=>setTimeout(r,400));let g=String((await A('/api/dashboard')).gross??'').trim(),n=PF(g);if(!Number.isFinite(n))return{ok:false,why:`Brutto nicht lesbar (${g||'—'})`};v.push([n,g])}
let dc=Math.max(...v.map(([,g])=>((g.match(/[.,](\d+)/)||[,''])[1]).length)),mn=Math.min(...v.map(x=>x[0])),mx=Math.max(...v.map(x=>x[0])),w=v.map(x=>x[1]).join(' ');
return mx-mn<=2*10**-dc+1e-9?{ok:true,why:`Brutto ruhig: ${w}`}:{ok:false,why:`Brutto schwankt: ${w}`}}
function PN(e){return`0x${e.index.toString(16).toUpperCase()}:0x${e.subIndex.toString(16).toUpperCase().padStart(2,'0')} ${e.name||''}`}
async function f33(){if(!C||!F)return M('Profil und DAD143-Verbindung erforderlich.');let d=[];for(let e of E()){let p=P.find(x=>x.Index===e.index&&x.SubIndex===e.subIndex);if(p&&CR(p))try{let v=await RP(p,1);if(String(v)!=String(e.value))d.push(`${p.IndexHex}:${p.SubIndexHex} ${v}→${e.value}`)}catch{}}$('pd').textContent=d.join('\n')||'Gleich.'}
async function f34(){if(!MC||!F)return M('Profil/Masterverbindung fehlt.');let a=E(),d=DS(),x0=$('ik').checked?EX():[];if(!a.length||!d.length)return M('Keine Auswahl.');
let ca=a.filter(e=>e.category==='cal');
if(ca.length&&!confirm(`Kalibrier-Einstellwerte auf alle ausgewählten Geräte schreiben?\n\n${ca.map(PN).join('\n')}`+(x0.length?`\n\nNICHT übertragen (Kalibrierbefehle und waagenspezifischer Nullpunkt):\n${x0.map(PN).join('\n')}`:'')+'\n\nDas Kalibriergewicht (Verstärkung) muss bei Bedarf auf jeder Waage einzeln kalibriert werden.'))return;
if(!confirm(`${d.length} DAD143 · ${a.length} Parameter übertragen?`))return;
// Zero calibration (0x2300:0x0A with 0, as in the calibration dialog) is the only calibration command the transfer may run, and only on request: each target scale must be empty.
let zc=!!ca.length&&confirm(`Nullpunkt auf jeder Zielwaage kalibrieren?\n\nAlle ${d.length} Zielwaagen müssen vollständig entlastet sein und ruhig stehen.\nDie Verstärkung (Kalibriergewicht) wird nicht verändert.\nAntwortet eine Waage nicht auf die Nullpunkt-Kalibrierung, wird die Übertragung sofort beendet.\n\nAbbrechen = Nullpunkt nicht kalibrieren.`);
let sv=confirm('Nach erfolgreicher Übertragung und Kontrolle dauerhaft im EEPROM speichern?\n\nGespeichert wird nur auf Geräten ohne Schreib- oder Prüffehler.\nAbbrechen = nur flüchtig schreiben.');
// A device that times out gets no further requests; a zero calibration that times out ends the whole run, a rejected one (SDO abort) only fails that device (seen on hardware: 1015 dropped to SAFE-OP after 0x2300:0x0A timed out).
let o=[],n=0,t=d.length*a.length,orig=DA,cs=$('cs').textContent,TO=/Timeout/i,stop=false;clearInterval(T);
for(let[xi,x]of d.entries()){let okc=0,er=[],sk=0,w=[],dead=false;try{await f35(x);$('cs').textContent=`Profilübertragung ${x}`;let cb=false;
for(let e of a){let p=P.find(y=>y.Index===e.index&&y.SubIndex===e.subIndex);PG(n++,t,`Gerät ${x} · ${okc+er.length+sk+1}/${a.length}`);if(!p||!CW(p))continue;
let c=e.category==='cal';if(c&&cb){sk++;continue}
// Calibration parameters need a fresh TAC unlock directly before every write, exactly like the calibration dialog (CWZ).
try{c?await CWZ(p,e.value):await WP(p,e.value,1);okc++;w.push([p,e])}catch(z){er.push(`${PN(e)}: ${z.message}`);if(c)cb=true;if(TO.test(z.message)){dead=true;break}}}
// Compare before the zero calibration: it rewrites 0x2300:0x02 Absolute zero with the target scale's own value.
let mm=[];if(!dead)for(let[p,e]of w)try{let v=await RP(p,1);if(String(v)!=String(e.value))mm.push(`${PN(e)}: ${v}≠${e.value}`)}catch(z){mm.push(`${PN(e)}: ${z.message}`);if(TO.test(z.message)){dead=true;break}}
let zs='';if(zc){if(cb||dead)zs=' · Nullpunkt übersprungen';else try{PG(n,t,`Gerät ${x} · Stillstand prüfen`);let st=await ST();L(`NULLPUNKT ${x} | ${st.why}`);if(!st.ok){zs=` · Nullpunkt NICHT kalibriert – Waage nicht ruhig (${st.why})`;throw 0}PG(n,t,`Gerät ${x} · Nullpunkt`);let az=CP(0x2300,2),a0=az?await RP(az,1).catch(()=>'?'):'?';await CWZ(CP(0x2300,0x0A),0);zs=` · Nullpunkt kalibriert${az?` (Absolute zero ${a0}→${await RP(az,1).catch(()=>'?')})`:''}`}catch(z){if(z!==0){er.push(`0x2300:0x0A Nullpunkt kalibrieren: ${z.message}`);zs=' · Nullpunkt-Fehler';if(TO.test(z.message))dead=stop=true}}}
let s='';if(dead)s=' · Gerät antwortet nicht mehr, abgebrochen';else if(sv){if(er.length||mm.length||sk)s=' · NICHT gespeichert';else try{await WP(CP(0x2004,2),0,1);s=' · EEPROM gespeichert'}catch(z){s=` · EEPROM-Fehler: ${z.message}`}}else s=' · nicht dauerhaft gespeichert';
o.push(`${x}: ${okc} OK · ${er.length} Fehler`+(sk?` · ${sk} übersprungen`:'')+(mm.length?` · ${mm.length} Abweichungen`:'')+zs+s);for(let z of[...er,...mm])o.push('   '+z);if(stop){let r=d.slice(xi+1);o.push(`ÜBERTRAGUNG BEENDET: Nullpunkt-Kalibrierung auf ${x} ohne Antwort. EtherCAT-Zustand von ${x} prüfen (z. B. SAFE-OP).`+(r.length?` Nicht bearbeitet: ${r.join(', ')}`:''));break}}catch(e){o.push(`${x}: ${e.message}`);n+=a.length}}
$('cs').textContent=cs;if(orig&&B.some(x=>x.address===orig&&x.isDAD143))try{await f35(orig)}catch{}PG(t,t,stop?'Abgebrochen':'Fertig');$('pd').textContent=o.join('\n');M(stop?'Übertragung abgebrochen – siehe Ergebnis.':'Übertragen.');V()}
// While the master is slow (bus scan took 176 s on 2026-10-07), do not queue further scans or dashboard polls behind a pending one.
let BZ=0,DZ=0;async function f25(){if(BZ)return M('Bus wird bereits eingelesen – bitte warten.');BZ=1;try{await f25x()}finally{BZ=0}}async function f9(){if(DZ)return;DZ=1;try{await f9x()}finally{DZ=0}}
async function f35(a){await A('/api/select-device',{method:'POST',body:JSON.stringify({address:a})});$('address').value=a;if(Z)Z.address=a;f5(true,true,a)}
function H(s){return parseInt(s.replace(/^0x/i,''),16)}async function f37(){try{let r=await A('/api/read',{method:'POST',body:JSON.stringify({index:H($('exIndex').value),sub:H($('exSub').value),length:+$('exLen').value,type:$('exType').value})});$('exValue').value=$('exType').value==='HEX'?r.raw:r.value}catch(e){M(e.message)}}async function f38(){let i=H($('exIndex').value),s=H($('exSub').value),e=i===0x2004;if(e&&!confirm('EEPROM schreiben?'))return;try{await WP({Index:i,SubIndex:s,ValueBytes:+$('exLen').value,DataTypeName:$('exType').value},$('exValue').value,1);M('Geschrieben.')}catch(e){M(e.message)}}function PF(x){return parseFloat(String(x).replace(',','.'))}function CP(i,s){return P.find(p=>p.Index===i&&p.SubIndex===s)}async function CE(){let p=CP(0x2300,3),t=await RP(p,1);await WP(p,t,1);return+t}async function CWZ(p,v){await CE();await WP(p,v,1)}async function CAL(){if(!C)return M('DAD143 auswählen.');let w=PF($('cw').value),x=PF($('cmax').value),n=PF($('cmin').value),u=$('cu').value;if(![w,x,n].every(Number.isFinite)||x<=n||w<=n||w>x)return M('Gewichtswerte prüfen.');if(w<x*.01)return M('Kalibriergewicht muss mindestens 1% vom Maximum sein.');if(w<x*.2&&!confirm(`Hinweis: Flintec empfiehlt ein Kalibriergewicht von mindestens 20% des Maximums.\n\n${w} ${u} entsprechen ${(w/x*100).toFixed(1)}% von ${x} ${u}.\n\nTrotzdem fortfahren?`))return;clearInterval(T);try{$('cs').textContent='Lese Dezimalpunkt…';let dp=+await RP(CP(0x2300,0x0B),1),f=10**dp,W=Math.round(w*f),X=Math.round(x*f),N=Math.round(n*f);$('cs').textContent=`DP ${dp} · Faktor ×${f}`;if(!confirm(`Kalibrierung starten?\n\nKalibriergewicht: ${w} ${u} → Parameter ${W}\nMaximum: ${x} ${u} → Parameter ${X}\nMinimum: ${n} ${u} → Parameter ${N}\nDezimalstellen: ${dp} · Faktor ×${f}\n\nFalls das Gerät versiegelt/geeicht ist, muss der Seal Switch für die Kalibrierung geöffnet sein.`)){$('cs').textContent='Abgebrochen';return}$('cs').textContent='Setze Messbereich…';await CWZ(CP(0x2300,0x07),X);await CWZ(CP(0x2300,0x08),N);if(!confirm(`Schritt 1 von 2 – Nullpunkt\n\nWaage vollständig entlasten. Es darf sich kein Gewicht auf der Waage befinden.\nWarten, bis der Messwert stabil ist, dann OK klicken.`)){$('cs').textContent='Abgebrochen – nicht gespeichert';return}$('cs').textContent='Prüfe Stillstand…';{let st=await ST();L('NULLPUNKT | '+st.why);if(!st.ok)throw Error('Waage nicht ruhig – Nullpunkt nicht kalibriert. '+st.why)}$('cs').textContent='Kalibriere Nullpunkt…';await CWZ(CP(0x2300,0x0A),0);if(!confirm(`Schritt 2 von 2 – Kalibriergewicht\n\nJetzt ${w} ${u} auflegen. Warten, bis das Gewicht vollständig aufliegt und der Messwert stabil ist, dann OK klicken.\n\nInterner Kalibrierwert: ${W}`)){$('cs').textContent='Abgebrochen – nicht gespeichert';return}$('cs').textContent=`Kalibriere Verstärkung mit ${w} ${u}…`;await CWZ(CP(0x2300,0x04),W);let b=+await RP(CP(0x2300,3),1);if(confirm('Kalibrierung ist durchgeführt.\n\nSoll sie jetzt dauerhaft im EEPROM gespeichert werden?')){$('cs').textContent='Speichere EEPROM…';await WP(CP(0x2004,2),0,1);let aa=+await RP(CP(0x2300,3),1);$('cs').textContent=`Fertig · ${w} ${u} · TAC ${b} → ${aa}`;alert(`Kalibrierung abgeschlossen und gespeichert.\nKalibriergewicht: ${w} ${u}\nParameterwert: ${W}\nTAC: ${b} → ${aa}`)}else{$('cs').textContent='Kalibriert, aber nicht dauerhaft gespeichert';alert('Kalibrierung durchgeführt, aber nicht im EEPROM gespeichert.')}}catch(e){$('cs').textContent='Fehler: '+e.message;M(e.message);alert('Kalibrierung abgebrochen:\n'+e.message)}finally{V()}}async function shutdown(){if(confirm('Beenden?'))try{await A('/api/shutdown',{method:'POST',body:'{}'});document.body.innerHTML='<div style="padding:50px">Beendet.</div>'}catch{}}
window.addEventListener('beforeunload',()=>{try{navigator.sendBeacon('/api/disconnect','{}')}catch(e){}});L('START | Version 1.9.5 (Build 9eb72d40) | SDO-Fehler zeigen den letzten Anfrageversuch.');init().catch(e=>M(e.message));
