const G=[0x2100,0x2200,0x2500,0x2600,0x2680,0x2700,0x2800,0x2d00];
// 0x2300 sub-indices that execute a calibration (zero 0x0A, gain 0x04) or hold the scale's own zero (0x02 Absolute zero). Never transfer them.
const XC=[0x02,0x04,0x0A],XN=/calib|gain|span|adjust|command|execute|absolute zero/i;
function KX(i,s,n){return i===0x2300&&(s===3||XC.includes(s)||XN.test(n||''))}
function K(p){if(p.Index===0x2300&&!KX(p.Index,p.SubIndex,p.EntryName)&&!/^Not used/.test(p.EntryName))return'cal';if(G.includes(p.Index))return'cfg';return''}
function Q(c,k){return P.filter(p=>CR(p)&&CW(p)&&((c&&K(p)==='cfg')||(k&&K(p)==='cal')))}function PG(n,t,s){$('pr').value=t?n/t*100:0;$('ps').textContent=s||`${n}/${t}`}async function f30(){if(!C)return M('DAD143 auswählen.');let a=Q($('pc').checked,$('pk').checked),o=[];if(!a.length)return M('Keine Auswahl.');for(let i=0;i<a.length;i++){let p=a[i];PG(i,a.length,`Lese ${i+1}/${a.length}`);try{o.push({index:p.Index,subIndex:p.SubIndex,name:p.EntryName,type:p.DataTypeName,value:await RP(p,1),category:K(p)})}catch{}}let f={formatVersion:1,device:'DAD143',created:new Date().toISOString(),parameters:o},u=URL.createObjectURL(new Blob([JSON.stringify(f,null,2)])),a1=document.createElement('a');a1.href=u;a1.download=`DAD143_${Date.now()}.flintecprofile`;a1.click();URL.revokeObjectURL(u);PG(a.length,a.length,`${o.length} gespeichert`)}async function f31(f){if(!f)return;try{let o=JSON.parse(await f.text());if(o.formatVersion!==1||!Array.isArray(o.parameters))throw Error('Ungültiges Profil');F=o;$('pi').textContent=`${f.name} · ${o.parameters.length} Parameter`;$('pd').textContent='Geladen.'}catch(e){F=null;M(e.message)}}
// Profiles saved by older versions still contain the calibration commands; filter them here as well.
function E(){return F?F.parameters.filter(e=>!KX(e.index,e.subIndex,e.name)&&((e.category==='cfg'&&$('ic').checked)||(e.category==='cal'&&$('ik').checked))):[]}
function EX(){return F?F.parameters.filter(e=>e.category==='cal'&&KX(e.index,e.subIndex,e.name)):[]}
// The zero calibration (0x2300:0x0A) hung 1014 and 1015 sporadically (no SDO answer, then SAFE-OP) and 1023 rejected it ("General error").
// Per manual 9.3 the DAD143 only accepts "Calibrate Zero" when the signal varied by at most NR increments (0x2100:0x0A) during NT ms (0x2100:0x0B).
// Send it only after the gross weight has been steady for NT + 0.5 s (at least 2 s) within min(NR, 2) display steps (step = 0x2300:0x0C · 10^-0x2300:0x0B).
const SL=ms=>new Promise(r=>setTimeout(r,ms));
async function RD(i,s,d){let p=CP(i,s);if(!p)return d;try{let v=+await RP(p,1);return Number.isFinite(v)?v:d}catch(e){if(/Timeout/i.test(e.message))throw e;return d}}
// A/D error or input out of range (Qualifier 0x2900:0D bits 7/1/0): no measuring signal, the device rejects the zero with "General error"
// (1023 on 2026-10-07 without load cell: Qualifier 128, A/D 0, mV/V 1). Checked first, so the command is not sent at all.
async function ST(){let q=await RD(0x2900,0x0D,0);if(q&0x83)return{ok:false,ad:true,why:`kein gültiges Messsignal – ${QB.filter(([b])=>(0x83>>b&1)&&(q>>b&1)).map(([,n])=>n).join(', ')} (Qualifier ${q}). Wägezelle angeschlossen?`};let nr=await RD(0x2100,0x0A,2),nt=await RD(0x2100,0x0B,1500),dp=await RD(0x2300,0x0B,0),ds=await RD(0x2300,0x0C,1),st=(ds||1)*10**-dp,tol=Math.min(Math.max(nr,1),2)*st+1e-9,k=Math.max(6,Math.ceil((Math.min(nt,10000)+500)/400)+1),v=[];
for(let i=0;i<k;i++){if(i)await SL(400);let g=String((await A('/api/dashboard')).gross??'').trim(),n=PF(g);if(!Number.isFinite(n))return{ok:false,why:`Brutto nicht lesbar (${g||'—'})`};v.push(n)}
let mn=Math.min(...v),mx=Math.max(...v),w=`${v.join(' ')} (Spanne ${+(mx-mn).toFixed(6)}, Grenze ${+tol.toFixed(6)}, NR ${nr} d, NT ${nt} ms)`;
return mx-mn<=tol?{ok:true,why:`Brutto ruhig: ${w}`}:{ok:false,why:`Brutto schwankt: ${w}`}}
// Save commands per parameter group (manual 9.12). Calibration (CS) increments the TAC and goes last.
const SN={1:'Analog',2:'Kalibrierung',3:'Setup',4:'Füllparameter',5:'Sollwerte'},AN=[0x01,0x02,0x03,0x15,0x19];
function SG(ps,zero){let g=new Set();for(let p of ps){let i=p.Index;if(i===0x2300||(i===0x2100&&p.SubIndex===0x12))g.add(2);if(i===0x2100&&AN.includes(p.SubIndex))g.add(1);else if(i===0x2100||i===0x2500||i===0x2d00)g.add(3);if(i===0x2200)g.add(4);if(i===0x2600||i===0x2680||i===0x2700||i===0x2800)g.add(5)}if(zero)g.add(2);return[3,4,5,1,2].filter(k=>g.has(k))}
// Device state for diagnosing a rejected zero calibration ("Ecat SDO: General error" on 1023, 1030, 1037 with a perfectly steady gross weight):
// Qualifier 0x2900:0D bits per manual 14.1.6, device status 0x2900:0A, internal mV/V 0x2900:12 and raw A/D sample 0x2900:07.
const QB=[[0,'Unterbereich'],[1,'Überbereich'],[2,'außerhalb Nullstellbereich'],[3,'Nullpunktmitte'],[4,'Stillstand'],[5,'Tara gesetzt'],[6,'Preset-Tara'],[7,'A/D-Fehler']];
async function DG(){let q=await RD(0x2900,0x0D,null),st=await RD(0x2900,0x0A,null),mv=await RD(0x2900,0x12,null),ad=await RD(0x2900,0x07,null);
return`Qualifier ${q??'?'}${q==null?'':` [${QB.filter(([b])=>q>>b&1).map(([,n])=>n).join(', ')||'–'}]`} · Status ${st??'?'} · mV/V-intern ${mv??'?'} · A/D ${ad??'?'}`}
function PN(e){return`0x${e.index.toString(16).toUpperCase()}:0x${e.subIndex.toString(16).toUpperCase().padStart(2,'0')} ${e.name||''}`}
async function f33(){if(!C||!F)return M('Profil und DAD143-Verbindung erforderlich.');let d=[];for(let e of E()){let p=P.find(x=>x.Index===e.index&&x.SubIndex===e.subIndex);if(p&&CR(p))try{let v=await RP(p,1);if(String(v)!=String(e.value))d.push(`${p.IndexHex}:${p.SubIndexHex} ${v}→${e.value}`)}catch{}}$('pd').textContent=d.join('\n')||'Gleich.'}
// Devices that failed in the last run; they can be retried with f36 ("Fehlgeschlagene erneut übertragen").
let FD=[];function RB(){let b=$('rt');if(b){b.disabled=!FD.length;b.textContent=`Fehlgeschlagene erneut übertragen${FD.length?` (${FD.length})`:''}`}}
async function f36(){if(!FD.length)return M('Keine fehlgeschlagenen Geräte.');await f34([...FD])}
async function f34(ds){if(!MC||!F)return M('Profil/Masterverbindung fehlt.');let a=E(),d=Array.isArray(ds)?ds:DS(),x0=$('ik').checked?EX():[];if(!a.length||!d.length)return M('Keine Auswahl.');
let ca=a.filter(e=>e.category==='cal');
if(ca.length&&!confirm(`Kalibrier-Einstellwerte auf alle ausgewählten Geräte schreiben?\n\n${ca.map(PN).join('\n')}`+(x0.length?`\n\nNICHT übertragen (Kalibrierbefehle und waagenspezifischer Nullpunkt):\n${x0.map(PN).join('\n')}`:'')+'\n\nDas Kalibriergewicht (Verstärkung) muss bei Bedarf auf jeder Waage einzeln kalibriert werden.'))return;
if(!confirm(`${d.length} DAD143 · ${a.length} Parameter übertragen?`))return;
// Zero calibration (0x2300:0x0A with 0, as in the calibration dialog) is the only calibration command the transfer may run, and only on request: each target scale must be empty.
let zc=!!ca.length&&confirm(`Nullpunkt auf jeder Zielwaage kalibrieren?\n\nAlle ${d.length} Zielwaagen müssen vollständig entlastet sein und ruhig stehen.\nDie Verstärkung (Kalibriergewicht) wird nicht verändert.\nWaagen mit Fehler werden übersprungen und können danach erneut übertragen werden.\n\nAbbrechen = Nullpunkt nicht kalibrieren.`);
let sv=confirm('Nach erfolgreicher Übertragung und Kontrolle dauerhaft im EEPROM speichern?\n\nGespeichert wird nur auf Geräten ohne Schreib- oder Prüffehler.\nAbbrechen = nur flüchtig schreiben.');
// A failing device is skipped and remembered for a manual retry. A device that times out gets no further requests (1015 dropped to SAFE-OP after 0x2300:0x0A timed out);
// only after three zero calibrations without answer in one run is the rest stopped, as that points to a general problem.
let o=[],n=0,t=d.length*a.length,orig=DA,cs=$('cs').textContent,TO=/Timeout/i,stop=false,zt=0;FD=[];clearInterval(T);
for(let[xi,x]of d.entries()){let okc=0,er=[],sk=0,w=[],dead=false,zok=!zc;try{await f35(x);$('cs').textContent=`Profilübertragung ${x}`;let cb=false;
for(let e of a){let p=P.find(y=>y.Index===e.index&&y.SubIndex===e.subIndex);PG(n++,t,`Gerät ${x} · ${okc+er.length+sk+1}/${a.length}`);if(!p||!CW(p))continue;
let c=e.category==='cal';if(c&&cb){sk++;continue}
// Calibration parameters need a fresh TAC unlock directly before every write, exactly like the calibration dialog (CWZ).
try{c?await CWZ(p,e.value):await WP(p,e.value,1);okc++;w.push([p,e])}catch(z){er.push(`${PN(e)}: ${z.message}`);if(c)cb=true;if(TO.test(z.message)){dead=true;break}}}
// Compare before the zero calibration: it rewrites 0x2300:0x02 Absolute zero with the target scale's own value.
let mm=[];if(!dead)for(let[p,e]of w)try{let v=await RP(p,1);if(String(v)!=String(e.value))mm.push(`${PN(e)}: ${v}≠${e.value}`)}catch(z){mm.push(`${PN(e)}: ${z.message}`);if(TO.test(z.message)){dead=true;break}}
let zs='';if(zc){if(cb||dead)zs=' · Nullpunkt übersprungen';else try{PG(n,t,`Gerät ${x} · Stillstand prüfen`);let st=await ST();L(`NULLPUNKT ${x} | ${st.why}`);if(!st.ok){zs=st.ad?` · Nullpunkt NICHT kalibriert – ${st.why}`:` · Nullpunkt NICHT kalibriert – Waage nicht ruhig (${st.why})`;throw 0}PG(n,t,`Gerät ${x} · Nullpunkt`);let az=CP(0x2300,2),a0=az?await RP(az,1).catch(()=>'?'):'?',dg=await DG().catch(()=>'');L(`NULLPUNKT ${x} | vorher ${dg}`);try{await CWZ(CP(0x2300,0x0A),0)}catch(z){if(!TO.test(z.message)){let d2=await DG().catch(()=>'');L(`NULLPUNKT ${x} | abgelehnt | ${d2}`);z.message+=` | ${d2}`}throw z}zok=true;zs=` · Nullpunkt kalibriert${az?` (Absolute zero ${a0}→${await RP(az,1).catch(()=>'?')})`:''}`}catch(z){if(z!==0){er.push(`0x2300:0x0A Nullpunkt kalibrieren: ${z.message}`);zs=' · Nullpunkt-Fehler';if(TO.test(z.message)){dead=true;stop=++zt>=3}}}}
let s='';if(dead)s=' · Gerät antwortet nicht mehr, übersprungen (Waage prüfen bzw. neu starten)';else if(sv){if(er.length||mm.length||sk)s=' · NICHT gespeichert';else{let g=SG(w.map(([p])=>p),zok&&zc),ok=[],bad=[];for(let k of g){let q=CP(0x2004,k);if(!q)continue;try{await WP(q,0,1);ok.push(SN[k])}catch(z){(k===1?o:er).push(`   ${SN[k]} (0x2004:0${k}): ${z.message}`);if(k!==1)bad.push(SN[k])}}s=bad.length?` · EEPROM-Fehler (${bad.join(', ')})`:` · EEPROM gespeichert (${ok.join(', ')})`}}else s=' · nicht dauerhaft gespeichert';
o.push(`${x}: ${okc} OK · ${er.length} Fehler`+(sk?` · ${sk} übersprungen`:'')+(mm.length?` · ${mm.length} Abweichungen`:'')+zs+s);for(let z of[...er,...mm])o.push('   '+z);if(dead||er.length||mm.length||sk||!zok)FD.push(x);if(stop){let r=d.slice(xi+1);FD.push(...r);o.push(`ÜBERTRAGUNG BEENDET: ${zt} Waagen ohne Antwort auf die Nullpunkt-Kalibrierung. EtherCAT-Zustand prüfen (z. B. SAFE-OP).`+(r.length?` Nicht bearbeitet: ${r.join(', ')}`:''));break}}catch(e){o.push(`${x}: ${e.message}`);FD.push(x);n+=a.length}}
if(FD.length)o.push(`\nFehlgeschlagen/übersprungen: ${FD.join(', ')}\nNach Prüfung (ggf. Waage neu starten) mit „Fehlgeschlagene erneut übertragen“ wiederholen.`);RB();
$('cs').textContent=cs;if(orig&&B.some(x=>x.address===orig&&x.isDAD143))try{await f35(orig)}catch{}PG(t,t,stop?'Abgebrochen':'Fertig');$('pd').textContent=o.join('\n');M(stop?'Übertragung abgebrochen – siehe Ergebnis.':'Übertragen.');V()}
// While the master is slow (bus scan took 176 s on 2026-10-07), do not queue further scans or dashboard polls behind a pending one.
let BZ=0,DZ=0;async function f25(){if(BZ)return M('Bus wird bereits eingelesen – bitte warten.');BZ=1;try{await f25x()}finally{BZ=0}}async function f9(){if(DZ)return;DZ=1;try{await f9x()}finally{DZ=0}}

// Filter optimisation (manual 9.4): try every filter level FL 1..8 (0x2100:0x04) in the chosen mode FM (0x2100:0x09) volatile, measure the
// gross weight of the empty scale and pick the weakest (fastest) level whose noise stays within the target. Two passes (up, down).
// On 1022 (2026-10-07) single passes showed 48-51 d while the other pass of the same level showed 3-7 d, and even FL 8 stayed at 5 d: slow
// wandering and disturbances, not noise a low-pass can remove. Therefore: noise = P5..P95 range after removing the linear drift, the better
// pass decides (a disturbed pass is marked "!"), drift is shown separately, and nothing is proposed when no level reaches the target.
// Nothing is saved without confirmation; cancel or error restores the old values.
const FT={0:[0,55,122,242,322,482,963,1923,3847],1:[0,47,93,140,187,233,280,327,373]},FC={0:['–',18,8,4,3,2,1,0.5,0.25],1:['–',19.7,9.8,6.5,4.9,3.9,3.2,2.8,2.5]};
function NS(v){let n=v.length,mx=(n-1)/2,my=v.reduce((a,b)=>a+b,0)/n,sxy=0,sxx=0;v.forEach((y,x)=>{sxy+=(x-mx)*(y-my);sxx+=(x-mx)**2});let b=sxx?sxy/sxx:0,r=v.map((y,x)=>y-my-b*(x-mx)).sort((a,c)=>a-c),q=f=>r[Math.min(n-1,Math.max(0,Math.round(f*(n-1))))];
return{ns:q(0.95)-q(0.05),dr:b*(n-1),sd:Math.sqrt(r.reduce((a,c)=>a+c*c,0)/n)}}
async function FO(){if(!C)return M('DAD143 auswählen.');let fl=CP(0x2100,4),fm=CP(0x2100,9),gp=CP(0x2900,1);if(!fl||!fm||!gp)return M('Filterparameter nicht gefunden.');
let tg=Math.max(1,PF($('fo_t').value)||1),sec=Math.max(3,PF($('fo_s').value)||8),m=+$('fo_m').value||0,o=[],x=DA;
if(!confirm(`Filter-Optimierung für DAD143 ${x}\n\nDie Waage muss leer sein und darf während der Messung nicht berührt werden.\nGetestet werden Filterstufen 1–8 (${m?'FIR':'IIR'}), je ${sec} s in zwei Durchgängen (ca. ${Math.ceil(16*(sec*1.3+2)/60)} min).\nDie Werte werden nur flüchtig geschrieben; gespeichert wird erst nach Bestätigung.`))return;
clearInterval(T);let F0=+await RP(fl,1),M0=+await RP(fm,1),dp=await RD(0x2300,0x0B,0),ds=await RD(0x2300,0x0C,1),d=(ds||1)*10**-dp,R={},cs=$('cs').textContent,done=false;
const val=k=>Math.min(...R[k].ns),bad=k=>R[k].ns.length>1&&Math.max(...R[k].ns)>3*Math.max(val(k),0.5)&&Math.max(...R[k].ns)-val(k)>2,f1=v=>v.toFixed(1);
const show=()=>{$('fo').textContent=[`DAD143 ${x} · Ziel ≤ ${tg} d (1 d = ${+d.toFixed(6)}) · vorher FM ${M0} / FL ${F0}`,'Unruhe = Spanne P5–P95 ohne Drift; maßgeblich ist der bessere Durchgang, ! = Störung in einem Durchgang','FL  Grenzfreq. Hz  Einschwingzeit ms  Unruhe d (1./2.)  Drift d (1./2.)  Std.-Abw. d',...Object.keys(R).sort((a,b)=>a-b).map(k=>`${k.padStart(2)}  ${String(FC[m][k]).padStart(14)}  ${String(FT[m][k]).padStart(17)}  ${(R[k].ns.map(f1).join(' / ')+(bad(k)?' !':'  ')).padStart(16)}  ${R[k].dr.map(f1).join(' / ').padStart(15)}  ${Math.min(...R[k].sd).toFixed(2).padStart(11)}`),...o].join('\n')};
try{$('cs').textContent='Filter-Optimierung';if(M0!==m)await WP(fm,m,1);
for(let pass=0;pass<2;pass++)for(let k of(pass?[8,7,6,5,4,3,2,1]:[1,2,3,4,5,6,7,8])){$('fs').textContent=`Durchgang ${pass+1}/2 · FL ${k}`;await WP(fl,k,1);await SL(Math.max(1000,2*FT[m][k]+500));
let v=[];for(let i=0;i<sec*10;i++){v.push(+await RP(gp,1));await SL(100)}let r=NS(v),mn=Math.min(...v),mx=Math.max(...v);
(R[k]??={ns:[],dr:[],sd:[]}).ns.push(r.ns/d);R[k].dr.push(r.dr/d);R[k].sd.push(r.sd/d);L(`FILTER ${x} | FM ${m} FL ${k} | Unruhe ${f1(r.ns/d)} d | Drift ${f1(r.dr/d)} d | Min/Max ${mn}/${mx} | n ${v.length}`);show()}
let best=[1,2,3,4,5,6,7,8].find(k=>val(k)<=tg+1e-9),lo=Math.min(...[1,2,3,4,5,6,7,8].map(val)),drf=Math.max(...Object.values(R).flatMap(r=>r.dr.map(Math.abs)));
if(best){o.push('',`Vorschlag: FL ${best} (schwächster Filter mit Unruhe ≤ ${tg} d, Einschwingzeit ${FT[m][best]} ms).`);show();
if(confirm(`${o.at(-1)}\n\nFM ${m} / FL ${best} übernehmen und dauerhaft speichern (Setup, 0x2004:03)?\nAbbrechen = vorherige Einstellung FM ${M0} / FL ${F0} wiederherstellen.`)){await WP(fl,best,1);await WP(CP(0x2004,3),0,1);done=true;o.push(`Übernommen und gespeichert: FM ${m} / FL ${best}.`);L(`FILTER ${x} | gespeichert FM ${m} FL ${best}`)}else o.push('Nicht übernommen.')}
else{o.push('',`Kein Filter erreicht ${tg} d (beste Stufe: ${f1(lo)} d). Ein stärkerer Filter hilft hier nicht – keine Änderung.`,drf>tg?`Der Messwert wandert langsam (Drift bis ${f1(drf)} d in ${sec} s): Ursache eher Temperatur, Kriechen, Luftzug oder mechanischer Kontakt.`:'Die Unruhe sinkt mit stärkerem Filter kaum: Ursache eher Vibration im sehr niedrigen Frequenzbereich, Kontakt oder Einstreuung (Schirmung, Erdung).');show();alert(o.slice(1).join('\n'))}
}catch(e){o.push('Abgebrochen: '+e.message);M(e.message)}
finally{if(!done){try{await WP(fl,F0,1);if(M0!==m)await WP(fm,M0,1);o.push(`Wiederhergestellt: FM ${M0} / FL ${F0} (nicht gespeichert, EEPROM unverändert).`)}catch(e){o.push('Wiederherstellen fehlgeschlagen: '+e.message)}}$('cs').textContent=cs;$('fs').textContent=done?'Gespeichert':'Bereit';show();V()}}
