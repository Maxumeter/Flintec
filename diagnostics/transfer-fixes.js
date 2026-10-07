const G=[0x2100,0x2200,0x2500,0x2600,0x2680,0x2700,0x2800,0x2d00];
// 0x2300 sub-indices that execute a calibration (zero 0x0A, gain 0x04) instead of storing a value. Never transfer them.
const XC=[0x04,0x0A],XN=/calib|gain|span|adjust|command|execute/i;
function KX(i,s,n){return i===0x2300&&(s===3||XC.includes(s)||XN.test(n||''))}
function K(p){if(p.Index===0x2300&&!KX(p.Index,p.SubIndex,p.EntryName)&&!/^Not used/.test(p.EntryName))return'cal';if(G.includes(p.Index))return'cfg';return''}
function Q(c,k){return P.filter(p=>CR(p)&&CW(p)&&((c&&K(p)==='cfg')||(k&&K(p)==='cal')))}function PG(n,t,s){$('pr').value=t?n/t*100:0;$('ps').textContent=s||`${n}/${t}`}async function f30(){if(!C)return M('DAD143 auswählen.');let a=Q($('pc').checked,$('pk').checked),o=[];if(!a.length)return M('Keine Auswahl.');for(let i=0;i<a.length;i++){let p=a[i];PG(i,a.length,`Lese ${i+1}/${a.length}`);try{o.push({index:p.Index,subIndex:p.SubIndex,name:p.EntryName,type:p.DataTypeName,value:await RP(p,1),category:K(p)})}catch{}}let f={formatVersion:1,device:'DAD143',created:new Date().toISOString(),parameters:o},u=URL.createObjectURL(new Blob([JSON.stringify(f,null,2)])),a1=document.createElement('a');a1.href=u;a1.download=`DAD143_${Date.now()}.flintecprofile`;a1.click();URL.revokeObjectURL(u);PG(a.length,a.length,`${o.length} gespeichert`)}async function f31(f){if(!f)return;try{let o=JSON.parse(await f.text());if(o.formatVersion!==1||!Array.isArray(o.parameters))throw Error('Ungültiges Profil');F=o;$('pi').textContent=`${f.name} · ${o.parameters.length} Parameter`;$('pd').textContent='Geladen.'}catch(e){F=null;M(e.message)}}
// Profiles saved by older versions still contain the calibration commands; filter them here as well.
function E(){return F?F.parameters.filter(e=>!KX(e.index,e.subIndex,e.name)&&((e.category==='cfg'&&$('ic').checked)||(e.category==='cal'&&$('ik').checked))):[]}
function EX(){return F?F.parameters.filter(e=>e.category==='cal'&&KX(e.index,e.subIndex,e.name)):[]}
function PN(e){return`0x${e.index.toString(16).toUpperCase()}:0x${e.subIndex.toString(16).toUpperCase().padStart(2,'0')} ${e.name||''}`}
async function f33(){if(!C||!F)return M('Profil und DAD143-Verbindung erforderlich.');let d=[];for(let e of E()){let p=P.find(x=>x.Index===e.index&&x.SubIndex===e.subIndex);if(p&&CR(p))try{let v=await RP(p,1);if(String(v)!=String(e.value))d.push(`${p.IndexHex}:${p.SubIndexHex} ${v}→${e.value}`)}catch{}}$('pd').textContent=d.join('\n')||'Gleich.'}
async function f34(){if(!MC||!F)return M('Profil/Masterverbindung fehlt.');let a=E(),d=DS(),x0=$('ik').checked?EX():[];if(!a.length||!d.length)return M('Keine Auswahl.');
let ca=a.filter(e=>e.category==='cal');
if(ca.length&&!confirm(`Kalibrier-Einstellwerte auf alle ausgewählten Geräte schreiben?\n\n${ca.map(PN).join('\n')}`+(x0.length?`\n\nNICHT übertragen (Kalibrierbefehle, nur pro Waage mit Gewicht ausführbar):\n${x0.map(PN).join('\n')}`:'')+'\n\nNullpunkt und Kalibriergewicht müssen danach auf jeder Waage einzeln kalibriert werden.'))return;
if(!confirm(`${d.length} DAD143 · ${a.length} Parameter übertragen?`))return;
// Zero calibration (0x2300:0x0A with 0, as in the calibration dialog) is the only calibration command the transfer may run, and only on request: each target scale must be empty.
let zc=!!ca.length&&confirm(`Nullpunkt auf jeder Zielwaage kalibrieren?\n\nAlle ${d.length} Zielwaagen müssen vollständig entlastet sein und ruhig stehen.\nDie Verstärkung (Kalibriergewicht) wird nicht verändert.\n\nAbbrechen = Nullpunkt nicht kalibrieren.`);
let sv=confirm('Nach erfolgreicher Übertragung und Kontrolle dauerhaft im EEPROM speichern?\n\nGespeichert wird nur auf Geräten ohne Schreib- oder Prüffehler.\nAbbrechen = nur flüchtig schreiben.');
let o=[],n=0,t=d.length*a.length,orig=DA,cs=$('cs').textContent;clearInterval(T);
for(let x of d){let okc=0,er=[],sk=0,w=[];try{await f35(x);$('cs').textContent=`Profilübertragung ${x}`;let cb=false;
for(let e of a){let p=P.find(y=>y.Index===e.index&&y.SubIndex===e.subIndex);PG(n++,t,`Gerät ${x} · ${okc+er.length+sk+1}/${a.length}`);if(!p||!CW(p))continue;
let c=e.category==='cal';if(c&&cb){sk++;continue}
// Calibration parameters need a fresh TAC unlock directly before every write, exactly like the calibration dialog (CWZ).
try{c?await CWZ(p,e.value):await WP(p,e.value,1);okc++;w.push([p,e])}catch(z){er.push(`${PN(e)}: ${z.message}`);if(c)cb=true}}
let zs='';if(zc){if(cb)zs=' · Nullpunkt übersprungen';else try{PG(n,t,`Gerät ${x} · Nullpunkt`);await CWZ(CP(0x2300,0x0A),0);zs=' · Nullpunkt kalibriert'}catch(z){er.push(`0x2300:0x0A Nullpunkt kalibrieren: ${z.message}`);zs=' · Nullpunkt-Fehler'}}
let mm=[];for(let[p,e]of w)try{let v=await RP(p,1);if(String(v)!=String(e.value))mm.push(`${PN(e)}: ${v}≠${e.value}`)}catch(z){mm.push(`${PN(e)}: ${z.message}`)}
let s='';if(sv){if(er.length||mm.length||sk)s=' · NICHT gespeichert';else try{await WP(CP(0x2004,2),0,1);s=' · EEPROM gespeichert'}catch(z){s=` · EEPROM-Fehler: ${z.message}`}}else s=' · nicht dauerhaft gespeichert';
o.push(`${x}: ${okc} OK · ${er.length} Fehler`+(sk?` · ${sk} übersprungen`:'')+(mm.length?` · ${mm.length} Abweichungen`:'')+zs+s);for(let z of[...er,...mm])o.push('   '+z)}catch(e){o.push(`${x}: ${e.message}`);n+=a.length}}
$('cs').textContent=cs;if(orig&&B.some(x=>x.address===orig&&x.isDAD143))try{await f35(orig)}catch{}PG(t,t,'Fertig');$('pd').textContent=o.join('\n');M('Übertragen.');V()}
