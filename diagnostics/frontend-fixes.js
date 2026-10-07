// v1.9.4: restored frontend replacements; no changes to calibration commands.
function L(s) {
  const node = $('log');
  const lines = (node.textContent + `[${new Date().toISOString()}] ${s}\n`).split('\n');
  node.textContent = lines.slice(-3001).join('\n');
  node.scrollTop = node.scrollHeight;
}
function exportLog() {
  const data = `Flintec Control Center 1.9.4\n${$('log').textContent}`;
  const url = URL.createObjectURL(new Blob([data], {type:'text/plain;charset=utf-8'}));
  const a = document.createElement('a');
  a.href = url; a.download = `Flintec_Log_${new Date().toISOString().replace(/[:.]/g,'-')}.txt`;
  a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
}
async function A(p, o = {}) {
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
async function f26(a, stayBus = false) {
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
