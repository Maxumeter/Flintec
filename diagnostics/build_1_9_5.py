"""Reproducible recovery build from an exact v1.9.3 binary; not a Go rebuild."""
import pathlib,sys,struct,hashlib,json,re
p=pathlib.Path(__file__).parent
sys.path.insert(0,str(p/'python-libs'))
import pefile,capstone
root=p.parent; source=root/'source'; dist=root/'release-1.9.5'
source.mkdir(exist_ok=True); dist.mkdir(exist_ok=True)
b=bytearray((p/'installed-1.9.3.exe').read_bytes())
assert hashlib.sha256(b).hexdigest()=='4211ec5384da2b1e260a16fccfb2c353ec1e62e9ffebb72ead4c7c1eb5aa1fcb'
pe=pefile.PE(data=b); base=pe.OPTIONAL_HEADER.ImageBase
html_offset=3682240; ref=0x4107a8
ptr,length=struct.unpack_from('<QQ',b,ref)
assert ptr==base+pe.get_rva_from_offset(html_offset) and length==26647
html=bytes(b[html_offset:html_offset+length]).decode('utf-8')
fix=(p/'frontend-fixes.js').read_text(encoding='utf-8')
def replace_between(s,start,end,new):
    assert s.count(start)==1,(start,s.count(start))
    i=s.index(start); j=s.index(end,i)
    return s[:i]+new+s[j:]
html=replace_between(html,'function L(s){','function M(s)',fix[fix.index('function L(s)'):fix.index('async function A(')])
html=replace_between(html,'async function A(p,o={})','function tab(id)',fix[fix.index('async function A('):fix.index('async function f26(')])
html=replace_between(html,'async function f26(a,stayBus=false)','const G=',fix[fix.index('async function f26('):]+'\n')
html=html.replace('let MC=false,C=false,','let DA=0,MC=false,C=false,')
html=html.replace('MC=!!master;C=!!device;', 'MC=!!master;C=!!device;DA=C?Number(address):0;')
html=html.replace("let cur=C?(+$('address').value||0):0;",'let cur=DA;')
html=html.replace("orig=+$('address').value||0;",'orig=DA;')
html=html.replace('Z=q;','Z={ip:q.ip,master:q.master,address:q.address};')
html=html.replace("catch(e){L('Dashboard: '+e.message)}",'catch(e){}')
html=html.replace("<button class=danger onclick=\"shutdown()\">Beenden</button>",'<button onclick="exportLog()">Log exportieren</button><span class=muted>Letzte 3000 Zeilen · Zugangsdaten werden nicht protokolliert</span><button class=danger onclick="shutdown()">Beenden</button>')
html=replace_between(html,'const G=','async function f35(',(p/'transfer-fixes.js').read_text(encoding='utf-8'))
html=html.replace('1.9.3','1.9.5')
html=html.replace('Flintec Control Center 1.9.4','Flintec Control Center 1.9.5')
html=html.replace('init().catch(e=>M(e.message));',"L('START | Version 1.9.5 (Build "+hashlib.sha256((p/'transfer-fixes.js').read_bytes()).hexdigest()[:8]+") | SDO-Fehler zeigen den letzten Anfrageversuch.');init().catch(e=>M(e.message));")
html=re.sub(r' {50,}','',html)
(source/'index.html').write_text(html,encoding='utf-8')
js=html.split('<script>',1)[1].split('</script>',1)[0]
(source/'app.js').write_text(js,encoding='utf-8')
payload=html.encode('utf-8')
align=lambda x,a:(x+a-1)//a*a
last=pe.sections[-1]; sec_off=last.get_file_offset()+40
assert sec_off+40<=pe.sections[0].PointerToRawData
raw_offset=align(len(b),pe.OPTIONAL_HEADER.FileAlignment)
rva=align(max(s.VirtualAddress+max(s.Misc_VirtualSize,s.SizeOfRawData) for s in pe.sections),pe.OPTIONAL_HEADER.SectionAlignment)
raw_size=align(len(payload),pe.OPTIONAL_HEADER.FileAlignment)
b.extend(b'\0'*(raw_offset-len(b))+payload+b'\0'*(raw_size-len(payload)))
b[sec_off:sec_off+40]=struct.pack('<8sIIIIIIHHI',b'.fccweb\0',len(payload),rva,raw_size,raw_offset,0,0,0,0,0x40000040)
struct.pack_into('<H',b,pe.FILE_HEADER.get_field_absolute_offset('NumberOfSections'),pe.FILE_HEADER.NumberOfSections+1)
struct.pack_into('<I',b,pe.OPTIONAL_HEADER.get_field_absolute_offset('SizeOfImage'),align(rva+len(payload),pe.OPTIONAL_HEADER.SectionAlignment))
struct.pack_into('<I',b,pe.OPTIONAL_HEADER.get_field_absolute_offset('SizeOfInitializedData'),pe.OPTIONAL_HEADER.SizeOfInitializedData+raw_size)
struct.pack_into('<QQ',b,ref,base+rva,len(payload))
b[ref+16:ref+32]=hashlib.sha256(payload).digest()[:16]
# Preserve the second failed read response; verified instruction and target.
off=pe.get_offset_from_rva(0x6c8234-base)
assert b[off:off+5]==bytes.fromhex('488b4c2460')
b[off:off+5]=bytes.fromhex('e91c000000')
version_count=b.count(b'1.9.3'); b=b.replace(b'1.9.3',b'1.9.5')
check=pefile.PE(data=b)
struct.pack_into('<I',b,check.OPTIONAL_HEADER.get_field_absolute_offset('CheckSum'),check.generate_checksum())
out=dist/'Flintec_ControlCenter_1.9.5_Portable.exe';out.write_bytes(b)
verified=pefile.PE(data=b)
assert verified.sections[-1].Name.rstrip(b'\0')==b'.fccweb'
assert b[raw_offset:raw_offset+len(payload)]==payload
# Machine instructions outside the already reviewed diagnostic patch must match.
original=(p/'installed-1.9.3.exe').read_bytes(); ts=pe.sections[0]; lo=ts.PointerToRawData; hi=lo+ts.SizeOfRawData
assert b[lo:off]==original[lo:off] and b[off+5:hi]==original[off+5:hi]
manifest={'version':'1.9.5','build_method':'Recovery patch of exact v1.9.3 binary, restored HTML source; no original Go sources available','base_sha256':hashlib.sha256(original).hexdigest(),'sha256':hashlib.sha256(b).hexdigest(),'html_bytes':len(payload),'version_literals_updated':version_count,'machine_code_change':'Only 5-byte SDO read error-selection patch at VA 0x6c8234; write code unchanged','changes':['Always confirm selected device through backend','Track confirmed address independently of input','Clear stale values on failed selection','Contextual request logs with sequence ID, object, type, length, write value, result, duration and calibration status','Log export and 3000-line bound','Suppress repeated dashboard errors; log recovery','Do not log login payloads','Profile transfer: never transfer calibration commands 0x2300:0x04 (gain) and 0x2300:0x0A (zero)','Profile transfer: fresh TAC unlock before every calibration parameter write','Profile transfer: stop calibration block per device after first error, read back and compare written values','Profile transfer: optional EEPROM save only on devices without errors'],'limitations':['Original calibration timeout not fixed','Log describes local API operation; backend does not expose both HTTP attempts together']}
(dist/'build-manifest.json').write_text(json.dumps(manifest,indent=2),encoding='utf-8')
(dist/'SHA256SUMS.txt').write_text(f'{manifest["sha256"]}  {out.name}\n')
print(json.dumps(manifest,indent=2))
