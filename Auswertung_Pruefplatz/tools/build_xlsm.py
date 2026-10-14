#!/usr/bin/env python3
"""Erzeugt Pruefplatz_Auswertung.xlsm (Startblatt mit Schaltflaechen + VBA-Modul).

Das VBA-Projekt (vbaProject.bin) wird nach MS-OVBA aus dem Quelltext
Pruefplatz_Auswertung.bas aufgebaut. Es enthaelt nur Quelltext und keinen
kompilierten P-Code (_VBA_PROJECT-Version 0xFFFF), daher kompiliert Excel
das Modul beim ersten Oeffnen selbst. Projektinformationen und Verweise
(stdole, Office) werden aus einer von Excel erzeugten Vorlage uebernommen
(z. B. examples/vbaProject.bin aus dem XlsxWriter-Quellpaket).

Aufruf:
    python3 build_xlsm.py <vorlage_vbaProject.bin> <modul.bas> <ziel.xlsm>

Benoetigt: xlsxwriter, olefile, oletools (nur zur Pruefung).
"""
import struct
import sys
import uuid

import olefile
import xlsxwriter

CODEPAGE = "cp1252"


# --------------------------------------------------------------------------
# MS-OVBA 2.4.1  Kompression
# --------------------------------------------------------------------------
def _compress_chunk(data):
    out = bytearray()
    pos = 0
    n = len(data)
    seen = {}                    # 3-Byte-Praefix -> Positionen im Chunk
    while pos < n:
        flag_pos = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= n:
                break
            # Bitaufteilung des Copy-Tokens haengt von der Position im Chunk ab
            bit_count = max((pos - 1).bit_length(), 4) if pos > 0 else 4
            max_len = (0xFFFF >> bit_count) + 3
            best_len, best_off = 0, 0
            key = data[pos:pos + 3]
            if len(key) == 3:
                for cand in reversed(seen.get(key, [])[-64:]):
                    if pos - cand > (1 << bit_count):
                        break
                    ln = 0
                    while (ln < max_len and pos + ln < n
                           and data[cand + ln] == data[pos + ln]):
                        ln += 1
                    if ln > best_len:
                        best_len, best_off = ln, pos - cand
                        if ln == max_len:
                            break
            step = best_len if best_len >= 3 else 1
            for q in range(pos, min(pos + step, n - 2)):
                seen.setdefault(data[q:q + 3], []).append(q)
            if best_len >= 3:
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack("<H", token)
                flags |= 1 << bit
            else:
                out.append(data[pos])
            pos += step
        out[flag_pos] = flags
    return bytes(out)


def ovba_compress(data):
    out = bytearray(b"\x01")
    for i in range(0, max(len(data), 1), 4096):
        chunk = data[i:i + 4096]
        comp = _compress_chunk(chunk)
        if len(comp) <= 4094:
            header = 0xB000 | (len(comp) + 2 - 3)
            out += struct.pack("<H", header) + comp
        else:  # unkomprimierter Chunk (muss genau 4096 Bytes haben)
            assert len(chunk) == 4096, "letzter Chunk nicht komprimierbar"
            out += struct.pack("<H", 0x3000 | (4096 + 2 - 3)) + chunk
    return bytes(out)


def ovba_decompress(data):
    from oletools.olevba import decompress_stream
    return bytes(decompress_stream(bytearray(data)))


# --------------------------------------------------------------------------
# Compound File Binary (MS-CFB, Version 3) - minimaler Schreiber
# --------------------------------------------------------------------------
ENDOFCHAIN, FREESECT, FATSECT, NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF


def _cfb_key(name):
    return (len(name), name.upper())


class _Node:
    def __init__(self, name, kind, data=b""):
        self.name, self.kind, self.data = name, kind, data   # kind: 1 Storage, 2 Stream, 5 Root
        self.children = []
        self.left = self.right = self.child = NOSTREAM
        self.start, self.size, self.sid = ENDOFCHAIN, 0, 0


def write_cfb(streams, path):
    """streams: dict 'A/B' -> bytes"""
    root = _Node("Root Entry", 5)
    storages = {"": root}
    for full in sorted(streams):
        parts = full.split("/")
        parent = root
        for i, p in enumerate(parts[:-1]):
            key = "/".join(parts[:i + 1])
            if key not in storages:
                st = _Node(p, 1)
                parent.children.append(st)
                storages[key] = st
            parent = storages[key]
        parent.children.append(_Node(parts[-1], 2, streams[full]))

    nodes = []

    def number(n):
        n.sid = len(nodes)
        nodes.append(n)
        for c in n.children:
            number(c)
    number(root)

    def tree(children):          # balancierter Binaerbaum nach CFB-Sortierung
        if not children:
            return NOSTREAM
        children = sorted(children, key=lambda c: _cfb_key(c.name))
        m = len(children) // 2
        mid = children[m]
        mid.left = tree(children[:m])
        mid.right = tree(children[m + 1:])
        return mid.sid
    for n in nodes:
        if n.kind in (1, 5):
            n.child = tree(n.children)

    SS, MSS, CUTOFF = 512, 64, 4096
    fat, minifat = [], []
    sectors = []                 # Inhalte der regulaeren Sektoren

    def alloc(data, table, size):
        count = max(1, (len(data) + size - 1) // size)
        first = len(table)
        for i in range(count):
            table.append(first + i + 1 if i < count - 1 else ENDOFCHAIN)
        return first, count

    # Mini-Stream fuer kleine Streams
    ministream = bytearray()
    for n in nodes:
        if n.kind == 2:
            n.size = len(n.data)
            if n.size < CUTOFF:
                if n.size == 0:
                    n.start = ENDOFCHAIN
                    continue
                first, cnt = alloc(n.data, minifat, MSS)
                n.start = first
                ministream += n.data.ljust(cnt * MSS, b"\x00")

    def put(data):
        first, cnt = alloc(data, fat, SS)
        for i in range(cnt):
            sectors.append(data[i * SS:(i + 1) * SS].ljust(SS, b"\x00"))
        return first

    for n in nodes:
        if n.kind == 2 and n.size >= CUTOFF:
            n.start = put(n.data)
    if ministream:
        root.start = put(bytes(ministream))
        root.size = len(ministream)
    minifat_bytes = b"".join(struct.pack("<I", x) for x in minifat)
    minifat_bytes = minifat_bytes.ljust(((len(minifat_bytes) + SS - 1) // SS) * SS, b"\xff")
    minifat_start = put(minifat_bytes) if minifat else ENDOFCHAIN
    minifat_count = len(minifat_bytes) // SS if minifat else 0

    def dir_entry(n):
        name = n.name.encode("utf-16-le") + b"\x00\x00"
        e = name.ljust(64, b"\x00")
        e += struct.pack("<HBB", len(name), n.kind, 1)        # Farbe: schwarz
        e += struct.pack("<III", n.left, n.right, n.child)
        e += b"\x00" * 16 + struct.pack("<I", 0) + b"\x00" * 16
        e += struct.pack("<IQ", n.start if n.kind != 1 else 0, n.size if n.kind != 1 else 0)
        return e
    dirdata = b"".join(dir_entry(n) for n in nodes)
    while len(dirdata) % SS:
        dirdata += (b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0)
                    + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM) + b"\x00" * 52)
    dir_start = put(dirdata)

    # FAT-Sektoren selbst eintragen
    nfat = 1
    while True:
        total = len(fat) + nfat
        if (total + 127) // 128 <= nfat:
            break
        nfat += 1
    fat_start = len(fat)
    fat += [FATSECT] * nfat
    fat_bytes = b"".join(struct.pack("<I", x) for x in fat)
    fat_bytes = fat_bytes.ljust(nfat * SS, b"\xff")
    for i in range(nfat):
        sectors.append(fat_bytes[i * SS:(i + 1) * SS])
    assert nfat <= 109

    hdr = bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 16
    hdr += struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\x00" * 6
    hdr += struct.pack("<IIIIIIIII", 0, nfat, dir_start, 0, CUTOFF,
                       minifat_start, minifat_count, ENDOFCHAIN, 0)
    difat = [fat_start + i for i in range(nfat)] + [FREESECT] * (109 - nfat)
    hdr += b"".join(struct.pack("<I", x) for x in difat)
    assert len(hdr) == 512
    with open(path, "wb") as f:
        f.write(hdr + b"".join(sectors))


# --------------------------------------------------------------------------
# VBA-Projekt aufbauen
# --------------------------------------------------------------------------
def rec(rid, data):
    return struct.pack("<HI", rid, len(data)) + data


def build_vba_project(template, bas_path, out_bin):
    t = olefile.OleFileIO(template)
    tdir = ovba_decompress(t.openstream("VBA/dir").read())
    # PROJECTINFORMATION + PROJECTREFERENCES bis vor PROJECTMODULES (0x000F) uebernehmen
    i = 0
    while True:
        rid, sz = struct.unpack_from("<HI", tdir, i)
        if rid == 0x000F:
            break
        i += 6 + (6 if rid == 0x0009 else sz)
    head = tdir[:i]

    def doc_header(stream):
        raw = t.openstream("VBA/" + stream).read()
        # Offset des Quelltexts aus dem dir-Stream der Vorlage ermitteln
        nm = stream.encode()
        p = tdir.index(rec(0x001A, nm))
        p = tdir.index(struct.pack("<HI", 0x0031, 4), p)
        off = struct.unpack_from("<I", tdir, p + 6)[0]
        return ovba_decompress(raw[off:]).decode(CODEPAGE)

    src = open(bas_path, encoding="utf-8", newline="").read()
    src = src.replace("\r\n", "\n").replace("\n", "\r\n")
    assert src.startswith('Attribute VB_Name = "Pruefplatz_Auswertung"')
    modules = [
        ("ThisWorkbook", 0x22, doc_header("ThisWorkbook")),
        ("Sheet1", 0x22, doc_header("Sheet1")),
        ("Pruefplatz_Auswertung", 0x21, src),
    ]

    body = rec(0x000F, struct.pack("<H", len(modules))) + rec(0x0013, struct.pack("<H", 0xFFFF))
    streams = {}
    for name, typ, code in modules:
        nb, nu = name.encode(CODEPAGE), name.encode("utf-16-le")
        body += rec(0x0019, nb) + rec(0x0047, nu)
        body += rec(0x001A, nb) + rec(0x0032, nu)
        body += rec(0x001C, b"") + rec(0x0048, b"")
        body += rec(0x0031, struct.pack("<I", 0))           # Quelltext ab Offset 0
        body += rec(0x001E, struct.pack("<I", 0))
        body += rec(0x002C, struct.pack("<H", 0xFFFF))
        body += rec(typ, b"")
        body += rec(0x002B, b"")
        streams["VBA/" + name] = ovba_compress(code.encode(CODEPAGE))
    body += rec(0x0010, b"")
    dirdata = head + body
    streams["VBA/dir"] = ovba_compress(dirdata)
    streams["VBA/_VBA_PROJECT"] = struct.pack("<HHBH", 0x61CC, 0xFFFF, 0x00, 0x0000)

    tproj = t.openstream("PROJECT").read().decode(CODEPAGE)
    keep = {k: v for k, v in (l.split("=", 1) for l in tproj.split("\r\n") if l[:3] in ("CMG", "DPB", "GC="))}
    proj = [
        'ID="{%s}"' % str(uuid.uuid4()).upper(),
        "Document=ThisWorkbook/&H00000000",
        "Document=Sheet1/&H00000000",
        "Module=Pruefplatz_Auswertung",
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        "CMG=" + keep["CMG"],
        "DPB=" + keep["DPB"],
        "GC=" + keep["GC"],
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
        "ThisWorkbook=0, 0, 0, 0, C",
        "Sheet1=0, 0, 0, 0, C",
        "Pruefplatz_Auswertung=0, 0, 0, 0, C",
        "",
    ]
    streams["PROJECT"] = "\r\n".join(proj).encode(CODEPAGE)
    wm = b""
    for name, _, _ in modules:
        wm += name.encode(CODEPAGE) + b"\x00" + name.encode("utf-16-le") + b"\x00\x00"
    streams["PROJECTwm"] = wm + b"\x00\x00"
    write_cfb(streams, out_bin)

    # Selbstpruefung: Rundreise der Kompression und Lesbarkeit der Datei
    for name, _, code in modules:
        assert ovba_decompress(streams["VBA/" + name]) == code.encode(CODEPAGE), name
    assert ovba_decompress(streams["VBA/dir"]) == dirdata


# --------------------------------------------------------------------------
# Arbeitsmappe mit Startblatt und Schaltflaechen
# --------------------------------------------------------------------------
def build_workbook(vba_bin, out_xlsm):
    wb = xlsxwriter.Workbook(out_xlsm)
    wb.set_vba_name("ThisWorkbook")
    ws = wb.add_worksheet("Start")
    ws.set_vba_name("Sheet1")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 3)
    ws.set_column("B:D", 14)
    ws.set_column("G:G", 90)

    title = wb.add_format({"bold": True, "font_size": 20, "font_color": "#1F4E79"})
    sub = wb.add_format({"font_color": "#767676"})
    info = wb.add_format({"font_color": "#595959", "text_wrap": True, "valign": "vcenter"})
    head = wb.add_format({"bold": True, "font_size": 12, "font_color": "#1F4E79"})
    warn = wb.add_format({"italic": True, "font_color": "#C00000"})
    txt = wb.add_format({"text_wrap": True, "valign": "top"})

    ws.write("B2", "Pruefplatz-Auswertung", title)
    ws.write("B3", "Grafische Auswertung und Analyse der Pruefplatz-CSV-Daten", sub)

    buttons = [
        ("B5", "1.  CSV importieren und auswerten", "Start_Auswertung",
         "CSV-Datei auswaehlen. Alle Auswertungsblaetter und Diagramme werden (neu) erzeugt."),
        ("B9", "2.  Neu berechnen", "Neu_Berechnen",
         "Nach Aenderungen im Blatt 'Einstellungen' (Grenzwerte, Einheiten, Zeitzone) - ohne erneuten Import."),
        ("B13", "Einstellungen zuruecksetzen", "Einstellungen_Zuruecksetzen",
         "Grenzwerte und Einheiten auf die Standardwerte zuruecksetzen."),
    ]
    for cell, caption, macro, text in buttons:
        ws.insert_button(cell, {"macro": macro, "caption": caption, "width": 260, "height": 48})
        row = int(cell[1:])
        ws.write(row, 6, text, info)

    ws.write("B18", "Bedienung", head)
    lines = [
        "1. Beim Oeffnen in der gelben Leiste 'Inhalt aktivieren' klicken (Makros zulassen).",
        "2. Schaltflaeche '1. CSV importieren und auswerten' klicken und die CSV-Datei des Pruefplatzes waehlen.",
        "3. Ergebnisse: Blaetter Uebersicht, Diagramme, Statistik, Wochen, NIO_Liste, Daten, Datenqualitaet.",
        "4. Grenzwerte im Blatt 'Einstellungen' anpassen und '2. Neu berechnen' klicken.",
    ]
    for i, l in enumerate(lines):
        ws.write(18 + i, 1, l)
    ws.write("B24", "Hinweis: Die Standard-Grenzwerte (Isolation >= 5000 MOhm, 5-6 R >= 100 MOhm, "
                    "Durchgang <= 10 Ohm) sind Annahmen und muessen an die Pruefspezifikation angepasst werden.", warn)
    ws.activate()
    wb.add_vba_project(vba_bin)
    wb.close()


if __name__ == "__main__":
    template, bas, target = sys.argv[1:4]
    vba_bin = target + ".vbaProject.bin"
    build_vba_project(template, bas, vba_bin)
    build_workbook(vba_bin, target)
    print("geschrieben:", target)
