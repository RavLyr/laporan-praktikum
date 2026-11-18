#!/usr/bin/env python3
"""Bangun laporan praktikum (.docx) dari spesifikasi JSON memakai template resmi.

Pakai: python build_laprak.py spec.json [-o keluaran.docx]
Skema spec: references/spec-schema.md
Profil format: assets/profil.json
"""
import argparse
import json
import os
import re
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, os.pardir, "assets")
TNR = "Times New Roman"

ISTILAH = [
    "access point", "default gateway", "gateway", "interface", "port",
    "router", "switch", "server", "client", "host", "hostname", "pool",
    "wireless", "username", "password", "passphrase", "browser",
    "firmware", "lightweight", "management", "guest", "staff",
    "end-to-end", "troubleshooting", "ping", "login", "setup", "subnet",
    "trainer board", "breadboard", "jumper", "software", "hardware",
    "forward bias", "reverse bias", "bias", "output", "input", "probe",
    "screenshot", "sheet", "datasheet", "power supply", "ground",
    "clock", "flip-flop", "truth table", "timer", "counter", "gate",
]


def el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn("w:" + k), str(v))
    return e


def load_profiles():
    with open(os.path.join(ASSETS, "profil.json"), encoding="utf-8") as f:
        return json.load(f)


# ---------- penomoran ----------

def patch_numbering(doc, bab):
    """Ganti nomor BAB pada tiap lvlText dan start level teratas."""
    numb = doc.part.numbering_part.element
    for lvl in numb.iter(qn("w:lvl")):
        fmt = lvl.find(qn("w:numFmt"))
        txt = lvl.find(qn("w:lvlText"))
        start = lvl.find(qn("w:start"))
        if txt is not None:
            v = txt.get(qn("w:val"))
            txt.set(qn("w:val"), re.sub(r"(?<![0-9%])5(?=[.%])", str(bab), v))
        if start is not None and lvl.get(qn("w:ilvl")) == "0":
            fv = fmt.get(qn("w:val")) if fmt is not None else ""
            if fv in ("upperRoman", "decimal") and start.get(qn("w:val")) == "5":
                start.set(qn("w:val"), str(bab))


def new_num(doc, levels, left=720, hanging=720):
    """Buat daftar bernomor baru agar penomoran mulai dari 1."""
    numb = doc.part.numbering_part.element
    aid = max(int(a.get(qn("w:abstractNumId"))) for a in numb.findall(qn("w:abstractNum"))) + 1
    nid = max(int(n.get(qn("w:numId"))) for n in numb.findall(qn("w:num"))) + 1

    a = el("w:abstractNum", abstractNumId=aid)
    a.append(el("w:multiLevelType", val="hybridMultilevel"))
    for i, (fmt, text) in enumerate(levels):
        lvl = el("w:lvl", ilvl=i)
        lvl.append(el("w:start", val=1))
        lvl.append(el("w:numFmt", val=fmt))
        lvl.append(el("w:lvlText", val=text))
        lvl.append(el("w:lvlJc", val="left"))
        pPr = el("w:pPr")
        pPr.append(el("w:ind", left=left + 720 * i, hanging=hanging if i == 0 else 360))
        lvl.append(pPr)
        a.append(lvl)
    numb.findall(qn("w:abstractNum"))[-1].addnext(a)

    n = el("w:num", numId=nid)
    n.append(el("w:abstractNumId", val=aid))
    numb.append(n)
    return nid


def set_num(p, num_id, ilvl=0):
    pPr = p._p.get_or_add_pPr()
    numPr = el("w:numPr")
    numPr.append(el("w:ilvl", val=ilvl))
    numPr.append(el("w:numId", val=num_id))
    pPr.insert(0, numPr)


def set_ind(p, **kw):
    pPr = p._p.get_or_add_pPr()
    old = pPr.find(qn("w:ind"))
    if old is not None:
        pPr.remove(old)
    pPr.append(el("w:ind", **kw))


# ---------- teks ----------

def split_italic(text, istilah):
    parts = []
    for chunk in re.split(r"(\*[^*]+\*)", text):
        if not chunk:
            continue
        if chunk.startswith("*") and chunk.endswith("*") and len(chunk) > 2:
            parts.append((chunk[1:-1], True))
        else:
            parts.extend(auto_italic(chunk, istilah))
    return parts


def auto_italic(text, istilah):
    if not istilah:
        return [(text, False)]
    # tanda hubung dihitung bagian kata agar "port-security" tidak terpenggal
    pat = re.compile(r"(?<![\w-])(" + "|".join(re.escape(t) for t in sorted(istilah, key=len, reverse=True)) + r")(?![\w-])", re.I)
    out, pos = [], 0
    for m in pat.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], False))
        out.append((m.group(0), True))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], False))
    return out or [(text, False)]


def add_runs(p, text, istilah, size=None, bold=False, color=None, font=TNR, no_italic=False):
    for chunk, ital in split_italic(text, istilah):
        r = p.add_run(chunk)
        r.font.name = font
        r._element.rPr.rFonts.set(qn("w:eastAsia"), font)
        r.italic = (ital or False) if no_italic else (ital or None)
        if bold:
            r.bold = True
        if size:
            r.font.size = size
        if color:
            r.font.color.rgb = RGBColor.from_string(color)
    return p


# ---------- pembangun ----------

class Builder:
    def __init__(self, spec):
        self.spec = spec
        self.bab = int(spec["bab"])
        self.istilah = [t for t in spec.get("istilah", ISTILAH)
                        if t not in spec.get("istilah_kecuali", [])]
        self.base = spec.get("base_dir", ".")
        profiles = load_profiles()
        name = spec.get("profil", "srwe")
        if name not in profiles:
            sys.exit(f"profil tidak dikenal: {name}. Pilihan: {', '.join(profiles)}")
        self.P = profiles[name]
        self.P.update(spec.get("profil_override", {}))
        self.doc = Document(os.path.join(ASSETS, self.P["template"]))
        patch_numbering(self.doc, self.bab)
        self.pending_break = False
        self.hnum = new_num(
            self.doc,
            [("decimal", f"{self.bab}.%1"), ("decimal", f"{self.bab}.%1.%2"),
             ("decimal", f"{self.bab}.%1.%2.%3"), ("decimal", "%4.")],
            left=self.P["heading_ind"]["left"],
            hanging=self.P["heading_ind"]["hanging"],
        )

    # -- primitif --

    def img_path(self, f):
        return f if os.path.isabs(f) else os.path.join(self.base, f)

    def heading(self, text, level):
        """level 0 = x.1, level 1 = x.1.1, level 2 = x.1.1.1"""
        p = self.doc.add_paragraph(style=self.P["heading"][level])
        add_runs(p, text, self.istilah, bold=True, color="000000", no_italic=True)
        set_num(p, self.hnum, level)
        set_ind(p, **self.P["heading_ind"])
        mark = el("w:rPr")
        mark.append(el("w:b"))
        mark.append(el("w:i", val="0"))
        mark.append(el("w:color", val="000000"))
        p._p.pPr.append(mark)
        if self.pending_break:
            p.paragraph_format.page_break_before = True
            self.pending_break = False
        return p

    def para(self, text, first_line=None):
        cfg = self.P["paragraf"]
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_after = Pt(cfg["after"])
        pf.line_spacing = cfg["line"]
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        fl = cfg["first_line"] if first_line is None else first_line
        if fl:
            set_ind(p, left=0, firstLine=fl)
        add_runs(p, text, self.istilah)
        return p

    def item(self, text, num_id):
        cfg = self.P["daftar"]
        p = self.doc.add_paragraph(style="List Paragraph")
        set_num(p, num_id)
        pf = p.paragraph_format
        pf.space_after = Pt(0)
        pf.line_spacing = self.P["paragraf"]["line"]
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        set_ind(p, **cfg)
        add_runs(p, text, self.istilah)
        return p

    def caption(self, kind, text):
        gambar = kind == "Gambar"
        style = self.P["caption_gambar_style"] if gambar else self.P["caption_tabel_style"]
        p = self.doc.add_paragraph(style=style)
        rata = self.P.get("caption_gambar_rata" if gambar else "caption_tabel_rata", "tengah")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if rata == "tengah" else WD_ALIGN_PARAGRAPH.LEFT
        if rata != "tengah":
            set_ind(p, left=0, firstLine=0)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(f"{kind} {self.bab}. ")
        r.font.name = TNR
        sfx = self.P["seq_suffix"]
        fld = el("w:fldSimple", instr=f" SEQ {kind}_{self.bab}.{sfx} \\* ARABIC ")
        inner = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rPr.append(OxmlElement("w:noProof"))
        inner.append(rPr)
        t = OxmlElement("w:t")
        t.text = "1"
        inner.append(t)
        fld.append(inner)
        p._p.append(fld)
        add_runs(p, " " + text, self.istilah)
        return p

    def figure(self, path, cap):
        path = self.img_path(path)
        if not os.path.exists(path):
            sys.exit(f"gambar tidak ada: {path}")
        g = self.P["gambar"]
        with Image.open(path) as im:
            w_px, h_px = im.size
        w = min(int(Cm(g["max_w"])), int(Cm(w_px / 96 * 2.54)))
        h = int(w * h_px / w_px)
        if h > int(Cm(g["max_h"])):
            w = int(w * int(Cm(g["max_h"])) / h)
        p = self.doc.add_paragraph(style=g["style"])
        pf = p.paragraph_format
        pf.keep_with_next = True
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_ind(p, left=0, firstLine=0)
        p.add_run().add_picture(path, width=w)
        if cap:
            self.caption("Gambar", cap)
        return p

    def code(self, lines):
        c = self.P["kode"]
        t = self.doc.add_table(rows=1, cols=1)
        t.style = self.P["tabel"]["style"]
        self.borders(t)
        t._tbl.tblPr.append(el("w:tblInd", w=c["indent"], type="dxa"))
        cell = t.cell(0, 0)
        for r in list(cell.paragraphs[0].runs):
            r._element.getparent().remove(r._element)
        for i, line in enumerate(lines):
            p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            r = p.add_run(line)
            r.font.name = c["font"]
            r._element.rPr.rFonts.set(qn("w:eastAsia"), c["font"])
            r.font.size = Pt(c["size"])
        self.doc.add_paragraph()
        return t

    def table(self, spec):
        cfg = self.P["tabel"]
        self.caption("Tabel", spec["caption"])
        header, rows = spec["header"], spec["rows"]
        t = self.doc.add_table(rows=1 + len(rows), cols=len(header))
        t.style = cfg["style"]
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = t._tbl.tblPr
        tblW = tblPr.find(qn("w:tblW"))
        if tblW is None:
            tblW = el("w:tblW")
            tblPr.append(tblW)
        tblW.set(qn("w:w"), str(cfg["width"]))
        tblW.set(qn("w:type"), "dxa" if cfg["width"] else "auto")
        self.borders(t)
        for j, text in enumerate(header):
            cell = t.rows[0].cells[j]
            if cfg.get("header_fill"):
                cell._tc.get_or_add_tcPr().append(
                    el("w:shd", val="clear", color="auto", fill=cfg["header_fill"]))
            self.fill(cell, str(text), bold=True, color=cfg.get("header_color"))
        for i, row in enumerate(rows, start=1):
            for j, text in enumerate(row):
                self.fill(t.rows[i].cells[j], str(text))
        for j, w in enumerate(spec.get("widths") or []):
            for row in t.rows:
                row.cells[j].width = Cm(w)
        self.doc.add_paragraph()
        return t

    def borders(self, t):
        if not self.P["tabel"].get("border"):
            return
        b = el("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            b.append(el("w:" + side, val="single", sz=4, space=0, color="000000"))
        t._tbl.tblPr.append(b)

    def fill(self, cell, text, bold=False, color=None):
        cell._tc.get_or_add_tcPr().append(el("w:vAlign", val="center"))
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        add_runs(p, text, self.istilah, bold=bold, color=color)

    def brk(self):
        if self.P.get("break_antar_bagian"):
            self.pending_break = True

    # -- blok generik --

    def blok(self, b, num_id=None):
        """Render satu blok isi. Kunci yang dikenali boleh digabung."""
        if b.get("teks"):
            if num_id:
                self.item(b["teks"], num_id)
            else:
                self.para(b["teks"])
        for par in b.get("paragraf", []):
            self.para(par)
        if b.get("kode"):
            self.code(b["kode"])
        for tb in listify(b.get("tabel")):
            self.table(tb)
        for g in gambar_list(b):
            self.figure(g["file"], g.get("caption", ""))
        for par in b.get("paragraf_akhir", []):
            self.para(par)

    def bagian(self, sec, level=0):
        """Bagian berisi sub-bagian bersarang, daftar bernomor, atau blok bebas."""
        if sec.get("judul"):
            self.heading(sec["judul"], level)
        for par in sec.get("pengantar", []):
            self.para(par)
        for b in sec.get("isi", []):
            self.blok(b)
        if sec.get("daftar"):
            nid = new_num(self.doc, [("decimal", "%1."), ("lowerLetter", "%2."),
                                     ("lowerRoman", "%3."), ("decimal", "%4.")],
                          **self.P["daftar"])
            for it in sec["daftar"]:
                if isinstance(it, str):
                    self.item(it, nid)
                else:
                    self.blok(it, nid)
        for b in sec.get("penutup", []):
            self.blok(b)
        for sub in sec.get("sub", []):
            self.bagian(sub, level + 1)

    def build(self, out):
        s = self.spec
        self.judul()
        for i, sec in enumerate(s["bagian"]):
            if i:
                self.brk()
            self.bagian(sec, 0)
        self.doc.save(out)
        return out

    def judul(self):
        p = self.doc.add_paragraph(style=self.P["judul_style"])
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.5
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_ind(p, left=0, firstLine=0)
        baris = self.spec["judul"].split("\n")
        if self.P["judul_bab"]:
            set_num(p, 1, 0)
            baris = [""] + baris
        for i, line in enumerate(baris):
            r = p.add_run()
            r.font.name = TNR
            r.bold = True
            if i:
                r._element.append(OxmlElement("w:br"))
            if line:
                t = OxmlElement("w:t")
                t.text = line
                t.set(qn("xml:space"), "preserve")
                r._element.append(t)


def listify(x):
    if not x:
        return []
    return x if isinstance(x, list) else [x]


def gambar_list(entry):
    g = entry.get("gambar")
    if not g:
        return []
    if isinstance(g, str):
        return [{"file": g, "caption": entry.get("caption", "")}]
    if isinstance(g, dict):
        return [g]
    return g


def build(spec, out):
    return Builder(spec).build(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("-o", "--output")
    a = ap.parse_args()
    with open(a.spec, encoding="utf-8") as f:
        spec = json.load(f)
    spec.setdefault("base_dir", os.path.dirname(os.path.abspath(a.spec)))
    out = a.output or spec.get("output") or "Laporan.docx"
    print(build(spec, out))


if __name__ == "__main__":
    main()
