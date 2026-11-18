#!/usr/bin/env python3
"""Uji mandiri generator laporan untuk tiap profil. Jalankan: python selfcheck.py"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm
from PIL import Image

import build_laprak as B


def spec(profil, bab):
    return {
        "profil": profil,
        "bab": bab,
        "judul": "Judul Modul Uji",
        "bagian": [
            {"judul": "Tujuan", "daftar": ["Praktikan mampu menjelaskan konsep."]},
            {"judul": "Percobaan", "sub": [{
                "judul": "Percobaan 1",
                "isi": [{"tabel": {"caption": "Alamat Perangkat",
                                   "header": ["Devices", "IP"],
                                   "rows": [["Router", "192.168.1.1"]],
                                   "widths": [2.9, 3.0]}}],
                "daftar": [
                    {"teks": "Berkas topologi diunduh.", "gambar": {"file": "a.png", "caption": "Topologi"}},
                    {"teks": "Switch dikonfigurasi.", "kode": ["Switch>enable", "Switch#configure terminal"]},
                ],
            }]},
            {"judul": "Hasil dan Analisis", "sub": [{
                "judul": "Percobaan 1",
                "isi": [{"paragraf": ["Seluruh tautan aktif dan port trunk menyala."]}],
            }]},
            {"judul": "Kesimpulan", "daftar": ["VLAN memisahkan domain siar."]},
        ],
    }


def cek(profil, bab, tmp):
    s = dict(spec(profil, bab), base_dir=tmp)
    out = os.path.join(tmp, f"{profil}.docx")
    b = B.Builder(s)
    b.build(out)
    P = b.P
    d = Document(out)
    styles = [p.style.name for p in d.paragraphs]

    assert P["judul_style"] in styles, "judul hilang"
    assert styles.count(P["heading"][0]) == 4, f"{profil}: heading tingkat 1 harus 4"
    assert styles.count(P["heading"][1]) == 2, f"{profil}: heading tingkat 2 harus 2"

    numb = d.part.numbering_part.element
    texts = {l.find(qn("w:lvlText")).get(qn("w:val")) for l in numb.iter(qn("w:lvl"))
             if l.find(qn("w:lvlText")) is not None}
    assert f"{bab}.%1" in texts, f"{profil}: lvlText heading tidak ditambal"

    fields = [f.get(qn("w:instr")) for f in d.element.body.iter(qn("w:fldSimple"))]
    sfx = P["seq_suffix"]
    assert any(f"SEQ Gambar_{bab}.{sfx} " in f for f in fields), f"{profil}: SEQ gambar salah"
    assert any(f"SEQ Tabel_{bab}.{sfx} " in f for f in fields), f"{profil}: SEQ tabel salah"

    t0 = d.tables[0]
    shd = t0.rows[0].cells[0]._tc.tcPr.find(qn("w:shd"))
    assert shd.get(qn("w:fill")) == P["tabel"]["header_fill"], f"{profil}: warna header tabel salah"
    assert t0.rows[0].cells[0].paragraphs[0].runs[0].bold, f"{profil}: header tabel tidak tebal"

    tcode = d.tables[1]
    assert tcode.cell(0, 0).paragraphs[0].runs[0].font.name == P["kode"]["font"], f"{profil}: font kode salah"
    assert len(tcode.cell(0, 0).paragraphs) == 2, f"{profil}: satu paragraf per baris perintah"

    g = P["gambar"]
    n_img = 0
    for e in d.element.body.iter(qn("wp:extent")):
        n_img += 1
        cx, cy = int(e.get("cx")), int(e.get("cy"))
        assert cx <= int(Cm(g["max_w"])) + 1, f"{profil}: gambar terlalu lebar"
        assert cy <= int(Cm(g["max_h"])) + 1, f"{profil}: gambar terlalu tinggi"
        assert abs(cx / cy - 1.8) < 0.02, f"{profil}: rasio gambar berubah"
    assert n_img == 1, f"{profil}: jumlah gambar salah"

    p = next(p for p in d.paragraphs if "dikonfigurasi" in p.text)
    assert any(r.italic and r.text.strip().lower() == "switch" for r in p.runs), \
        f"{profil}: istilah asing tidak dimiringkan"

    brks = sum(1 for p in d.paragraphs if p.paragraph_format.page_break_before)
    assert brks == 3, f"{profil}: pemisah halaman harus 3, dapat {brks}"

    for p in d.paragraphs:
        if p.style.name in P["heading"]:
            assert not any(r.italic for r in p.runs), f"{profil}: heading tidak boleh miring"

    print(f"  {profil} OK")


def main():
    profiles = B.load_profiles()
    with tempfile.TemporaryDirectory() as tmp:
        Image.new("RGB", (1800, 1000), (200, 220, 240)).save(os.path.join(tmp, "a.png"))
        for i, profil in enumerate(profiles):
            cek(profil, 5 + i, tmp)
    print("selfcheck OK")


if __name__ == "__main__":
    main()
