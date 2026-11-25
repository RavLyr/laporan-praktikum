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
import lint_laprak as L


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
                    {"teks": "Hasil ping dicatat.", "gambar": {"file": "belum-ada.png", "caption": "Hasil Ping"}},
                ],
            }]},
            {"judul": "Hasil dan Analisis", "sub": [{
                "judul": "Percobaan 1",
                "isi": [{"paragraf": [f"Seluruh tautan aktif dan port trunk menyala, seperti pada Gambar {bab}.1 dan Gambar {bab}.9."],
                         "persamaan": {"rumus": "Galat = |Vu - Vt| / Vt x 100%",
                                       "keterangan": ["dengan Vu nilai ukur dan Vt nilai teori."]}}],
            }]},
            {"judul": "Kesimpulan", "daftar": ["VLAN memisahkan domain siar."]},
            {"judul": "Daftar Pustaka", "pustaka": [
                "Zeta, A. (2020). *Jaringan Komputer*. Penerbit.",
                "Alpha, B. (2019). *Elektronika Dasar*. Penerbit.",
            ]},
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
    assert styles.count(P["heading"][0]) == 5, f"{profil}: heading tingkat 1 harus 5"
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
    assert brks == 4, f"{profil}: pemisah halaman harus 4, dapat {brks}"

    for p in d.paragraphs:
        if p.style.name in P["heading"]:
            assert not any(r.italic for r in p.runs), f"{profil}: heading tidak boleh miring"

    # caption: "Gambar 5." menempel langsung ke nomor SEQ, tanpa spasi
    caps = [p for p in d.paragraphs if p.text.startswith("Gambar")]
    assert caps and caps[0].runs[0].text == f"Gambar {bab}.", f"{profil}: awalan caption salah"

    # kode: spasi 1, before/after 0
    for p in tcode.cell(0, 0).paragraphs:
        pf = p.paragraph_format
        assert pf.line_spacing == 1.0 and pf.space_after == 0 and pf.space_before == 0, \
            f"{profil}: spasi kode harus 1 dengan before/after 0"

    # gambar yang belum ada jadi penanda, build tetap jalan
    assert any("belum-ada.png" in w for w in b.peringatan), f"{profil}: gambar hilang tidak diperingatkan"
    assert any("[ISI: gambar belum-ada.png]" in t.cell(0, 0).text for t in d.tables), \
        f"{profil}: penanda gambar hilang tidak ada"

    # persamaan bernomor di kanan
    eq = next(p for p in d.paragraphs if "Galat" in p.text)
    assert eq.text.endswith(f"({bab}.1)"), f"{profil}: nomor persamaan salah: {eq.text!r}"

    # daftar pustaka urut abjad, indent menggantung
    pus = [p for p in d.paragraphs if "Penerbit." in p.text]
    assert [p.text[:5] for p in pus] == ["Alpha", "Zeta,"], f"{profil}: pustaka tidak urut abjad"
    ind = pus[0]._p.pPr.find(qn("w:ind"))
    assert ind.get(qn("w:hanging")) == "709", f"{profil}: pustaka tanpa indent menggantung"

    # Word diminta memperbarui field saat dibuka
    assert d.settings.element.find(qn("w:updateFields")) is not None, f"{profil}: updateFields tidak ada"

    # lint menangkap penanda, rujukan salah, dan tabel yang tidak dirujuk
    temuan = L.lint(out, bab)
    assert any("penanda tersisa" in t for t in temuan), f"{profil}: lint tidak menangkap penanda"
    assert any(f"Gambar {bab}.9 tidak ada" in t for t in temuan), f"{profil}: lint tidak menangkap rujukan salah"
    assert any(f"Tabel {bab}.1 tidak dirujuk" in t for t in temuan), f"{profil}: lint tidak menangkap tabel tak dirujuk"
    assert not any(f"Gambar {bab}.1 tidak dirujuk" in t for t in temuan), f"{profil}: lint salah menandai gambar yang dirujuk"

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
