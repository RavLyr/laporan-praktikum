#!/usr/bin/env python3
"""Periksa laporan praktikum (.docx): penanda tersisa, rujukan Gambar/Tabel, dan pola gaya.

Pakai: python lint_laprak.py Laporan.docx [--bab N]
Dipanggil otomatis oleh build_laprak.py setelah dokumen dibangun.
Hasilnya daftar temuan untuk ditimbang satu per satu, bukan untuk diganti otomatis.
"""
import argparse
import re
import sys

from docx import Document

PENANDA = re.compile(r"\[(ISI|CEK SUMBER)\s*:[^\]]*\]")
RUJUKAN = re.compile(r"\b(Gambar|Tabel)\s+(\d+)\.(\d+)\b")
CAPTION = re.compile(r"^(Gambar|Tabel)\s+(\d+)\.")
GAYA = [
    (re.compile("[—–]"), "em dash atau en dash"),
    (re.compile(r"\b(di mana|yang mana)\b", re.I), "'di mana' atau 'yang mana' sebagai penghubung"),
    (re.compile(r"\b(melakukan|dilakukan)\b", re.I), "nominalisasi 'melakukan/dilakukan'"),
    (re.compile(r"\b(kami|penulis)\b", re.I), "kata ganti orang pertama"),
    (re.compile(r"\bmembuktikan\b", re.I), "klaim 'membuktikan'"),
]


def potong(teks, m, lebar=40):
    a, b = max(0, m.start() - lebar), min(len(teks), m.end() + lebar)
    return ("..." if a else "") + teks[a:b].strip() + ("..." if b < len(teks) else "")


def kumpulkan(doc):
    """Kembalikan (prosa, caption) dari paragraf badan dan sel tabel."""
    prosa, caption = [], []

    def lihat(p):
        teks = p.text.strip()
        if not teks:
            return
        if "caption" in p.style.name.lower() and CAPTION.match(teks):
            caption.append(teks)
        else:
            prosa.append(teks)

    for p in doc.paragraphs:
        lihat(p)
    for t in doc.tables:
        for row in t.rows:
            seen = set()
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                for p in cell.paragraphs:
                    lihat(p)
    return prosa, caption


def lint(path, bab=None):
    doc = Document(path)
    prosa, caption = kumpulkan(doc)
    temuan = []

    jumlah = {"Gambar": 0, "Tabel": 0}
    for c in caption:
        m = CAPTION.match(c)
        jumlah[m.group(1)] += 1
        if bab is None:
            bab = int(m.group(2))

    for teks in prosa:
        for m in PENANDA.finditer(teks):
            temuan.append(f"penanda tersisa: {m.group(0)}")

    dirujuk = {"Gambar": set(), "Tabel": set()}
    for teks in prosa:
        for m in RUJUKAN.finditer(teks):
            jenis, b, n = m.group(1), int(m.group(2)), int(m.group(3))
            if bab is not None and b != bab:
                temuan.append(f"rujukan {m.group(0)} memakai BAB {b}, dokumen BAB {bab}: {potong(teks, m)}")
            elif n < 1 or n > jumlah[jenis]:
                temuan.append(f"rujukan {m.group(0)} tidak ada (hanya {jumlah[jenis]} {jenis.lower()}): {potong(teks, m)}")
            else:
                dirujuk[jenis].add(n)

    for jenis in ("Gambar", "Tabel"):
        for n in range(1, jumlah[jenis] + 1):
            if n not in dirujuk[jenis]:
                temuan.append(f"{jenis} {bab}.{n} tidak dirujuk di prosa")

    for teks in prosa:
        for pola, nama in GAYA:
            for m in pola.finditer(teks):
                temuan.append(f"gaya ({nama}): {potong(teks, m)}")

    return temuan


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--bab", type=int)
    a = ap.parse_args()
    temuan = lint(a.docx, a.bab)
    for t in temuan:
        print("PERINGATAN: " + t)
    print(f"lint selesai: {len(temuan)} temuan")
    sys.exit(1 if temuan else 0)


if __name__ == "__main__":
    main()
