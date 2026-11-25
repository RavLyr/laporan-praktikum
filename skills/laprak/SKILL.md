---
name: laprak
description: "Membuat laporan praktikum (.docx) yang persis mengikuti format resmi kampus: heading bernomor otomatis, caption field SEQ, tabel data berwarna, tabel kode, gambar beserta caption. Untuk praktikum Teknik Komputer: profil jaringan komputer dan profil elektronika dasar/sistem digital. Dipakai saat pengguna memberi modul praktikum plus folder screenshot dan minta laporan jadi. Trigger: laprak, laporan praktikum, teknik komputer, jaringan komputer, elektronika dasar, sistem digital, bikin laporan modul."
---

# Laporan Praktikum

Hasilkan `.docx` yang identik struktur dengan template resmi. Yang berganti antar modul
hanya **nomor BAB, judul, dan isi**. Format tidak diubah.

## Profil

| Profil | Template asal | Ciri |
|---|---|---|
| `srwe` | Format Laporan Prak SRWE Modul 4 | judul berawalan "BAB V", caption gambar `captiongambar`, header tabel `FF89FF`, gambar maks 11,67 cm |
| `eldas` | ELDAS2026_MODUL4_Kel26 | judul tanpa "BAB", seluruh caption style `Caption`, header tabel `E8EEF5`, gambar maks 14 cm, heading sampai empat tingkat |

Pilih lewat `"profil"` di spec. Daftar lengkap kunci ada di `assets/profil.json`;
timpa sebagian lewat `"profil_override"` kalau dosen minta variasi kecil.

Profil baru: salin `.docx` acuan ke `assets/`, tambahkan entri di `assets/profil.json`,
jalankan `scripts/selfcheck.py`.

## Alur

Path `scripts/`, `assets/`, `references/` di bawah relatif terhadap folder skill ini, bukan folder kerja pengguna.

1. Baca modul (markdown/PDF/docx) dan daftar gambar yang diberikan pengguna.
2. Tentukan profil. Kalau ragu, tanya. Jangan menebak.
3. Petakan gambar ke langkah percobaan. Urutan berkas biasanya urutan langkah;
   konfirmasi kalau tidak jelas.
4. Tulis `spec.json` sesuai `references/spec-schema.md`.
5. Tulis prosa Indonesia sesuai `references/bahasa-id.md`. **Wajib dibaca**, bagian
   ini yang paling sering gagal.
6. Bangun:
   ```bash
   python scripts/build_laprak.py spec.json -o "Laporan-Modul-N.docx"
   ```
   Butuh `python-docx` dan `pillow`. Setelah build, `scripts/lint_laprak.py` berjalan otomatis
   dan mencetak `PERINGATAN` untuk penanda `[ISI: ...]` yang tersisa, rujukan Gambar atau
   Tabel yang nomornya tidak ada, gambar atau tabel yang tidak dirujuk, dan pola gaya.
   Tangani tiap temuan satu per satu, jangan diganti otomatis. Gambar yang belum ada diganti
   kotak `[ISI: gambar ...]`; tambahkan `--ketat` bila ingin build berhenti.
7. Finalisasi dan verifikasi visual. Pilih sesuai mesin:
   - **Windows dengan Microsoft Word** (paling baik): perbarui semua field, simpan, dan ekspor PDF.
     Pengguna tidak perlu lagi menekan Ctrl+A lalu F9.
     ```bash
     powershell -NoProfile -ExecutionPolicy Bypass -File scripts/word_final.ps1 -Docx Laporan-Modul-N.docx -Pdf cek.pdf
     ```
     Dari WSL, salin dulu ke folder Windows dan panggil `powershell.exe` dengan path Windows.
   - **Tanpa Word** (Linux, sandbox): `soffice --headless --convert-to pdf Laporan-Modul-N.docx`
     lalu `pdftoppm -png -r 70 Laporan-Modul-N.pdf cek/p`. Nomor caption di render ini bisa
     masih "1"; Word menawarkan pembaruan field saat dokumen dibuka (`updateFields` sudah dipasang).
   Baca PDF atau PNG hasilnya dan periksa halaman berisi tabel, kode, dan gambar.

## Aturan yang tidak boleh dilanggar

- Nomor gambar dan tabel memakai field `SEQ`, jangan diketik manual. Nomor baru muncul
  setelah field diperbarui di Word (`word_final.ps1`, atau Word menawarkannya saat dibuka).
- Data yang belum ada ditulis `[ISI: ...]`, sumber yang belum pasti `[CEK SUMBER: ...]`.
  Jangan mengarang angka hasil percobaan.
- Nomor BAB menyetir seluruh penomoran. Cukup set `"bab": N`; skrip menambal `numbering.xml`.
- Ukuran gambar dibatasi profil, rasio asli dijaga, rata tengah, caption tepat di bawahnya.
- Caption tabel diletakkan **di atas** tabel.
- Heading tidak pernah miring walaupun memuat istilah asing.
- Istilah asing dimiringkan di prosa. `*teks*` memaksa miring.
- Header dan footer dibiarkan kosong, sama seperti template.
- Jangan menulis `.docx` dari nol. Selalu lewat template di `assets/`, karena di situlah
  styles, numbering, `sectPr`, dan font tertanam berada.

## Berkas

- `assets/srwe.docx`, `assets/eldas.docx` — template kosong hasil pelucutan dokumen resmi. Jangan ditimpa.
- `assets/profil.json` — parameter format tiap profil.
- `scripts/build_laprak.py` — generator.
- `scripts/lint_laprak.py` — pemeriksa rujukan, penanda, dan gaya; juga bisa dipakai pada .docx lain: `python scripts/lint_laprak.py Laporan.docx`.
- `scripts/word_final.ps1` — perbarui field, simpan, dan ekspor PDF lewat Word (Windows).
- `scripts/selfcheck.py` — uji mandiri seluruh profil.
- `references/spec-schema.md` — bentuk spec.
- `references/struktur-docx.md` — angka format hasil pembedahan kedua template.
- `references/bahasa-id.md` — kaidah menulis agar tidak terbaca sebagai keluaran AI.

## Uji mandiri

```bash
python scripts/selfcheck.py
```
Membangun dokumen contoh tiap profil lalu memeriksa style, penomoran, warna header,
ukuran gambar, field SEQ, pemisah halaman, kemiringan heading, spasi kode, penanda gambar
kosong, nomor persamaan, urutan pustaka, dan hasil lint.
