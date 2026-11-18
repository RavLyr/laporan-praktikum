---
name: laprak
description: "Membuat laporan praktikum (.docx) yang persis mengikuti format resmi kampus: heading bernomor otomatis, caption field SEQ, tabel data berwarna, tabel kode, gambar beserta caption. Mendukung profil SRWE/jaringan dan Elektronika Dasar/Sistem Digital. Dipakai saat pengguna memberi modul praktikum plus folder screenshot dan minta laporan jadi. Trigger: laprak, laporan praktikum, SRWE, ELDAS, sistem digital, bikin laporan modul."
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
   Butuh `python-docx` dan `pillow`.
7. Verifikasi: ekspor PDF lewat Word COM, render beberapa halaman, periksa visual.
   ```bash
   cp out.docx /mnt/c/Users/<USER>/cek.docx
   powershell.exe -NoProfile -Command "\$w=New-Object -ComObject Word.Application; \$w.Visible=\$false; \$w.DisplayAlerts=0; \$d=\$w.Documents.Open('C:\Users\<USER>\cek.docx',[ref]\$false,[ref]\$true); \$d.Fields.Update(); \$d.ExportAsFixedFormat('C:\Users\<USER>\cek.pdf',17); \$d.Close(0); \$w.Quit()"
   pdftoppm -png -r 70 /mnt/c/Users/<USER>/cek.pdf /tmp/opencode/cek/p
   ```
   Nama berkas Windows harus baru tiap ekspor, kalau tidak Word memakai cache.

## Aturan yang tidak boleh dilanggar

- Nomor gambar dan tabel memakai field `SEQ`, jangan diketik manual. Nomor baru muncul
  setelah `Fields.Update()` di Word.
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
- `scripts/selfcheck.py` — uji mandiri seluruh profil.
- `references/spec-schema.md` — bentuk spec.
- `references/struktur-docx.md` — angka format hasil pembedahan kedua template.
- `references/bahasa-id.md` — kaidah menulis agar tidak terbaca sebagai keluaran AI.

## Uji mandiri

```bash
python scripts/selfcheck.py
```
Membangun dokumen contoh tiap profil lalu memeriksa style, penomoran, warna header,
ukuran gambar, field SEQ, pemisah halaman, dan kemiringan heading.
