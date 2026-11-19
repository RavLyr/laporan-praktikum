# Jalur B: bangun laporan dari spec

Dipakai bila tidak ada dokumen terbuka dan pengguna memberi modul plus folder screenshot.
Generator ada di repo `RavLyr/laporan-praktikum` (folder `skills/laprak`). Format dikunci oleh
kode dan template `.docx`, jadi konsisten antar modul. Prinsip utama, analisis, sitasi, dan
gaya bahasa di SKILL.md tetap berlaku untuk prosa di spec.

## 1. Siapkan repo

Cari salinan lokal dulu, baru clone:

| Lingkungan | Lokasi |
|---|---|
| Skill `laprak` terpasang berdampingan (`npx skills add RavLyr/laporan-praktikum` memasang keduanya) | folder `../laprak` relatif terhadap skill ini; path `skills/laprak/...` di bawah dibaca relatif terhadap folder itu |
| Agent di laptop pengguna | folder repo yang sudah ada, atau skill `laprak` di folder skill agent (misalnya `~/.claude/skills/`) |
| Sandbox tanpa salinan lokal | `git clone --depth 1 https://github.com/RavLyr/laporan-praktikum` ke folder kerja sendiri, bukan folder pengguna |

Lalu `pip install python-docx pillow` (tambah `--break-system-packages` di sandbox Linux bila pip menolak).

Cek kemampuan versi repo yang dipakai, sekali:

- Ada `scripts/lint_laprak.py`: lint berjalan otomatis setelah build, gambar yang belum ada jadi
  kotak `[ISI: ...]`, dan blok `persamaan` serta `pustaka` tersedia.
- Tidak ada: versi lama. Gambar yang belum ada membuat build berhenti, jadi jangan masukkan
  gambar yang berkasnya belum ada; tulis penandanya sebagai paragraf. Persamaan dan pustaka
  ditulis sebagai `paragraf` biasa. Jalankan lint manual di langkah 6.

Bila repo tidak bisa diambil atau strukturnya berubah sehingga langkah di bawah gagal, jangan
mengarang skrip dan jangan menambal diam-diam. Laporkan satu kalimat, lalu pakai Jalur A dengan
skill `docx`.

## 2. Baca acuan repo

`skills/laprak/references/spec-schema.md` wajib dibaca sebelum menulis spec.
`bahasa-id.md` boleh dibaca sebagai pelengkap; bila berbeda, Bagian 5 SKILL.md yang dipakai.

## 3. Profil

Tentukan `eldas` atau `srwe` dari ciri di `assets/profil.json`. Bila ragu, tanya satu kali.
Profil baru: kerjakan dengan Jalur A, lalu tawarkan menambah profil di repo.

## 4. Tulis spec.json

- Petakan screenshot ke langkah percobaan. Urutan berkas biasanya urutan langkah; konfirmasi
  bila tidak jelas.
- Data kosong ditulis `[ISI: ...]`, sumber belum pasti `[CEK SUMBER: ...]`, bukan dikarang.
- Hitungan (galat, rata-rata, gain) dihitung dengan kode dulu, hasilnya baru masuk spec.
- Rumus memakai blok `persamaan` (nomor dihitung skrip: 5.1, 5.2, ...). Rujuk sebagai
  "Persamaan (5.1)".
- Daftar Pustaka memakai bagian dengan kunci `pustaka`; skrip mengurutkan abjad dan memberi
  indent menggantung. Judul buku ditulis `*miring*`.
- Rujukan "Gambar 5.n" dan "Tabel 5.n" ditulis manual sesuai urutan kemunculan caption.

## 5. Bangun

```bash
python skills/laprak/scripts/build_laprak.py spec.json -o Laporan-Modul-N.docx
```

## 6. Tangani PERINGATAN

Skrip mencetak `PERINGATAN` dari lint: penanda tersisa, rujukan yang nomornya tidak ada,
gambar atau tabel yang tidak dirujuk, dan pola gaya. Tangani satu per satu: betulkan spec lalu
bangun ulang. Penanda `[ISI: ...]` yang memang menunggu data pengguna dibiarkan dan masuk
daftar cek di penutup. Pola gaya ditimbang, tidak diganti otomatis.

Versi repo tanpa lint: jalankan pemeriksaan yang sama secara manual (Lint rujukan di langkah 8
SKILL.md) pada teks spec.

## 7. Finalisasi dan render

| Mesin | Cara |
|---|---|
| Windows dengan Microsoft Word | `powershell -NoProfile -ExecutionPolicy Bypass -File skills/laprak/scripts/word_final.ps1 -Docx Laporan-Modul-N.docx -Pdf cek.pdf`. Field diperbarui dan disimpan, jadi nomor caption sudah final. |
| Tanpa Word (sandbox Linux) | `soffice --headless --convert-to pdf` lalu `pdftoppm -png -r 70`. Nomor caption di render bisa masih "1"; Word menawarkan pembaruan field saat dokumen dibuka. |

Cek `soffice` atau Word dulu sebelum memakai; jangan menganggap keduanya ada. Baca hasil render
dan periksa halaman berisi tabel, kode, persamaan, dan gambar. Kirim `.docx` ke pengguna dengan
SendUserFile bila tool itu ada, atau sebutkan path lengkapnya.

## Catatan format

Caption di Jalur B memakai format repo: teks "Gambar 5." lalu field `SEQ Gambar_5.`, bukan
`STYLEREF 1 \s` seperti Bagian 9.1. Hasil cetaknya sama ("Gambar 5.1"). Jangan menandai ini
sebagai pelanggaran Bagian 9.1.
