# laporan-praktikum

Skill [OpenCode](https://opencode.ai) / agent coding untuk membuat **laporan praktikum `.docx`** yang mengikuti format resmi kampus. Kamu beri modul praktikum dan folder screenshot, agent menulis spec JSON dan prosa Indonesia, lalu skrip membangun `.docx` dari template resmi.

## Fitur

- Heading bernomor otomatis (BAB V, 5.1, 5.1.1) dari satu nilai `"bab"`.
- Caption gambar dan tabel memakai field `SEQ` Word, bukan ketikan manual.
- Tabel data berwarna, tabel kode (Courier New), gambar rata tengah dengan rasio asli dan caption tepat di bawahnya.
- Istilah asing dimiringkan otomatis di prosa, tidak di heading.
- Dibangun dari template di `assets/`, jadi styles, numbering, dan font identik dengan dokumen resmi.
- Dua profil bawaan:

| Profil | Untuk | Ciri |
|---|---|---|
| `srwe` | SRWE / jaringan | judul berawalan "BAB V", header tabel `FF89FF`, gambar maks 11,67 cm |
| `eldas` | Elektronika Dasar / Sistem Digital | judul tanpa "BAB", header tabel `E8EEF5`, gambar maks 14 cm, heading sampai 4 tingkat |

## Struktur

```
skills/laprak/
├── SKILL.md                  # instruksi untuk agent (alur, aturan wajib)
├── assets/
│   ├── srwe.docx             # template kosong profil SRWE
│   ├── eldas.docx            # template kosong profil ELDAS
│   └── profil.json           # parameter format tiap profil
├── scripts/
│   ├── build_laprak.py       # generator spec.json -> .docx
│   └── selfcheck.py          # uji mandiri semua profil
└── references/
    ├── spec-schema.md        # bentuk spec.json
    ├── struktur-docx.md      # angka format hasil pembedahan template
    └── bahasa-id.md          # kaidah menulis prosa Indonesia
```

## Instalasi

Skill dibaca agent dari folder skills. Salin folder `skills/laprak` ke lokasi yang sesuai.

**OpenCode (global):**

```bash
git clone https://github.com/RavLyr/laporan-praktikum.git
mkdir -p ~/.config/opencode/skills
cp -r laporan-praktikum/skills/laprak ~/.config/opencode/skills/
```

**OpenCode (per proyek):** salin ke `.opencode/skills/laprak` di repo proyekmu.

**Claude Code:** salin ke `~/.claude/skills/laprak`.

Restart agent. Skill `laprak` akan muncul di daftar skill.

### Dependensi

Python 3.9+ dengan `python-docx` dan `pillow`:

```bash
pip install python-docx pillow
```

Cek instalasi:

```bash
python skills/laprak/scripts/selfcheck.py
# srwe OK
# eldas OK
# selfcheck OK
```

## Cara pakai

### Lewat agent (disarankan)

1. Siapkan modul praktikum (markdown, PDF, atau docx) dan folder screenshot, urut sesuai langkah.
2. Minta agent, misalnya: `bikin laprak modul 4 profil eldas, modul di modul4.pdf, screenshot di ./shots`.
3. Agent memilih profil (tanya kalau ragu), memetakan gambar ke langkah, menulis `spec.json` dan prosa, lalu membangun `.docx`.
4. Buka hasil di Word, tekan **Ctrl+A lalu F9** (update field) supaya nomor gambar dan tabel muncul.

### Manual

```bash
python skills/laprak/scripts/build_laprak.py spec.json -o "Laporan-Modul-4.docx"
```

Contoh `spec.json` minimal:

```json
{
  "profil": "eldas",
  "bab": 4,
  "judul": "Gerbang Logika Dasar",
  "base_dir": ".",
  "bagian": [
    {"judul": "Tujuan", "daftar": ["Praktikan mampu menjelaskan cara kerja gerbang AND."]},
    {"judul": "Hasil Percobaan dan Analisis", "sub": [
      {"judul": "Percobaan 1", "isi": [
        {"paragraf": ["Rangkaian diuji dengan empat kombinasi input."],
         "tabel": {"caption": "Tabel Kebenaran", "header": ["A", "B", "Y"], "rows": [["0", "0", "0"], ["1", "1", "1"]]},
         "gambar": [{"file": "shots/01.png", "caption": "Rangkaian AND"}]}
      ]}
    ]},
    {"judul": "Kesimpulan", "daftar": ["Keluaran AND bernilai 1 hanya jika kedua input 1."]}
  ]
}
```

Skema lengkap: [`references/spec-schema.md`](skills/laprak/references/spec-schema.md).

## Aturan penting

- Jangan menulis `.docx` dari nol. Selalu lewat template di `assets/`.
- Jangan menimpa file di `assets/`.
- Nomor gambar dan tabel selalu field `SEQ`.
- Caption tabel di atas tabel, caption gambar di bawah gambar.
- Baca `references/bahasa-id.md` sebelum menulis prosa supaya tidak terbaca seperti keluaran AI.

## Menambah profil baru

1. Salin `.docx` acuan (dokumen resmi yang sudah dikosongkan) ke `skills/laprak/assets/`.
2. Tambah entri di `skills/laprak/assets/profil.json`. Pakai profil lain sebagai contoh.
3. Jalankan `python skills/laprak/scripts/selfcheck.py`.
4. Hapus nama dan data pribadi dari properti dokumen sebelum commit (File > Info > Inspect Document).

## Verifikasi visual (opsional, Windows/WSL)

`SKILL.md` menjelaskan ekspor PDF lewat Word COM dari WSL. Ganti `<USER>` dengan nama user Windows kamu.

## Catatan

- Template dikosongkan dari isi dan metadata penulis. Format kampus bisa berubah per semester, jadi cocokkan dengan panduan terbaru dosen atau asisten.
- Hasil tetap perlu dibaca ulang. Skill menyusun format dan draf prosa, bukan menggantikan data percobaanmu.
