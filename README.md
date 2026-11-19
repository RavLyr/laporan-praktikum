# laporan-praktikum

**Laprak Teknik Komputer.** Skill untuk agent coding yang membuat **laporan praktikum `.docx`** praktikum Teknik Komputer (jaringan, elektronika, sistem digital) sesuai format resmi kampus. Kamu beri modul praktikum dan folder screenshot. Agent menyusun `spec.json` dan prosa Indonesia, lalu skrip membangun `.docx` dari template resmi, lengkap dengan heading bernomor, caption `SEQ`, tabel berwarna, dan gambar.

Format skill-nya mengikuti standar [Agent Skills](https://agentskills.io), jadi satu paket yang sama jalan di **Claude Code, Codex, OpenCode, Antigravity**, dan puluhan agent lain.

| Agent | Kompatibilitas | Pasang | Panggil |
|---|---|---|---|
| Claude Code | sesuai dokumentasi | `~/.claude/skills/` | `/laprak` |
| Codex (CLI, IDE, app) | sesuai dokumentasi | `~/.agents/skills/` | `$laprak` |
| OpenCode | sesuai dokumentasi | `~/.config/opencode/skills/` | minta langsung atau `@laprak` |
| Antigravity (IDE, CLI) | sesuai dokumentasi | `~/.gemini/config/skills/` | minta langsung |
| Cursor, Cline, Copilot, Kiro, dll | lewat `npx skills` | otomatis | sesuai agent |

Folder dan cara panggil diambil dari dokumentasi masing-masing agent. Skill ini baru diuji lewat `npx skills` (terdeteksi) dan di OpenCode. Di agent lain mengikuti dokumentasinya, jadi kabari lewat issue bila ada yang tidak jalan. Detail dan catatan versi ada di [Pemasangan](#pemasangan).

## Mulai cepat

```bash
# 1. pasang (pilih agent saat ditanya, tambah -g untuk semua proyek)
npx skills add RavLyr/laporan-praktikum -g

# 2. pasang dependensi Python
pip install python-docx pillow

# 3. di agent-mu
#    "bikin laprak profil eldas, bab 4. modul: ./modul4.pdf, screenshot: ./shots"
```

Buka `.docx` hasilnya di Word dan setujui tawaran memperbarui field (atau tekan **Ctrl+A lalu F9**) supaya nomor gambar dan tabel muncul, lalu baca ulang. Di Windows, agent bisa melakukannya otomatis lewat `scripts/word_final.ps1`. Panduan lengkap: [docs/CARA-PAKAI.md](docs/CARA-PAKAI.md).

## Daftar isi

- [Fitur](#fitur)
- [Cara kerja](#cara-kerja)
- [Prasyarat](#prasyarat)
- [Pemasangan](#pemasangan)
- [Pemakaian](#pemakaian)
- [Profil format](#profil-format)
- [Menyesuaikan](#menyesuaikan)
- [Struktur repo](#struktur-repo)
- [Pertanyaan umum](#pertanyaan-umum)
- [Batasan](#batasan)

## Fitur

- Heading bernomor otomatis (BAB V, 5.1, 5.1.1) dari satu nilai `"bab"`.
- Caption gambar dan tabel memakai field `SEQ` Word, bukan ketikan manual.
- Tabel data berwarna, tabel kode (Courier New), gambar rata tengah dengan rasio asli dan caption tepat di bawahnya.
- Istilah asing dimiringkan otomatis di prosa, tidak di heading.
- Selalu dibangun dari template di `assets/`, jadi style, penomoran, dan font identik dengan dokumen resmi.
- Pedoman menulis prosa Indonesia yang tidak terbaca kaku atau seperti keluaran AI ([bahasa-id.md](skills/laprak/references/bahasa-id.md)).
- Uji mandiri untuk semua profil: `python skills/laprak/scripts/selfcheck.py`.

## Cara kerja

```
modul praktikum + screenshot + data hasil
              │  kamu kasih ke agent
              ▼
   agent menulis spec.json dan prosa
              │  scripts/build_laprak.py + template di assets/
              ▼
        Laporan.docx
              │  kamu buka di Word, Ctrl+A, F9, baca ulang
              ▼
      laporan siap dikumpulkan
```

Agent hanya mengisi **isi** (`spec.json`). **Format** datang dari template, jadi hasilnya konsisten antar modul. Yang berganti per modul hanya nomor BAB, judul, dan isi.

Agent memuat `SKILL.md` saat permintaanmu cocok dengan `description`-nya (misalnya "laprak", "laporan praktikum", "teknik komputer", "jaringan", "elektronika", "sistem digital"), atau saat kamu memanggilnya langsung.

## Prasyarat

| Kebutuhan | Keterangan |
|---|---|
| Agent coding | salah satu dari tabel di atas, dengan akses shell untuk menjalankan skrip |
| Python 3.9+ | `python3` di macOS dan Linux, `python` di Windows |
| `python-docx`, `pillow` | `pip install python-docx pillow` |
| Node.js 22.20+ | hanya bila memakai `npx skills` |
| Microsoft Word | untuk memperbarui field `SEQ` dan memeriksa hasil. LibreOffice belum diuji |

Agent menjalankan skrip di mesin yang sama dengan agent-nya. Kalau Python-mu ada di virtualenv, aktifkan sebelum menjalankan agent.

## Pemasangan

Semua agent memakai folder `laprak` yang sama. Beda hanya lokasi tujuan.

### Cara cepat: `npx skills`

```bash
npx skills add RavLyr/laporan-praktikum              # proyek ini saja
npx skills add RavLyr/laporan-praktikum -g           # semua proyek
npx skills add RavLyr/laporan-praktikum -a claude-code -a codex -g   # agent tertentu
```

Perintah ini ([vercel-labs/skills](https://github.com/vercel-labs/skills)) mendeteksi agent yang terpasang dan menaruh skill di folder yang benar. Alat bantu lain: `npx skills list`, `npx skills update`, `npx skills remove laprak`.

### Manual

```bash
git clone https://github.com/RavLyr/laporan-praktikum.git
cp -r laporan-praktikum/skills/laprak <FOLDER-TUJUAN>/
```

| Agent | Global (semua proyek) | Per proyek |
|---|---|---|
| Claude Code | `~/.claude/skills/` | `.claude/skills/` |
| Codex | `~/.agents/skills/` | `.agents/skills/` |
| OpenCode | `~/.config/opencode/skills/` | `.opencode/skills/` |
| Antigravity | `~/.gemini/config/skills/` | `.agents/skills/` |

Catatan per agent:

- **Claude Code.** Skill juga bisa dipanggil lewat `/laprak`. Kalau tidak muncul, jalankan `/reload-skills` atau mulai ulang sesi.
- **Codex.** Mendeteksi perubahan folder otomatis, mulai ulang bila belum muncul. Ketik `/skills` untuk melihat daftar, lalu `$laprak` untuk memanggil. Folder per proyek dicari dari direktori kerja sampai akar repo.
- **OpenCode.** Selain `~/.config/opencode/skills/`, OpenCode juga membaca `~/.claude/skills/`, `~/.agents/skills/`, `.claude/skills/`, dan `.agents/skills/`. Jadi pemasangan untuk Claude Code atau Codex otomatis terbaca juga di OpenCode. Jangan menaruh skill yang sama di dua tempat, karena ID yang bentrok ditimpa sumber yang didaftarkan belakangan.
- **Antigravity.** Folder global di atas mengikuti `npx skills`. Codelab resmi Google menyebut Antigravity CLI bisa memakai `~/.gemini/antigravity-cli/skills/`. Kalau skill tidak terdeteksi di CLI, salin ke folder itu.
- **Satu pemasangan untuk beberapa agent.** `~/.agents/skills/` dibaca Codex dan OpenCode (serta banyak agent lain di daftar `npx skills`). Kalau memakai keduanya, cukup satu salinan di sana. Claude Code memakai folder sendiri.
- **Agent lain.** Pakai `npx skills add`. Daftar agent dan foldernya ada di [docs agents](https://github.com/vercel-labs/skills/blob/main/docs/agents.md).

### Memeriksa pemasangan

Mulai ulang agent bila perlu, lalu:

| Agent | Cara cek |
|---|---|
| Claude Code, Codex | ketik `/skills` |
| OpenCode, Antigravity | tanya "skill apa saja yang tersedia?" |
| Semua | `python <folder-skill>/laprak/scripts/selfcheck.py` harus mencetak `selfcheck OK` |

## Pemakaian

### 1. Siapkan bahan

- Modul praktikum (markdown, PDF, atau docx).
- Folder screenshot, nama berkas berurutan sesuai langkah.
- Data hasil percobaan (angka, output, tabel). Agent tidak bisa menebak data praktikummu.
- Pilihan [profil](#profil-format): `srwe` atau `eldas`.

### 2. Minta ke agent

```
bikin laprak profil eldas, bab 4.
modul: ./modul4.pdf
screenshot: ./shots (urut sesuai langkah)
data hasil: ./data-hasil.md
output: Laporan-Modul-4.docx
```

Panggil lewat `/laprak` (Claude Code), `$laprak` (Codex), `@laprak` (OpenCode), atau minta langsung di agent mana pun. Kalau agent tidak menemukan skrip, sebutkan lokasi skill-nya. Path `scripts/`, `assets/`, dan `references/` relatif terhadap folder `laprak`, bukan folder kerjamu.

### 3. Periksa hasil

1. Buka `.docx` di Word, **Ctrl+A lalu F9**.
2. Cocokkan rujukan silang ("seperti pada Gambar 5.2") dengan nomor sebenarnya.
3. Periksa data, gambar, dan caption, lalu baca analisisnya.
4. Perlu perbaikan? Ubah `spec.json` atau minta agent memperbaikinya, lalu build ulang. Jangan menambal `.docx` berulang kali.

### Tanpa agent

```bash
python skills/laprak/scripts/build_laprak.py spec.json -o "Laporan-Modul-4.docx"
```

Contoh siap jalan: [`docs/contoh/spec.json`](docs/contoh/spec.json).

```bash
python skills/laprak/scripts/build_laprak.py docs/contoh/spec.json -o contoh.docx
```

Contoh `spec.json` minimal:

```json
{
  "profil": "eldas",
  "bab": 4,
  "judul": "Gerbang Logika Dasar",
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

Skema lengkap: [spec-schema.md](skills/laprak/references/spec-schema.md).

## Profil format

| Profil | Untuk | Ciri |
|---|---|---|
| `srwe` | praktikum jaringan komputer (router, switch, wireless) | judul berawalan "BAB V" dua baris, caption gambar `captiongambar`, header tabel `FF89FF`, gambar maks 11,67 cm |
| `eldas` | praktikum elektronika dasar dan sistem digital (rangkaian, gerbang logika) | judul tanpa "BAB", seluruh caption style `Caption`, header tabel `E8EEF5`, gambar maks 14 cm, heading sampai 4 tingkat |

Ragu memilih? Lihat contoh laporan dari asisten atau dosen. Ada kata "BAB" di judul dan tabel pink berarti `srwe`.

## Menyesuaikan

| Kebutuhan | Cara | Rujukan |
|---|---|---|
| Dosen minta variasi kecil (font kode, lebar gambar, warna header) | `profil_override` di `spec.json` | [CARA-PAKAI, bagian 6](docs/CARA-PAKAI.md) |
| Istilah asing salah dimiringkan | `istilah` atau `istilah_kecuali` | [CARA-PAKAI, bagian 6](docs/CARA-PAKAI.md) |
| Format kampus lain | tambah profil baru | [Menambah profil](#menambah-profil-baru) |
| Gaya tulisan beda | edit atau tambahkan aturan di prompt | [bahasa-id.md](skills/laprak/references/bahasa-id.md) |

`profil_override` menimpa dangkal: menulis `"tabel": {"header_fill": "CCCCCC"}` saja menghapus kunci `tabel` lain dan build gagal dengan `KeyError: 'style'`. Selalu tulis sub-objeknya lengkap. Contoh lengkap ada di panduan.

### Menambah profil baru

1. Kosongkan dokumen resmi (sisakan style, margin, penomoran) dan simpan sebagai `skills/laprak/assets/<nama>.docx`.
2. Hapus nama dan data pribadi dari properti dokumen (Word: File > Info > Periksa Dokumen).
3. Tambah entri `<nama>` di `skills/laprak/assets/profil.json`. Salin `srwe` atau `eldas` sebagai titik awal.
4. Jalankan `python skills/laprak/scripts/selfcheck.py`, lalu periksa visual di Word.

## Struktur repo

```
README.md
docs/
├── CARA-PAKAI.md             # panduan lengkap: bahan, prompt, cek, penyesuaian
└── contoh/                   # spec.json + gambar contoh siap build
skills/laprak/
├── SKILL.md                  # instruksi untuk agent (alur, aturan wajib)
├── assets/
│   ├── srwe.docx             # template kosong profil srwe (jaringan)
│   ├── eldas.docx            # template kosong profil eldas (elektronika, digital)
│   └── profil.json           # parameter format tiap profil
├── scripts/
│   ├── build_laprak.py       # generator spec.json -> .docx
│   └── selfcheck.py          # uji mandiri semua profil
└── references/
    ├── spec-schema.md        # bentuk spec.json
    ├── struktur-docx.md      # angka format hasil pembedahan template
    └── bahasa-id.md          # kaidah menulis prosa Indonesia
```

Aturan yang dipegang skill, dan sebaiknya tidak dilanggar:

- Jangan menulis `.docx` dari nol. Selalu lewat template di `assets/`.
- Jangan menimpa berkas di `assets/`.
- Nomor gambar dan tabel selalu field `SEQ`.
- Caption tabel di atas tabel, caption gambar di bawah gambar.

## Pertanyaan umum

**Agent tidak mengenali skill.** Cek folder pemasangan untuk agent-mu, mulai ulang, lalu `/skills` atau tanya daftar skill. Di OpenCode, pastikan tidak ada ID `laprak` ganda di folder yang berbeda.

**Nomor gambar tampil "Gambar 5." tanpa angka.** Field `SEQ` belum diperbarui. Word: Ctrl+A, F9.

**`ModuleNotFoundError: No module named 'docx'`.** Jalankan `pip install python-docx pillow` di Python yang dipakai agent.

**Agent menulis path skrip yang salah.** Sebutkan lokasi folder `laprak`, atau jalankan build manual seperti di atas.

**Bisa dipakai tanpa agent?** Bisa. `build_laprak.py` berdiri sendiri, cukup `spec.json` dan gambar.

**Verifikasi visual otomatis?** `SKILL.md` menjelaskan ekspor PDF lewat Word COM dari WSL. Itu khusus Windows atau WSL. Ganti `<USER>` dengan nama user Windows kamu. Di macOS dan Linux, periksa manual di Word.

**Kenapa skill tidak membuatkan cover?** Template hanya mencakup badan laporan. Lengkapi cover, nama, NIM, dan kelompok sendiri.

## Batasan

- Format kampus bisa berubah per semester. Cocokkan dengan panduan terbaru dosen atau asisten.
- Hasil tetap perlu dibaca ulang. Skill menyusun format dan draf prosa, bukan menggantikan data percobaanmu. Pahami yang kamu kumpulkan.
- Dua profil bawaan mengikuti template yang ada. Format lain butuh [profil baru](#menambah-profil-baru).
- Belum diuji langsung di Claude Code, Codex, dan Antigravity. Kompatibilitasnya mengikuti standar Agent Skills dan dokumentasi tiap agent.
- Folder skill tiap agent bisa berubah antar versi. Bila berbeda dari tabel, ikuti dokumentasi resmi agent-nya atau pakai `npx skills add`.
