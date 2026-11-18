# Skema spec.json

Dokumen dibangun dari daftar **bagian** bersarang. Tidak ada bagian yang dipaksakan,
jadi satu skema melayani praktikum jaringan, elektronika dasar, sistem digital, dan modul lain
dengan struktur serupa.

```jsonc
{
  "profil": "srwe",               // "srwe" atau "eldas", lihat assets/profil.json
  "bab": 5,                       // menyetir "BAB V", "5.1", "Gambar 5.x"
  "judul": "Judul Modul",         // "\n" memecah jadi beberapa baris
  "output": "Laporan.docx",       // opsional, bisa ditimpa lewat -o
  "base_dir": ".",                // akar path gambar relatif; default folder spec
  "istilah": ["gateway"],         // opsional, menimpa daftar bawaan auto-italic
  "profil_override": {},          // opsional, menimpa sebagian kunci profil

  "bagian": [ /* lihat di bawah */ ]
}
```

## Bagian

Tiap bagian menghasilkan satu heading, lalu isinya dalam urutan tetap:
`pengantar` → `isi` → `daftar` → `penutup` → `sub`.

```jsonc
{
  "judul": "Percobaan",           // opsional; tanpa judul berarti tanpa heading
  "pengantar": ["paragraf pembuka"],
  "isi": [ /* blok */ ],
  "daftar": [ /* string atau blok, dinomori 1, 2, 3 */ ],
  "penutup": [ /* blok */ ],
  "sub": [ /* bagian bersarang, otomatis turun satu tingkat heading */ ]
}
```

Bagian tingkat teratas dipisah halaman baru secara otomatis.
Kedalaman heading maksimal tiga tingkat: `5.1` → `5.1.1` → `5.1.1.1`.

## Blok

Satu blok boleh menggabungkan beberapa kunci. Urutan render selalu:
`teks` → `paragraf` → `kode` → `tabel` → `gambar` → `paragraf_akhir`.

```jsonc
{
  "teks": "Satu kalimat langkah.",          // jadi butir bernomor di dalam "daftar"
  "paragraf": ["paragraf biasa"],
  "kode": ["Router>enable", "Router#configure terminal"],
  "tabel": {
    "caption": "Data Hasil Pengukuran",
    "header": ["Sub-Percobaan", "Hasil Ukur"],
    "rows": [["Percobaan 1.1", "0,7 V"]],
    "widths": [2.9, 3.0]                    // cm, opsional
  },
  "gambar": [{"file": "02-percobaan-1/01.png", "caption": "Topologi"}],
  "paragraf_akhir": ["paragraf setelah gambar"]
}
```

`gambar` boleh berupa string, satu objek, atau daftar objek. `tabel` boleh satu objek
atau daftar objek.

## Contoh ringkas: jaringan (profil `srwe`)

```jsonc
"bagian": [
  {"judul": "Tujuan", "daftar": ["..."]},
  {"judul": "Percobaan", "sub": [
    {"judul": "Percobaan 1",
     "isi": [{"tabel": {...}}],
     "daftar": [{"teks": "...", "gambar": {...}}, {"teks": "...", "kode": [...]}]}
  ]},
  {"judul": "Hasil dan Analisis", "sub": [
    {"judul": "Percobaan 1", "isi": [{"paragraf": ["..."], "gambar": [...]}]}
  ]},
  {"judul": "Kesimpulan", "daftar": ["..."]}
]
```

## Contoh ringkas: elektronika dasar (profil `eldas`)

```jsonc
"bagian": [
  {"judul": "Tujuan", "daftar": ["..."]},
  {"judul": "Alat dan Bahan", "sub": [
    {"judul": "Trainer Board",
     "isi": [{"gambar": {...}, "paragraf_akhir": ["penjelasan", "(sumber URL)"]}]}
  ]},
  {"judul": "Langkah Kerja", "sub": [
    {"judul": "Modul Software", "sub": [
      {"judul": "Percobaan 1", "daftar": [{"teks": "...", "gambar": {...}}]}
    ]}
  ]},
  {"judul": "Hasil Percobaan dan Analisis", "sub": [
    {"judul": "Modul Software", "sub": [
      {"judul": "Percobaan 1", "isi": [{"gambar": {...}}, {"tabel": {...}}, {"paragraf": ["..."]}]}
    ]}
  ]},
  {"judul": "Kesimpulan", "daftar": ["..."]}
]
```

## Catatan

- `*miring*` memaksa cetak miring. Sisanya dimiringkan otomatis bila cocok daftar `istilah`.
- Tiap `daftar` memperoleh penomoran sendiri yang mulai dari 1.
- Rujukan silang ditulis manual, misal "seperti pada Gambar 5.4". Cocokkan nomornya
  setelah `Fields.Update()` di Word.
- Modul tanpa kode tinggal tidak memakai kunci `kode`. Tidak ada yang perlu dimatikan.
