# Nomor halaman (mode dokumen lengkap saja)

Aturan baku karya ilmiah Indonesia. Bila template atau aturan kampus berbeda, template dan
aturan kampus menang.

- **Bagian awal** (Kata Pengantar, Abstrak, Daftar Isi, Daftar Tabel, Daftar Gambar): angka
  Romawi kecil (i, ii, iii, ...) di bawah tengah. Cover dan Halaman Pengesahan tetap dihitung
  urutannya, tetapi nomornya tidak ditampilkan.
- **Bagian inti** (Bab I sampai Bab Penutup): angka Arab, mulai dari 1 di halaman pertama Bab I.
  - Halaman awal bab (yang memuat judul BAB): nomor di bawah tengah.
  - Halaman lanjutan bab: nomor di pojok kanan atas.
- **Bagian akhir** (Daftar Pustaka, Lampiran): nomor Arab melanjutkan bagian inti tanpa
  mengulang dari 1, posisi seperti halaman lanjutan (pojok kanan atas).

## Implementasi di Word

Pisahkan bagian dengan Section Break (Next Page), matikan Link to Previous di header dan footer
tiap bagian, atur Format Page Number (Romawi kecil atau Arab, Start at), dan pakai Different
First Page untuk halaman awal bab. Nomor halaman memakai field, bukan teks ketikan.

Selesai bila tiap bagian punya format dan posisi nomor sesuai aturan, dan halaman pertama
Bab I bernomor 1.

Bila pengguna bertanya cara mengaturnya di Word atau Google Docs, beri langkah teknis memakai
Section Break (Next Page) dan mematikan Link to Previous.
