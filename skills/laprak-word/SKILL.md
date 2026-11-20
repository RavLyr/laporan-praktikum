---
name: laprak-word
description: "Mengisi, melengkapi, merapikan, atau membangun laporan praktikum (laprak) Teknik Komputer dalam format Word sesuai template dosen, dengan penomoran akademik (caption bab.urutan, heading bertingkat, nomor halaman). Dipakai saat template laporan terbuka di Word, pengguna melampirkan .docx laporan, atau memberi modul plus screenshot dan minta laporan jadi. Juga untuk membuat kerangka atau mengecek penomoran laporan, makalah, proposal, atau skripsi. Trigger: laprak, laporan praktikum, isi BAB, eldas, srwe, penomoran gambar, caption, daftar pustaka laporan."
---

# Laporan Praktikum (Word) dengan penomoran akademik

Skill ini mengisi template laporan praktikum dan memastikan penomorannya benar, dalam satu kali pakai. Hasilnya harus langsung bisa dikumpulkan: format sesuai template, isi berasal dari data pengguna, dan tidak terbaca seperti jawaban chat.

## Kapan dipakai
- Pengguna meminta mengerjakan, melengkapi, atau merapikan laporan praktikum atau salah satu bagiannya.
- Dokumen terbuka berupa template dosen (judul BAB, heading dengan petunjuk dalam kurung, kolom kosong).
- Pengguna meminta membuat kerangka bab, merapikan, atau mengecek penomoran laporan, makalah, proposal, atau skripsi.

## Dua mode
- **Mode isi BAB** (default untuk laprak): template sudah ada, tugasnya mengisi satu BAB. Langkah 1 sampai 8 berlaku. Dari Bagian 9, yang berlaku hanya 9.1 dan 9.2. Jangan menyentuh section break, nomor halaman, cover, daftar isi, atau daftar gambar kecuali diminta.
- **Mode dokumen lengkap** (kerangka, perapian, atau pengecekan laporan, makalah, proposal, skripsi): seluruh Bagian 9 berlaku, termasuk nomor halaman di `references/nomor-halaman.md`.

## Jalur kerja
Pilih jalur di awal berdasarkan apa yang ada, lalu sebutkan satu kalimat jalur mana yang dipakai.

| Keadaan | Jalur |
|---|---|
| Template terbuka di Word dan tool add-in tersedia (`execute_office_js`, `verify_doc`, `verify_doc_visual`) | **A. Edit langsung** di dokumen terbuka |
| Template terbuka di Word di Windows, tanpa tool add-in, tetapi agent punya shell (misalnya Claude Code) | **A. Edit live lewat COM**: pakai `scripts/word_live.ps1` dari skill `laprak` (lihat bagian "Mengedit dokumen yang sudah terbuka di Word" di SKILL.md-nya). Backup `.bak` dibuat otomatis, caption memakai `Range.InsertCaption`, berkas yang terbuka ditimpa |
| Pengguna melampirkan .docx yang harus diisi atau diperbaiki, tanpa tool add-in | **A. Edit salinan** dengan skill `docx`, kembalikan sebagai .docx, sebutkan satu kalimat bahwa verifikasi visual lewat render |
| Tidak ada dokumen; pengguna memberi modul plus folder screenshot dan minta laporan jadi | **B. Bangun dari spec**: baca `references/jalur-b.md` dan ikuti langkahnya |
| Dokumen terbuka tetapi pengguna ingin dibangun ulang bersih dari template | tanya satu kali, default A |

Kedua jalur memakai aturan yang sama: Prinsip utama, analisis, gaya bahasa, sitasi, penomoran, dan pengecekan di langkah 8. Yang berbeda hanya cara menulis ke dokumen.

## Prinsip utama
1. **Data pengguna adalah acuan.** Angka, rumus, screenshot, dan foto dari pengguna dipakai apa adanya. Jangan mengarang data hasil percobaan, nilai komponen, kutipan, nomor halaman, atau referensi. Data yang belum ada ditandai `[ISI: ...]` dan masuk ke daftar cek di akhir.
2. **Template adalah hukum.** Ikuti struktur, urutan, dan style yang sudah ada. Jangan mengubah margin, font, atau style kecuali pengguna meminta. Urutan prioritas bila ada pertentangan: ketentuan dosen atau kampus di dokumen, lalu template, lalu Bagian 9, lalu Format umum di bawah. Format umum adalah default untuk semua mata kuliah bila template atau dosen tidak menetapkan lain.
3. **Tulisan, bukan chat.** Isi dokumen tidak boleh memuat komentar kepada pengguna, basa-basi, atau kalimat seperti "berikut adalah". Catatan untuk pengguna hanya ada di balasan chat.
4. **Salah di sumber, bilang.** Kalau rumus atau data pengguna tampak keliru, beri tahu singkat di chat dan jangan mengubahnya diam-diam.

## Format umum
Berlaku untuk semua mata kuliah bila template atau dosen tidak menetapkan lain.
- **Halaman**: A4, margin kiri 4, atas 4, kanan 3, bawah 3 cm.
- **Teks**: Times New Roman 12 pt di paragraf, heading, nomor heading, daftar, dan isi tabel. Rata kiri-kanan, spasi 1,5.
- **Spasi paragraf**: before dan after **0** di semua paragraf, termasuk heading, caption, dan sel tabel.
- **Page break**: setiap sub-bab level 2 dimulai di halaman baru. Contoh: setelah 12.1 selesai, page break, lalu 12.2; setelah 12.2 selesai, page break, lalu 12.3. Sub-bab level 3 (12.2.1) tidak dipisah halaman.
- **Paragraf biasa**: indent baris pertama 36 pt, left indent 0, sehingga sejajar dengan teks heading sub-bab.
- **Heading**: Heading 1 rata tengah dan tebal. Heading 2 dan 3 tebal dengan nomor otomatis, li 36 pt fi −36 pt. Nomor bab dan sub-bab tidak diubah.
- **Daftar** (Tujuan, langkah kerja, Kesimpulan): li 36 pt fi −36 pt; nomor mepet kiri, teks mulai di 36 pt.
- **Caption gambar**: di bawah gambar, TNR 10 hitam, tidak miring, rata tengah, spasi 1,5, before/after 0. Paragraf gambar di atasnya spasi 1, before/after 0.
- **Caption tabel**: di atas tabel, TNR 10 hitam, tidak miring, spasi 1, before/after 0, li 36 pt fi −36 pt.
- **Tabel data**: `Table Grid`, TNR 12, header tebal dan rata tengah, isi rata tengah, spasi 1,5, tanpa indent. Isi tabel data (alamat IP, VLAN) tidak diubah saat merapikan tabel lain.
- **Kode CLI**: tabel 1x1 `Table Grid`, Courier New 10, spasi 1, before/after 0, tidak miring.
- **Miring**: istilah bahasa Inggris di paragraf isi (*port security*, *native VLAN*, *sticky*, *DHCP snooping*). Tidak miring: nama merek, produk, atau perangkat (Cisco, Packet Tracer, Command Prompt, Rogue Laptop), heading, caption, tabel, dan kode.
- **Placeholder**: gambar atau data yang belum ada ditandai `[ISI: ...]` (di Jalur B berupa kotak abu-abu bergaris putus). Struktur dan isi template tidak dihapus tanpa diminta.

## Aturan isi
- Jangan menambah bagian Dasar Teori di dokumen apa pun kecuali diminta eksplisit.
- Langkah kerja tiap percobaan atau konfigurasi (Percobaan 1, 2, dan seterusnya) wajib dipisah. Jangan digabung atau diringkas menjadi satu daftar.
- Tujuan: tiap butir cukup dua kalimat.
- Kesimpulan murni menjawab Tujuan, tanpa penanda `[ISI: ...]` dan tanpa data kelompok (alamat IP, nilai ukur, nama kelompok).
- Bagian yang datanya belum ada (analisis, screenshot) dibiarkan dengan penanda saja; Kesimpulan tetap ditulis dari Tujuan.

## Langkah kerja

### 1. Baca dokumen dan inventarisasi
- Baca outline dan seluruh teks dokumen. Catat: nomor BAB, judul modul, heading beserta petunjuk dalam kurung (misalnya "Minimal 5"), butir bernomor kosong, tabel, dan gambar yang sudah ada.
- Baca ketentuan format di dokumen (ukuran halaman, font, spasi, margin, aturan caption, kotak kode). Periksa halaman lewat `section.pageSetup` bila tersedia. Laporkan bila tidak cocok, jangan diubah sendiri.
- Daftar apa yang sudah diberikan pengguna: modul, data, screenshot, foto. Daftar apa yang belum.
- Nilai yang dibaca dari foto atau screenshot (layar osiloskop, multimeter, simulasi) disalin dulu ke tabel kerja singkat di chat. Nilai yang kabur atau ambigu ditanyakan, tidak ditebak. Angka yang sudah dikonfirmasi dipakai apa adanya.

### 2. Tanya dulu bila kurang jelas
Maksimal dua pertanyaan penting, dalam satu pesan, hanya untuk hal yang benar-benar menghalangi penulisan (nilai ukur, foto mana untuk percobaan mana, tugas pendahuluan bila template memintanya). Contoh: "Data VCE tiap percobaan sudah lengkap? Foto mana yang untuk Percobaan 1?" Jangan bertanya hal yang bisa dibaca dari dokumen. Jangan mulai menulis analisis tanpa data hasil. Bila data sudah cukup untuk sebagian bagian, kerjakan bagian itu dulu dan tandai sisanya `[ISI: ...]` daripada menunggu jawaban.

### 3. Tulis per bagian
Sebelum menulis, tulis rencana bagian di chat (satu baris tiap bagian). Di Jalur A dengan add-in, satu bagian per panggilan `execute_office_js`. Setiap penulisan idempoten: baca dulu heading bagian target sampai heading berikutnya. Bila bagian sudah berisi, ganti isinya, jangan menambahkan di bawahnya, supaya tidak ada bagian ganda.

Urutan umum (Eldas): Tujuan, Alat dan Bahan, Langkah Kerja, Hasil Percobaan dan Analisis, Kesimpulan, lalu Daftar Pustaka bila ada sitasi.

| Bagian | Isi | Bentuk |
|---|---|---|
| Tujuan | Sesuai jumlah minimal di petunjuk. Tiap butir paling banyak dua kalimat, diawali kata kerja (mengamati, menentukan, menganalisis). | Daftar bernomor |
| Alat dan Bahan | Satu sub-heading per alat. Gambar, caption, satu paragraf fungsi singkat. Sumber gambar internet ditulis di bawahnya. | Sub-heading, tanpa bullet |
| Langkah Kerja | Dibagi Modul Software dan Modul Hardware bila template memakainya. Langkah ringkas, urut, satu tindakan per langkah, dipisah per percobaan. Gambar secukupnya. | Daftar bernomor, mulai dari 1 tiap percobaan |
| Hasil dan Analisis | Tiap percobaan: gambar atau tabel hasil, lalu paragraf analisis. | Heading tingkat empat per percobaan, tanpa bullet |
| Kesimpulan | Sesuai jumlah minimal. Menjawab tiap Tujuan, tanpa penanda dan tanpa data kelompok. | Daftar bernomor |
| Daftar Pustaka | Hanya sumber yang benar-benar dikutip. APA. | Urut abjad, indent menggantung |

Petunjuk dalam kurung pada heading (misalnya "(Minimal 5)") dihapus setelah bagiannya terisi. Sisanya, teks heading tidak diubah.

### 4. Menulis analisis (bagian yang paling sering gagal)
Susun tiap paragraf analisis dengan urutan: apa yang teramati (dengan angka), pola atau perbandingan, penyebab menurut prinsip rangkaian, lalu kaitan ke tujuan. Contoh arah untuk transistor: bagaimana VCE berubah terhadap VCC pada IB tetap, apakah transistor berada di daerah saturasi atau aktif, dan apa artinya bagi LED menyala atau mati.

- Hitung dulu dengan kode bila ada perhitungan (selisih, rata-rata, persentase galat, gain). Tulis hasil yang sudah dicek, jangan menghitung di kepala. Galat relatif memakai |nilai ukur - nilai teori| / nilai teori x 100%, ditulis sebagai persamaan bernomor bila rumusnya baru muncul. Desimal mengikuti resolusi alat ukur, konsisten di seluruh tabel.
- Setiap paragraf analisis merujuk satu gambar atau tabel yang menjadi dasarnya, dan setiap angka di paragraf harus bisa ditemukan di data pengguna atau hasil hitung tadi.
- Hubungkan penyebab hanya bila bisa dijelaskan dari data atau teori yang disitasi. Bila hasil menyimpang dari teori, tulis penyimpangannya dan sebab yang masuk akal (toleransi komponen, resistansi kontak, kesalahan baca alat), diberi kata hati-hati seperti "kemungkinan".
- Hindari klaim "membuktikan". Laporan praktikum melaporkan pengamatan.

### 5. Gaya bahasa
Bahasa Indonesia baku tapi luwes, seperti mahasiswa yang menulis rapi. Bagian ini sumber utama aturan gaya untuk kedua jalur.
- Satu gagasan per kalimat. Variasikan panjang: sisipkan kalimat pendek (lima sampai delapan kata) di antara kalimat panjang. Cek bahwa selisih terpanjang dan terpendek dalam satu paragraf lebih dari lima kata.
- Hindari nominalisasi berlebih: "mengatur", bukan "melakukan pengaturan"; "menguji", bukan "dilakukan pengujian".
- Hindari "di mana" dan "yang mana" sebagai penghubung, pleonasme ("agar supaya", "saling satu sama lain"), dan konjungsi berpasangan bertubi.
- Hindari klise: "sangat penting untuk dicatat", "dalam era modern ini", "secara keseluruhan dapat disimpulkan bahwa".
- Tanpa em dash dan en dash di badan teks. Pakai koma, titik dua, kurung, atau pecah kalimat.
- Tanpa kata ganti orang pertama ("kami", "penulis").
- Langkah kerja memakai pola kalimat yang seragam dalam satu daftar (semua imperatif atau semua pasif). Hasil dan analisis memakai kalimat aktif.
- Istilah asing belum baku ditulis miring (*collector*, *breadboard*, *waveform*, *ground*). Singkatan teknis (LED, BJT, DC, VCC) tidak dimiringkan. Heading tidak pernah dimiringkan.
- Imbuhan mengikuti KBBI: mengonfigurasi, memverifikasi, menganalisis.
- Istilah teknis yang jarang dipakai dijelaskan sekali, singkat, pada kemunculan pertama.

### 6. Tabel, persamaan, kode, gambar, rujukan
Aturan caption dan nomornya ada di Bagian 9. Di sini yang berlaku untuk isi:
- Setelah menyisipkan caption, baca ulang teksnya. Bila field gagal dibuat, laporkan dan jangan mengetik nomor diam-diam tanpa memberi tahu pengguna.
- Setiap gambar dan tabel dirujuk minimal sekali di prosa ("seperti pada Gambar 6.2", "Tabel 6.1 memuat..."), dengan huruf kapital pada Gambar dan Tabel.
- Persamaan ditulis di baris sendiri, rata tengah, dengan nomor di kanan: (6.1), (6.2). Simbol dijelaskan satu kali setelah persamaan. Rumus dari pengguna dipakai apa adanya.
- Tabel data: header tebal, isi rata tengah, angka dengan desimal konsisten dan satuan di header. Ikuti gaya tabel yang sudah ada di dokumen (baca `style` dan `headerRowCount` dari tabel sibling).
- Kode program atau perintah: dalam tabel satu sel `Table Grid` (bukan shape), Courier New 10 pt, spasi 1, before 0 after 0, warna hitam, tidak miring.
- Gambar: rata tengah, rasio asli dijaga, lebar tidak melebihi area teks (maksimal sekitar 14 cm untuk profil Eldas). Paragraf gambar tidak dipisah dari caption.

### 7. Sumber dan sitasi
- Teori, angka, dan data selain milik pengguna hanya dari sumber tepercaya: buku teks universitas, jurnal, datasheet resmi produsen, situs pemerintah atau universitas. Jangan Wikipedia, blog, atau situs tanpa penulis jelas.
- Sitasi dalam teks (Penulis, tahun). Daftar pustaka APA di akhir. Setiap sumber di daftar pustaka dikutip di teks, dan sebaliknya.
- Jangan menulis sitasi dari ingatan. Verifikasi lewat pencarian web atau gunakan sumber yang diberikan pengguna. Bila tidak yakin, tulis `[CEK SUMBER: ...]` dan masukkan ke daftar cek.
- Tujuan dan Kesimpulan tidak membutuhkan sitasi. Pembahasan boleh, secukupnya.

### 8. Verifikasi
- Setelah tiap bagian, baca ulang teks dan style yang baru ditulis, lalu bandingkan dengan paragraf template di sekitarnya (font, ukuran, spasi, before dan after, indent). Acuannya template, bukan angka tetap.
- Setelah semua selesai: di Jalur A dengan add-in, `verify_doc` untuk struktur dan penomoran, `verify_doc_visual` untuk tampilan, lalu perbarui field (caption, daftar isi bila ada). Di Jalur A dengan berkas atau Jalur B, render ke PDF lalu periksa (lihat `references/jalur-b.md` langkah 7).
- Cek penomoran sesuai Bagian 9: caption berbentuk bab.urutan, nomor mulai lagi tiap bab, heading memakai style Heading, dan (mode dokumen lengkap) nomor halaman tiap bagian benar.
- **Lint rujukan.** Kumpulkan semua paragraf prosa dan sel tabel, lalu periksa:
  1. Penanda tersisa, pola `\[(ISI|CEK SUMBER)\s*:[^\]]*\]`.
  2. Rujukan, pola `\b(Gambar|Tabel)\s+(\d+)\.(\d+)\b`: nomor BAB harus sama dengan BAB dokumen, dan nomor urut harus ada. Jumlah gambar dan tabel dihitung dari paragraf caption (style `Caption` atau `captiongambar`) yang diawali "Gambar" atau "Tabel".
  3. Setiap gambar dan tabel dirujuk minimal sekali di prosa.
  4. Pola gaya: em dash atau en dash, "di mana" dan "yang mana", "melakukan" dan "dilakukan", "kami" dan "penulis", "membuktikan".

  Cara menjalankannya: di Jalur A dengan add-in, lewat `execute_office_js` dengan membaca teks isi dokumen. Bila ada berkas .docx dan Python (Jalur A dengan berkas, Jalur B), pakai `python lint_laprak.py Laporan.docx --bab N` dari folder `skills/laprak/scripts` repo bila tersedia. Jalur B menjalankannya otomatis setelah build.

  Hasil lint dilaporkan sebagai daftar temuan. Tiap temuan ditimbang, tidak diganti otomatis. Rujukan yang nomornya salah dibetulkan hanya bila jelas gambar mana yang dimaksud. Penanda `[ISI:` dan `[CEK SUMBER:` yang tersisa dan petunjuk dalam kurung tidak boleh tertinggal tanpa dilaporkan.

### 9. Penomoran akademik
Aturan baku karya ilmiah Indonesia. Bila template atau aturan kampus berbeda, template dan aturan kampus menang.

#### 9.1 Caption, tabel, gambar (berlaku di kedua mode)
- Tabel dan gambar bernomor per bab: Tabel 2.1 berarti tabel pertama di Bab II, Gambar 3.2 berarti gambar kedua di Bab III.
- Judul tabel di atas tabel. Judul gambar di bawah gambar.
- Caption wajib fitur caption Word, bukan teks ketikan. Cara baku: `Range.InsertCaption` dengan label manual berisi nomor bab, misalnya "Gambar 6." atau "Tabel 6." (posisi Below untuk gambar, Above untuk tabel). Hasilnya field `SEQ Gambar_6. \* ARABIC`. Word menyisipkan spasi antara label dan nomor; hapus spasi itu sehingga teks menjadi "Gambar 6.1 Keterangan". Lewat COM, matikan sementara `Options.SmartCutPaste` saat menghapus spasi itu, karena kalau tidak Word mengembalikannya. `word_live.ps1 -Aksi caption` sudah melakukan semuanya.
- Lewat add-in tanpa InsertCaption, sisipkan teks "Gambar 6." lalu field `SEQ Gambar_6. \* ARABIC` tanpa spasi di antaranya. Hasilnya identik dengan InsertCaption.
- Bila template sudah memakai pola field lain (misalnya `STYLEREF 1 \s` + `SEQ ... \s 1`), ikuti template. Setelah menyisipkan, perbarui semua field dan baca ulang teksnya.
- Rujukan di badan teks ditulis "Gambar 6.1" dan "Tabel 6.2", sama persis dengan caption.
- Jalur B memakai format caption repo (teks "Gambar 5." lalu `SEQ Gambar_5.`). Hasil cetaknya sama; jangan dianggap pelanggaran.
- Selesai bila seluruh caption berbentuk bab.urutan dan urutannya mulai lagi tiap bab.

#### 9.2 Hierarki bab dan sub-bab
- Template laprak umumnya memakai format desimal (misalnya 6.3.1.1 untuk Eldas). Bila template atau pengguna memakai format desimal, pakai itu secara konsisten dan hindari bullet untuk struktur bab.
- Untuk kerangka dokumen baru tanpa template: BAB I (Romawi besar, kapital, tebal, rata tengah), lalu A., B., C. (tebal), 1., 2., 3. (tebal), a., b., c., 1), 2), 3), a), b), c).
- Heading memakai style Heading 1, 2, 3 (bukan teks tebal biasa) supaya bisa dibuat daftar isi. Penomoran selalu otomatis lewat daftar bertingkat Word, tidak diketik manual.
- Selesai bila `verify_doc` (atau pembacaan style di berkas) menunjukkan style Heading terpasang dan nomor daftar berurutan tanpa lompat.

#### 9.3 Nomor halaman (mode dokumen lengkap saja)
Baca `references/nomor-halaman.md`.

## Hal yang tidak dilakukan
- Tidak menghapus gambar, tabel, atau paragraf yang menjadi jangkar gambar.
- Tidak menghapus bagian "Laporan", "Catatan", atau "Kontak Asisten" di akhir template tanpa izin. Itu ketentuan dosen. Tanyakan apakah perlu dibuang sebelum dikumpulkan.
- Tidak mengubah teks tracked change atau komentar orang lain.
- Tidak menyalin teks dari laporan lain. Ada aturan plagiarisme di ketentuan.
- Tidak menyentuh section break, nomor halaman, cover, dan daftar isi pada mode isi BAB kecuali diminta.

## Variasi profil
- **Eldas**: judul tanpa "BAB", heading sampai empat tingkat (6.3.1.1), caption seluruhnya style `Caption`, header tabel terang (`E8EEF5`) dengan teks tebal hitam, tabel hasil boleh muncul di Hasil dan Analisis.
- **SRWE**: judul "BAB V", Tujuan, Percobaan, Hasil dan Analisis, Kesimpulan, ada tabel kode, caption gambar style `captiongambar`, header tabel berwarna (`FF89FF`) dengan teks putih bila template menggunakannya.
Bila profil tidak jelas, baca template dulu dan tanyakan satu pertanyaan. Bila ternyata profil baru (mata kuliah lain), kerjakan mengikuti template dan di akhir tawarkan menyimpan cirinya sebagai profil baru.

## Penutup setiap tugas
Balasan chat singkat berisi:
1. Bagian apa yang sudah diisi (dengan tautan heading bila ada).
2. Yang perlu dicek atau dilengkapi pengguna: foto atau screenshot yang belum ada, data yang belum diberikan, sumber yang perlu diverifikasi, bagian ketentuan di akhir template yang perlu dihapus, dan nomor caption yang perlu dilihat setelah pembaruan field.
3. Bila ada data atau rumus pengguna yang tampak janggal, satu kalimat penjelasan.
