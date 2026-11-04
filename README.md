# SIP-PKKS 2026 — Google Apps Script Web App
Aplikasi administrasi bukti PKKS untuk SMP Negeri 3 Cisarua. Paket ini sudah berisi 102 item bukti awal yang diekstrak dari 12 worksheet Excel yang diberikan, termasuk kode, deskripsi, tautan, dan keterangan yang tersedia.

## Isi paket
- `Code.gs` — backend, database, dashboard, CRUD, ekspor CSV, import Excel
- `Config.gs` — identitas sekolah dan daftar email berdasarkan peran
- `SeedData.gs` — struktur 12 Map dan data bukti awal dari Excel
- `Index.html`, `Styles.html`, `Scripts.html` — antarmuka mobile-first
- `appsscript.json` — konfigurasi project
- `Template_Administrasi...xlsx` tidak disertakan di project; pilih file tersebut di fitur Import Excel jika ingin memperbarui data.

## Instalasi
1. Buka https://script.google.com/ dan buat **New project**.
2. Buat file dengan nama persis seperti file pada paket. Untuk `.gs`, salin isi file ke file script. Untuk `.html`, buat HTML file dengan nama tanpa ekstensi `.html` di Apps Script (mis. `Index`, `Styles`, `Scripts`) lalu salin isinya.
3. Aktifkan tampilan `Show "appsscript.json" manifest file` di Project Settings lalu ganti manifest dengan isi `appsscript.json` dalam paket.
4. Di `Config.gs`, ganti `ganti-email-admin@sekolah.sch.id` dengan alamat email Google administrator sebenarnya. Tambahkan email kepala sekolah di `principalEmails` dan tim di `teamEmails`. Simpan.
5. Di Apps Script, buka **Services (+)**, tambahkan **Drive API** (Advanced Google Service) untuk fitur Import Excel. Jika diminta, aktifkan Google Drive API pada Google Cloud project terkait.
6. Dari dropdown fungsi, jalankan `setupApp` satu kali. Berikan izin akses Google Sheets dan Google Drive. Fungsi ini membuat database spreadsheet dan mengisi 102 item bukti awal. Lihat Execution log untuk ID database.
7. Deploy → New deployment → pilih **Web app**.
8. Untuk lingkungan Google Workspace sekolah, pilih **Execute as: Me** dan batasi akses ke pengguna dalam domain organisasi jika tersedia. Pastikan identitas email pengguna dapat dibaca oleh Apps Script pada lingkungan domain Anda. Jika `Session.getActiveUser().getEmail()` kosong, gunakan deployment/domain policy yang menyediakan identitas pengguna; jangan menonaktifkan pemeriksaan akses hanya untuk menghilangkan pesan.
9. Deploy, salin URL Web App, lalu buka di HP. Untuk setiap perubahan kode, gunakan **Deploy → Manage deployments → Edit → New version**.

## Pengaturan peran
- `adminEmails`: boleh impor, tambah/hapus bukti, mengubah seluruh data.
- `principalEmails`: boleh melihat laporan, memeriksa status, mengisi catatan dan rekomendasi.
- `teamEmails`: boleh memperbarui tautan, keterangan, dan status sesuai tugas.
- Email yang tidak terdaftar hanya dapat melihat data.
Pastikan daftar email berisi alamat sebenarnya. Jangan mengisi data pribadi/akun palsu.

## Fitur
- Dashboard progres 12 Map dan status bukti
- Pencarian, filter status, detail bukti, buka tautan
- Edit status, tautan, keterangan, tanggal pemeriksaan, pemeriksa, catatan, rekomendasi
- Tambah/hapus bukti untuk administrator
- Riwayat aktivitas perubahan
- Laporan, cetak browser, ekspor CSV
- Import Excel `.xlsx` melalui Drive API; opsi pembaruan mempertahankan status dan catatan verifikasi
- Layout responsif untuk HP/tablet/desktop dan toggle tema

## Catatan penting
- Data bukti awal diambil dari worksheet Excel. Status awal semuanya **Belum diperiksa**; jangan menganggap link yang ada berarti bukti valid.
- Data dan perubahan disimpan di Google Sheets, bukan hanya di browser.
- Menu import mengubah data referensi untuk kode bukti yang cocok jika opsi pembaruan dicentang. Status verifikasi, tanggal pemeriksaan, catatan, dan rekomendasi dipertahankan.
- `setupApp` idempotent untuk sheet yang sudah terisi; tidak mengosongkan data yang ada.
- Fitur ini belum dideploy ke akun Google pengguna. Deployment harus dilakukan oleh administrator karena membutuhkan otorisasi akun sekolah.
- Uji dahulu pada salinan database sebelum penggunaan resmi.

## Pengujian awal
1. Jalankan `setupApp`; pastikan `Bukti_PKKS` memiliki 102 baris data (di luar header).
2. Buka URL Web App di HP dan komputer.
3. Periksa Map 1 sampai Map 12, pencarian kode seperti `1.1.a`, dan tautan yang tersedia.
4. Login dengan email admin yang terdaftar; ubah status satu bukti, simpan, muat ulang, pastikan tersimpan.
5. Uji email yang tidak terdaftar: harus tidak dapat melakukan perubahan.
6. Uji ekspor CSV, cetak laporan, dan import Excel pada salinan data terlebih dahulu.
