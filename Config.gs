const APP_CONFIG = {
  appName: 'SIP-PKKS 2026',
  schoolName: 'SMP Negeri 3 Cisarua',
  schoolArea: 'Kecamatan Cisarua, Kabupaten Bandung Barat, Provinsi Jawa Barat',
  year: 2026,
  timezone: 'Asia/Jakarta',
  // GANTI dengan email Google administrator sekolah sebelum deployment.
  adminEmails: ['firmanfathurahman87@guru.smp.belajar.id'],
  // Tambahkan email kepala sekolah/tim jika ingin menetapkan peran secara khusus.
  principalEmails: [],
  teamEmails: []
};
const SHEETS = {
  evidence: 'Bukti_PKKS',
  users: 'Pengguna',
  activity: 'Riwayat_Aktivitas',
  settings: 'Pengaturan',
  maps: 'Struktur_Map'
};
const EVIDENCE_HEADERS = [
  'bukti_id','map_id','nama_map','judul_map','bidang','kode_indikator','nama_indikator',
  'kode_bukti','deskripsi_bukti','url_dokumen','status_bukti','keterangan',
  'tanggal_pemeriksaan','pemeriksa','catatan_evaluasi','rekomendasi','tanggal_dibuat','tanggal_diperbarui'
];
