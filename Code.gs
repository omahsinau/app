/**
 * ============================================================
 *  BIMBINGAN BELAJAR OMAH SINAU (OSIN) — Portal Admin, Guru & Orang Tua
 *  Backend: Google Apps Script + Google Spreadsheet
 * ============================================================
 *  Cara pakai singkat (detail di PANDUAN.md):
 *  1. Buat Google Spreadsheet baru → Extensions → Apps Script.
 *  2. Tempel file ini ke Code.gs, buat file HTML bernama "index"
 *     lalu tempel isi index.html.
 *  3. Pilih fungsi setup() → Run (izinkan akses).
 *  4. Deploy → New deployment → Web app
 *       Execute as: Me | Who has access: Anyone
 *     (setiap ubah kode: Manage deployments → Edit → New version)
 *  5. Buka URL /exec → login admin (owner): admin / omahsinau123 (segera ganti).
 * ============================================================
 */

// Kosongkan jika script dibuat dari dalam spreadsheet (Extensions → Apps Script).
const SS_ID = '';
const SESSION_SECONDS = 21600; // 6 jam
const TZ = 'Asia/Jakarta';

// Kolom baru selalu DITAMBAH DI BELAKANG agar sheet lama tetap cocok.
const SCHEMA = {
  Pengaturan:   ['key', 'value'],
  Users:        ['id', 'username', 'passwordHash', 'nama', 'role'], // role: owner | akademik | admin2
  Paket:        ['id', 'nama', 'kategori', 'harga', 'feeGuru', 'satuan', 'urut', 'aktif', 'feeUtama', 'feeAndalan', 'feeMaster', 'waktuBayar'], // feeGuru = fee Mitra Tentor · waktuBayar: Awal / Akhir
  Siswa:        ['id', 'nama', 'panggilan', 'jk', 'tglLahir', 'jenjang', 'kelas', 'sekolah', 'paketId', 'homeVisit', 'ortu', 'waOrtu', 'alamat', 'tglMulai', 'status', 'sppNominal', 'kodeOrtu', 'catatan', 'waSiswa', 'tempatLes', 'jadwalDiminta', 'paketDiminta', 'tglDaftar'],
  Guru:         ['id', 'nama', 'wa', 'status', 'pin', 'mapel', 'tarifPaket', 'catatan', 'jenjang', 'metodeBayar', 'noRekening', 'atasNama', 'pendidikan', 'alamat', 'tglGabung'],
  Jadwal:       ['id', 'hari', 'jamMulai', 'jamSelesai', 'tipe', 'guruId', 'siswaIds', 'lokasi', 'catatan'],
  Absensi:      ['id', 'tanggal', 'siswaId', 'status', 'guruId', 'catatan', 'jam', 'jadwalId'], // 1 baris = 1 siswa hadir di 1 sesi guru
  Perkembangan: ['id', 'tanggal', 'siswaId', 'guruId', 'materi', 'deskripsi'],
  SPP:          ['id', 'noKuitansi', 'siswaId', 'periode', 'nominal', 'tanggalBayar', 'metode', 'untuk', 'catatan', 'jenis'], // jenis: Les (bulanan) / Pendaftaran / Buku / Lain
  Gaji:         ['id', 'guruId', 'periode', 'pokok', 'transport', 'bonus', 'potongan', 'total', 'rincian', 'catatan', 'tanggalBayar', 'sesi', 'jenjang'],
  Raport:       ['id', 'siswaId', 'periode', 'kelas', 'paket', 'hadir', 'pertemuan', 'hasil', 'arahan', 'guruId', 'status', 'tanggal', 'dibuat'],
  Keuangan:     ['id', 'tanggal', 'tipe', 'kategori', 'keterangan', 'nominal', 'metode', 'dibuat'], // tipe: Pemasukan / Pengeluaran / Simpanan
  PendaftarTentor: ['id', 'tanggal', 'nama', 'wa', 'email', 'pendidikan', 'alamat', 'mapel', 'pengalaman', 'ketersediaan', 'motivasi', 'status', 'catatan', 'token', 'guruId'],
  Materi:       ['id', 'tanggal', 'judul', 'jenis', 'jenjang', 'kelas', 'mapel', 'deskripsi', 'fileId', 'url', 'namaFile', 'mime', 'ukuran', 'oleh', 'olehNama', 'olehTipe'] // bank soal & materi (file di Google Drive)
};
const ADMIN_TABLES = ['Paket', 'Siswa', 'Guru', 'Jadwal', 'Absensi', 'Perkembangan', 'SPP', 'Gaji', 'Raport', 'Keuangan', 'PendaftarTentor', 'Materi'];
const JENIS_MATERI = ['Materi', 'Latihan Soal', 'Kunci Jawaban', 'Lainnya'];
const JENJANG_MATERI = ['TK', 'SD', 'SMP', 'SMA', 'Umum'];
const MAX_FILE_MB = 10;
const HEAD_COLOR = '#1E8A4C';

// Hak akses admin. owner = semua. akademik = jadwal/siswa/sesi/raport. admin2 = akademik + tagihan & kuitansi ortu.
const AKADEMIK_TABLES = ['Siswa', 'Guru', 'Jadwal', 'Absensi', 'Perkembangan', 'Raport'];
// Materi dibaca semua admin, tapi diubah lewat aksi saveMateri / delMateri (karena ada file di Drive)
const ACCESS = {
  owner:    { read: ADMIN_TABLES, write: ADMIN_TABLES.filter(t => t !== 'Materi') },
  akademik: { read: ['Paket'].concat(AKADEMIK_TABLES, ['Materi']), write: AKADEMIK_TABLES },
  admin2:   { read: ['Paket'].concat(AKADEMIK_TABLES, ['SPP', 'Materi']), write: AKADEMIK_TABLES.concat(['SPP']) }
};
const GURU_PRIVATE = ['tarifPaket', 'metodeBayar', 'noRekening', 'atasNama', 'jenjang']; // hanya owner yang boleh ubah
const levelOf_ = role => (role === 'akademik' || role === 'admin2') ? role : 'owner';

const DEFAULT_SETTINGS = {
  namaLembaga: 'Bimbingan Belajar Omah Sinau',
  singkatan: 'OSIN',
  tagline: 'Bimbel TK · SD · SMP · SMA',
  alamat: 'Perumahan Regency Jabal Blok B No. 9 Kaliwungu, Kab. Kendal',
  kota: 'Kaliwungu',
  noWa: '081227206753',
  kontak: 'Dian',
  tahunAjaran: '2026/2027',
  jatuhTempo: '10',
  rekening: '',
  portalUrl: '',
  namaKepala: 'Risca Dian Pratiwi, S.Pd',
  namaAkademik: 'Robiatul Adawiyah, S.M',
  namaPenerima: 'Risca Dian Pratiwi',
  noKuitansiAwal: '1',
  ttdKepala: '',
  ttdAkademik: '',
  ttdPenerima: '',
  jatuhTempoAkhir: '28',
  tagihAkhirMulai: '26',
  biayaPendaftaran: '50000',
  biayaBuku: '40000',
  tipeKelas: 'Klasikal\nRombel\nPrivat Bimbel\nPrivat Rumah\nPrivat Bimbel Khusus\nHome Visit',
  lokasiList: 'Rumah siswa\nJabal\nJagalan\nCitran\nPatebon\nKertomulyo',
  saldoAwalKas: '0',
  saldoAwalPeriode: '',
  mapelList: 'Calistung\nTematik\nMatematika\nBahasa Indonesia\nBahasa Inggris\nIPA\nIPS\nPKn\nBahasa Jawa\nLainnya',
  materiFolderId: ''
};

// Paket belajar (sesuai daftar paket). Harga & fee guru diisi admin di menu Paket & Fee Guru.
const DEFAULT_PAKET = [
  ['Klasikal TK SD 1-2 anak', 'Klasikal'],
  ['Klasikal TK SD 3-4 anak', 'Klasikal'],
  ['Klasikal SD SMP 1-2 anak', 'Klasikal'],
  ['Klasikal SD & SMP 3-4 anak', 'Klasikal'],
  ['Privat Klasikal Khusus 2 anak', 'Klasikal Khusus'],
  ['Klasikal Khusus 3-4 anak', 'Klasikal Khusus'],
  ['Privat Bimbel TK & SD', 'Privat'],
  ['Privat Bimbel SMP', 'Privat'],
  ['Privat Bimbel SMA', 'Privat'],
  ['Privat (Home/Umum) TK & SD', 'Privat Home'],
  ['Privat (Home/Umum) SMP', 'Privat Home'],
  ['Privat (Home/Umum) SMA', 'Privat Home'],
  ['Uang Transport (Home Visit / Luar Bimbel)', 'Transport']
];

// Data guru contoh (bisa diubah / ditambah di Portal Admin).
const DEFAULT_GURU = [
  'Qoshimatul Awaliyah', 'Inarotul Maftukhah', 'Osy Siswi Utami', 'Risca Dian P', 'Nurul', 'Wilma Alfa Serra',
  'Miftah', 'Adelia', 'Auliya', 'Rumah', 'Makde Siti', 'Luthfa', 'Fitri Oktaviani', 'Citra Della', 'Aisyah Latifah',
  'Enggar', 'Siva', 'Silvia', 'Devi Wanda', 'Nabil Roy M', 'Ficky Nurul', 'Baharuddin', 'Sofiyatus Salma',
  'Lisa Silvia', 'Haiyul', 'Wahyu Wicaksono', 'Bu Mahromah', 'Indah', 'Meliza', 'Eliza', 'Lisa Chily',
  'Agustina', 'Shinta Isnaini'
];

/* ---------------- ENTRY POINTS ---------------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Portal Omah Sinau')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let res;
  try { res = route_(JSON.parse(e.postData.contents)); }
  catch (err) { res = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function api(reqJson) {
  try { return JSON.stringify(route_(JSON.parse(reqJson))); }
  catch (err) { return JSON.stringify({ ok: false, error: String(err && err.message || err) }); }
}

/* ---------------- ROUTER ---------------- */

function route_(req) {
  const a = req.action;
  if (a === 'ping') return { ok: true, time: new Date().toISOString() };
  if (a === 'login') return login_(req);
  if (a === 'publicSettings') return { ok: true, settings: publicSettings_(true) };
  if (a === 'daftarMurid') return { ok: true, id: withLock_(() => daftarMurid_(req.data || {})) };
  if (a === 'daftarTentor') return { ok: true, id: withLock_(() => daftarTentor_(req.data || {})) };
  if (a === 'dataTentorInfo') return { ok: true, info: dataTentorInfo_(req.token2) };
  if (a === 'dataTentorSubmit') return { ok: true, guru: withLock_(() => dataTentorSubmit_(req.token2, req.data || {})) };

  const s = getSession_(req.token);
  if (!s) return { ok: false, error: 'SESSION', message: 'Sesi berakhir, silakan login kembali.' };
  if (a === 'logout') { CacheService.getScriptCache().remove('s_' + req.token); return { ok: true }; }

  if (s.role === 'admin') {
    const lv = levelOf_(s.level);
    const can = (t, w) => { checkTable_(t); if (ACCESS[lv][w ? 'write' : 'read'].indexOf(t) < 0) throw new Error('Akses ditolak untuk menu ini'); };
    const owner = () => { if (lv !== 'owner') throw new Error('Khusus owner'); };
    switch (a) {
      case 'bootstrap':      return { ok: true, data: bootstrap_(lv) };
      case 'save':           can(req.table, 1); return { ok: true, row: cleanOut_(lv, req.table, withLock_(() => upsertMany_(req.table, [cleanIn_(lv, req.table, req.row)]))[0]) };
      case 'saveMany':       can(req.table, 1); return { ok: true, rows: withLock_(() => upsertMany_(req.table, (req.rows || []).map(r => cleanIn_(lv, req.table, r)))).map(r => cleanOut_(lv, req.table, r)) };
      case 'remove':         can(req.table, 1); withLock_(() => removeRow_(req.table, req.id)); return { ok: true };
      case 'saveSettings':   owner(); { const st = Object.assign({}, req.settings || {}); delete st.materiFolderId; withLock_(() => saveSettings_(st)); } return { ok: true, settings: getSettings_() };
      case 'changePassword': return changePassword_(s, req);
      case 'saveSesi':       return Object.assign({ ok: true }, withLock_(() => saveSesi_(req.guruId, req.tanggal, req.jam, req.siswaIds, req.old, false, req.jadwalId)));
      case 'delSesi':        withLock_(() => delSesi_(req)); return { ok: true };
      case 'aktivasiSiswa':  return Object.assign({ ok: true }, withLock_(() => aktivasiSiswa_(req.siswa || {}, req.biaya || [])));
      case 'saveUser':       owner(); return { ok: true, users: withLock_(() => saveUser_(req.user || {})) };
      case 'delUser':        owner(); return { ok: true, users: withLock_(() => delUser_(s, req.id)) };
      case 'saveMateri':     return { ok: true, row: withLock_(() => saveMateri_(req.row || {}, req.file, pembuat_(s))) };
      case 'delMateri':      withLock_(() => delMateri_(req.id, pembuat_(s))); return { ok: true };
    }
  }
  if (s.role === 'guru') {
    switch (a) {
      case 'guruData':         return { ok: true, data: guruData_(s) };
      case 'guruSaveSesi':     return Object.assign({ ok: true }, withLock_(() => saveSesi_(s.uid, req.tanggal, req.jam, req.siswaIds, req.old ? Object.assign({}, req.old, { guruId: s.uid }) : null, true, req.jadwalId, req.notes)));
      case 'guruDelSesi':      withLock_(() => delSesi_({ guruId: s.uid, tanggal: req.tanggal, jam: req.jam })); return { ok: true };
      case 'guruSaveProgress': return { ok: true, rows: withLock_(() => guruSaveProgress_(s, req.rows || [])) };
      case 'guruDelProgress':  withLock_(() => guruDelProgress_(s, req.id)); return { ok: true };
      case 'guruSaveRaport':   return { ok: true, row: withLock_(() => guruSaveRaport_(s, req.row || {})) };
      case 'gantiPassword':    return withLock_(() => gantiPassword_(s, req));
      case 'guruSaveMateri':   return { ok: true, row: withLock_(() => saveMateri_(req.row || {}, req.file, pembuat_(s))) };
      case 'guruDelMateri':    withLock_(() => delMateri_(req.id, pembuat_(s))); return { ok: true };
    }
  }
  if (s.role === 'ortu') {
    if (a === 'me') return { ok: true, data: me_(s) };
    if (a === 'gantiPassword') return withLock_(() => gantiPassword_(s, req));
  }
  return { ok: false, error: 'Akses ditolak' };
}

// data keluar ke admin non-owner: sembunyikan rekening & fee khusus guru
function cleanOut_(lv, table, row) {
  if (lv === 'owner' || table !== 'Guru' || !row) return row;
  const o = Object.assign({}, row); GURU_PRIVATE.forEach(k => { if (k !== 'jenjang') delete o[k]; }); return o;
}
// data masuk dari admin non-owner: kolom rahasia guru tidak boleh diubah
function cleanIn_(lv, table, row) {
  const r = Object.assign({}, row || {});
  if (lv !== 'owner' && table === 'Guru') GURU_PRIVATE.forEach(k => delete r[k]);
  return r;
}

/* ---------------- AUTH ---------------- */

function login_(req) {
  const role = String(req.role || '');
  const user = String(req.user || '').trim();
  const pass = String(req.pass || '').trim();
  if (!user || !pass) return { ok: false, error: 'Lengkapi data login.' };

  let profile = null;
  if (role === 'admin') {
    const u = readTable_('Users').find(r => r.username.toLowerCase() === user.toLowerCase());
    if (u && u.passwordHash === hash_(pass)) profile = { id: u.id, nama: u.nama, username: u.username, level: levelOf_(u.role) };
  } else if (role === 'guru') {
    lengkapiDataGuru();
    const g = readTable_('Guru').find(r => r.id.toUpperCase() === user.toUpperCase());
    if (g && String(g.pin) === pass && g.status !== 'Nonaktif') profile = { id: g.id, nama: g.nama };
  } else if (role === 'ortu') {
    lengkapiDataSiswa();
    const s = readTable_('Siswa').find(r => r.id.toUpperCase() === user.toUpperCase());
    if (s && String(s.kodeOrtu) === pass && s.status !== 'Berhenti' && s.status !== 'Baru') profile = { id: s.id, nama: s.nama, panggilan: s.panggilan };
  }
  if (!profile) return { ok: false, error: role === 'admin' ? 'Username atau password salah.' : role === 'guru' ? 'ID guru atau PIN salah.' : 'ID siswa atau kode akses salah.' };

  const token = Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('s_' + token, JSON.stringify({ role: role, uid: profile.id, level: profile.level || '' }), SESSION_SECONDS);
  return { ok: true, token: token, role: role, profile: profile };
}

function getSession_(token) {
  if (!token) return null;
  const v = CacheService.getScriptCache().get('s_' + token);
  if (!v) return null;
  CacheService.getScriptCache().put('s_' + token, v, SESSION_SECONDS);
  return JSON.parse(v);
}

function changePassword_(s, req) {
  const u = readTable_('Users').find(r => r.id === s.uid);
  if (!u || u.passwordHash !== hash_(String(req.oldPass || ''))) return { ok: false, error: 'Password lama salah.' };
  if (String(req.newPass || '').length < 6) return { ok: false, error: 'Password baru minimal 6 karakter.' };
  u.passwordHash = hash_(String(req.newPass));
  withLock_(() => upsertMany_('Users', [u]));
  return { ok: true };
}

// guru ganti PIN / ortu ganti kode akses (admin tetap bisa melihat di daftar)
function gantiPassword_(s, req) {
  const oldP = String(req.oldPass || '').trim(), newP = String(req.newPass || '').trim();
  if (newP.length < 4) return { ok: false, error: 'Password baru minimal 4 karakter.' };
  if (!/^[A-Za-z0-9]+$/.test(newP)) return { ok: false, error: 'Gunakan huruf/angka saja (tanpa spasi).' };
  if (s.role === 'guru') {
    const g = readTable_('Guru').find(r => r.id === s.uid);
    if (!g || String(g.pin) !== oldP) return { ok: false, error: 'PIN lama salah.' };
    upsertMany_('Guru', [{ id: g.id, pin: newP }]);
  } else {
    const x = readTable_('Siswa').find(r => r.id === s.uid);
    if (!x || String(x.kodeOrtu) !== oldP) return { ok: false, error: 'Kode akses lama salah.' };
    upsertMany_('Siswa', [{ id: x.id, kodeOrtu: newP }]);
  }
  return { ok: true };
}

/* ---- akun admin (dikelola owner) ---- */
const safeUsers_ = () => readTable_('Users').map(u => ({ id: u.id, username: u.username, nama: u.nama, role: levelOf_(u.role) }));
function saveUser_(u) {
  const all = readTable_('Users');
  const username = String(u.username || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,}$/.test(username)) throw new Error('Username minimal 3 karakter (huruf kecil/angka).');
  if (all.some(x => x.username.toLowerCase() === username && x.id !== u.id)) throw new Error('Username sudah dipakai.');
  const role = levelOf_(u.role);
  const ex = u.id ? all.find(x => x.id === u.id) : null;
  if (ex && levelOf_(ex.role) === 'owner' && role !== 'owner' && all.filter(x => levelOf_(x.role) === 'owner').length < 2) throw new Error('Harus ada minimal 1 akun owner.');
  const row = { username: username, nama: String(u.nama || username), role: role };
  if (ex) row.id = ex.id;
  if (u.password) { if (String(u.password).length < 6) throw new Error('Password minimal 6 karakter.'); row.passwordHash = hash_(String(u.password)); }
  else if (!ex) throw new Error('Isi password untuk akun baru.');
  upsertMany_('Users', [row]);
  return safeUsers_();
}
function delUser_(s, id) {
  if (id === s.uid) throw new Error('Tidak bisa menghapus akun sendiri.');
  const all = readTable_('Users'), u = all.find(x => x.id === id);
  if (u && levelOf_(u.role) === 'owner' && all.filter(x => levelOf_(x.role) === 'owner').length < 2) throw new Error('Harus ada minimal 1 akun owner.');
  removeRow_('Users', id);
  return safeUsers_();
}

function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

const today_ = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
const daysAgo_ = n => Utilities.formatDate(new Date(Date.now() - n * 864e5), TZ, 'yyyy-MM-dd');

/* ---------------- ADMIN ---------------- */

function bootstrap_(lv) {
  lengkapiDataSiswa();
  lengkapiDataGuru();
  const db = {};
  ACCESS[lv].read.forEach(t => db[t] = readTable_(t));
  if (lv !== 'owner') {
    db.Guru = db.Guru.map(g => { const o = Object.assign({}, g); GURU_PRIVATE.forEach(k => { if (k !== 'jenjang') delete o[k]; }); return o; });
    db.Paket = db.Paket.map(p => { const o = Object.assign({}, p); ['feeGuru', 'feeUtama', 'feeAndalan', 'feeMaster'].forEach(k => delete o[k]); if (lv === 'akademik') delete o.harga; return o; });
  }
  const out = { db: db, settings: getSettings_(), level: lv };
  if (lv === 'owner') out.users = safeUsers_();
  return out;
}

/* ---------------- REGISTRASI (tanpa login) ---------------- */
const str_ = (v, n) => String(v == null ? '' : v).trim().slice(0, n || 200);
function rateLimit_(key, max) {
  const c = CacheService.getScriptCache(), k = 'rl_' + key, n = Number(c.get(k) || 0);
  if (n >= max) throw new Error('Terlalu banyak pendaftaran, coba lagi beberapa menit lagi.');
  c.put(k, String(n + 1), 600);
}
function kelasDariForm_(k) {
  k = String(k || '');
  let m;
  if (/TK|PAUD/i.test(k)) return { jenjang: 'TK', kelas: '' };
  if ((m = k.match(/SD\D*(\d+)/i))) return { jenjang: 'SD', kelas: m[1] };
  if ((m = k.match(/SMP\D*(\d+)/i))) return { jenjang: 'SMP', kelas: String(+m[1] + 6) };
  if ((m = k.match(/SMA\D*(XII|XI|X)/i))) return { jenjang: 'SMA', kelas: { X: '10', XI: '11', XII: '12' }[m[1].toUpperCase()] };
  return { jenjang: '', kelas: k };
}
function daftarMurid_(d) {
  if (d.website) throw new Error('Gagal');
  ['nama', 'ortu', 'waOrtu'].forEach(k => { if (!str_(d[k])) throw new Error('Lengkapi nama, nama orang tua & nomor WhatsApp.'); });
  rateLimit_('murid', 30);
  const kl = kelasDariForm_(d.kelas);
  const row = {
    nama: str_(d.nama, 100), ortu: str_(d.ortu, 100), paketDiminta: str_(d.paket, 60), jenjang: kl.jenjang, kelas: kl.kelas,
    sekolah: str_(d.sekolah, 100), tempatLes: str_(d.tempatLes, 80), jadwalDiminta: str_(d.jadwal, 120), alamat: str_(d.alamat, 200),
    waOrtu: str_(d.waOrtu, 20), waSiswa: str_(d.waSiswa, 20), jk: d.jk === 'L' ? 'L' : 'P', status: 'Baru', tglDaftar: today_(), homeVisit: /home/i.test(d.paket) ? 'Ya' : 'Tidak',
    catatan: str_(d.catatan, 300)
  };
  return upsertMany_('Siswa', [row])[0].id;
}
function daftarTentor_(d) {
  if (d.website) throw new Error('Gagal');
  ['nama', 'wa'].forEach(k => { if (!str_(d[k])) throw new Error('Lengkapi nama & nomor WhatsApp.'); });
  rateLimit_('tentor', 20);
  const row = { tanggal: today_(), status: 'Baru' };
  ['nama', 'wa', 'email', 'pendidikan', 'alamat', 'mapel', 'pengalaman', 'ketersediaan', 'motivasi'].forEach(k => row[k] = str_(d[k], k === 'pengalaman' || k === 'motivasi' ? 600 : 200));
  return upsertMany_('PendaftarTentor', [row])[0].id;
}
function pendaftarByToken_(token) {
  token = String(token || '');
  if (token.length < 10) throw new Error('Link tidak valid.');
  const p = readTable_('PendaftarTentor').find(x => x.token === token);
  if (!p || !/Diterima/.test(p.status)) throw new Error('Link tidak valid atau sudah tidak berlaku.');
  return p;
}
function dataTentorInfo_(token) {
  const p = pendaftarByToken_(token);
  if (p.guruId) return { done: true, nama: p.nama, guruId: p.guruId };
  return { nama: p.nama, wa: p.wa, pendidikan: p.pendidikan, alamat: p.alamat };
}
function dataTentorSubmit_(token, d) {
  const p = pendaftarByToken_(token);
  if (p.guruId) throw new Error('Data sudah pernah dikirim. ID guru: ' + p.guruId);
  ['nama', 'wa', 'metodeBayar', 'noRekening'].forEach(k => { if (!str_(d[k])) throw new Error('Lengkapi semua kolom wajib.'); });
  const pin = rand4_();
  const g = upsertMany_('Guru', [{
    nama: str_(d.nama, 100), wa: str_(d.wa, 20), pendidikan: str_(d.pendidikan, 150), alamat: str_(d.alamat, 200),
    metodeBayar: str_(d.metodeBayar, 30), noRekening: str_(d.noRekening, 60), atasNama: str_(d.atasNama, 100),
    status: /Menunggu/.test(p.status) ? 'Menunggu Murid' : 'Aktif', pin: pin, jenjang: 'Mitra Tentor', mapel: p.mapel, tglGabung: today_(), tarifPaket: '', catatan: ''
  }])[0];
  upsertMany_('PendaftarTentor', [{ id: p.id, guruId: g.id }]);
  return { id: g.id, pin: pin, nama: g.nama };
}

/* ---------------- AKTIVASI MURID BARU ---------------- */
function aktivasiSiswa_(sw, biaya) {
  const ex = readTable_('Siswa').find(x => x.id === sw.id);
  if (!ex) throw new Error('Data siswa tidak ditemukan');
  const allow = ['nama', 'panggilan', 'jk', 'jenjang', 'kelas', 'sekolah', 'paketId', 'homeVisit', 'ortu', 'waOrtu', 'waSiswa', 'alamat', 'tglMulai', 'sppNominal', 'kodeOrtu', 'catatan', 'tempatLes'];
  const row = { id: ex.id, status: 'Aktif' };
  allow.forEach(k => { if (sw[k] != null) row[k] = str_(sw[k], 300); });
  if (!row.tglMulai) row.tglMulai = today_();
  if (!row.kodeOrtu && !ex.kodeOrtu) row.kodeOrtu = rand4_();
  const siswa = upsertMany_('Siswa', [row])[0];
  const per = siswa.tglMulai.slice(0, 7);
  const spp = upsertMany_('SPP', (biaya || []).filter(b => Number(b.nominal) > 0).map(b => ({
    siswaId: siswa.id, periode: per, nominal: String(Math.round(Number(b.nominal))), tanggalBayar: '', metode: '', jenis: ['Pendaftaran', 'Buku'].indexOf(b.jenis) > -1 ? b.jenis : 'Lain',
    untuk: str_(b.untuk || ('Biaya ' + String(b.jenis || '').toLowerCase()), 120), catatan: ''
  })));
  return { siswa: siswa, spp: spp };
}

/* ---------------- HELPER PAKET / SPP ---------------- */

function paketMap_() {
  const m = {}; readTable_('Paket').forEach(p => m[p.id] = p); return m;
}
function transportPaket_(pm) {
  return Object.keys(pm).map(k => pm[k]).find(p => p.kategori === 'Transport' && p.aktif !== 'Tidak') || null;
}
function tagihanSiswa_(s, pm) {
  if (Number(s.sppNominal)) return Number(s.sppNominal);
  const p = pm[s.paketId];
  const t = s.homeVisit === 'Ya' ? transportPaket_(pm) : null;
  return (p ? Number(p.harga) || 0 : 0) + (t ? Number(t.harga) || 0 : 0);
}

/* ---------------- GURU ---------------- */

function diajarMap_(uid) {
  const jadwal = readTable_('Jadwal').filter(j => j.guruId === uid);
  const diajar = {};
  jadwal.forEach(j => String(j.siswaIds).split(',').forEach(x => { if (x.trim()) diajar[x.trim()] = 1; }));
  return { jadwal: jadwal, diajar: diajar };
}

function guruData_(s) {
  const g = readTable_('Guru').find(r => r.id === s.uid);
  if (!g) throw new Error('Data guru tidak ditemukan');
  const dm = diajarMap_(s.uid), diajar = dm.diajar;
  const pm = paketMap_();
  const siswa = readTable_('Siswa').filter(x => x.status === 'Aktif')
    .map(x => ({ id: x.id, nama: x.nama, panggilan: x.panggilan, jk: x.jk, jenjang: x.jenjang, kelas: x.kelas, paketId: x.paketId, paket: (pm[x.paketId] || {}).nama || '', homeVisit: x.homeVisit, diajar: diajar[x.id] ? '1' : '' }));
  const cutoff = daysAgo_(62), cutoffP = daysAgo_(400).slice(0, 7);
  const safe = { id: g.id, nama: g.nama, wa: g.wa, status: g.status, mapel: g.mapel, jenjang: g.jenjang };
  return {
    guru: safe,
    settings: publicSettings_(true),
    siswa: siswa,
    Jadwal: dm.jadwal,
    Absensi: readTable_('Absensi').filter(r => r.tanggal >= cutoff && (r.guruId === s.uid || diajar[r.siswaId])),
    Perkembangan: readTable_('Perkembangan').filter(r => r.tanggal >= cutoff && (r.guruId === s.uid || diajar[r.siswaId])),
    Raport: readTable_('Raport').filter(r => r.periode >= cutoffP && (r.guruId === s.uid || diajar[r.siswaId])),
    Gaji: readTable_('Gaji').filter(r => r.guruId === s.uid && r.tanggalBayar),
    Materi: readTable_('Materi')
  };
}

/* ---------------- BANK SOAL & MATERI ----------------
 * File disimpan di Google Drive: folder "OSIN - Bank Soal & Materi" / <Mapel> / <Kelas>.
 * Admin (semua level) & guru bisa unggah. Guru hanya bisa mengubah/menghapus unggahan sendiri.
 */
function pembuat_(s) {
  if (s.role === 'guru') { const g = readTable_('Guru').find(r => r.id === s.uid) || {}; return { tipe: 'guru', id: s.uid, nama: g.nama || s.uid }; }
  const u = readTable_('Users').find(r => r.id === s.uid) || {};
  return { tipe: 'admin', id: s.uid, nama: u.nama || u.username || 'Admin' };
}
const kelasMateri_ = (jenjang, kelas) => !kelas ? jenjang + (jenjang === 'Umum' ? '' : ' (semua kelas)') : jenjang === 'TK' ? 'TK ' + kelas : jenjang + ' Kelas ' + kelas;
function subFolder_(parent, name) { const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function materiRoot_() {
  const id = getSettings_().materiFolderId;
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const f = DriveApp.createFolder('OSIN - Bank Soal & Materi');
  saveSettings_({ materiFolderId: f.getId() });
  return f;
}
const materiFolder_ = (mapel, jenjang, kelas) => subFolder_(subFolder_(materiRoot_(), mapel), kelasMateri_(jenjang, kelas));
function bisaUbahMateri_(m, by) { return by.tipe === 'admin' || (m.olehTipe === 'guru' && m.oleh === by.id); }
function saveMateri_(r, file, by) {
  const judul = str_(r.judul, 150), mapel = str_(r.mapel, 60), jenjang = str_(r.jenjang, 10), kelas = str_(r.kelas, 10);
  if (!judul) throw new Error('Isi judul materi');
  if (!mapel) throw new Error('Pilih mata pelajaran');
  if (JENJANG_MATERI.indexOf(jenjang) < 0) throw new Error('Pilih jenjang');
  const ex = r.id ? readTable_('Materi').find(x => x.id === r.id) : null;
  if (r.id && !ex) throw new Error('Materi tidak ditemukan');
  if (ex && !bisaUbahMateri_(ex, by)) throw new Error('Hanya bisa mengubah materi yang kamu unggah sendiri');
  const o = ex ? Object.assign({}, ex) : { tanggal: today_(), fileId: '', namaFile: '', mime: '', ukuran: '', oleh: by.id, olehNama: by.nama, olehTipe: by.tipe };
  Object.assign(o, { judul: judul, jenis: JENIS_MATERI.indexOf(r.jenis) > -1 ? r.jenis : 'Materi', jenjang: jenjang, kelas: jenjang === 'Umum' ? '' : kelas, mapel: mapel, deskripsi: str_(r.deskripsi, 1000), url: str_(r.url, 500) });
  if (o.url && !/^https?:\/\//i.test(o.url)) o.url = 'https://' + o.url;
  if (file && file.b64) {
    const bytes = Utilities.base64Decode(String(file.b64));
    if (bytes.length > MAX_FILE_MB * 1048576) throw new Error('Ukuran file maksimal ' + MAX_FILE_MB + ' MB');
    const blob = Utilities.newBlob(bytes, str_(file.mime, 100) || 'application/octet-stream', str_(file.name, 150) || judul);
    const f = materiFolder_(o.mapel, o.jenjang, o.kelas).createFile(blob);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
    catch (e) { try { f.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e2) {} }
    if (o.fileId) { try { DriveApp.getFileById(o.fileId).setTrashed(true); } catch (e) {} }
    Object.assign(o, { fileId: f.getId(), namaFile: f.getName(), mime: blob.getContentType(), ukuran: String(bytes.length) });
  } else if (ex && o.fileId && (ex.mapel !== o.mapel || ex.jenjang !== o.jenjang || ex.kelas !== o.kelas)) {
    try { DriveApp.getFileById(o.fileId).moveTo(materiFolder_(o.mapel, o.jenjang, o.kelas)); } catch (e) {}
  }
  if (!o.fileId && !o.url) throw new Error('Unggah file atau isi link materi');
  return upsertMany_('Materi', [o])[0];
}
function delMateri_(id, by) {
  const m = readTable_('Materi').find(x => x.id === id);
  if (!m) return;
  if (!bisaUbahMateri_(m, by)) throw new Error('Hanya bisa menghapus materi yang kamu unggah sendiri');
  if (m.fileId) { try { DriveApp.getFileById(m.fileId).setTrashed(true); } catch (e) {} }
  removeRow_('Materi', id);
}

/* ---------------- SESI (guru mengisi sesi + siswa yang hadir) ---------------- */
// Kunci 1 sesi = guruId + tanggal + jam. Simpan = ganti semua baris sesi itu dengan daftar siswa baru.
function saveSesi_(guruId, tanggal, jam, siswaIds, old, isGuru, jadwalId, notes) {
  tanggal = String(tanggal || ''); jam = String(jam || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(tanggal)) throw new Error('Tanggal tidak valid');
  if (isGuru && tanggal > today_()) throw new Error('Tanggal tidak boleh melebihi hari ini');
  if (!guruId || !readTable_('Guru').some(g => g.id === guruId)) throw new Error('Guru tidak ditemukan');
  const ada = {}; readTable_('Siswa').forEach(x => { if (!isGuru || x.status === 'Aktif') ada[x.id] = 1; });
  const ids = (siswaIds || []).map(String).filter((id, i, arr) => ada[id] && arr.indexOf(id) === i);
  if (!ids.length) throw new Error('Centang minimal satu siswa yang hadir');
  const keys = [{ guruId: guruId, tanggal: tanggal, jam: jam }];
  if (old && old.tanggal) keys.push({ guruId: old.guruId || guruId, tanggal: String(old.tanggal), jam: String(old.jam || '') });
  deleteWhere_('Absensi', r => keys.some(k => r.guruId === k.guruId && r.tanggal === k.tanggal && String(r.jam || '') === k.jam));
  const rows = upsertMany_('Absensi', ids.map(id => ({ tanggal: tanggal, siswaId: id, status: 'Hadir', guruId: guruId, catatan: '', jam: jam, jadwalId: str_(jadwalId, 40) })));
  // catatan perkembangan per siswa (diisi bersamaan dengan sesi)
  let perkembangan = [];
  if (notes && typeof notes === 'object') {
    const ex = readTable_('Perkembangan').filter(p => p.guruId === guruId && p.tanggal === tanggal);
    const up = [], del = [];
    ids.forEach(sid => {
      const n = notes[sid] || {}, e = ex.find(p => p.siswaId === sid), d = str_(n.deskripsi, 1000);
      if (d) up.push({ id: e ? e.id : '', tanggal: tanggal, siswaId: sid, guruId: guruId, materi: str_(n.materi, 120), deskripsi: d });
      else if (e) del.push(e.id);
    });
    up.forEach(r => { if (!r.id) delete r.id; });
    if (del.length) deleteWhere_('Perkembangan', p => del.indexOf(p.id) > -1);
    perkembangan = upsertMany_('Perkembangan', up);
  }
  return { rows: rows, perkembangan: perkembangan };
}
function delSesi_(k) {
  const jam = String(k.jam || '');
  deleteWhere_('Absensi', r => r.guruId === k.guruId && r.tanggal === k.tanggal && String(r.jam || '') === jam);
}

function guruSaveProgress_(s, rows) {
  const aktif = {}; readTable_('Siswa').forEach(x => { if (x.status === 'Aktif') aktif[x.id] = 1; });
  const existing = readTable_('Perkembangan');
  const out = rows.filter(r => aktif[r.siswaId] && String(r.deskripsi || '').trim()).map(r => {
    const ex = r.id ? existing.find(e => e.id === r.id && e.guruId === s.uid) : null;
    const o = { tanggal: r.tanggal || today_(), siswaId: r.siswaId, guruId: s.uid, materi: String(r.materi || ''), deskripsi: String(r.deskripsi).trim() };
    if (ex) o.id = ex.id;
    return o;
  });
  return upsertMany_('Perkembangan', out);
}

function guruDelProgress_(s, id) {
  const ex = readTable_('Perkembangan').find(e => e.id === id && e.guruId === s.uid);
  if (ex) removeRow_('Perkembangan', id);
}

function guruSaveRaport_(s, r) {
  const siswa = readTable_('Siswa').find(x => x.id === r.siswaId && x.status === 'Aktif');
  if (!siswa) throw new Error('Siswa tidak ditemukan / tidak aktif');
  if (!/^\d{4}-\d{2}$/.test(String(r.periode || ''))) throw new Error('Periode tidak valid');
  const ex = readTable_('Raport').find(x => x.siswaId === r.siswaId && x.periode === r.periode);
  if (ex && ex.status === 'Terbit') throw new Error('Raport ini sudah diterbitkan admin. Hubungi admin bila perlu diubah.');
  const o = {
    siswaId: r.siswaId, periode: r.periode, kelas: String(r.kelas || ''), paket: String(r.paket || ''),
    hadir: String(r.hadir || ''), pertemuan: String(r.pertemuan || ''), hasil: String(r.hasil || ''), arahan: String(r.arahan || ''),
    guruId: s.uid, status: 'Draft', tanggal: (ex && ex.tanggal) || '', dibuat: today_()
  };
  if (ex) o.id = ex.id;
  return upsertMany_('Raport', [o])[0];
}

/* ---------------- ORANG TUA ---------------- */

function me_(s) {
  const siswa = readTable_('Siswa').find(r => r.id === s.uid);
  if (!siswa) throw new Error('Data siswa tidak ditemukan');
  const mine = t => readTable_(t).filter(r => r.siswaId === s.uid);
  const guru = readTable_('Guru');
  const gNama = id => { const g = guru.find(x => x.id === id); return g ? g.nama : ''; };
  const pm = paketMap_();
  const jadwal = readTable_('Jadwal')
    .filter(j => String(j.siswaIds).split(',').map(x => x.trim()).indexOf(s.uid) > -1)
    .map(j => Object.assign({}, j, { siswaIds: '', guruNama: gNama(j.guruId) }));
  const safe = Object.assign({}, siswa); delete safe.kodeOrtu;
  safe.paket = (pm[siswa.paketId] || {}).nama || '';
  safe.tagihan = String(tagihanSiswa_(siswa, pm));
  safe.waktuBayar = (pm[siswa.paketId] || {}).waktuBayar || 'Awal';
  return {
    role: 'ortu',
    siswa: safe,
    settings: publicSettings_(false),
    Absensi: mine('Absensi').map(r => Object.assign({}, r, { guruNama: gNama(r.guruId) })),
    SPP: mine('SPP'),
    Raport: mine('Raport').filter(r => r.status === 'Terbit'),
    Perkembangan: mine('Perkembangan').map(p => Object.assign({}, p, { guruNama: gNama(p.guruId) })),
    Jadwal: jadwal
  };
}

/* ---------------- SETTINGS ---------------- */

function getSettings_() {
  const o = Object.assign({}, DEFAULT_SETTINGS);
  readTable_('Pengaturan').forEach(r => { if (r.key) o[r.key] = r.value; });
  return o;
}
function publicSettings_(noImages) {
  const s = getSettings_();
  const keys = ['namaLembaga', 'singkatan', 'tagline', 'alamat', 'kota', 'noWa', 'kontak', 'rekening', 'jatuhTempo', 'jatuhTempoAkhir', 'tagihAkhirMulai', 'tahunAjaran', 'portalUrl',
    'namaKepala', 'namaAkademik', 'namaPenerima', 'biayaPendaftaran', 'biayaBuku', 'tipeKelas', 'lokasiList', 'mapelList'];
  if (!noImages) keys.push('ttdKepala', 'ttdAkademik', 'ttdPenerima');
  const o = {}; keys.forEach(k => o[k] = s[k]);
  try { o.appUrl = ScriptApp.getService().getUrl(); } catch (e) { o.appUrl = ''; }
  return o;
}
function saveSettings_(obj) {
  const sh = sheet_('Pengaturan');
  const cur = getSettings_();
  Object.keys(obj).forEach(k => cur[k] = obj[k]);
  const rows = Object.keys(cur).map(k => [k, String(cur[k])]);
  rows.forEach(r => { if (r[1].length > 49000) throw new Error('Gambar "' + r[0] + '" terlalu besar. Gunakan gambar yang lebih kecil.'); });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
}

/* ---------------- SHEET HELPERS ---------------- */

function ss_() { return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

const _checked = {};
function sheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  const h = SCHEMA[name];
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground(HEAD_COLOR).setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), h.length).setNumberFormat('@');
  } else if (!_checked[name]) {
    // migrasi: tambahkan kolom baru di belakang bila sheet lama
    const cur = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0].filter(String);
    if (cur.length < h.length) {
      sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground(HEAD_COLOR).setFontColor('#FFFFFF');
      sh.getRange(1, cur.length + 1, sh.getMaxRows(), h.length - cur.length).setNumberFormat('@');
    }
  }
  _checked[name] = 1;
  return sh;
}

function checkTable_(t) { if (ADMIN_TABLES.indexOf(t) < 0) throw new Error('Tabel tidak dikenal: ' + t); }

function readTable_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const h = SCHEMA[name];
  return sh.getRange(2, 1, last - 1, h.length).getDisplayValues()
    .filter(r => r[0] !== '')
    .map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
}

function newId_(name, existing) {
  const seq = (prefix, digits) => {
    let max = 0;
    existing.forEach(r => { const m = String(r.id).match(new RegExp('^' + prefix + '(\\d+)$')); if (m && +m[1] > max) max = +m[1]; });
    return prefix + String(max + 1).padStart(digits, '0');
  };
  if (name === 'Siswa') return seq('OS', 3);
  if (name === 'Guru') return seq('G', 2);
  if (name === 'Paket') return seq('P', 2);
  return name.slice(0, 2).toUpperCase() + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

// Nomor kuitansi: OSIN/yymmdd/urut  (contoh OSIN/260929/283)
function nextKuitansi_(row, existing) {
  let max = Math.max(0, (Number(getSettings_().noKuitansiAwal) || 1) - 1);
  existing.forEach(r => { const m = String(r.noKuitansi || '').match(/\/(\d+)$/); if (m && +m[1] > max) max = +m[1]; });
  const d = String(row.tanggalBayar || today_()).replace(/-/g, '').slice(2, 8);
  return 'OSIN/' + d + '/' + String(max + 1).padStart(3, '0');
}

function upsertMany_(name, rows) {
  const sh = sheet_(name);
  const h = SCHEMA[name];
  const existing = readTable_(name);
  const idx = {}; existing.forEach((r, i) => idx[r.id] = i + 2);
  const out = [], appends = [];
  rows.forEach(row => {
    const r = Object.assign({}, row);
    const needNo = () => name === 'SPP' && r.tanggalBayar && !r.noKuitansi;
    const needTok = () => name === 'PendaftarTentor' && /Diterima/.test(r.status || '') && !r.token;
    if (!r.id || !idx[r.id]) {
      if (!r.id) r.id = newId_(name, existing.concat(out));
      if (needNo()) r.noKuitansi = nextKuitansi_(r, existing.concat(out));
      if (needTok()) r.token = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
      appends.push(h.map(k => r[k] == null ? '' : String(r[k])));
    } else {
      const old = existing[idx[r.id] - 2];
      const merged = Object.assign({}, old, r);
      Object.assign(r, merged);
      if (needNo()) r.noKuitansi = nextKuitansi_(r, existing.concat(out));
      if (needTok()) r.token = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
      sh.getRange(idx[r.id], 1, 1, h.length).setNumberFormat('@').setValues([h.map(k => r[k] == null ? '' : String(r[k]))]);
    }
    const clean = {}; h.forEach(k => clean[k] = r[k] == null ? '' : String(r[k]));
    out.push(clean);
  });
  if (appends.length) {
    const start = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(start, 1, appends.length, h.length).setNumberFormat('@').setValues(appends);
  }
  return out;
}

function removeRow_(name, id) {
  const sh = sheet_(name);
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getDisplayValues().map(r => r[0]);
  const i = ids.indexOf(String(id));
  if (i > 0) sh.deleteRow(i + 1);
}

function deleteWhere_(name, pred) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const h = SCHEMA[name];
  const vals = sh.getRange(2, 1, last - 1, h.length).getDisplayValues();
  const rows = [];
  vals.forEach((r, i) => { if (r[0] === '') return; const o = {}; h.forEach((k, j) => o[k] = r[j]); if (pred(o)) rows.push(i + 2); });
  rows.reverse().forEach(n => sh.deleteRow(n));
  return rows.length;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------------- INPUT LANGSUNG DI GOOGLE SHEET ----------------
 * Siswa & Guru boleh diketik langsung di sheet. Kolom kosong (id, kode/PIN, status, dst)
 * terisi otomatis saat portal admin dibuka / ada login, atau jalankan fungsi ini manual.
 */
function fillSheet_(name, fillers, idPrefix, digits) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const h = SCHEMA[name];
  const rng = sh.getRange(2, 1, last - 1, h.length);
  const vals = rng.getDisplayValues();
  const c = k => h.indexOf(k);
  let max = 0;
  vals.forEach(r => { const m = String(r[0]).match(new RegExp('^' + idPrefix + '(\\d+)$')); if (m && +m[1] > max) max = +m[1]; });
  let changed = 0;
  vals.forEach(r => {
    if (!String(r[c('nama')]).trim()) return;
    const before = r.join('|');
    if (!r[0]) r[0] = idPrefix + String(++max).padStart(digits, '0');
    Object.keys(fillers).forEach(k => { if (!r[c(k)]) r[c(k)] = fillers[k](); });
    if (r.join('|') !== before) changed++;
  });
  if (changed) rng.setNumberFormat('@').setValues(vals);
  return changed;
}
const rand4_ = () => String(Math.floor(1000 + Math.random() * 9000));
function lengkapiDataSiswa() {
  return fillSheet_('Siswa', { kodeOrtu: rand4_, status: () => 'Aktif', tglMulai: today_, jk: () => 'P', homeVisit: () => 'Tidak' }, 'OS', 3);
}
function lengkapiDataGuru() {
  return fillSheet_('Guru', { pin: rand4_, status: () => 'Aktif', jenjang: () => 'Mitra Tentor' }, 'G', 2);
}

/* ---------------- SETUP (jalankan sekali) ---------------- */

// Klasikal dibayar awal bulan; privat & home visit dibayar akhir bulan (tgl 26–28)
const waktuDefault_ = (nama, kat) => /privat|home|transport/i.test(String(nama) + ' ' + String(kat)) ? 'Akhir' : 'Awal';

function setup() {
  Object.keys(SCHEMA).forEach(sheet_);
  const def = ss_().getSheetByName('Sheet1') || ss_().getSheetByName('Lembar1');
  if (def && def.getLastRow() === 0 && ss_().getSheets().length > 1) ss_().deleteSheet(def);
  if (!readTable_('Users').length) {
    upsertMany_('Users', [{ username: 'admin', passwordHash: hash_('omahsinau123'), nama: 'Owner Omah Sinau', role: 'owner' }]);
  }
  if (!readTable_('Paket').length) {
    upsertMany_('Paket', DEFAULT_PAKET.map((p, i) => ({ nama: p[0], kategori: p[1], harga: '0', feeGuru: '0', satuan: 'Per bulan', urut: String(i + 1), aktif: 'Ya', feeUtama: '', feeAndalan: '', feeMaster: '', waktuBayar: waktuDefault_(p[0], p[1]) })));
  }
  // migrasi: paket lama tanpa waktu bayar
  const pk = readTable_('Paket').filter(p => !p.waktuBayar).map(p => ({ id: p.id, waktuBayar: waktuDefault_(p.nama, p.kategori) }));
  if (pk.length) upsertMany_('Paket', pk);
  if (!readTable_('Guru').length) {
    upsertMany_('Guru', DEFAULT_GURU.map(n => ({ nama: n, wa: '', status: 'Aktif', pin: rand4_(), mapel: '', tarifPaket: '', catatan: '', jenjang: 'Mitra Tentor', tglGabung: today_() })));
  }
  saveSettings_({}); // menulis pengaturan default yang belum ada
  materiRoot_();     // folder Google Drive untuk bank soal & materi (minta izin Drive)
  lengkapiDataSiswa();
  lengkapiDataGuru();
  Logger.log('Setup selesai. Login owner: admin / omahsinau123 (segera ganti).');
}

function resetAdminPassword() {
  const u = readTable_('Users').find(r => r.username === 'admin');
  if (u) { u.passwordHash = hash_('omahsinau123'); u.role = 'owner'; upsertMany_('Users', [u]); }
  else upsertMany_('Users', [{ username: 'admin', passwordHash: hash_('omahsinau123'), nama: 'Owner Omah Sinau', role: 'owner' }]);
  Logger.log('Password admin direset ke: omahsinau123');
}
