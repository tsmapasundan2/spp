// SPP SMA Pasundan 2 Tasikmalaya: penghubung website dan Google Sheet.
// Tempel seluruh isi file ini di Extensions > Apps Script pada Google Sheet Anda.

// Kode untuk membuat akun admin pertama. Kosong = tanpa kode: pembuat akun pertama otomatis menjadi admin.
const SETUP_CODE = '';

const SESSION_SECONDS = 6 * 60 * 60;
const MONTHS = ['Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des', 'Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun'];

// Kolom tiap lembar: [kunci data, judul kolom, tipe: s teks, n angka, b ya/kosong]
const SHEETS = {
  Siswa: { key: 'students', cols: [['id', 'ID', 's'], ['nis', 'NIS', 's'], ['nama', 'Nama', 's'], ['kelas', 'Kelas', 's'], ['spp', 'SPP per Bulan', 'n'], ['status', 'Status', 's'], ['ket', 'Keterangan', 's'], ['baru', 'Siswa Baru', 'b']] },
  Pembayaran: {
    key: 'payments',
    cols: [['id', 'ID', 's'], ['receiptNo', 'No. Kuitansi', 's'], ['studentId', 'ID Siswa', 's'], ['year', 'Tahun Ajaran (awal)', 'n'], ['month', 'Bulan ke- (1=Juli)', 'n'], ['amount', 'Nominal', 'n'], ['date', 'Tanggal', 's'], ['method', 'Metode', 's'], ['officer', 'Bendahara', 's'], ['note', 'Catatan', 's']],
    extra: [['Nama Siswa', (o, st) => (st[o.studentId] || {}).nama || ''], ['Bulan SPP', o => MONTHS[(o.month || 1) - 1] || '']]
  },
  Biaya: {
    key: 'fees',
    cols: [['id', 'ID', 's'], ['receiptNo', 'No. Kuitansi', 's'], ['studentId', 'ID Siswa', 's'], ['year', 'Tahun Ajaran (awal)', 'n'], ['kind', 'Jenis', 's'], ['amount', 'Nominal', 'n'], ['date', 'Tanggal', 's'], ['method', 'Metode', 's'], ['officer', 'Bendahara', 's'], ['note', 'Catatan', 's']],
    extra: [['Nama Siswa', (o, st) => (st[o.studentId] || {}).nama || '']]
  }
};

function doGet() {
  return ContentService.createTextOutput('API SPP aktif.');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try { out = route_(req); } finally { lock.releaseLock(); }
  } catch (err) {
    out = { error: 'Terjadi kesalahan di server: ' + err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function route_(req) {
  const me = userOf_(req.token);
  switch (req.action) {
    case 'rev': return { rev: rev_() };
    case 'load': return { rev: rev_(), hasUsers: users_().length > 0, me, data: readAll_() };
    case 'login': return login_(req);
    case 'setup': return setup_(req);
    case 'logout': if (req.token) CacheService.getScriptCache().remove('t_' + req.token); return { ok: true };
    case 'save': return me ? save_(req, me) : { unauthorized: true };
    case 'listUsers': return isAdmin_(me) ? { users: users_().map(publicUser_) } : { error: 'Hanya admin.' };
    case 'addUser': return isAdmin_(me) ? addUser_(req) : { error: 'Hanya admin.' };
    case 'deleteUser': return isAdmin_(me) ? deleteUser_(req, me) : { error: 'Hanya admin.' };
    case 'setPassword': return me ? setPassword_(req, me) : { unauthorized: true };
    default: return { error: 'Permintaan tidak dikenal.' };
  }
}

// ---------- data ----------

const props_ = () => PropertiesService.getScriptProperties();
const rev_ = () => props_().getProperty('rev') || '0';

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
const conv_ = (v, t) => t === 'n' ? (Number(v) || 0) : t === 'b' ? (v === 'ya' || v === true) : (v === null || v === undefined ? '' : String(v));

function readSection_(name) {
  const def = SHEETS[name], sh = sheet_(name), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, def.cols.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => { const o = {}; def.cols.forEach((c, i) => { o[c[0]] = conv_(r[i], c[2]); }); return o; });
}

function readKeyValue_() {
  const sh = sheet_('Pengaturan'), last = sh.getLastRow(), out = {};
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(r => { try { out[r[0]] = JSON.parse(r[1]); } catch (e) { /* abaikan */ } });
  return out;
}

function readAll_() {
  const kv = readKeyValue_();
  const data = { v: 2, settings: kv.settings || {}, seq: kv.seq || {} };
  Object.keys(SHEETS).forEach(n => { data[SHEETS[n].key] = readSection_(n); });
  return data;
}

function writeSection_(name, arr, studentsById) {
  const def = SHEETS[name], sh = sheet_(name), extra = def.extra || [];
  const head = def.cols.map(c => c[1]).concat(extra.map(x => x[0]));
  sh.clearContents();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (!arr.length) return;
  const rows = arr.map(o => def.cols.map(c => c[2] === 'b' ? (o[c[0]] ? 'ya' : '') : (o[c[0]] === undefined || o[c[0]] === null ? '' : o[c[0]]))
    .concat(extra.map(x => x[1](o, studentsById))));
  const rng = sh.getRange(2, 1, rows.length, head.length);
  rng.setNumberFormat('General');
  head.forEach((h, i) => { if (i >= def.cols.length || def.cols[i][2] === 's') sh.getRange(2, i + 1, rows.length, 1).setNumberFormat('@'); });
  rng.setValues(rows);
}

function writeKeyValue_(settings, seq) {
  const sh = sheet_('Pengaturan');
  sh.clearContents();
  sh.getRange(1, 1, 1, 2).setValues([['Kunci', 'Nilai (JSON)']]).setFontWeight('bold');
  sh.getRange(2, 1, 2, 2).setNumberFormat('@').setValues([['settings', JSON.stringify(settings)], ['seq', JSON.stringify(seq)]]);
}

const digest_ = s => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s));

function save_(req, me) {
  const base = rev_();
  if (String(req.rev) !== base) return { conflict: true, rev: base, data: readAll_() };
  const d = req.data || {};
  const cur = readAll_();
  const settings = me.role === 'admin' && d.settings ? d.settings : cur.settings;
  const next = { settings, seq: d.seq || {} };
  Object.keys(SHEETS).forEach(n => { next[SHEETS[n].key] = Array.isArray(d[SHEETS[n].key]) ? d[SHEETS[n].key] : cur[SHEETS[n].key]; });
  const studentsById = {};
  next.students.forEach(s => { studentsById[s.id] = s; });
  const p = props_();
  Object.keys(SHEETS).forEach(n => {
    const arr = next[SHEETS[n].key], h = digest_(JSON.stringify(arr));
    if (p.getProperty('h_' + n) !== h) { writeSection_(n, arr, studentsById); p.setProperty('h_' + n, h); }
  });
  writeKeyValue_(next.settings, next.seq);
  const rev = String(Number(base) + 1);
  p.setProperty('rev', rev);
  return { ok: true, rev };
}

// ---------- akun ----------

function users_() {
  const sh = sheet_('Akun'), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 6).getValues().filter(r => r[0] !== '')
    .map(r => ({ id: String(r[0]), username: String(r[1]), nama: String(r[2]), role: String(r[3]), salt: String(r[4]), hash: String(r[5]) }));
}
function saveUsers_(list) {
  const sh = sheet_('Akun');
  sh.clearContents();
  sh.getRange(1, 1, 1, 6).setValues([['ID', 'Nama Pengguna', 'Nama', 'Peran', 'Salt', 'Hash Sandi']]).setFontWeight('bold');
  if (!list.length) return;
  const rng = sh.getRange(2, 1, list.length, 6);
  rng.setNumberFormat('@');
  rng.setValues(list.map(u => [u.id, u.username, u.nama, u.role, u.salt, u.hash]));
}
const publicUser_ = u => ({ id: u.id, username: u.username, nama: u.nama, role: u.role });
const isAdmin_ = me => !!me && me.role === 'admin';

function hash_(pw, salt) {
  let h = salt + pw;
  for (let i = 0; i < 200; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  return h;
}
function userOf_(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get('t_' + token);
  return raw ? JSON.parse(raw) : null;
}
function startSession_(u) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  const me = publicUser_(u);
  CacheService.getScriptCache().put('t_' + token, JSON.stringify(me), SESSION_SECONDS);
  return { token, user: me };
}
function checkNewUser_(r, users) {
  if (!String(r.nama || '').trim()) return 'Nama wajib diisi.';
  if (!/^[a-z0-9._-]{3,}$/.test(r.username || '')) return 'Nama pengguna minimal 3 karakter: huruf kecil, angka, titik, strip.';
  if (users.some(u => u.username === r.username)) return 'Nama pengguna sudah dipakai.';
  if (String(r.password || '').length < 6) return 'Kata sandi minimal 6 karakter.';
  return '';
}
function makeUser_(r, role) {
  const salt = Utilities.getUuid();
  return { id: 'u' + Utilities.getUuid().slice(0, 8), username: r.username, nama: r.nama.trim(), role, salt, hash: hash_(r.password, salt) };
}

function tooManyTries_(key) {
  const c = CacheService.getScriptCache();
  return Number(c.get(key) || 0) >= 5;
}
function addTry_(key) {
  const c = CacheService.getScriptCache();
  c.put(key, String(Number(c.get(key) || 0) + 1), 600);
}

function login_(r) {
  const username = String(r.username || '').trim().toLowerCase(), key = 'f_' + username;
  if (tooManyTries_(key)) return { error: 'Terlalu banyak percobaan. Coba lagi 10 menit lagi.' };
  const u = users_().find(x => x.username === username);
  if (!u || hash_(String(r.password || ''), u.salt) !== u.hash) { addTry_(key); return { error: 'Nama pengguna atau kata sandi salah.' }; }
  return startSession_(u);
}

function setup_(r) {
  const users = users_();
  if (users.length) return { error: 'Admin sudah ada. Masuk dengan akun Anda.' };
  if (tooManyTries_('f_setup')) return { error: 'Terlalu banyak percobaan. Coba lagi 10 menit lagi.' };
  if (SETUP_CODE && String(r.code || '').trim() !== SETUP_CODE) { addTry_('f_setup'); return { error: 'Kode pendaftaran salah.' }; }
  r.username = String(r.username || '').trim().toLowerCase();
  const bad = checkNewUser_(r, users);
  if (bad) return { error: bad };
  const u = makeUser_(r, 'admin');
  saveUsers_([u]);
  return startSession_(u);
}

function addUser_(r) {
  const users = users_();
  r.username = String(r.username || '').trim().toLowerCase();
  const bad = checkNewUser_(r, users);
  if (bad) return { error: bad };
  users.push(makeUser_(r, r.role === 'admin' ? 'admin' : 'petugas'));
  saveUsers_(users);
  return { ok: true };
}

function deleteUser_(r, me) {
  const users = users_(), u = users.find(x => x.id === r.id);
  if (!u) return { error: 'Akun tidak ditemukan.' };
  if (u.id === me.id) return { error: 'Anda tidak bisa menghapus akun sendiri.' };
  if (u.role === 'admin' && users.filter(x => x.role === 'admin').length < 2) return { error: 'Admin terakhir tidak bisa dihapus.' };
  saveUsers_(users.filter(x => x.id !== u.id));
  return { ok: true };
}

function setPassword_(r, me) {
  const users = users_(), u = users.find(x => x.id === r.userId);
  if (!u) return { error: 'Akun tidak ditemukan.' };
  if (String(r.password || '').length < 6) return { error: 'Kata sandi minimal 6 karakter.' };
  if (u.id === me.id) {
    if (hash_(String(r.current || ''), u.salt) !== u.hash) return { error: 'Kata sandi saat ini salah.' };
  } else if (me.role !== 'admin') {
    return { error: 'Hanya admin.' };
  }
  u.salt = Utilities.getUuid();
  u.hash = hash_(r.password, u.salt);
  saveUsers_(users);
  return { ok: true };
}

// Lupa sandi admin: isi dua nilai di bawah, pilih fungsi resetSandiAdmin di bar atas editor, lalu klik Jalankan.
// Setelah berhasil, masuk ke website dengan sandi sementara itu, ganti sandinya di Pengaturan, lalu kosongkan lagi dua nilai ini.
function resetSandiAdmin() {
  const USERNAME = '';
  const SANDI_SEMENTARA = '';
  if (!USERNAME || SANDI_SEMENTARA.length < 6) throw new Error('Isi USERNAME dan SANDI_SEMENTARA (minimal 6 karakter) dulu.');
  const users = users_(), u = users.find(x => x.username === USERNAME.toLowerCase());
  if (!u) throw new Error('Nama pengguna tidak ditemukan di lembar Akun.');
  u.salt = Utilities.getUuid();
  u.hash = hash_(SANDI_SEMENTARA, u.salt);
  saveUsers_(users);
  CacheService.getScriptCache().remove('f_' + u.username);
  console.log('Sandi untuk ' + u.username + ' sudah diatur ulang.');
}
