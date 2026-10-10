// TPFX Sales Workspace - Google Apps Script backend (database = Google Sheets)
// Spreadsheet database: memakai sheet tempat script ini dibuat; jika script berdiri sendiri, dibuat otomatis.

function ss_() {
  let s = SpreadsheetApp.getActiveSpreadsheet();
  if (s) return s;

  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('SS_ID');

  if (id) return SpreadsheetApp.openById(id);

  s = SpreadsheetApp.create('TPFX Sales Database');
  p.setProperty('SS_ID', s.getId());

  return s;
}

function users_() {
  const s = ss_().getSheetByName('Users');
  if (!s) throw new Error('Sheet Users belum ada. Jalankan fungsi setup dulu.');
  return s;
}

function data_() {
  const s = ss_().getSheetByName('Data');
  if (!s) throw new Error('Sheet Data belum ada. Jalankan fungsi setup dulu.');
  return s;
}

// Mengambil sheet; jika belum ada, dibuat otomatis lengkap dengan header.
// Jadi chat/feed tetap jalan walau setup() belum sempat dijalankan.
function sheetEnsure_(name, headers) {
  const b = ss_();
  let sh = b.getSheetByName(name);
  if (!sh) {
    sh = b.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

function posts_() {
  return sheetEnsure_('Posts', ['id', 'username', 'nama', 'text', 'ts', 'imageUrl']);
}

function messages_() {
  return sheetEnsure_('Messages', ['id', 'from', 'to', 'text', 'ts']);
}

// Keep row numbers stable for active chat cursors while removing message contents.
function redactMessageRow_(sheet, rowNumber) {
  sheet.getRange(rowNumber, 2, 1, 4).clearContent();
}

function chatSummary_() {
  return sheetEnsure_('ChatSummary', [
    'key', 'type', 'target', 'text', 'ts', 'from', 'row'
  ]);
}

function chatThreadKey_(type, from, target) {
  if (type == 'personal') {
    return 'personal:' + [String(from).toLowerCase(), String(target).toLowerCase()]
      .sort().join('|');
  }
  return type + ':' + String(target);
}

function rebuildChatSummary_() {
  const messages = messages_();
  const lastRow = messages.getLastRow();
  const summary = chatSummary_();
  const latest = {};
  if (lastRow > 1) {
    messages.getRange(2, 1, lastRow - 1, 5).getValues().forEach((row, index) => {
      const from = String(row[1] || '');
      const to = String(row[2] || '');
      if (!String(row[0] || '').trim() || (!from && !to)) return;
      let type, target;
      if (to == 'global') {
        type = 'global';
        target = 'global';
      } else if (to.indexOf('group:') == 0) {
        type = 'group';
        target = to.slice(6);
      } else {
        type = 'personal';
        target = to;
      }
      const key = chatThreadKey_(type, from, target);
      const sheetRow = index + 2;
      if (!latest[key] || sheetRow > latest[key][6]) {
        latest[key] = [
          key, type, target, String(row[3] || ''), Number(row[4]) || 0,
          from, sheetRow
        ];
      }
    });
  }
  const oldLastRow = summary.getLastRow();
  if (oldLastRow > 1) summary.getRange(2, 1, oldLastRow - 1, 7).clearContent();
  const rows = Object.keys(latest).map(key => latest[key]);
  if (rows.length) summary.getRange(2, 1, rows.length, 7).setValues(rows);
}

function updateChatSummary_(type, target, from, text, ts, row) {
  const summary = chatSummary_();
  const key = chatThreadKey_(type, from, target);
  const lastRow = summary.getLastRow();
  const keys = lastRow > 1
    ? summary.getRange(2, 1, lastRow - 1, 1).getValues()
    : [];
  const index = keys.findIndex(value => String(value[0]) == key);
  const values = [key, type, target, text, ts, from, row];
  if (index < 0) summary.appendRow(values);
  else summary.getRange(index + 2, 1, 1, 7).setValues([values]);
}

function deleteChatSummary_(type, from, target) {
  const summary = chatSummary_();
  const key = chatThreadKey_(type, from, target);
  const rows = summary.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) == key) summary.deleteRow(i + 1);
  }
}

function chatSummaryRows_() {
  return chatSummary_().getDataRange().getValues().slice(1);
}

function postLikes_() {
  return sheetEnsure_('PostLikes', ['postId', 'username', 'ts']);
}

function postPins_() {
  return sheetEnsure_('PostPins', ['postId', 'username', 'ts']);
}

function postComments_() {
  return sheetEnsure_('PostComments', ['id', 'postId', 'username', 'text', 'ts']);
}

function chatGroups_() {
  return sheetEnsure_('ChatGroups', ['id', 'name', 'owner', 'members_json', 'ts']);
}


// ============================================================
// SESSION
// ============================================================

const DAYS = 30;


// ============================================================
// DATABASE CALON NASABAH
// ============================================================

// TEMPEL seluruh record array lama (dari `const DATA=[...]` di database.html
// commit 388de7a) ke dalam array di bawah ini, tanpa mengubah isi record.
// Array harus berisi SELURUH daftar calon nasabah, semua kategori/kelompok.
// Format: [{ n: 'Nama kelompok', r: [ [nomor, pekerjaan, telepon], ... ] }, ...]
// Jangan menaruh isi array ini di file dokumentasi atau commit Git baru.

const DATABASE_DATA = [
  /* tempel seluruh record array lama di sini */
];


// ============================================================
// SETUP
// ============================================================

// Jalankan sekali jika sheet Users, Data, Posts & Messages belum ada.
//
// Aman dijalankan ulang.
// Akun manager akan dibuat jika belum ada.
// Jika akun sudah ada, nama profil TIDAK ditimpa.
// Password akun yang sudah ada juga TIDAK diubah.

function setup() {

  const b = ss_();
  const properties = PropertiesService.getScriptProperties();

  if (!b.getSheetByName('Users')) {
    b.insertSheet('Users')
      .appendRow(['username', 'password', 'nama', 'aktif', 'role']);
  }

  if (!b.getSheetByName('Data')) {
    b.insertSheet('Data')
      .appendRow(['username', 'seq', 'json']);
  }

  if (!b.getSheetByName('Posts')) {
    b.insertSheet('Posts').appendRow([
      'id', 'username', 'nama', 'text', 'ts', 'imageUrl'
    ]);
  }

  const postsSheet = b.getSheetByName('Posts');
  if (postsSheet.getLastColumn() < 6) postsSheet.getRange(1, 6).setValue('imageUrl');

  if (!b.getSheetByName('PostLikes')) {
    b.insertSheet('PostLikes').appendRow(['postId', 'username', 'ts']);
  }

  if (!b.getSheetByName('PostPins')) {
    b.insertSheet('PostPins').appendRow(['postId', 'username', 'ts']);
  }

  if (!b.getSheetByName('PostComments')) {
    b.insertSheet('PostComments').appendRow([
      'id', 'postId', 'username', 'text', 'ts'
    ]);
  }

  if (!b.getSheetByName('Messages')) {
    b.insertSheet('Messages').appendRow([
      'id', 'from', 'to', 'text', 'ts'
    ]);
  }

  if (!b.getSheetByName('ChatGroups')) {
    b.insertSheet('ChatGroups').appendRow([
      'id', 'name', 'owner', 'members_json', 'ts'
    ]);
  }

  chatSummary_();
  if (properties.getProperty('CHAT_SUMMARY_VERSION') != '1') {
    rebuildChatSummary_();
    properties.setProperty('CHAT_SUMMARY_VERSION', '1');
  }

  const usersSheet = b.getSheetByName('Users');
  if (!String(usersSheet.getRange(1, 6).getValue() || '').trim()) {
    usersSheet.getRange(1, 6).setValue('photoUrl');
  }

  // Existing manager passwords are preserved. For a missing manager account,
  // configure its one-time initial password in Script Properties.
  ensureManager_(
    'kevin',
    managerInitialPassword_('INITIAL_BM_PASSWORD', 'kevin'),
    'Kevin Ricardo',
    'BM'
  );
  ensureManager_(
    'roby',
    managerInitialPassword_('INITIAL_DM_PASSWORD', 'roby'),
    'Roby',
    'DM'
  );
}

function managerInitialPassword_(propertyName, username) {
  if (findUser_(username)) return '';
  const password = PropertiesService.getScriptProperties().getProperty(propertyName);
  if (!password || password.length < 12) {
    throw new Error(
      'Set Script Property ' + propertyName +
      ' to a temporary password of at least 12 characters before setup().'
    );
  }
  return password;
}


// ============================================================
// ENSURE MANAGER
// ============================================================

// Membuat / memperbaiki akun manager.
// Nama yang sudah diedit melalui Edit Profile TIDAK ditimpa.
// Password akun yang sudah ada TIDAK diubah.

function ensureManager_(username, password, nama, role) {

  const r = findUser_(username);

  if (!r) {
    ensureUser_(username, password, nama, role);
    return;
  }

  const sh = users_();

  // Pertahankan nama yang telah diubah oleh pemilik akun.
  // Setup hanya memastikan akun aktif dan role tetap benar.
  sh.getRange(r.idx, 4).setValue('Y');
  sh.getRange(r.idx, 5).setValue(String(role).toUpperCase());
}


// ============================================================
// API
// ============================================================

const API_VERSION = '2026-10-09-team-social';

function doGet() {
  return out_({
    ok: true,
    app: 'TPFX Sales API',
    version: API_VERSION,
    endpoints: Object.keys(api_())
  });
}


function out_(o) {
  return ContentService
    .createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}


function api_() {
  return {
    login: login,
    logout: logout,
    ubahPassword: ubahPassword,
    loadState: loadState,
    saveState: saveState,
    dashboardDM: dashboardDM,
    uploadImg: uploadImg,
    listUsers: listUsers,
    createUser: createUser,
    resetBDOPassword: resetBDOPassword,
    updateBDOTarget: updateBDOTarget,
    updateProfile: updateProfile,
    updateProfilePhoto: updateProfilePhoto,
    deleteUser: deleteUser,
    databaseData: databaseData,
    teamLeads: teamLeads,
    chatNotificationSummary: chatNotificationSummary,
    checkSession: checkSession,
    socialFeed: socialFeed,
    recentPosts: recentPosts,
    createPost: createPost,
    deletePost: deletePost,
    togglePostLike: togglePostLike,
    togglePostPin: togglePostPin,
    postComments: postComments,
    addPostComment: addPostComment,
    chatContacts: chatContacts,
    chatHistory: chatHistory,
    chatSync: chatSync,
    sendChat: sendChat,
    chatGroups: chatGroups,
    createChatGroup: createChatGroup,
    deleteChatGroup: deleteChatGroup,
    groupHistory: groupHistory,
    sendGroupChat: sendGroupChat,
    globalChatHistory: globalChatHistory,
    sendGlobalChat: sendGlobalChat
  };
}


function doPost(e) {

  try {

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Request kosong');
    }

    const req = JSON.parse(e.postData.contents);
    const f = api_()[req.fn];

    if (!f) {
      throw new Error('Fungsi tidak dikenal: ' + req.fn + ' (deployment belum versi terbaru?)');
    }

    let result = f.apply(null, req.args || []);

    // Jaminan: frontend selalu menerima properti `result`
    // (undefined akan hilang saat JSON.stringify).
    if (result === undefined) result = { ok: 1 };

    return out_({ result: result });

  } catch (err) {

    return out_({ error: String(err.message || err) });

  }
}


// ============================================================
// PASSWORD / HASH
// ============================================================

function hash_(u, p) {
  return 'h:' +
    Utilities.base64Encode(
      Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        String(u).toLowerCase() + ':' + p + ':tpfx'
      )
    );
}


// ============================================================
// USER
// ============================================================

function findUser_(u) {

  const v = users_().getDataRange().getValues();

  for (let i = 1; i < v.length; i++) {

    if (
      String(v[i][0]).toLowerCase() == String(u).toLowerCase() &&
      String(v[i][3]).toUpperCase() == 'Y'
    ) {
      v[i].idx = i + 1;
      return v[i];
    }

  }

  return null;
}


function passOk_(r, pw) {

  if (!r) return false;

  const st = String(r[1]);

  return st.indexOf('h:') == 0
    ? st === hash_(r[0], pw)
    : st === String(pw);
}


function user_(t) {

  if (!t) return null;

  const p = PropertiesService.getScriptProperties();
  const r = p.getProperty('t_' + t);

  if (!r) return null;

  const parts = r.split('|');
  const u = parts[0];
  const exp = Number(parts[1]);

  if (!u || !exp || Date.now() > exp) {
    p.deleteProperty('t_' + t);
    return null;
  }

  return u;
}


function role_(t) {

  const u = user_(t);

  if (!u) return null;

  const r = findUser_(u);

  return r ? String(r[4] || 'BDO').toUpperCase() : null;
}


function ensureUser_(username, password, nama, role) {

  const existing = findUser_(username);

  if (existing) return false;

  users_().appendRow([
    String(username).trim(),
    hash_(username, password),
    String(nama).trim(),
    'Y',
    String(role).toUpperCase()
  ]);

  return true;
}


// ============================================================
// LOGIN
// ============================================================

function login(u, pw) {

  const r = findUser_(u);

  if (!r || !passOk_(r, pw)) {
    return { err: 1 };
  }

  // Kalau password lama masih plaintext, otomatis diubah menjadi hash.
  if (String(r[1]).indexOf('h:') != 0) {
    users_().getRange(r.idx, 2).setValue(hash_(r[0], pw));
  }

  const t = Utilities.getUuid();

  // Session 30 hari
  PropertiesService
    .getScriptProperties()
    .setProperty('t_' + t, r[0] + '|' + (Date.now() + DAYS * 864e5));

  return { token: t };
}


// ============================================================
// LOGOUT
// ============================================================

function logout(t) {

  if (t) {
    PropertiesService.getScriptProperties().deleteProperty('t_' + t);
  }

  return { ok: 1 };
}


// ============================================================
// UBAH PASSWORD
// ============================================================

function ubahPassword(t, lama, baru) {

  const u = user_(t);

  if (!u) return { err: 'Sesi habis, login ulang' };

  const r = findUser_(u);

  if (!r) return { err: 'Akun tidak ditemukan' };

  if (!passOk_(r, lama)) return { err: 'Password lama salah' };

  if (String(baru).length < 6) {
    return { err: 'Password baru minimal 6 karakter' };
  }

  users_().getRange(r.idx, 2).setValue(hash_(r[0], baru));

  return { ok: 1 };
}


// ============================================================
// LOAD STATE
// ============================================================

function loadState(t) {

  const u = user_(t);

  if (!u) return { err: 'AUTH' };

  // Perpanjang session otomatis setiap kali user aktif membuka data.
  PropertiesService
    .getScriptProperties()
    .setProperty('t_' + t, u + '|' + (Date.now() + DAYS * 864e5));

  const rows = data_()
    .getDataRange()
    .getValues()
    .slice(1)
    .filter(r => String(r[0]) == String(u))
    .sort((a, b) => Number(a[1]) - Number(b[1]));

  let d = '';

  try {

    d = rows.map(r => String(r[2] || '')).join('');

    const state = d ? JSON.parse(d) : null;

    const me = findUser_(u);

    if (!me) return { err: 'AUTH' };

    const targetOverride = PropertiesService.getScriptProperties()
      .getProperty(activityTargetKey_(u));

    return {
      state: state,
      name: me[2],
      user: u,
      role: String(me[4] || 'BDO').toUpperCase(),
      target: targetOverride === null
        ? (state && state.target || 10)
        : Number(targetOverride),
      photoUrl: String(me[5] || '')
    };

  } catch (e) {

    // Jangan hapus data jika JSON rusak.
    return {
      err: 'Data akun tidak valid. Data belum dihapus. Hubungi admin untuk pemulihan data.'
    };

  }
}


// ============================================================
// SAVE STATE
// ============================================================

function saveState(t, json) {

  const u = user_(t);

  if (!u) return { err: 'AUTH' };

  // ==========================================================
  // VALIDASI JSON
  // ==========================================================
  // Sangat penting: jangan menyentuh database sebelum JSON valid.

  let parsed;

  try {
    parsed = JSON.parse(String(json || ''));
  } catch (e) {
    return { err: 'JSON data tidak valid. Data lama dipertahankan.' };
  }

  if (!parsed || typeof parsed !== 'object') {
    return { err: 'Format data tidak valid. Data lama dipertahankan.' };
  }

  // ==========================================================
  // PERPANJANG SESSION
  // ==========================================================

  PropertiesService
    .getScriptProperties()
    .setProperty('t_' + t, u + '|' + (Date.now() + DAYS * 864e5));

  // ==========================================================
  // LOCK
  // ==========================================================

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {

    const s = data_();
    const v = s.getDataRange().getValues();

    // Simpan nomor baris data lama.
    const oldRows = [];

    for (let i = 1; i < v.length; i++) {
      if (String(v[i][0]) == String(u)) {
        oldRows.push(i + 1);
      }
    }

    // ========================================================
    // TULIS DATA BARU TERLEBIH DAHULU
    // ========================================================
    // Maksimum 40.000 karakter per cell agar aman terhadap
    // batas cell Google Sheets.

    const rows = [];
    // Pertahankan target yang ditetapkan manager (BM/DM).
    const targetOverride = PropertiesService.getScriptProperties()
      .getProperty(activityTargetKey_(u));

    if (targetOverride !== null) parsed.target = Number(targetOverride);

    const raw = JSON.stringify(parsed);

    for (let i = 0, n = 0; i < raw.length; i += 40000) {
      rows.push([u, n++, raw.slice(i, i + 40000)]);
    }

    if (!rows.length) {
      rows.push([u, 0, '{}']);
    }

    const startRow = s.getLastRow() + 1;

    s.getRange(startRow, 1, rows.length, 3).setValues(rows);

    // ========================================================
    // BARU SETELAH DATA BARU BERHASIL DITULIS,
    // DATA LAMA DIHAPUS
    // ========================================================

    for (let i = oldRows.length - 1; i >= 0; i--) {
      s.deleteRow(oldRows[i]);
    }

  } finally {

    lock.releaseLock();

  }

  return { ok: 1 };
}


// ============================================================
// DASHBOARD BM / DM
// ============================================================

function dashboardDM(t) {

  const u = user_(t);

  if (!u) return { err: 1 };

  const callerRole = role_(t);

  if (callerRole != 'DM' && callerRole != 'BM') {
    return { err: 1 };
  }

  const ch = {};

  data_()
    .getDataRange()
    .getValues()
    .slice(1)
    .forEach(r => {
      (ch[r[0]] = ch[r[0]] || []).push(r);
    });

  // BM dan DM menggunakan dashboard yang sama.
  // Yang ditampilkan adalah akun aktif BDO/Sales.

  const rows = users_()
    .getDataRange()
    .getValues()
    .slice(1)
    .filter(r =>
      String(r[3]).toUpperCase() == 'Y' &&
      !['BM', 'DM'].includes(String(r[4]).toUpperCase())
    )
    .map(r => {

      const j = (ch[r[0]] || [])
        .sort((a, b) => a[1] - b[1])
        .map(x => x[2])
        .join('');

      let s = {};

      try {
        s = j ? JSON.parse(j) : {};
      } catch (e) {
        s = {};
      }

      const targetOverride = PropertiesService.getScriptProperties()
        .getProperty(activityTargetKey_(r[0]));

      return {
        user: r[0],
        nama: r[2],
        state: {
          target: targetOverride === null ? (s.target || 10) : Number(targetOverride),
          leads: s.leads || [],
          acts: s.acts || []
        }
      };

    });

  return { rows: rows };
}


// ============================================================
// LIST USERS
// ============================================================

function listUsers(t) {

  const r = role_(t);

  if (r != 'BM' && r != 'DM') {
    return { err: 1 };
  }

  const rows = users_()
    .getDataRange()
    .getValues()
    .slice(1)
    .filter(x => String(x[3]).toUpperCase() == 'Y')
    .map(x => ({
      username: String(x[0]),
      nama: String(x[2]),
      role: String(x[4] || 'BDO').toUpperCase(),
      photoUrl: String(x[5] || '')
    }));

  return { rows: rows };
}


// ============================================================
// CREATE USER
// ============================================================

function createUser(t, username, password, nama, role) {

  const caller = role_(t);

  if (caller != 'BM' && caller != 'DM') {
    return { err: 'Akses ditolak' };
  }

  username = String(username || '').trim();
  password = String(password || '');
  nama = String(nama || '').trim();
  role = String(role || 'BDO').toUpperCase();

  if (!username || !nama || !password) {
    return { err: 'Data akun belum lengkap' };
  }

  if (!/^[a-zA-Z0-9._-]{3,30}$/.test(username)) {
    return { err: 'Username 3-30 karakter: huruf, angka, titik, _ atau -' };
  }

  if (password.length < 6) {
    return { err: 'Password minimal 6 karakter' };
  }

  if (!['BM', 'DM', 'BDO', 'SALES'].includes(role)) {
    return { err: 'Role tidak valid' };
  }

  if (findUser_(username)) {
    return { err: 'Username sudah digunakan' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {

    if (findUser_(username)) {
      return { err: 'Username sudah digunakan' };
    }

    users_().appendRow([
      username,
      hash_(username, password),
      nama,
      'Y',
      role
    ]);

  } finally {

    lock.releaseLock();

  }

  return { ok: 1 };
}


// ============================================================
// RESET PASSWORD & TARGET AKTIVITAS BDO (BM / DM)
// ============================================================

// Target aktivitas disimpan sebagai override di Script Properties
// agar tidak tertimpa state lama saat aplikasi BDO sinkron otomatis.

function activityTargetKey_(username) {
  return 'activity_target_' + String(username).toLowerCase();
}


// Hanya untuk akun role BDO. Semua sesi aktif target dicabut.
// Sampaikan password sementara kepada BDO secara pribadi.

function resetBDOPassword(t, username, newPassword) {

  const caller = user_(t);

  if (!caller) return { err: 'AUTH' };

  const callerRole = role_(t);

  if (callerRole != 'BM' && callerRole != 'DM') {
    return { err: 'Akses ditolak' };
  }

  username = String(username || '').trim();
  newPassword = String(newPassword || '');

  if (!username) return { err: 'Username BDO tidak valid' };

  if (newPassword.length < 6 || newPassword.length > 128) {
    return { err: 'Password harus terdiri dari 6 sampai 128 karakter' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {

    const target = findUser_(username);

    if (!target || String(target[4]).toUpperCase() != 'BDO') {
      return { err: 'Akun BDO tidak ditemukan atau tidak aktif' };
    }

    users_().getRange(target.idx, 2).setValue(hash_(target[0], newPassword));

    const properties = PropertiesService.getScriptProperties();
    const all = properties.getProperties();
    const targetUsername = String(target[0]).toLowerCase();

    Object.keys(all).forEach(key => {

      if (key.indexOf('t_') != 0) return;

      const sessionUsername = String(all[key]).split('|')[0];

      if (sessionUsername.toLowerCase() == targetUsername) {
        properties.deleteProperty(key);
      }

    });

  } finally {

    lock.releaseLock();

  }

  return { ok: 1 };
}


function updateBDOTarget(t, username, targetValue) {

  const caller = user_(t);

  if (!caller) return { err: 'AUTH' };

  const callerRole = role_(t);

  if (callerRole != 'BM' && callerRole != 'DM') {
    return { err: 'Akses ditolak' };
  }

  username = String(username || '').trim();

  const target = Number(targetValue);

  if (!username) return { err: 'Username BDO tidak valid' };

  if (!Number.isInteger(target) || target < 1 || target > 1000) {
    return { err: 'Target harus bilangan bulat antara 1 sampai 1000' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {

    const account = findUser_(username);

    if (!account || String(account[4]).toUpperCase() != 'BDO') {
      return { err: 'Akun BDO tidak ditemukan atau tidak aktif' };
    }

    PropertiesService.getScriptProperties()
      .setProperty(activityTargetKey_(account[0]), String(target));

  } finally {

    lock.releaseLock();

  }

  return { ok: 1, target: target };
}


// ============================================================
// UPDATE PROFILE
// ============================================================

// BM / DM dapat mengubah nama profil akunnya sendiri.
// Username tidak diubah.

function updateProfile(t, nama) {

  const username = user_(t);

  if (!username) return { err: 'AUTH' };

  const callerRole = role_(t);

  if (callerRole != 'BM' && callerRole != 'DM') {
    return { err: 'Akses ditolak' };
  }

  nama = String(nama || '').trim();

  if (!nama) return { err: 'Nama tidak boleh kosong' };

  if (nama.length > 100) return { err: 'Nama maksimal 100 karakter' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {

    const akun = findUser_(username);

    if (!akun) return { err: 'AUTH' };

    users_().getRange(akun.idx, 3).setValue(nama);

  } finally {

    lock.releaseLock();

  }

  return { ok: 1, nama: nama };
}

function validImageUrl_(url) {
  const value = String(url || '');
  return /^https:\/\/drive\.google\.com\/(?:thumbnail\?id=[A-Za-z0-9_-]+&sz=w\d{1,4}|uc\?export=view&id=[A-Za-z0-9_-]+)$/i
    .test(value);
}

function updateProfilePhoto(t, photoUrl) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account) return { err: 'AUTH' };

  photoUrl = String(photoUrl || '').trim();
  if (!validImageUrl_(photoUrl)) {
    return { err: 'URL foto profil tidak valid' };
  }

  users_().getRange(account.idx, 6).setValue(photoUrl);
  return { ok: 1, photoUrl: photoUrl };
}


// ============================================================
// DELETE USER
// ============================================================

// BM / DM dapat menghapus akun user lain.
// Akun yang sedang login tidak dapat menghapus dirinya sendiri.
//
// Saat akun dihapus:
// 1. Session user dihapus.
// 2. Seluruh data/activity user dihapus dari Data.
// 3. Posting dan chat user dihapus (Posts & Messages).
// 4. Akun user dihapus dari Users.

function deleteUser(t, username) {

  const caller = user_(t);

  if (!caller) return { err: 'AUTH' };

  const callerRole = role_(t);

  if (callerRole != 'BM' && callerRole != 'DM') {
    return { err: 'Akses ditolak' };
  }

  username = String(username || '').trim();

  if (!username) return { err: 'Username akun tidak valid' };

  if (username.toLowerCase() == caller.toLowerCase()) {
    return { err: 'Akun yang sedang digunakan tidak dapat dihapus' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {

    const target = findUser_(username);

    if (!target) {
      return { err: 'Akun tidak ditemukan atau sudah dihapus' };
    }

    const targetUsername = String(target[0]).toLowerCase();

    // ========================================================
    // HAPUS SESSION USER
    // ========================================================

    const scriptProperties = PropertiesService.getScriptProperties();
    const properties = scriptProperties.getProperties();

    Object.keys(properties).forEach(function(key) {

      if (key.indexOf('t_') != 0) return;

      const sessionUsername = String(properties[key]).split('|')[0];

      if (sessionUsername.toLowerCase() == targetUsername) {
        scriptProperties.deleteProperty(key);
      }

    });

    // ========================================================
    // HAPUS DATA USER
    // ========================================================

    const data = data_();
    const rows = data.getDataRange().getValues();

    // Hapus dari bawah ke atas agar indeks baris tetap valid.
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]).toLowerCase() == targetUsername) {
        data.deleteRow(i + 1);
      }
    }

    // ========================================================
    // HAPUS POSTING & CHAT USER
    // ========================================================

    purgeSocialData_(targetUsername);

    // ========================================================
    // HAPUS OVERRIDE TARGET AKTIVITAS
    // ========================================================

    PropertiesService.getScriptProperties()
      .deleteProperty(activityTargetKey_(targetUsername));

    // ========================================================
    // HAPUS USER
    // ========================================================

    users_().deleteRow(target.idx);

  } finally {

    lock.releaseLock();

  }

  return { ok: 1 };
}


// ============================================================
// UPLOAD IMAGE
// ============================================================

function uploadImg(t, dataUrl) {

  const u = user_(t);

  if (!u) throw new Error('auth');
  if (!findUser_(u)) throw new Error('auth');

  const it = DriveApp.getFoldersByName('TPFX Bukti Aktivitas');

  const folder = it.hasNext()
    ? it.next()
    : DriveApp.createFolder('TPFX Bukti Aktivitas');

  const raw = String(dataUrl || '');

  if (raw.length > 8 * 1024 * 1024) {
    throw new Error('Ukuran gambar terlalu besar');
  }

  const match = raw.match(/^data:image\/(?:jpeg|png|webp);base64,([A-Za-z0-9+/]+={0,2})$/i);
  if (!match) {
    throw new Error('Format gambar tidak valid');
  }

  const f = folder.createFile(
    Utilities.newBlob(
      Utilities.base64Decode(match[1]),
      'image/jpeg',
      u + '_' + Date.now() + '.jpg'
    )
  );

  f.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  return 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1200';
}


// ============================================================
// DATABASE & CEK SESI
// ============================================================

// Kedua endpoint memvalidasi token melalui loadState;
// database hanya dikirim kepada akun dengan sesi aktif.
// JANGAN membuat endpoint publik untuk database ini.

// Menggabungkan semua kontak dari DATABASE_DATA, lalu membaginya bergiliran
// ke SETIAP akun aktif (aktif = 'Y') dengan role BDO di sheet Users.
// Akun BDO baru otomatis ikut terbagi. Jumlah kontak tiap BDO berbeda
// paling banyak satu. Pembagian dibatalkan bila kontak kurang dari jumlah
// BDO, atau ada nomor kosong/duplikat.

function databaseData(t) {

  const session = loadState(t);

  if (!session || session.err) return { err: 'AUTH' };

  const activeBDOs = users_().getDataRange().getValues().slice(1)
    .filter(r =>
      String(r[3]).toUpperCase() == 'Y' &&
      String(r[4]).toUpperCase() == 'BDO'
    )
    .map(r => ({
      username: String(r[0]),
      name: String(r[2] || r[0]),
      rows: []
    }));

  if (!activeBDOs.length) {
    return { err: 'Belum ada akun BDO aktif untuk pembagian kontak' };
  }

  const contacts = [];

  DATABASE_DATA.forEach(group => {

    if (!group || !Array.isArray(group.r)) {
      throw new Error('Format data database lama tidak valid');
    }

    group.r.forEach(row => {

      if (!Array.isArray(row) || row.length < 3) {
        throw new Error('Format kontak database lama tidak valid');
      }

      contacts.push(row);

    });

  });

  if (contacts.length < activeBDOs.length) {
    return {
      err: 'Jumlah kontak (' + contacts.length +
        ') lebih sedikit daripada BDO aktif (' + activeBDOs.length +
        '); semua BDO belum dapat menerima minimal satu kontak'
    };
  }

  const phoneNumbers = {};

  contacts.forEach(contact => {

    const phone = String(contact[2] || '').replace(/\D/g, '');

    if (!phone) throw new Error('Kontak database tidak memiliki nomor telepon');

    if (phoneNumbers[phone]) {
      throw new Error('Nomor kontak duplikat ditemukan; pembagian dibatalkan');
    }

    phoneNumbers[phone] = true;

  });

  contacts.forEach((contact, index) => {
    activeBDOs[index % activeBDOs.length].rows.push(contact);
  });

  return {
    rows: activeBDOs.map(bdo => ({
      n: bdo.name,
      r: bdo.rows
    }))
  };
}

function teamLeads(t) {
  const username = user_(t);
  const caller = username && findUser_(username);
  if (!caller) return { err: 'AUTH' };

  const callerRole = String(caller[4] || '').toUpperCase();
  if (callerRole != 'BDO' && callerRole != 'SALES') {
    return { err: 'Akses ditolak' };
  }

  const owners = {};
  users_().getDataRange().getValues().slice(1).forEach(r => {
    const role = String(r[4] || '').toUpperCase();
    if (String(r[3]).toUpperCase() == 'Y' &&
        (role == 'BDO' || role == 'SALES')) {
      owners[String(r[0]).toLowerCase()] = String(r[2] || r[0]);
    }
  });

  const chunks = {};
  data_().getDataRange().getValues().slice(1).forEach(r => {
    const key = String(r[0] || '').toLowerCase();
    if (owners[key]) (chunks[key] = chunks[key] || []).push(r);
  });

  const rows = [];
  Object.keys(chunks).forEach(key => {
    const savedRows = chunks[key].sort((a, b) => Number(a[1]) - Number(b[1]));
    let state;
    try {
      state = JSON.parse(savedRows.map(r => String(r[2] || '')).join(''));
    } catch (e) {
      throw new Error('Data lead tidak valid untuk akun ' + owners[key]);
    }
    if (!state || !Array.isArray(state.leads)) {
      throw new Error('Format daftar lead tidak valid untuk akun ' + owners[key]);
    }

    state.leads.forEach(lead => {
      if (!lead || typeof lead != 'object') {
        throw new Error('Format lead tidak valid untuk akun ' + owners[key]);
      }
      rows.push({
        id: lead.id,
        name: String(lead.name || ''),
        phone: String(lead.phone || ''),
        product: String(lead.product || ''),
        status: String(lead.status || 'new'),
        interest: String(lead.interest || ''),
        last: String(lead.last || ''),
        next: String(lead.next || ''),
        notes: String(lead.notes || ''),
        c: String(lead.c || ''),
        owner: owners[key]
      });
    });
  });

  rows.sort((a, b) => String(b.c).localeCompare(String(a.c)));
  return { rows: rows };
}


function checkSession(t) {

  const session = loadState(t);

  if (!session || session.err) {
    return { err: 'AUTH' };
  }

  return { ok: 1 };
}


// ============================================================
// SOCIAL FEED
// ============================================================

function socialUsers_() {
  const users = {};
  users_().getDataRange().getValues().slice(1).forEach(r => {
    const key = String(r[0] || '').toLowerCase();
    if (key && String(r[3]).toUpperCase() == 'Y') {
      users[key] = {
        nama: String(r[2] || r[0]),
        role: String(r[4] || 'BDO').toUpperCase(),
        photoUrl: String(r[5] || '')
      };
    }
  });
  return users;
}

function postData_(limit) {
  const sheet = posts_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const startRow = limit ? Math.max(2, lastRow - limit + 1) : 2;
  const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, 6).getValues();
  const selected = {};
  rows.forEach(r => { selected[String(r[0])] = r; });

  if (limit) {
    postPins_().getDataRange().getValues().slice(1).forEach(pin => {
      const id = String(pin[0] || '');
      if (!id || selected[id]) return;
      const match = sheet.getRange(2, 1, lastRow - 1, 1)
        .createTextFinder(id).matchEntireCell(true).findNext();
      if (match) selected[id] = sheet.getRange(match.getRow(), 1, 1, 6).getValues()[0];
    });
  }

  return Object.keys(selected).map(id => selected[id])
    .filter(r => String(r[3] || '').trim() || String(r[5] || '').trim())
    .map(r => ({
      id: String(r[0]),
      username: String(r[1]),
      nama: String(r[2]),
      text: String(r[3] || ''),
      ts: Number(r[4]) || 0,
      imageUrl: String(r[5] || ''),
      kind: 'post'
    }));
}

function enrichPosts_(rows, viewer) {
  const likes = {};
  const liked = {};
  postLikes_().getDataRange().getValues().slice(1).forEach(r => {
    const postId = String(r[0] || '');
    if (!postId) return;
    likes[postId] = (likes[postId] || 0) + 1;
    if (String(r[1] || '').toLowerCase() == String(viewer).toLowerCase()) {
      liked[postId] = true;
    }
  });

  const pins = {};
  postPins_().getDataRange().getValues().slice(1).forEach(r => {
    const postId = String(r[0] || '');
    if (postId) pins[postId] = true;
  });

  const comments = {};
  postComments_().getDataRange().getValues().slice(1).forEach(r => {
    const postId = String(r[1] || '');
    if (postId) comments[postId] = (comments[postId] || 0) + 1;
  });

  const profiles = socialUsers_();
  return rows.map(post => {
    const profile = profiles[String(post.username).toLowerCase()] || {};
    return Object.assign({}, post, {
      nama: post.nama || profile.nama || post.username,
      role: profile.role || '',
      profilePhoto: profile.photoUrl || '',
      likes: likes[post.id] || 0,
      liked: !!liked[post.id],
      pinned: !!pins[post.id],
      commentsCount: comments[post.id] || 0
    });
  }).sort((a, b) =>
    Number(b.pinned) - Number(a.pinned) || Number(b.ts) - Number(a.ts)
  );
}

function socialFeed(t) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account) return { err: 'AUTH' };

  const feed = enrichPosts_(postData_(100), username);
  const activeUsers = socialUsers_();
  const chunks = {};
  data_().getDataRange().getValues().slice(1).forEach(r => {
    const key = String(r[0] || '').toLowerCase();
    if (activeUsers[key]) (chunks[key] = chunks[key] || []).push(r);
  });

  const labels = {
    call: 'Mencatat panggilan',
    wa: 'Menghubungi calon nasabah',
    meet: 'Melakukan pertemuan',
    post: 'Mencatat aktivitas promosi',
    shot: 'Mengunggah bukti aktivitas'
  };
  const since = Date.now() - 30 * 864e5;

  Object.keys(chunks).forEach(key => {
    const savedRows = chunks[key].sort((a, b) => Number(a[1]) - Number(b[1]));
    let state;
    try {
      state = JSON.parse(savedRows.map(r => String(r[2] || '')).join(''));
    } catch (e) {
      Logger.log('Feed melewati state JSON tidak valid untuk ' + key);
      return;
    }

    (state.acts || []).forEach(activity => {
      const ts = Number(activity.ts) || 0;
      if (!ts || ts < since) return;
      const profile = activeUsers[key];
      feed.push({
        id: 'activity-' + key + '-' + String(activity.id || ts),
        username: savedRows[0][0],
        nama: profile.nama,
        role: profile.role,
        profilePhoto: profile.photoUrl,
        text: labels[activity.type] || 'Mencatat aktivitas penjualan',
        ts: ts,
        kind: 'activity'
      });
    });
  });

  feed.sort((a, b) =>
    Number(b.pinned) - Number(a.pinned) || Number(b.ts) - Number(a.ts)
  );
  return { rows: feed.slice(0, 100) };
}

function recentPosts(t) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account) return { err: 'AUTH' };
  return { rows: enrichPosts_(postData_(100), username) };
}

function createPost(t, text, imageUrl) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account) return { err: 'AUTH' };

  text = String(text || '').trim();
  imageUrl = String(imageUrl || '').trim();
  if (!text && !imageUrl) return { err: 'Isi postingan atau foto wajib diisi' };
  if (text.length > 500) return { err: 'Postingan maksimal 500 karakter' };
  if (imageUrl && !validImageUrl_(imageUrl)) {
    return { err: 'URL foto postingan tidak valid' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    posts_().appendRow([
      Utilities.getUuid(), String(account[0]), String(account[2]),
      text, Date.now(), imageUrl
    ]);
  } finally {
    lock.releaseLock();
  }
  return { ok: 1 };
}

function deletePost(t, postId) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account) return { err: 'AUTH' };
  postId = String(postId || '').trim();
  if (!postId) return { err: 'ID postingan tidak valid' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = posts_();
    const values = sheet.getDataRange().getValues();
    const index = values.findIndex((r, i) => i > 0 && String(r[0]) == postId);
    if (index < 1) return { err: 'Postingan tidak ditemukan' };
    if (String(values[index][1]).toLowerCase() != String(username).toLowerCase()) {
      return { err: 'Anda hanya dapat menghapus postingan sendiri' };
    }
    sheet.deleteRow(index + 1);
    deletePostMetadata_(postId);
  } finally {
    lock.releaseLock();
  }
  return { ok: 1 };
}

function deletePostMetadata_(postId) {
  [
    { sheet: postLikes_(), postColumn: 1 },
    { sheet: postPins_(), postColumn: 1 },
    { sheet: postComments_(), postColumn: 2 }
  ].forEach(item => {
    const values = item.sheet.getDataRange().getValues();
    for (let i = values.length - 1; i >= 1; i--) {
      if (String(values[i][item.postColumn - 1]) == String(postId)) {
        item.sheet.deleteRow(i + 1);
      }
    }
  });
}

function togglePostLike(t, postId) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account) return { err: 'AUTH' };
  postId = String(postId || '').trim();
  if (!postData_().some(post => post.id == postId)) {
    return { err: 'Postingan tidak ditemukan' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = postLikes_();
    const rows = sheet.getDataRange().getValues();
    const index = rows.findIndex((r, i) =>
      i > 0 && String(r[0]) == postId &&
      String(r[1]).toLowerCase() == String(username).toLowerCase()
    );
    if (index > 0) {
      sheet.deleteRow(index + 1);
      return { ok: 1, liked: false };
    }
    sheet.appendRow([postId, username, Date.now()]);
    return { ok: 1, liked: true };
  } finally {
    lock.releaseLock();
  }
}

function togglePostPin(t, postId) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account) return { err: 'AUTH' };
  if (!['BM', 'DM'].includes(String(account[4]).toUpperCase())) {
    return { err: 'Hanya BM/DM yang dapat menyematkan postingan' };
  }
  postId = String(postId || '').trim();
  if (!postData_().some(post => post.id == postId)) {
    return { err: 'Postingan tidak ditemukan' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = postPins_();
    const rows = sheet.getDataRange().getValues();
    const index = rows.findIndex((r, i) => i > 0 && String(r[0]) == postId);
    if (index > 0) {
      sheet.deleteRow(index + 1);
      return { ok: 1, pinned: false };
    }
    sheet.appendRow([postId, username, Date.now()]);
    return { ok: 1, pinned: true };
  } finally {
    lock.releaseLock();
  }
}

function postComments(t, postId) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account) return { err: 'AUTH' };
  postId = String(postId || '').trim();
  if (!postData_().some(post => post.id == postId)) {
    return { err: 'Postingan tidak ditemukan' };
  }

  const profiles = socialUsers_();
  const rows = postComments_().getDataRange().getValues().slice(1)
    .filter(r => String(r[1]) == postId)
    .slice(-100)
    .map(r => {
      const profile = profiles[String(r[2]).toLowerCase()] || {};
      return {
        id: String(r[0]),
        username: String(r[2]),
        nama: profile.nama || String(r[2]),
        role: profile.role || '',
        photoUrl: profile.photoUrl || '',
        text: String(r[3]),
        ts: Number(r[4]) || 0
      };
    });
  return { rows: rows };
}

function addPostComment(t, postId, text) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account) return { err: 'AUTH' };
  postId = String(postId || '').trim();
  text = String(text || '').trim();
  if (!text) return { err: 'Komentar tidak boleh kosong' };
  if (text.length > 500) return { err: 'Komentar maksimal 500 karakter' };
  if (!postData_().some(post => post.id == postId)) {
    return { err: 'Postingan tidak ditemukan' };
  }
  postComments_().appendRow([
    Utilities.getUuid(), postId, username, text, Date.now()
  ]);
  return { ok: 1 };
}


// ============================================================
// CHAT
// ============================================================

function conversationMatches_(row, type, target, username) {
  const from = String(row[1] || '').toLowerCase();
  const to = String(row[2] || '').toLowerCase();
  if (type == 'global') return to == 'global';
  if (type == 'group') return to == ('group:' + target).toLowerCase();
  return (from == username.toLowerCase() && to == target.toLowerCase()) ||
    (from == target.toLowerCase() && to == username.toLowerCase());
}

function chatMessageObject_(row, rowNumber, profiles, type) {
  const from = String(row[1] || '');
  const profile = profiles[from.toLowerCase()] || {};
  const message = {
    id: String(row[0]),
    from: from,
    nama: profile.nama || from,
    role: profile.role || '',
    photoUrl: profile.photoUrl || '',
    text: String(row[3] || ''),
    ts: Number(row[4]) || 0,
    row: rowNumber
  };
  if (type == 'personal') message.to = String(row[2] || '');
  return message;
}

function conversationPage_(type, target, username, beforeRow, profiles, pageSize) {
  const messages = messages_();
  const cursor = messages.getLastRow();
  const before = Number(beforeRow) > 1
    ? Math.min(cursor + 1, Math.floor(Number(beforeRow)))
    : cursor + 1;
  let end = before - 1;
  const found = [];
  const blockSize = 5000;
  while (end >= 2 && found.length <= pageSize) {
    const start = Math.max(2, end - blockSize + 1);
    const values = messages.getRange(start, 1, end - start + 1, 5).getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      if (conversationMatches_(values[i], type, target, username)) {
        found.push(chatMessageObject_(values[i], start + i, profiles, type));
        if (found.length > pageSize) break;
      }
    }
    end = start - 1;
  }
  const hasMore = found.length > pageSize;
  const rows = found.slice(0, pageSize).reverse();
  return { rows: rows, cursor: cursor, hasMore: hasMore };
}

function chatSync(t, type, target, afterRow) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account) return { err: 'AUTH' };

  type = String(type || '');
  target = String(target || '').trim();
  if (type == 'personal') {
    if (!target || target.toLowerCase() == username.toLowerCase() || !findUser_(target)) {
      return { err: 'Akun chat tidak ditemukan' };
    }
  } else if (type == 'group') {
    const group = chatGroups_().getDataRange().getValues().slice(1)
      .find(row => String(row[0]) == target);
    if (!group) return { err: 'Grup tidak ditemukan' };
    if (groupMembers_(group).indexOf(String(username).toLowerCase()) < 0) {
      return { err: 'Anda bukan anggota grup ini' };
    }
  } else if (type != 'global') {
    return { err: 'Jenis chat tidak valid' };
  }

  const messages = messages_();
  const lastRow = messages.getLastRow();
  const cursor = Math.max(1, Math.floor(Number(afterRow) || 1));
  if (cursor >= lastRow) return { rows: [], cursor: lastRow };

  const start = Math.max(2, cursor + 1);
  const end = Math.min(lastRow, start + 999);
  const values = messages.getRange(start, 1, end - start + 1, 5).getValues();
  const profiles = socialUsers_();
  const rows = [];
  values.forEach((row, index) => {
    if (conversationMatches_(row, type, target, username)) {
      rows.push(chatMessageObject_(row, start + index, profiles, type));
    }
  });
  return { rows: rows, cursor: end, hasMore: end < lastRow };
}

function chatContacts(t) {

  const username = user_(t);

  if (!username || !findUser_(username)) return { err: 'AUTH' };

  const contacts = {};

  users_().getDataRange().getValues().slice(1).forEach(r => {

    const other = String(r[0]);

    if (
      String(r[3]).toUpperCase() != 'Y' ||
      other.toLowerCase() == username.toLowerCase()
    ) return;

    contacts[other.toLowerCase()] = {
      username: other,
      nama: String(r[2] || other),
      role: String(r[4] || 'BDO').toUpperCase(),
      photoUrl: String(r[5] || ''),
      lastText: '',
      lastTs: 0,
      lastFrom: ''
    };

  });

  chatSummaryRows_().forEach(r => {
    if (String(r[1]) != 'personal') return;
    const key = String(r[0]);
    const match = key.slice('personal:'.length).split('|');
    if (match.length != 2) return;
    const other = match[0] == username.toLowerCase() ? match[1] :
      match[1] == username.toLowerCase() ? match[0] : '';
    const contact = contacts[other];
    if (!contact) return;
    contact.lastText = String(r[3] || '');
    contact.lastTs = Number(r[4]) || 0;
    contact.lastFrom = String(r[5] || '');
  });

  return {
    rows: Object.keys(contacts)
      .map(k => contacts[k])
      .sort((a, b) => b.lastTs - a.lastTs || a.nama.localeCompare(b.nama))
  };
}


function chatHistory(t, peer, beforeRow) {

  const username = user_(t);

  if (!username || !findUser_(username)) return { err: 'AUTH' };
  const profiles = socialUsers_();

  peer = String(peer || '').trim();

  if (
    !peer ||
    peer.toLowerCase() == username.toLowerCase() ||
    !findUser_(peer)
  ) return { err: 'Akun chat tidak ditemukan' };

  return conversationPage_(
    'personal', peer, username, beforeRow, profiles, 50
  );
}


function sendChat(t, peer, text) {

  const username = user_(t);

  if (!username || !findUser_(username)) return { err: 'AUTH' };

  peer = String(peer || '').trim();

  if (
    !peer ||
    peer.toLowerCase() == username.toLowerCase() ||
    !findUser_(peer)
  ) return { err: 'Akun chat tidak ditemukan' };

  text = String(text || '').trim();

  if (!text) return { err: 'Pesan tidak boleh kosong' };

  if (text.length > 1000) return { err: 'Pesan maksimal 1000 karakter' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {

    const messages = messages_();
    const ts = Date.now();
    messages.appendRow([
      Utilities.getUuid(),
      username,
      peer,
      text,
      ts
    ]);
    updateChatSummary_('personal', peer, username, text, ts, messages.getLastRow());

  } finally {

    lock.releaseLock();

  }

  return { ok: 1 };
}


// ============================================================
// CHAT GRUP
// ============================================================

function groupMembers_(row) {
  const members = JSON.parse(String(row[3] || '[]'));
  if (!Array.isArray(members)) throw new Error('Daftar anggota grup tidak valid.');
  return members.map(name => String(name).toLowerCase());
}


function chatGroups(t) {

  const username = user_(t);
  const account = username && findUser_(username);

  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };

  const lastMessages = {};
  chatSummaryRows_().forEach(r => {
    if (String(r[1]) != 'group') return;
    lastMessages[String(r[2])] = {
      text: String(r[3] || ''),
      ts: Number(r[4]) || 0,
      from: String(r[5] || '')
    };
  });

  const rows = chatGroups_().getDataRange().getValues().slice(1)
    .filter(r => groupMembers_(r).indexOf(String(username).toLowerCase()) >= 0)
    .map(r => {
      const id = String(r[0]);
      const members = groupMembers_(r);
      const last = lastMessages[id] || {};
      return {
        id: id,
        name: String(r[1]),
        owner: String(r[2] || ''),
        memberCount: members.length,
        lastText: String(last.text || ''),
        lastTs: Number(last.ts) || 0,
        lastFrom: String(last.from || '')
      };
    })
    .sort((a, b) => b.lastTs - a.lastTs || a.name.localeCompare(b.name));

  return { rows: rows };
}


function createChatGroup(t, name, selectedMembers) {

  const username = user_(t);
  const account = username && findUser_(username);

  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };

  name = String(name || '').trim();

  if (!name) return { err: 'Nama grup tidak boleh kosong' };
  if (name.length > 60) return { err: 'Nama grup maksimal 60 karakter' };
  if (!Array.isArray(selectedMembers)) return { err: 'Daftar anggota grup tidak valid' };

  const active = {};

  users_().getDataRange().getValues().slice(1).forEach(r => {
    if (String(r[3]).toUpperCase() == 'Y') {
      active[String(r[0]).toLowerCase()] = String(r[0]);
    }
  });

  const members = [String(username)];

  selectedMembers.forEach(value => {
    const key = String(value || '').trim().toLowerCase();
    if (!key || key == String(username).toLowerCase()) return;
    if (!active[key]) throw new Error('Anggota grup tidak aktif atau tidak ditemukan: ' + key);
    if (members.map(x => x.toLowerCase()).indexOf(key) < 0) members.push(active[key]);
  });

  if (members.length < 2) return { err: 'Pilih minimal satu anggota lain untuk grup' };
  if (members.length > 50) return { err: 'Grup maksimal 50 anggota' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {

    const id = Utilities.getUuid();

    chatGroups_().appendRow([id, name, username, JSON.stringify(members), Date.now()]);

    return { ok: 1, groupId: id };

  } finally {

    lock.releaseLock();

  }
}


function groupHistory(t, groupId, beforeRow) {

  const username = user_(t);
  const account = username && findUser_(username);

  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };

  groupId = String(groupId || '').trim();

  const group = chatGroups_().getDataRange().getValues().slice(1)
    .find(r => String(r[0]) == groupId);

  if (!group) return { err: 'Grup tidak ditemukan' };

  if (groupMembers_(group).indexOf(String(username).toLowerCase()) < 0) {
    return { err: 'Anda bukan anggota grup ini' };
  }

  return conversationPage_(
    'group', groupId, username, beforeRow, socialUsers_(), 50
  );
}


function sendGroupChat(t, groupId, text) {

  const username = user_(t);
  const account = username && findUser_(username);

  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };

  groupId = String(groupId || '').trim();

  const group = chatGroups_().getDataRange().getValues().slice(1)
    .find(r => String(r[0]) == groupId);

  if (!group) return { err: 'Grup tidak ditemukan' };

  if (groupMembers_(group).indexOf(String(username).toLowerCase()) < 0) {
    return { err: 'Anda bukan anggota grup ini' };
  }

  text = String(text || '').trim();

  if (!text) return { err: 'Pesan tidak boleh kosong' };
  if (text.length > 1000) return { err: 'Pesan maksimal 1000 karakter' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {

    const messages = messages_();
    const ts = Date.now();
    messages.appendRow([
      Utilities.getUuid(), username, 'group:' + groupId, text, ts
    ]);
    updateChatSummary_('group', groupId, username, text, ts, messages.getLastRow());

  } finally {

    lock.releaseLock();

  }

  return { ok: 1 };
}

function deleteChatGroup(t, groupId) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account) return { err: 'AUTH' };
  groupId = String(groupId || '').trim();
  if (!groupId) return { err: 'ID grup tidak valid' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const groups = chatGroups_();
    const rows = groups.getDataRange().getValues();
    const index = rows.findIndex((r, i) => i > 0 && String(r[0]) == groupId);
    if (index < 1) return { err: 'Grup tidak ditemukan' };
    const owner = String(rows[index][2] || '');
    if (owner.toLowerCase() != String(username).toLowerCase() &&
        !['BM', 'DM'].includes(String(account[4]).toUpperCase())) {
      return { err: 'Hanya pembuat grup atau BM/DM yang dapat menghapus grup' };
    }

    const messages = messages_();
    const messageRows = messages.getDataRange().getValues();
    for (let i = messageRows.length - 1; i >= 1; i--) {
      if (String(messageRows[i][2]) == 'group:' + groupId) {
        redactMessageRow_(messages, i + 1);
      }
    }
    deleteChatSummary_('group', groupId, groupId);
    groups.deleteRow(index + 1);
  } finally {
    lock.releaseLock();
  }
  return { ok: 1 };
}

function globalChatHistory(t, unusedTarget, beforeRow) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account) return { err: 'AUTH' };

  return conversationPage_(
    'global', 'global', username, beforeRow, socialUsers_(), 50
  );
}

function sendGlobalChat(t, unusedTarget, text) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account) return { err: 'AUTH' };

  text = String(text || '').trim();
  if (!text) return { err: 'Pesan tidak boleh kosong' };
  if (text.length > 1000) return { err: 'Pesan maksimal 1000 karakter' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const messages = messages_();
    const ts = Date.now();
    messages.appendRow([
      Utilities.getUuid(), username, 'global', text, ts
    ]);
    updateChatSummary_('global', 'global', username, text, ts, messages.getLastRow());
  } finally {
    lock.releaseLock();
  }
  return { ok: 1 };
}


// Ringkasan chat personal + grup dalam satu request (untuk notifikasi browser).
// lastFrom dipakai frontend untuk membedakan pesan masuk dari pesan sendiri.

function chatNotificationSummary(t) {

  const username = user_(t);
  const account = username && findUser_(username);

  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };

  const userKey = String(username).toLowerCase();
  const profiles = socialUsers_();

  const contacts = {};

  users_().getDataRange().getValues().slice(1).forEach(r => {

    const peer = String(r[0] || '');

    if (String(r[3]).toUpperCase() != 'Y' || peer.toLowerCase() == userKey) return;

    contacts[peer.toLowerCase()] = {
      username: peer,
      nama: profiles[peer.toLowerCase()].nama || peer,
      role: profiles[peer.toLowerCase()].role || 'BDO',
      photoUrl: profiles[peer.toLowerCase()].photoUrl || '',
      lastText: '',
      lastTs: 0,
      lastFrom: ''
    };

  });

  const groups = {};

  chatGroups_().getDataRange().getValues().slice(1).forEach(r => {

    const id = String(r[0]);
    const members = groupMembers_(r);

    if (members.indexOf(userKey) < 0) return;

    groups[id] = {
      id: id,
      name: String(r[1]),
      owner: String(r[2] || ''),
      memberCount: members.length,
      lastText: '',
      lastTs: 0,
      lastFrom: ''
    };

  });

  const global = {
    id: 'global',
    lastText: '',
    lastTs: 0,
    lastFrom: ''
  };

  chatSummaryRows_().forEach(r => {
    const type = String(r[1] || '');
    const target = String(r[2] || '');
    const from = String(r[5] || '');
    if (type == 'global') {
      global.lastText = String(r[3] || '');
      global.lastTs = Number(r[4]) || 0;
      global.lastFrom = from;
    } else if (type == 'group') {
      const group = groups[target];
      if (group) {
        group.lastText = String(r[3] || '');
        group.lastTs = Number(r[4]) || 0;
        group.lastFrom = from;
      }
    } else if (type == 'personal') {
      const match = String(r[0]).slice('personal:'.length).split('|');
      const peer = match[0] == userKey ? match[1] :
        match[1] == userKey ? match[0] : '';
      const contact = contacts[peer];
      if (contact) {
        contact.lastText = String(r[3] || '');
        contact.lastTs = Number(r[4]) || 0;
        contact.lastFrom = from;
      }
    }
  });

  return {
    contacts: Object.keys(contacts).map(key => contacts[key])
      .sort((a, b) => b.lastTs - a.lastTs || a.nama.localeCompare(b.nama)),
    groups: Object.keys(groups).map(key => groups[key])
      .sort((a, b) => b.lastTs - a.lastTs || a.name.localeCompare(b.name)),
    global: global
  };
}


// ============================================================
// PURGE DATA SOSIAL (dipanggil dari deleteUser)
// ============================================================

// Menghapus posting & pesan milik akun, mengeluarkan akun dari grup,
// dan menghapus grup yang tersisa kurang dari dua anggota.

function purgeSocialData_(username) {

  const key = String(username).toLowerCase();

  const posts = posts_();
  const postRows = posts.getDataRange().getValues();
  const ownedPostIds = {};
  postRows.slice(1).forEach(r => {
    if (String(r[1]).toLowerCase() == key) ownedPostIds[String(r[0])] = true;
  });

  for (let i = postRows.length - 1; i >= 1; i--) {
    if (String(postRows[i][1]).toLowerCase() == key) posts.deleteRow(i + 1);
  }

  [
    { sheet: postLikes_(), postColumn: 1, userColumn: 2 },
    { sheet: postPins_(), postColumn: 1, userColumn: 2 },
    { sheet: postComments_(), postColumn: 2, userColumn: 3 }
  ].forEach(item => {
    const values = item.sheet.getDataRange().getValues();
    for (let i = values.length - 1; i >= 1; i--) {
      const authoredByDeletedUser =
        String(values[i][item.userColumn - 1]).toLowerCase() == key;
      const attachedToDeletedPost =
        !!ownedPostIds[String(values[i][item.postColumn - 1])];
      if (authoredByDeletedUser || attachedToDeletedPost) {
        item.sheet.deleteRow(i + 1);
      }
    }
  });

  const messages = messages_();
  const messageRows = messages.getDataRange().getValues();

  for (let i = messageRows.length - 1; i >= 1; i--) {
    if (
      String(messageRows[i][1]).toLowerCase() == key ||
      String(messageRows[i][2]).toLowerCase() == key
    ) redactMessageRow_(messages, i + 1);
  }

  const groups = chatGroups_();
  const groupRows = groups.getDataRange().getValues();

  for (let i = groupRows.length - 1; i >= 1; i--) {

    const members = groupMembers_(groupRows[i]).filter(member => member != key);

    if (members.length < 2) {

      const groupId = 'group:' + String(groupRows[i][0]);
      const currentMessages = messages.getDataRange().getValues();

      for (let j = currentMessages.length - 1; j >= 1; j--) {
        if (String(currentMessages[j][2]) == groupId) {
          redactMessageRow_(messages, j + 1);
        }
      }

      groups.deleteRow(i + 1);

    } else {

      groups.getRange(i + 1, 4).setValue(JSON.stringify(members));

      if (String(groupRows[i][2]).toLowerCase() == key) {
        groups.getRange(i + 1, 3).setValue(members[0]);
      }

    }

  }

  rebuildChatSummary_();
}