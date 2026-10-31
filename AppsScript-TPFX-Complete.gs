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

// Script Properties punya kuota total ~500 KB. Token login lama menumpuk
// karena tidak selalu logout, sehingga penyimpanan bisa penuh dan upload gagal
// saat sesi hendak diperpanjang. Simpan beberapa sesi terbaru per akun saja.
const SESSION_KEEP = 8;

function touchSession_(t, u) {
  if (!t || !u) return;
  const props = PropertiesService.getScriptProperties();
  const value = u + '|' + (Date.now() + DAYS * 864e5);
  try {
    props.setProperty('t_' + t, value);
    return;
  } catch (e) {
    // Penyimpanan properti penuh: bersihkan sesi lama lalu coba sekali lagi.
  }
  try {
    pruneSessions_();
    props.setProperty('t_' + t, value);
  } catch (e) {
    // Token saat ini tetap valid walau perpanjangan sesi gagal.
  }
}

function createSession_(u) {
  const props = PropertiesService.getScriptProperties();
  const t = Utilities.getUuid();
  const value = u + '|' + (Date.now() + DAYS * 864e5);
  try {
    props.setProperty('t_' + t, value);
  } catch (e) {
    pruneSessions_();
    try {
      props.setProperty('t_' + t, value);
    } catch (e2) {
      throw new Error(
        'Penyimpanan sesi penuh. Minta manager membuka database.html dan ' +
        'menekan "Rapikan database", lalu coba login lagi.'
      );
    }
  }
  return t;
}

function pruneSessions_() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const now = Date.now();
  const active = {};
  let removed = 0;

  Object.keys(all).forEach(key => {
    if (key.indexOf('t_') !== 0) return;
    const raw = String(all[key] || '');
    const cut = raw.lastIndexOf('|');
    const user = cut < 0 ? '' : raw.slice(0, cut);
    const exp = Number(raw.slice(cut + 1));
    if (!user || !exp || exp <= now) {
      props.deleteProperty(key);
      removed++;
      return;
    }
    (active[user] = active[user] || []).push({ key: key, exp: exp });
  });

  Object.keys(active).forEach(user => {
    const list = active[user].sort((a, b) => b.exp - a.exp);
    list.slice(SESSION_KEEP).forEach(item => {
      props.deleteProperty(item.key);
      removed++;
    });
  });

  let remaining = 0;
  Object.keys(active).forEach(user => {
    remaining += Math.min(active[user].length, SESSION_KEEP);
  });

  return { removed: removed, remaining: remaining };
}


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
  pushSubs_();
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

const API_VERSION = '2026-10-29-storage-maintenance';

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
    importContacts: importContacts,
    setContactStatus: setContactStatus,
    addContactToPipeline: addContactToPipeline,
    storageReport: storageReport,
    storageCleanup: storageCleanup,
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
    sendGlobalChat: sendGlobalChat,
    pushConfig: pushConfig,
    subscribePush: subscribePush,
    unsubscribePush: unsubscribePush,
    testPush: testPush
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

  const t = createSession_(r[0]);

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
  touchSession_(t, u);

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

  touchSession_(t, u);

  // ==========================================================
  // LOCK
  // ==========================================================

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {

    try {

      writeStateRows_(u, parsed);

    } catch (e) {

      // Data baru sudah tertulis; yang gagal hanya menghapus baris lama.
      // Jangan tulis ulang supaya tidak menambah salinan.
      if (e && e.saved) {
        try { recoverStorage_(); } catch (e2) {}
        return { ok: 1 };
      }

      if (!isQuotaError_(e)) throw storageError_(e);

      // Penyimpanan penuh: rapikan dulu, lalu ulangi sekali.
      try { recoverStorage_(); } catch (e2) {}

      try {
        writeStateRows_(u, parsed);
      } catch (e3) {
        throw storageError_(e3);
      }

    }

  } finally {

    lock.releaseLock();

  }

  return { ok: 1 };
}

function writeStateRows_(u, parsed) {

  let saved = false;

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

    saved = true;

    // ========================================================
    // BARU SETELAH DATA BARU BERHASIL DITULIS,
    // DATA LAMA DIHAPUS
    // ========================================================

    for (let i = oldRows.length - 1; i >= 0; i--) {
      s.deleteRow(oldRows[i]);
    }

  } catch (e) {

    if (saved && e && typeof e === 'object') e.saved = true;

    throw e;

  }
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

  // Upload bisa menjadi aktivitas pertama setelah app lama dibiarkan terbuka,
  // jadi sesi diperpanjang di sini juga (tanpa menggagalkan upload bila penuh).
  touchSession_(t, u);

  let folder;

  try {
    const it = DriveApp.getFoldersByName('TPFX Bukti Aktivitas');
    folder = it.hasNext()
      ? it.next()
      : DriveApp.createFolder('TPFX Bukti Aktivitas');
  } catch (e) {
    throw driveError_(e);
  }

  const raw = String(dataUrl || '');

  if (raw.length > 8 * 1024 * 1024) {
    throw new Error('Ukuran gambar terlalu besar');
  }

  const match = raw.match(/^data:image\/(?:jpeg|png|webp);base64,([A-Za-z0-9+/]+={0,2})$/i);
  if (!match) {
    throw new Error('Format gambar tidak valid');
  }

  try {
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
  } catch (e) {
    throw driveError_(e);
  }
}

function isQuotaError_(e) {
  const msg = String((e && e.message) || e);
  return /exceeded|quota|maximum|too many cells|too large|limit/i.test(msg);
}

function storageError_(e) {
  const msg = String((e && e.message) || e);
  // Jangan bungkus dua kali pesan yang sudah ramah untuk pengguna.
  if (msg.indexOf('Penyimpanan Google penuh') === 0 ||
      msg.indexOf('Penyimpanan Google Drive penuh') === 0) {
    return e;
  }
  if (isQuotaError_(e)) {
    return new Error(
      'Penyimpanan Google penuh atau melewati batas. Ruang sudah dicoba ' +
      'dirapikan otomatis; jika masih gagal, minta manager membuka ' +
      'database.html lalu menekan "Rapikan database". Detail: ' + msg
    );
  }
  return e instanceof Error ? e : new Error(msg);
}

function driveError_(e) {
  const msg = String((e && e.message) || e);
  if (/quota|exceeded|storage/i.test(msg)) {
    return new Error(
      'Penyimpanan Google Drive penuh. Hapus file bukti aktivitas lama atau ' +
      'kosongkan Drive, lalu coba unggah lagi. Detail: ' + msg
    );
  }
  return e instanceof Error ? e : new Error(msg);
}

// Perbaikan ringan yang aman dijalankan otomatis saat penyimpanan menolak
// penulisan baru: buang sesi kedaluwarsa, salinan Data ganda, lalu kembalikan
// baris/kolom kosong yang sudah terlanjur mengunci kuota sel.
function recoverStorage_() {
  const sessions = pruneSessions_();
  const data = compactDataSheet_();
  const trimmed = trimSheets_();
  return { sessions: sessions, data: data, trimmed: trimmed };
}

// ============================================================
// KAPASITAS & PERAWATAN DATABASE
// ============================================================

const SHEET_CELL_LIMIT = 10000000;
const PROPERTY_STORAGE_LIMIT = 524288;

function storageReport(t) {

  // Sengaja tidak memakai loadState: laporan kapasitas harus tetap bisa dibuka
  // walau JSON state manager sendiri sedang bermasalah.
  const username = user_(t);

  if (!username) return { err: 'AUTH' };

  const me = findUser_(username);

  if (!me) return { err: 'AUTH' };

  const role = String(me[4] || '').toUpperCase();

  if (role != 'BM' && role != 'DM') {
    return { err: 'Hanya manager yang dapat memeriksa kapasitas database' };
  }

  const book = ss_();
  const sheets = book.getSheets().map(sh => {
    const allocatedRows = sh.getMaxRows();
    const allocatedColumns = sh.getMaxColumns();
    return {
      name: sh.getName(),
      rows: sh.getLastRow(),
      columns: sh.getLastColumn(),
      allocatedRows: allocatedRows,
      allocatedColumns: allocatedColumns,
      allocatedCells: allocatedRows * allocatedColumns
    };
  }).sort((a, b) => b.allocatedCells - a.allocatedCells);

  const allocatedCells = sheets.reduce((n, s) => n + s.allocatedCells, 0);

  const properties = PropertiesService.getScriptProperties().getProperties();
  const propBytes = JSON.stringify(properties).length;
  const now = Date.now();
  let sessions = 0;
  let expiredSessions = 0;

  Object.keys(properties).forEach(key => {
    if (key.indexOf('t_') !== 0) return;
    sessions++;
    const raw = String(properties[key] || '');
    const exp = Number(raw.slice(raw.lastIndexOf('|') + 1));
    if (!exp || exp <= now) expiredSessions++;
  });

  let drive = { available: false };
  try {
    const folders = DriveApp.getFoldersByName('TPFX Bukti Aktivitas');
    const folder = folders.hasNext() ? folders.next() : null;
    let files = 0;
    let bytes = 0;
    let truncated = false;
    if (folder) {
      const list = folder.getFiles();
      while (list.hasNext()) {
        const file = list.next();
        files++;
        bytes += Number(file.getSize() || 0);
        if (files >= 5000) {
          truncated = true;
          break;
        }
      }
    }
    drive = {
      available: true,
      used: Number(DriveApp.getStorageUsed() || 0),
      limit: Number(DriveApp.getStorageLimit() || 0),
      folderFiles: files,
      folderBytes: bytes,
      truncated: truncated
    };
  } catch (e) {
    drive = { available: false, error: String((e && e.message) || e) };
  }

  const dataSheet = book.getSheetByName('Data');
  const dataRows = dataSheet ? Math.max(0, dataSheet.getLastRow() - 1) : 0;

  const warnings = [];

  if (allocatedCells >= SHEET_CELL_LIMIT * 0.8) {
    warnings.push('Sel spreadsheet terpakai ' + Math.round(allocatedCells / SHEET_CELL_LIMIT * 100) + '% (batas 10.000.000).');
  }
  if (propBytes >= PROPERTY_STORAGE_LIMIT * 0.8) {
    warnings.push('Penyimpanan properti/sesi terpakai ' + Math.round(propBytes / PROPERTY_STORAGE_LIMIT * 100) + '% (batas 500 KB).');
  }
  if (expiredSessions > 0) {
    warnings.push(expiredSessions + ' sesi kedaluwarsa belum dibersihkan.');
  }
  if (drive.available && drive.limit && drive.used / drive.limit >= 0.9) {
    warnings.push('Penyimpanan Google Drive terpakai ' + Math.round(drive.used / drive.limit * 100) + '%.');
  }
  if (drive.available && drive.truncated) {
    warnings.push('Jumlah file bukti lebih dari 5.000; ukuran folder dihitung sebagian.');
  }

  return {
    ok: 1,
    spreadsheet: { name: book.getName(), url: book.getUrl() },
    sheets: sheets,
    allocatedCells: allocatedCells,
    cellLimit: SHEET_CELL_LIMIT,
    cellPercent: Math.round(allocatedCells / SHEET_CELL_LIMIT * 1000) / 10,
    properties: {
      keys: Object.keys(properties).length,
      bytes: propBytes,
      limitBytes: PROPERTY_STORAGE_LIMIT,
      percent: Math.round(propBytes / PROPERTY_STORAGE_LIMIT * 1000) / 10,
      sessions: sessions,
      expiredSessions: expiredSessions
    },
    data: { rows: dataRows },
    drive: drive,
    warnings: warnings
  };
}

function compactDataSheet_() {

  const sheet = data_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return { removed: 0, users: 0, skipped: 0 };

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const byUser = {};

  values.forEach((row, index) => {
    const user = String(row[0] || '');
    if (!user) return;
    (byUser[user] = byUser[user] || []).push({
      row: index + 2,
      seq: Number(row[1]),
      json: String(row[2] || '')
    });
  });

  const rowsToDelete = [];
  let users = 0;
  let skipped = 0;

  Object.keys(byUser).forEach(user => {
    const list = byUser[user].sort((a, b) => a.row - b.row);
    const sets = [];
    let current = null;
    let lastSeq = -1;

    list.forEach(item => {
      if (!current || item.seq === 0 || item.seq <= lastSeq) {
        current = [];
        sets.push(current);
      }
      current.push(item);
      lastSeq = item.seq;
    });

    let keep = null;

    for (let i = sets.length - 1; i >= 0; i--) {
      const set = sets[i];
      let valid = true;
      set.forEach((item, index) => {
        if (item.seq !== index) valid = false;
      });
      if (!valid) continue;
      try {
        const parsed = JSON.parse(set.map(item => item.json).join(''));
        if (parsed && typeof parsed === 'object' &&
            Array.isArray(parsed.leads) && Array.isArray(parsed.acts)) {
          keep = set;
          break;
        }
      } catch (e) {}
    }

    if (!keep) {
      skipped++;
      return;
    }

    const keepRows = {};
    keep.forEach(item => { keepRows[item.row] = true; });
    const extra = list.filter(item => !keepRows[item.row]);

    if (extra.length) {
      users++;
      extra.forEach(item => rowsToDelete.push(item.row));
    }
  });

  rowsToDelete.sort((a, b) => b - a).forEach(row => sheet.deleteRow(row));

  return { removed: rowsToDelete.length, users: users, skipped: skipped };
}

function trimSheets_() {
  const report = [];

  ss_().getSheets().forEach(sh => {
    const maxRows = sh.getMaxRows();
    const lastRow = Math.max(1, sh.getLastRow());
    let removedRows = 0;
    let removedColumns = 0;

    if (maxRows - lastRow > 200) {
      removedRows = maxRows - lastRow - 100;
      sh.deleteRows(lastRow + 1, removedRows);
    }

    const maxColumns = sh.getMaxColumns();
    const lastColumn = Math.max(1, sh.getLastColumn());

    if (maxColumns - lastColumn > 5) {
      removedColumns = maxColumns - lastColumn - 5;
      sh.deleteColumns(lastColumn + 1, removedColumns);
    }

    if (removedRows || removedColumns) {
      report.push({
        sheet: sh.getName(),
        removedRows: removedRows,
        removedColumns: removedColumns
      });
    }
  });

  return report;
}

function storageCleanup(t) {

  const username = user_(t);

  if (!username) return { err: 'AUTH' };

  const me = findUser_(username);

  if (!me) return { err: 'AUTH' };

  const role = String(me[4] || '').toUpperCase();

  if (role != 'BM' && role != 'DM') {
    return { err: 'Hanya manager yang dapat merapikan database' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const sessions = pruneSessions_();
    const data = compactDataSheet_();
    const trimmed = trimSheets_();
    return { ok: 1, sessions: sessions, data: data, trimmed: trimmed };
  } finally {
    lock.releaseLock();
  }
}


// ============================================================
// DATABASE & CEK SESI
// ============================================================

// Kedua endpoint memvalidasi token melalui loadState;
// database hanya dikirim kepada akun dengan sesi aktif.
// JANGAN membuat endpoint publik untuk database ini.

// Kontak disimpan di sheet Contacts. Setiap baris punya assignedTo (username
// BDO) supaya pembagian stabil: impor baru hanya menambah beban BDO yang
// paling sedikit, bukan mengacak ulang kontak yang sudah dibagikan.
// DATABASE_DATA lama tetap dibaca sebagai fallback dan dimigrasikan otomatis
// saat impor pertama.

const CONTACT_HEADERS = [
  'id', 'group', 'name', 'seq', 'job', 'phone',
  'assignedTo', 'createdBy', 'createdAt', 'status', 'statusAt'
];

const CONTACT_STATUSES = [
  'new', 'contacted', 'noanswer', 'inactive',
  'answered', 'responded', 'pipeline', 'lost'
];

function contacts_() {
  const sheet = sheetEnsure_('Contacts', CONTACT_HEADERS);
  if (sheet.getLastColumn() < CONTACT_HEADERS.length) {
    const have = sheet.getLastColumn()
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      : [];
    const matches = have.length === 0 || have.every((h, i) =>
      i < CONTACT_HEADERS.length && String(h) == CONTACT_HEADERS[i]
    );
    if (matches) {
      sheet.getRange(1, have.length + 1, 1, CONTACT_HEADERS.length - have.length)
        .setValues([CONTACT_HEADERS.slice(have.length)]);
    }
  }
  return sheet;
}

function activeBdos_() {
  return users_().getDataRange().getValues().slice(1)
    .filter(r =>
      String(r[3]).toUpperCase() == 'Y' &&
      String(r[4]).toUpperCase() == 'BDO'
    )
    .map(r => ({
      username: String(r[0]),
      name: String(r[2] || r[0]),
      key: String(r[0]).toLowerCase()
    }));
}

function normalizePhone_(value) {
  let d = String(value == null ? '' : value).replace(/\D/g, '');
  if (!d) return '';
  if (d.indexOf('00') == 0) d = d.slice(2);
  if (d.indexOf('620') == 0) d = '62' + d.slice(3);
  else if (d.indexOf('0') == 0) d = '62' + d.slice(1);
  else if (d.indexOf('8') == 0) d = '62' + d;
  return d;
}

function legacyContacts_() {
  const rows = [];
  if (!Array.isArray(DATABASE_DATA)) return rows;
  DATABASE_DATA.forEach(group => {
    if (!group || !Array.isArray(group.r)) return;
    group.r.forEach(row => {
      if (!Array.isArray(row) || row.length < 3) return;
      rows.push({
        group: String(group.n || ''),
        name: '',
        seq: row[0],
        job: String(row[1] || ''),
        phone: normalizePhone_(row[2]),
        assignedTo: '',
        status: ''
      });
    });
  });
  return rows.filter(r => r.phone);
}

function pickLeastLoaded_(bdos, counts) {
  let best = bdos[0];
  bdos.forEach(bdo => {
    if ((counts[bdo.key] || 0) < (counts[best.key] || 0)) best = bdo;
  });
  return best;
}

function migrateLegacyContacts_(sheet, bdos, username) {
  if (sheet.getLastRow() > 1) return;
  const legacy = legacyContacts_();
  if (!legacy.length) return;
  const counts = {};
  const values = legacy.map((row, index) => {
    const target = pickLeastLoaded_(bdos, counts);
    counts[target.key] = (counts[target.key] || 0) + 1;
    return [
      'c_' + Date.now() + '_' + index,
      row.group,
      row.name,
      row.seq,
      row.job,
      row.phone,
      target.username,
      String(username || ''),
      Date.now(),
      '',
      0
    ];
  });
  sheet.getRange(2, 1, values.length, CONTACT_HEADERS.length).setValues(values);
}

function importContacts(t, group, rows) {
  const session = loadState(t);
  if (!session || session.err) return { err: session.err || 'AUTH' };

  const role = String(session.role || '').toUpperCase();
  if (role != 'BM' && role != 'DM') {
    return { err: 'Impor kontak hanya untuk akun manager (BM/DM)' };
  }

  const bdos = activeBdos_();
  if (!bdos.length) {
    return { err: 'Belum ada akun BDO aktif untuk pembagian kontak' };
  }

  if (!Array.isArray(rows) || !rows.length) {
    return { err: 'Tidak ada kontak untuk diimpor' };
  }
  if (rows.length > 5000) {
    return { err: 'Maksimal 5000 kontak per impor' };
  }

  const groupName = String(group || '').trim().slice(0, 80) ||
    ('Impor ' + Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy'
    ));

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const sheet = contacts_();
    migrateLegacyContacts_(sheet, bdos, session.user);

    const lastRow = sheet.getLastRow();
    const existing = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, CONTACT_HEADERS.length).getValues()
      : [];

    const counts = {};
    const seen = {};

    existing.forEach(r => {
      const phone = normalizePhone_(r[5]);
      if (phone) seen[phone] = true;
      const owner = String(r[6] || '').toLowerCase();
      const bdo = bdos.find(b => b.key == owner);
      if (bdo) counts[bdo.key] = (counts[bdo.key] || 0) + 1;
    });

    const values = [];
    let skipped = 0;
    let duplicate = 0;
    let invalid = 0;

    rows.forEach(row => {
      if (!Array.isArray(row)) {
        invalid++;
        skipped++;
        return;
      }

      const rowGroup = String(row[0] || '').trim().slice(0, 80) || groupName;
      const name = String(row[1] == null ? '' : row[1]).trim().slice(0, 120);
      const seq = String(row[2] == null ? '' : row[2]).trim();
      const job = String(row[3] == null ? '' : row[3]).trim().slice(0, 120);
      const phone = normalizePhone_(row[4]);

      if (!/^62\d{8,13}$/.test(phone)) {
        invalid++;
        skipped++;
        return;
      }
      if (seen[phone]) {
        duplicate++;
        skipped++;
        return;
      }

      seen[phone] = true;
      const target = pickLeastLoaded_(bdos, counts);
      counts[target.key] = (counts[target.key] || 0) + 1;

      values.push([
        'c_' + Date.now() + '_' + values.length,
        rowGroup,
        name,
        seq,
        job,
        phone,
        target.username,
        String(session.user || ''),
        Date.now(),
        '',
        0
      ]);
    });

    if (values.length) {
      try {
        sheet.getRange(
          sheet.getLastRow() + 1, 1, values.length, CONTACT_HEADERS.length
        ).setValues(values);
      } catch (e) {
        if (!isQuotaError_(e)) throw e;
        // Tanpa ini impor selalu gagal ketika kuota sel sudah penuh:
        // rapikan dulu, lalu tulis sekali lagi.
        recoverStorage_();
        sheet.getRange(
          sheet.getLastRow() + 1, 1, values.length, CONTACT_HEADERS.length
        ).setValues(values);
      }
    }

    return {
      ok: 1,
      group: groupName,
      added: values.length,
      skipped: skipped,
      duplicate: duplicate,
      invalid: invalid,
      total: lastRow - 1 + values.length,
      perBdo: bdos.map(bdo => ({
        username: bdo.username,
        name: bdo.name,
        count: counts[bdo.key] || 0
      }))
    };
  } catch (e) {
    throw storageError_(e);
  } finally {
    lock.releaseLock();
  }
}

function normalizeContactStatus_(value) {
  const key = String(value == null ? '' : value).toLowerCase().trim();
  return CONTACT_STATUSES.indexOf(key) >= 0 ? key : 'new';
}

function contactRecords_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet
    .getRange(2, 1, lastRow - 1, CONTACT_HEADERS.length)
    .getValues()
    .map((r, index) => ({
      row: index + 2,
      group: String(r[1] || ''),
      name: String(r[2] || ''),
      seq: r[3],
      job: String(r[4] || ''),
      phone: normalizePhone_(r[5]),
      assignedTo: String(r[6] || ''),
      status: normalizeContactStatus_(r[9]),
      statusAt: Number(r[10] || 0)
    }));
}

function contactByPhone_(sheet, phone) {
  const target = normalizePhone_(phone);
  if (!target) return null;
  return contactRecords_(sheet).find(r => r.phone == target) || null;
}

function canManageContact_(session, record) {
  const role = String(session.role || '').toUpperCase();
  if (role == 'BM' || role == 'DM') return true;
  return String(record.assignedTo || '').toLowerCase() ==
    String(session.user || '').toLowerCase();
}

function ensureContactsMigrated_(sheet, bdos, username) {
  if (sheet.getLastRow() > 1) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    migrateLegacyContacts_(sheet, bdos, username);
  } finally {
    lock.releaseLock();
  }
}

function databaseData(t) {

  const session = loadState(t);

  if (!session || session.err) return { err: session.err || 'AUTH' };

  const bdos = activeBdos_();

  if (!bdos.length) {
    return { err: 'Belum ada akun BDO aktif untuk pembagian kontak' };
  }

  const byKey = {};
  const counts = {};
  bdos.forEach(bdo => {
    byKey[bdo.key] = bdo;
    counts[bdo.key] = 0;
  });

  const sheet = contacts_();
  ensureContactsMigrated_(sheet, bdos, session.user);

  let records = contactRecords_(sheet);

  if (!records.length) {
    records = legacyContacts_().map(r => ({
      group: r.group,
      name: r.name,
      seq: r.seq,
      job: r.job,
      phone: r.phone,
      assignedTo: '',
      status: 'new'
    }));
  }

  const seen = {};
  records = records.filter(r => {
    if (!/^62\d{8,13}$/.test(r.phone)) return false;
    if (seen[r.phone]) return false;
    seen[r.phone] = true;
    return true;
  });

  records.forEach(r => {
    const key = String(r.assignedTo || '').toLowerCase();
    if (byKey[key]) counts[key]++;
  });

  const rowsByKey = {};
  const groups = [];
  bdos.forEach(bdo => {
    rowsByKey[bdo.key] = [];
  });

  records.forEach(r => {
    let key = String(r.assignedTo || '').toLowerCase();
    if (!byKey[key]) {
      const target = pickLeastLoaded_(bdos, counts);
      key = target.key;
      counts[key]++;
    }
    if (r.group && groups.indexOf(r.group) < 0) groups.push(r.group);
    rowsByKey[key].push([
      r.seq, r.job, r.phone, r.name, normalizeContactStatus_(r.status)
    ]);
  });

  bdos.forEach(bdo => {
    rowsByKey[bdo.key].forEach((row, index) => {
      if (String(row[0] == null ? '' : row[0]).trim() === '') row[0] = index + 1;
    });
  });

  return {
    rows: bdos.map(bdo => ({
      n: bdo.name,
      u: bdo.username,
      r: rowsByKey[bdo.key]
    })),
    role: String(session.role || 'BDO').toUpperCase(),
    user: String(session.user || ''),
    name: String(session.name || ''),
    groups: groups,
    total: records.length
  };
}

function setContactStatus(t, phone, status) {

  const session = loadState(t);

  if (!session || session.err) return { err: session.err || 'AUTH' };

  const key = String(status == null ? '' : status).toLowerCase().trim();

  if (CONTACT_STATUSES.indexOf(key) < 0) return { err: 'Status kontak tidak valid' };

  if (key == 'pipeline') {
    return { err: 'Gunakan tombol Sambung ke pipeline' };
  }

  const target = normalizePhone_(phone);

  if (!/^62\d{8,13}$/.test(target)) return { err: 'Nomor kontak tidak valid' };

  const bdos = activeBdos_();

  if (!bdos.length) {
    return { err: 'Belum ada akun BDO aktif untuk pembagian kontak' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {

    const sheet = contacts_();
    migrateLegacyContacts_(sheet, bdos, session.user);

    const record = contactByPhone_(sheet, target);

    if (!record) {
      return { err: 'Kontak tidak ditemukan. Muat ulang halaman.' };
    }

    if (!canManageContact_(session, record)) {
      return { err: 'Kontak ini bukan milik Anda' };
    }

    const at = Date.now();
    sheet.getRange(record.row, 10, 1, 2).setValues([[key, at]]);

    return { ok: 1, phone: target, status: key, statusAt: at };
  } finally {
    lock.releaseLock();
  }
}

function addContactToPipeline(t, phone) {

  const session = loadState(t);

  if (!session || session.err) return { err: session.err || 'AUTH' };

  const role = String(session.role || '').toUpperCase();

  if (role != 'BDO') {
    return { err: 'Pipeline hanya untuk akun BDO pemilik kontak' };
  }

  const target = normalizePhone_(phone);

  if (!/^62\d{8,13}$/.test(target)) return { err: 'Nomor kontak tidak valid' };

  const bdos = activeBdos_();

  if (!bdos.length) {
    return { err: 'Belum ada akun BDO aktif untuk pembagian kontak' };
  }

  const sheet = contacts_();

  ensureContactsMigrated_(sheet, bdos, session.user);

  const record = contactByPhone_(sheet, target);

  if (!record) {
    return { err: 'Kontak tidak ditemukan. Muat ulang halaman.' };
  }

  if (String(record.assignedTo || '').toLowerCase() !=
      String(session.user || '').toLowerCase()) {
    return { err: 'Kontak ini bukan milik Anda' };
  }

  const state = session.state && typeof session.state == 'object'
    ? session.state
    : {};

  if (!Array.isArray(state.leads)) state.leads = [];
  if (!Array.isArray(state.acts)) state.acts = [];
  if (state.target == null) state.target = session.target || 10;

  const existing = state.leads.find(l =>
    l && normalizePhone_(l.phone) == target
  );

  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const tomorrow = Utilities.formatDate(
    new Date(now.getTime() + 864e5), tz, 'yyyy-MM-dd'
  );

  let lead = existing;

  if (!lead) {
    lead = {
      id: Date.now(),
      c: today,
      log: [],
      name: record.name || ('Kontak ' + target.slice(-4)),
      phone: target,
      product: '',
      source: 'WhatsApp',
      interest: 'Sedang',
      status: 'warm',
      last: today,
      next: tomorrow,
      notes: [record.group, record.job].filter(Boolean).join(' · ')
    };
    state.leads.unshift(lead);
    state.acts.unshift({
      id: Date.now() + 1,
      type: 'wa',
      text: 'Kontak database tersambung: ' + lead.name,
      leadId: lead.id,
      ts: Date.now()
    });
  }

  const saved = saveState(t, JSON.stringify(state));

  if (!saved || saved.err) {
    return { err: (saved && saved.err) || 'Gagal menyimpan pipeline' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const contactsSheet = contacts_();
    const fresh = contactByPhone_(contactsSheet, target);
    if (fresh) {
      contactsSheet.getRange(fresh.row, 10, 1, 2)
        .setValues([['pipeline', Date.now()]]);
    }
  } finally {
    lock.releaseLock();
  }

  return {
    ok: 1,
    phone: target,
    name: lead.name,
    leadId: lead.id,
    status: 'pipeline',
    already: existing ? 1 : 0
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

  // Postingan BM/DM adalah pengumuman: kirim push ke seluruh akun aktif.
  if (['BM', 'DM'].includes(String(account[4] || '').toUpperCase())) {
    const recipients = users_().getDataRange().getValues().slice(1)
      .filter(r => String(r[3]).toUpperCase() == 'Y')
      .map(r => String(r[0] || ''))
      .filter(name => name && name.toLowerCase() != String(username).toLowerCase());
    pushQuiet_(() => pushToUsers_(recipients, {
      title: String(account[2] || 'Manajemen') + ' · Pengumuman',
      body: text || 'Postingan baru dengan foto',
      url: './index.html#feed',
      tag: 'tpfx-feed-' + String(account[0]).toLowerCase()
    }));
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

  const senderName = chatSenderName_(username);
  pushQuiet_(() => pushToUsers_([peer], {
    title: senderName,
    body: text,
    url: './index.html#chat=personal:' + encodeURIComponent(username),
    tag: 'tpfx-dm-' + String(username).toLowerCase()
  }));

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

  const senderName = chatSenderName_(username);
  const groupName = String(group[1] || 'Grup');
  const recipients = groupMembers_(group)
    .filter(name => name != String(username).toLowerCase());
  pushQuiet_(() => pushToUsers_(recipients, {
    title: senderName + ' · ' + groupName,
    body: text,
    url: './index.html#chat=group:' + encodeURIComponent(groupId),
    tag: 'tpfx-group-' + groupId
  }));

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

  const senderName = chatSenderName_(username);
  const recipients = users_().getDataRange().getValues().slice(1)
    .filter(r => String(r[3]).toUpperCase() == 'Y')
    .map(r => String(r[0] || ''))
    .filter(name => name && name.toLowerCase() != String(username).toLowerCase());
  pushQuiet_(() => pushToUsers_(recipients, {
    title: senderName + ' · Grup Global',
    body: text,
    url: './index.html#chat=global',
    tag: 'tpfx-global'
  }));

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
// NOTIFIKASI HP (WEB PUSH)
// ============================================================
// Pengiriman push dilakukan oleh relay Vercel (/api/push) karena Apps Script
// tidak dapat menandatangani VAPID/ES256. Apps Script hanya menyimpan
// langganan dan meneruskan permintaan dengan secret.
//
// Sekali saja, jalankan dari editor Apps Script:
//   setPushConfig('SECRET_DARI_VERCEL', 'https://domain-anda.vercel.app')

const PUSH_SECRET_PROP = 'PUSH_RELAY_SECRET';
const PUSH_ORIGIN_PROP = 'PUSH_APP_ORIGIN';
const PUSH_CONTACT_PROP = 'PUSH_CONTACT';
const PUSH_MAX_SUBS_PER_USER = 10;

function pushSubs_() {
  return sheetEnsure_('PushSubs', [
    'endpoint', 'p256dh', 'auth', 'username', 'origin', 'ua', 'ts'
  ]);
}

function setPushConfig(secret, appOrigin, contact) {
  secret = String(secret || '').trim();
  appOrigin = String(appOrigin || '').trim().replace(/\/+$/, '');
  if (secret.length < 16) {
    throw new Error('Secret minimal 16 karakter.');
  }
  if (!/^https:\/\/[a-z0-9.-]+$/i.test(appOrigin)) {
    throw new Error('Origin harus https://domain tanpa path.');
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PUSH_SECRET_PROP, secret);
  props.setProperty(PUSH_ORIGIN_PROP, appOrigin);
  props.setProperty(
    PUSH_CONTACT_PROP,
    String(contact || '').trim() ||
      ('mailto:admin@' + appOrigin.replace(/^https:\/\//i, ''))
  );
  return { ok: 1, origin: appOrigin };
}

function pushConfig(t) {
  const username = user_(t);
  if (!username) return { err: 'AUTH' };
  const props = PropertiesService.getScriptProperties();
  return {
    enabled: !!(props.getProperty(PUSH_SECRET_PROP) && props.getProperty(PUSH_ORIGIN_PROP)),
    origin: String(props.getProperty(PUSH_ORIGIN_PROP) || '')
  };
}

function subscribePush(t, json) {
  const username = user_(t);
  if (!username) return { err: 'AUTH' };

  let sub;
  try {
    sub = JSON.parse(String(json || ''));
  } catch (e) {
    return { err: 'Langganan push tidak valid' };
  }

  const endpoint = String(sub && sub.endpoint || '');
  if (!/^https:\/\/\S+$/.test(endpoint) || endpoint.length > 500) {
    return { err: 'Endpoint push tidak valid' };
  }

  const keys = (sub && sub.keys) || {};
  const p256dh = String(keys.p256dh || '');
  const auth = String(keys.auth || '');
  if (!p256dh || !auth || p256dh.length > 200 || auth.length > 100) {
    return { err: 'Kunci langganan push tidak lengkap' };
  }

  const props = PropertiesService.getScriptProperties();
  const origin = String(props.getProperty(PUSH_ORIGIN_PROP) || '');
  if (!origin) return { err: 'Push belum dikonfigurasi manager' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const sheet = pushSubs_();
    const rows = sheet.getDataRange().getValues();
    const userKey = String(username).toLowerCase();

    // Satu endpoint hanya boleh muncul sekali; baris lama dengan endpoint yang
    // sama (milik akun ini maupun akun lain di perangkat yang sama) diganti.
    const staleRows = [];
    const mine = [];
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) == endpoint) {
        staleRows.push(i + 1);
      } else if (String(rows[i][3] || '').toLowerCase() == userKey) {
        mine.push({ row: i + 1, ts: Number(rows[i][6]) || 0 });
      }
    }
    staleRows.sort((a, b) => b - a).forEach(row => sheet.deleteRow(row));

    // Sisakan ruang untuk langganan baru, buang yang paling lama lebih dulu.
    // Baris dihapus dari nomor terbesar agar nomor baris lain tidak bergeser.
    mine.sort((a, b) => (a.ts - b.ts) || (a.row - b.row));
    const excess = mine.length - (PUSH_MAX_SUBS_PER_USER - 1);
    if (excess > 0) {
      mine.slice(0, excess)
        .sort((a, b) => b.row - a.row)
        .forEach(item => sheet.deleteRow(item.row));
    }

    sheet.appendRow([
      endpoint,
      p256dh,
      auth,
      username,
      origin,
      String(sub.userAgent || sub.ua || '').slice(0, 200),
      Date.now()
    ]);

    return { ok: 1 };
  } catch (e) {
    throw storageError_(e);
  } finally {
    lock.releaseLock();
  }
}

function unsubscribePush(t, endpoint) {
  const username = user_(t);
  if (!username) return { err: 'AUTH' };
  endpoint = String(endpoint || '');
  if (!endpoint) return { err: 'Endpoint push tidak valid' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = pushSubs_();
    const rows = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) != endpoint) continue;
      if (String(rows[i][3]).toLowerCase() != String(username).toLowerCase()) continue;
      sheet.deleteRow(i + 1);
    }
  } finally {
    lock.releaseLock();
  }
  return { ok: 1 };
}

function removePushEndpoints_(endpoints) {
  if (!endpoints || !endpoints.length) return 0;
  const wanted = {};
  endpoints.forEach(value => {
    const key = String(value || '');
    if (key) wanted[key] = true;
  });
  const sheet = pushSubs_();
  const rows = sheet.getDataRange().getValues();
  let removed = 0;
  for (let i = rows.length - 1; i >= 1; i--) {
    if (!wanted[String(rows[i][0])]) continue;
    sheet.deleteRow(i + 1);
    removed++;
  }
  return removed;
}

function chatSenderName_(username) {
  const account = findUser_(username);
  return String((account && account[2]) || username || 'Anggota TPFX');
}

function pushQuiet_(fn) {
  try {
    fn();
  } catch (e) {
    console.error('Notifikasi push gagal:', e);
  }
}

function pushToUsers_(usernames, note) {
  const props = PropertiesService.getScriptProperties();
  const secret = String(props.getProperty(PUSH_SECRET_PROP) || '');
  const origin = String(props.getProperty(PUSH_ORIGIN_PROP) || '');
  if (!secret || !origin || !usernames || !usernames.length) {
    return { sent: 0, failed: 0, subscriptions: 0 };
  }

  const wanted = {};
  usernames.forEach(name => {
    const key = String(name || '').toLowerCase();
    if (key) wanted[key] = true;
  });

  const subs = [];
  const seen = {};
  pushSubs_().getDataRange().getValues().slice(1).forEach(row => {
    const owner = String(row[3] || '').toLowerCase();
    if (!wanted[owner]) return;
    const endpoint = String(row[0] || '');
    const p256dh = String(row[1] || '');
    const auth = String(row[2] || '');
    if (!endpoint || !p256dh || !auth || seen[endpoint]) return;
    seen[endpoint] = true;
    subs.push({ endpoint: endpoint, keys: { p256dh: p256dh, auth: auth } });
  });

  if (!subs.length) return { sent: 0, failed: 0, subscriptions: 0 };

  const payload = {
    notification: {
      title: String(note.title || 'TPFX Workspace').slice(0, 90),
      body: String(note.body || '').slice(0, 180),
      url: String(note.url || './index.html').slice(0, 300),
      tag: String(note.tag || 'tpfx').slice(0, 60)
    },
    subs: subs.slice(0, 100)
  };

  let response;
  try {
    response = UrlFetchApp.fetch(origin + '/api/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-push-secret': secret },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('Relay push tidak dapat dihubungi:', e);
    return { sent: 0, failed: subs.length, subscriptions: subs.length };
  }

  let data = {};
  try {
    data = JSON.parse(response.getContentText());
  } catch (e) {}

  if (Array.isArray(data.gone) && data.gone.length) {
    try { removePushEndpoints_(data.gone); } catch (e) {}
  }

  return {
    sent: Number(data.sent) || 0,
    failed: Number(data.failed) || 0,
    subscriptions: subs.length
  };
}

function testPush(t) {
  const username = user_(t);
  if (!username) return { err: 'AUTH' };
  const result = pushToUsers_([username], {
    title: 'TPFX Workspace',
    body: 'Notifikasi HP aktif. Pesan baru akan muncul di sini.',
    url: './index.html#chat=global',
    tag: 'tpfx-test'
  });
  return { ok: 1, sent: result.sent, failed: result.failed, subscriptions: result.subscriptions };
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