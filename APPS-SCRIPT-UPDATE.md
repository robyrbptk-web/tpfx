# Pembaruan Google Apps Script

File `Pasted text #1.txt` tidak saya ubah karena merupakan lampiran read-only. Terapkan perubahan berikut pada project Apps Script yang sudah ada.

## 1. Ganti fungsi `ensureManager_`

Ganti fungsi lama dengan fungsi ini. Nama BM/DM yang telah diubah lewat profil tidak akan dikembalikan ke nama awal setiap kali `setup()` dijalankan.

```javascript
function ensureManager_(username, password, nama, role) {
  const r = findUser_(username);

  if (!r) {
    ensureUser_(username, password, nama, role);
    return;
  }

  const sh = users_();

  // Pertahankan nama yang telah diubah oleh pemilik akun.
  sh.getRange(r.idx, 4).setValue('Y');
  sh.getRange(r.idx, 5).setValue(String(role).toUpperCase());
}
```

## 2. Ganti fungsi `api_()`

Ganti fungsi `api_()` lama dengan fungsi ini:

```javascript
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
    deleteUser: deleteUser,
    databaseData: databaseData,
    publicDatabaseData: publicDatabaseData,
    checkSession: checkSession,
    socialFeed: socialFeed,
    recentPosts: recentPosts,
    createPost: createPost,
    chatContacts: chatContacts,
    chatHistory: chatHistory,
    sendChat: sendChat,
    chatGroups: chatGroups,
    createChatGroup: createChatGroup,
    groupHistory: groupHistory,
    sendGroupChat: sendGroupChat
  };
}
```

Jangan menghapus endpoint lain yang sudah ada.

## 3. Reset password dan target aktivitas BDO

Tempel dua endpoint berikut ke Code.gs. Reset password memakai helper hash yang sudah digunakan login (`hash_(username, password)`), hanya mengizinkan target role BDO, dan mencabut semua sesi aktif target. Sampaikan password sementara kepada BDO secara pribadi.

Target aktivitas disimpan sebagai override di Script Properties agar tidak tertimpa state lama saat aplikasi BDO melakukan sinkronisasi otomatis.

```javascript
function activityTargetKey_(username) {
  return 'activity_target_' + String(username).toLowerCase();
}


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
```

**Pertahankan target saat state dimuat/disimpan:**

1. Pada `loadState(t)`, tambahkan properti `target` pada object hasil sukses. Ambil nilainya dari override jika ada:

   ```javascript
   const targetOverride = PropertiesService.getScriptProperties()
     .getProperty(activityTargetKey_(u));
   ```

   Di dalam object hasil sukses:

   ```javascript
   target: targetOverride === null ? (state && state.target || 10) : Number(targetOverride)
   ```

2. Pada `saveState(t, json)`, di dalam blok `try` setelah lock didapatkan dan sebelum membuat `raw`, pertahankan override yang sudah ditetapkan manager:

   ```javascript
   const targetOverride = PropertiesService.getScriptProperties()
     .getProperty(activityTargetKey_(u));
   if (targetOverride !== null) parsed.target = Number(targetOverride);
   const raw = JSON.stringify(parsed);
   ```

   Ganti baris lama `const raw = String(json);` dengan snippet tersebut, agar penyimpanan BDO tidak menimpa target dari BM/DM.

3. Pada `dashboardDM(t)`, saat mengisi `state.target` di dalam `.map(r => ...)`, gunakan override:

   ```javascript
   const targetOverride = PropertiesService.getScriptProperties()
     .getProperty(activityTargetKey_(r[0]));
   // Dalam object state:
   target: targetOverride === null ? (s.target || 10) : Number(targetOverride),
   ```

   Letakkan deklarasi `targetOverride` sebelum `return` object untuk akun tersebut.

4. Pada `deleteUser`, sebelum `users_().deleteRow(target.idx)`, bersihkan override agar tidak tertinggal:

   ```javascript
   PropertiesService.getScriptProperties()
     .deleteProperty(activityTargetKey_(targetUsername));
   ```

   `targetUsername` pada fungsi hapus yang ada sudah lowercase.

Simpan Apps Script, lalu deploy versi baru Web App yang sama. Uji target di BM/DM dan login BDO kembali untuk melihat target terbaru; reset password mengeluarkan sesi BDO lama.

## 4. Tambahkan fungsi backend

Tempel dua fungsi berikut di Code.gs, misalnya setelah fungsi `createUser`.

```javascript
function updateProfile(t, nama) {
  const username = user_(t);

  if (!username) {
    return { err: 'AUTH' };
  }

  const callerRole = role_(t);

  if (callerRole != 'BM' && callerRole != 'DM') {
    return { err: 'Akses ditolak' };
  }

  nama = String(nama || '').trim();

  if (!nama) {
    return { err: 'Nama tidak boleh kosong' };
  }

  if (nama.length > 100) {
    return { err: 'Nama maksimal 100 karakter' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const akun = findUser_(username);

    if (!akun) {
      return { err: 'AUTH' };
    }

    users_().getRange(akun.idx, 3).setValue(nama);
  } finally {
    lock.releaseLock();
  }

  return { ok: 1, nama: nama };
}


function deleteUser(t, username) {
  const caller = user_(t);

  if (!caller) {
    return { err: 'AUTH' };
  }

  const callerRole = role_(t);

  if (callerRole != 'BM' && callerRole != 'DM') {
    return { err: 'Akses ditolak' };
  }

  username = String(username || '').trim();

  if (!username) {
    return { err: 'Username akun tidak valid' };
  }

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
    const scriptProperties = PropertiesService.getScriptProperties();
    const properties = scriptProperties.getProperties();

    Object.keys(properties).forEach(function(key) {
      if (key.indexOf('t_') != 0) return;

      const sessionUsername = String(properties[key]).split('|')[0];

      if (sessionUsername.toLowerCase() == targetUsername) {
        scriptProperties.deleteProperty(key);
      }
    });

    const data = data_();
    const rows = data.getDataRange().getValues();

    // Hapus data akun dari bawah ke atas agar indeks baris tetap valid.
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]).toLowerCase() == targetUsername) {
        data.deleteRow(i + 1);
      }
    }

    purgeSocialData_(targetUsername);
    users_().deleteRow(target.idx);
  } finally {
    lock.releaseLock();
  }

  return { ok: 1 };
}
```

## 5. Sajikan database publik dari Apps Script

`database.html` dapat dibuka tanpa login dan memanggil endpoint `publicDatabaseData`. Karena endpoint ini tidak memakai sesi, jangan hubungkan endpoint tersebut ke seluruh daftar pribadi atau langsung mengembalikan `DATABASE_DATA`. Buat whitelist terpisah yang hanya berisi kontak yang pemilik datanya telah setujui untuk dipublikasikan. Jangan menaruh isi whitelist atau nomor telepon ke file dokumentasi atau Git.

Data lama dapat diambil dari versi Git sebelum perubahan ini:

```powershell
git show 388de7a:database.html
```

Jangan salin seluruh array `DATA` lama sebagai data publik. Buat whitelist kosong berikut di Code.gs, lalu isi secara manual hanya dengan record yang sudah mendapat persetujuan publik. Bentuk setiap record tetap `[nomor, pekerjaan, telepon]`, dan setiap kelompok BDO berbentuk `{ n: 'Nama BDO', r: [ ...record... ] }`, sama seperti format lama:

```javascript
const PUBLIC_DATABASE_DATA = [];
```

Jangan tambahkan kontak ke whitelist tanpa persetujuan yang relevan. Simpan datanya hanya di project Apps Script.

Pertahankan `databaseData(t)` lama untuk pemanggilan berautentikasi yang mungkin masih digunakan aplikasi. Tambahkan endpoint publik berikut ke Code.gs dan daftarkan `publicDatabaseData` di `api_()`:

```javascript
function publicDatabaseData() {
  return { rows: PUBLIC_DATABASE_DATA };
}


function checkSession(t) {
  const session = loadState(t);

  if (!session || session.err) {
    return { err: 'AUTH' };
  }

  return { ok: 1 };
}
```

`databaseData(t)` dan `checkSession(t)` tetap memvalidasi token melalui `loadState`; hanya endpoint baru `publicDatabaseData()` yang tidak memakai sesi, dan endpoint tersebut wajib mengembalikan whitelist saja. Pastikan deployment Web App mengizinkan akses publik sesuai kebutuhan aplikasi; ini berarti setiap record dalam whitelist dapat dilihat siapa pun yang memiliki URL database. Simpan Apps Script dan **deploy versi baru** pada deployment Web App yang digunakan aplikasi.

## 6. Feed sosial dan chat di `index.html`

Tambahkan sheet penyimpanan berikut ke fungsi `setup()` setelah pembuatan sheet `Data`:

```javascript
  if (!b.getSheetByName('Posts')) {
    b.insertSheet('Posts').appendRow([
      'id', 'username', 'nama', 'text', 'ts'
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
```

Tambahkan helper dan endpoint berikut ke Code.gs:

```javascript
function posts_() {
  const sheet = ss_().getSheetByName('Posts');
  if (!sheet) throw new Error('Sheet Posts belum ada. Jalankan setup dulu.');
  return sheet;
}


function messages_() {
  const sheet = ss_().getSheetByName('Messages');
  if (!sheet) throw new Error('Sheet Messages belum ada. Jalankan setup dulu.');
  return sheet;
}


function socialFeed(t) {
  const username = user_(t);
  if (!username || !findUser_(username)) return { err: 'AUTH' };

  const feed = posts_().getDataRange().getValues().slice(1)
    .filter(r => String(r[3] || '').trim())
    .map(r => ({
      id: String(r[0]),
      username: String(r[1]),
      nama: String(r[2]),
      text: String(r[3]),
      ts: Number(r[4]),
      kind: 'post'
    }));

  const activeUsers = {};
  users_().getDataRange().getValues().slice(1).forEach(r => {
    if (String(r[3]).toUpperCase() == 'Y') {
      activeUsers[String(r[0]).toLowerCase()] = String(r[2] || r[0]);
    }
  });

  const chunks = {};
  data_().getDataRange().getValues().slice(1).forEach(r => {
    const key = String(r[0]).toLowerCase();
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
    const rows = chunks[key].sort((a, b) => Number(a[1]) - Number(b[1]));
    let state;
    try {
      state = JSON.parse(rows.map(r => String(r[2] || '')).join(''));
    } catch (e) {
      Logger.log('Feed melewati state JSON tidak valid untuk ' + key);
      return;
    }

    (state.acts || []).forEach(a => {
      const ts = Number(a.ts);
      if (!ts || ts < since) return;
      feed.push({
        id: 'activity-' + key + '-' + String(a.id || ts),
        username: rows[0][0],
        nama: activeUsers[key],
        text: labels[a.type] || 'Mencatat aktivitas penjualan',
        ts: ts,
        kind: 'activity'
      });
    });
  });

  feed.sort((a, b) => b.ts - a.ts);
  return { rows: feed.slice(0, 100) };
}


function createPost(t, text) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account) return { err: 'AUTH' };

  text = String(text || '').trim();
  if (!text) return { err: 'Isi postingan tidak boleh kosong' };
  if (text.length > 500) return { err: 'Postingan maksimal 500 karakter' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    posts_().appendRow([
      Utilities.getUuid(), String(account[0]), String(account[2]), text, Date.now()
    ]);
  } finally {
    lock.releaseLock();
  }
  return { ok: 1 };
}


function recentPosts(t) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account || String(account[3]).toUpperCase() != 'Y') {
    return { err: 'AUTH' };
  }

  const sheet = posts_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [] };

  const startRow = Math.max(2, lastRow - 99);
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, 5)
    .getValues();

  const rows = values
    .filter(r => String(r[3] || '').trim())
    .map(r => ({
      id: String(r[0]),
      username: String(r[1]),
      nama: String(r[2]),
      text: String(r[3]),
      ts: Number(r[4]) || 0,
      kind: 'post'
    }))
    .sort((a, b) => b.ts - a.ts);

  return { rows: rows };
}


function chatContacts(t) {
  const username = user_(t);
  if (!username || !findUser_(username)) return { err: 'AUTH' };

  const contacts = {};
  users_().getDataRange().getValues().slice(1).forEach(r => {
    const other = String(r[0]);
    if (String(r[3]).toUpperCase() != 'Y' ||
        other.toLowerCase() == username.toLowerCase()) return;
    contacts[other.toLowerCase()] = {
      username: other,
      nama: String(r[2] || other),
      role: String(r[4] || 'BDO').toUpperCase(),
      lastText: '',
      lastTs: 0
    };
  });

  messages_().getDataRange().getValues().slice(1).forEach(r => {
    const from = String(r[1]);
    const to = String(r[2]);
    if (from.toLowerCase() != username.toLowerCase() &&
        to.toLowerCase() != username.toLowerCase()) return;
    const other = (from.toLowerCase() == username.toLowerCase() ? to : from).toLowerCase();
    const contact = contacts[other];
    const ts = Number(r[4]) || 0;
    if (contact && ts >= contact.lastTs) {
      contact.lastText = String(r[3]);
      contact.lastTs = ts;
    }
  });

  return {
    rows: Object.keys(contacts).map(k => contacts[k])
      .sort((a, b) => b.lastTs - a.lastTs || a.nama.localeCompare(b.nama))
  };
}


function chatHistory(t, peer) {
  const username = user_(t);
  if (!username || !findUser_(username)) return { err: 'AUTH' };

  peer = String(peer || '').trim();
  if (!peer || peer.toLowerCase() == username.toLowerCase() ||
      !findUser_(peer)) return { err: 'Akun chat tidak ditemukan' };

  const rows = messages_().getDataRange().getValues().slice(1)
    .filter(r =>
      (String(r[1]).toLowerCase() == username.toLowerCase() &&
       String(r[2]).toLowerCase() == peer.toLowerCase()) ||
      (String(r[1]).toLowerCase() == peer.toLowerCase() &&
       String(r[2]).toLowerCase() == username.toLowerCase())
    )
    .sort((a, b) => Number(a[4]) - Number(b[4]))
    .slice(-100)
    .map(r => ({
      id: String(r[0]),
      from: String(r[1]),
      to: String(r[2]),
      text: String(r[3]),
      ts: Number(r[4])
    }));

  return { rows: rows };
}


function sendChat(t, peer, text) {
  const username = user_(t);
  if (!username || !findUser_(username)) return { err: 'AUTH' };

  peer = String(peer || '').trim();
  if (!peer || peer.toLowerCase() == username.toLowerCase() ||
      !findUser_(peer)) return { err: 'Akun chat tidak ditemukan' };

  text = String(text || '').trim();
  if (!text) return { err: 'Pesan tidak boleh kosong' };
  if (text.length > 1000) return { err: 'Pesan maksimal 1000 karakter' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    messages_().appendRow([
      Utilities.getUuid(), username, peer, text, Date.now()
    ]);
  } finally {
    lock.releaseLock();
  }
  return { ok: 1 };
}


function chatGroups_() {
  const sheet = ss_().getSheetByName('ChatGroups');
  if (!sheet) throw new Error('Sheet ChatGroups belum ada. Jalankan setup dulu.');
  return sheet;
}


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
  messages_().getDataRange().getValues().slice(1).forEach(r => {
    const target = String(r[2] || '');
    if (target.indexOf('group:') != 0) return;
    const id = target.slice(6);
    if (!lastMessages[id] || Number(r[4]) >= lastMessages[id].ts) {
      lastMessages[id] = { text: String(r[3]), ts: Number(r[4]) || 0 };
    }
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
        memberCount: members.length,
        lastText: String(last.text || ''),
        lastTs: Number(last.ts) || 0
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


function groupHistory(t, groupId) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };

  groupId = String(groupId || '').trim();
  const groups = chatGroups_().getDataRange().getValues().slice(1);
  const group = groups.find(r => String(r[0]) == groupId);
  if (!group) return { err: 'Grup tidak ditemukan' };
  if (groupMembers_(group).indexOf(String(username).toLowerCase()) < 0) {
    return { err: 'Anda bukan anggota grup ini' };
  }

  const names = {};
  users_().getDataRange().getValues().slice(1).forEach(r => {
    const key = String(r[0] || '').toLowerCase();
    if (key) names[key] = String(r[2] || r[0]);
  });

  const rows = messages_().getDataRange().getValues().slice(1)
    .filter(r => String(r[2]) == 'group:' + groupId)
    .sort((a, b) => Number(a[4]) - Number(b[4]))
    .slice(-200)
    .map(r => ({
      id: String(r[0]),
      from: String(r[1]),
      nama: names[String(r[1]).toLowerCase()] || String(r[1]),
      text: String(r[3]),
      ts: Number(r[4]) || 0
    }));

  return { rows: rows };
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
    messages_().appendRow([
      Utilities.getUuid(), username, 'group:' + groupId, text, Date.now()
    ]);
  } finally {
    lock.releaseLock();
  }
  return { ok: 1 };
}


function purgeSocialData_(username) {
  const key = String(username).toLowerCase();
  const posts = posts_();
  const postRows = posts.getDataRange().getValues();
  for (let i = postRows.length - 1; i >= 1; i--) {
    if (String(postRows[i][1]).toLowerCase() == key) posts.deleteRow(i + 1);
  }

  const messages = messages_();
  const messageRows = messages.getDataRange().getValues();
  for (let i = messageRows.length - 1; i >= 1; i--) {
    if (String(messageRows[i][1]).toLowerCase() == key ||
        String(messageRows[i][2]).toLowerCase() == key) {
      messages.deleteRow(i + 1);
    }
  }

  const groups = chatGroups_();
  const groupRows = groups.getDataRange().getValues();
  for (let i = groupRows.length - 1; i >= 1; i--) {
    const members = groupMembers_(groupRows[i])
      .filter(member => member != key);
    if (members.length < 2) {
      const groupId = 'group:' + String(groupRows[i][0]);
      const currentMessages = messages.getDataRange().getValues();
      for (let j = currentMessages.length - 1; j >= 1; j--) {
        if (String(currentMessages[j][2]) == groupId) messages.deleteRow(j + 1);
      }
      groups.deleteRow(i + 1);
    } else {
      groups.getRange(i + 1, 4).setValue(JSON.stringify(members));
      if (String(groupRows[i][2]).toLowerCase() == key) {
        groups.getRange(i + 1, 3).setValue(members[0]);
      }
    }
  }
}
```

Tambahkan pemanggilan `purgeSocialData_(targetUsername);` ke fungsi `deleteUser`, sebelum `users_().deleteRow(target.idx)`, agar posting dan chat ikut terhapus saat akun dihapus.

Chat menggunakan percakapan personal yang sudah ada dan grup dengan beberapa anggota. Semua role dapat membuat grup, mengirim pesan personal, dan mengirim pesan ke grup yang mereka ikuti. Grup disimpan di sheet `ChatGroups`; pesan grup disimpan di sheet `Messages` dengan kolom `to` bernilai `group:<id grup>`. Akses riwayat dan pengiriman grup divalidasi di Apps Script berdasarkan keanggotaan. `purgeSocialData_` menghapus pesan milik akun yang dihapus, mengeluarkan akun itu dari grup, dan menghapus grup yang tersisa kurang dari dua anggota.

Setelah endpoint ini dan semua perubahan di atas ditempel, jalankan `setup()` sekali untuk membuat sheet `Posts`, `Messages`, dan `ChatGroups`, simpan, lalu **deploy versi baru** Web App. Fungsi `setup()` hanya menambahkan sheet yang belum ada; tidak menghapus pesan personal yang sudah tersimpan di `Messages`.

## 7. Diagnostik jika login atau chat mendapat respons kosong

Frontend mengharapkan setiap request Apps Script mengembalikan JSON dengan properti `result`. Jika melihat pesan **“Apps Script tidak mengembalikan hasil”**, periksa hal berikut:

1. Fungsi `loadState(t)` harus selalu mengembalikan `{ err: 'AUTH' }` untuk token tidak valid, atau object `{ state, name, user, role }` untuk sesi aktif. Jangan menghapus `return` pada hasil sesi.
2. Fungsi `doPost(e)` harus membungkus hasil fungsi API sebagai `{ result: f.apply(null, req.args || []) }`.
3. Pastikan `loadState`, `socialFeed`, `recentPosts`, `createPost`, `chatContacts`, `chatHistory`, `sendChat`, `chatGroups`, `createChatGroup`, `groupHistory`, dan `sendGroupChat` terdaftar pada `api_()` dan masing-masing mengembalikan hasil.
4. Jalankan `setup()` agar sheet `Posts`, `Messages`, dan `ChatGroups` tersedia, lalu deploy sebagai **versi baru**. Menyimpan kode saja tidak memperbarui deployment Web App.
5. Keluar dari aplikasi, login kembali, lalu buka Feed/Chat. Token yang lama dapat tidak berlaku setelah akun dihapus atau sesi berakhir.

## 8. Feed dan chat terasa lambat

Frontend menampilkan postingan/pesan secara langsung sambil menyimpan ke Apps Script. Feed mengambil postingan terbaru dari `recentPosts` setiap 10 detik saat tab Feed terbuka; endpoint ini hanya membaca maksimal 100 baris terakhir dari sheet `Posts`, tidak memindai sheet `Data`. Tambahkan `recentPosts` pada `api_()` dan tempel fungsi di atas sebelum deploy versi baru.

Riwayat chat diperbarui setiap 6 detik ketika halaman terlihat. Endpoint `chatHistory` dan `groupHistory` saat ini membaca sheet `Messages`; bila sheet tumbuh besar, request akan makin lambat. Jangan menambahkan cache berumur panjang pada riwayat karena dapat membuat pesan baru terlambat terlihat. Batasi sementara jumlah baris Messages hanya jika Anda juga menerapkan pagination agar pesan lama tidak tersembunyi.

Endpoint sesi dapat diuji tanpa kredensial dengan memanggil `loadState` menggunakan token kosong: respons yang diharapkan adalah `{ "result": { "err": "AUTH" } }`, bukan respons tanpa `result`.

**Catatan riwayat Git:** data lama pernah ada di commit `388de7a`. Menghapusnya dari versi terbaru tidak menghapus data dari riwayat commit. Jika repository dapat diakses publik, data di commit lama masih dapat dilihat melalui history; penghapusan dari history memerlukan perubahan riwayat Git dan koordinasi sebelum force-push.
