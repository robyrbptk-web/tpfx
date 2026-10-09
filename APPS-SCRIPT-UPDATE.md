# Pembaruan Google Apps Script

Versi backend lengkap yang menggabungkan fitur di bawah tersedia di [AppsScript-TPFX-Complete.gs](./AppsScript-TPFX-Complete.gs). File sumber yang dibagikan di chat tidak diubah; file baru tersebut dibuat sebagai salinan lengkap dengan perubahan.

**Penting sebelum mengganti Code.gs:** lampiran sumber yang diterima berisi placeholder kosong `DATABASE_DATA`. Jika Code.gs yang sedang terpasang memuat daftar kontak asli, salin isi array `DATABASE_DATA` yang asli ke file lengkap ini terlebih dahulu agar daftar itu tidak hilang setelah deploy. Jangan menyalin ulang `setup()` sebelum memeriksa akun yang sudah ada; fungsi `setup()` membuat sheet tambahan dan menjaga nama/password akun lama.

Untuk database baru yang belum mempunyai akun manager, isi Script Properties `INITIAL_BM_PASSWORD` dan `INITIAL_DM_PASSWORD` dengan password sementara minimal 12 karakter sebelum menjalankan `setup()`, kemudian hapus kedua property setelah akun dibuat dan atur password melalui proses yang aman. Jangan simpan password di repository.

File ini tetap memuat penjelasan perubahan dan endpoint per bagian sebagai referensi. Untuk instalasi penuh, gunakan file `.gs` lengkap; tidak perlu menempelkan endpoint satu per satu dari bagian bawah.

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

## 5. Database calon nasabah

`database.html` memanggil endpoint `databaseData(TOKEN)` yang memerlukan sesi login. Token dibaca dari `localStorage` yang sama dengan workspace. Tab **Database** di navbar membuka halaman ini; jika sesi kedaluwarsa, halaman menyediakan tautan login. Data kontak tidak diletakkan di frontend atau Git. Pastikan sumber `DATABASE_DATA` di Apps Script berisi **seluruh** daftar calon nasabah, termasuk semua kategori, sebelum mengganti implementasi `databaseData(t)` di bawah. Fungsi menggabungkan record dari semua kelompok sumber lalu membaginya bergiliran kepada **setiap** akun di sheet `Users` yang statusnya `Y` dan role-nya `BDO`; akun baru (termasuk Marsiana, Yudha, dan BDO baru lainnya) otomatis masuk jika memenuhi kedua syarat. Semua akun login dapat melihat semua kelompok, sesuai pilihan pengguna. Setiap kelompok dikembalikan dengan nama akun BDO di sheet `Users`, sehingga bukan hanya nama dari kelompok BDO lama yang muncul. Jumlah kontak tiap BDO berbeda paling banyak satu, dan pemeriksaan menghentikan respons bila kontak kurang untuk memberi setiap BDO minimal satu atau bila nomor kosong/duplikat ditemukan. Label tab menampilkan nama BDO dan jumlah kontak untuk memudahkan verifikasi.

```javascript
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
```

Penyimpanan status kontak pada `database.html` memakai nomor kontak sebagai kunci lokal, agar status mengikuti kontak jika pembagian otomatis berubah setelah jumlah BDO aktif bertambah atau berkurang. Daftar nomor hanya diterima oleh browser setelah sesi berhasil divalidasi; jangan membuat endpoint publik untuk database ini.

### Bagikan lead yang diinput BDO/Sales

Lead pada menu **Leads** disimpan per akun di sheet `Data`, dengan baris berformat `[username, timestamp, potongan_json_state]`. Tambahkan endpoint berikut agar semua akun aktif ber-role `BDO` atau `SALES` dapat melihat dan menghubungi lead seluruh anggota tim. Endpoint hanya membaca data dan tetap memerlukan sesi login.

```javascript
function teamLeads(t) {
  const username = user_(t);
  const caller = username && findUser_(username);
  if (!caller || String(caller[3]).toUpperCase() != 'Y') {
    return { err: 'AUTH' };
  }
  if (!['BDO', 'SALES'].includes(String(caller[4]).toUpperCase())) {
    return { err: 'Akses ditolak' };
  }

  const owners = {};
  users_().getDataRange().getValues().slice(1).forEach(r => {
    const role = String(r[4] || '').toUpperCase();
    if (String(r[3]).toUpperCase() == 'Y' && (role == 'BDO' || role == 'SALES')) {
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
    const state = JSON.parse(savedRows.map(r => String(r[2] || '')).join(''));
    if (!state || !Array.isArray(state.leads)) {
      throw new Error('Format state lead tidak valid untuk akun ' + key);
    }

    state.leads.forEach(lead => {
      if (!lead || typeof lead != 'object') {
        throw new Error('Format lead tidak valid untuk akun ' + key);
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
```

Setelah menambahkan fungsi ke `api_()` dan Code.gs, deploy versi Web App terbaru. Di workspace, buka **Leads → Lead Tim** untuk mencari seluruh lead BDO/Sales dan memakai tombol **WhatsApp**. Tombol tersebut membuka percakapan dengan nomor lead; akun lain tidak dapat mengubah atau menghapus data lead pemiliknya.

## 6. Feed sosial dan chat di `index.html`

Tambahkan sheet penyimpanan berikut ke fungsi `setup()` setelah pembuatan sheet `Data`:

```javascript
  if (!b.getSheetByName('Posts')) {
    b.insertSheet('Posts').appendRow([
      'id', 'username', 'nama', 'text', 'ts', 'photoUrl'
    ]);
  }
  const postsSheet = b.getSheetByName('Posts');
  if (postsSheet.getLastColumn() < 6) postsSheet.getRange(1, 6).setValue('photoUrl');

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
  const usersSheet = b.getSheetByName('Users');
  if (usersSheet.getLastColumn() < 6) usersSheet.getRange(1, 6).setValue('photoUrl');
```

`Users` tetap memakai kolom A=username, C=nama, D=status, E=role; kolom F yang ditambahkan hanya menyimpan URL foto profil. `Posts` tetap mempertahankan kolom lama dan menambahkan URL foto di kolom F. `setup()` menambahkan sheet/kolom saja dan tidak menghapus data lama.

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
      lastTs: 0,
      lastFrom: ''
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
      contact.lastFrom = from;
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
      lastMessages[id] = {
        text: String(r[3]),
        ts: Number(r[4]) || 0,
        from: String(r[1] || '')
      };
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

### Fitur foto, interaksi feed, chat global, dan hapus grup

Tambahkan endpoint berikut ke `api_()` (jangan hapus endpoint lama): `updateProfilePhoto`, `deletePost`, `togglePostLike`, `togglePostPin`, `postComments`, `addPostComment`, `deleteChatGroup`, `globalChatHistory`, dan `sendGlobalChat`. Jalankan `setup()` satu kali untuk membuat `PostLikes`, `PostPins`, `PostComments`, serta menambahkan header kolom foto bila belum ada. Foto feed/profil diunggah lewat endpoint `uploadImg` yang sudah ada; endpoint baru hanya menyimpan URL Google Drive hasil upload.

#### Endpoint interaksi feed

Tempel helper dan endpoint ini di Code.gs. Ganti fungsi `socialFeed`, `recentPosts`, dan `createPost` versi lama dengan implementasi di bawah. Postingan hanya dapat dihapus oleh pemiliknya; BM/DM dapat menyematkan dan melepas sematan postingan. Komentar dan suka tersedia untuk semua akun aktif.

```javascript
function updateProfilePhoto(t, photoUrl) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };
  photoUrl = String(photoUrl || '').trim();
  if (!/^https:\/\/.+/i.test(photoUrl) || photoUrl.length > 2000) {
    return { err: 'URL foto profil tidak valid' };
  }
  users_().getRange(account.idx, 6).setValue(photoUrl);
  return { ok: 1, photoUrl: photoUrl };
}

function socialUsers_() {
  const users = {};
  users_().getDataRange().getValues().slice(1).forEach(r => {
    const key = String(r[0] || '').toLowerCase();
    if (key) users[key] = {
      nama: String(r[2] || r[0]),
      role: String(r[4] || 'BDO').toUpperCase(),
      photoUrl: String(r[5] || '')
    };
  });
  return users;
}

function enrichSocialPosts_(rows, viewer) {
  const likes = {};
  const likedByViewer = {};
  postsLikes_().getDataRange().getValues().slice(1).forEach(r => {
    const id = String(r[0] || '');
    if (!id) return;
    likes[id] = (likes[id] || 0) + 1;
    if (String(r[1] || '').toLowerCase() == String(viewer).toLowerCase()) {
      likedByViewer[id] = true;
    }
  });
  const pins = {};
  postsPins_().getDataRange().getValues().slice(1).forEach(r => {
    pins[String(r[0] || '')] = true;
  });
  const commentCounts = {};
  postComments_().getDataRange().getValues().slice(1).forEach(r => {
    const id = String(r[1] || '');
    if (id) commentCounts[id] = (commentCounts[id] || 0) + 1;
  });
  const users = socialUsers_();
  return rows.map(post => {
    const profile = users[String(post.username).toLowerCase()] || {};
    return Object.assign({}, post, {
      nama: post.nama || profile.nama || post.username,
      role: profile.role || '',
      profilePhoto: profile.photoUrl || '',
      likes: likes[String(post.id)] || 0,
      liked: !!likedByViewer[String(post.id)],
      pinned: !!pins[String(post.id)],
      commentsCount: commentCounts[String(post.id)] || 0
    });
  }).sort((a, b) =>
    Number(b.pinned) - Number(a.pinned) || Number(b.ts) - Number(a.ts)
  );
}

function postsLikes_() {
  const sheet = ss_().getSheetByName('PostLikes');
  if (!sheet) throw new Error('Sheet PostLikes belum ada. Jalankan setup().');
  return sheet;
}

function postsPins_() {
  const sheet = ss_().getSheetByName('PostPins');
  if (!sheet) throw new Error('Sheet PostPins belum ada. Jalankan setup().');
  return sheet;
}

function postComments_() {
  const sheet = ss_().getSheetByName('PostComments');
  if (!sheet) throw new Error('Sheet PostComments belum ada. Jalankan setup().');
  return sheet;
}

function socialPostRows_(limit) {
  const sheet = posts_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const startRow = Math.max(2, lastRow - limit + 1);
  const selected = {};
  sheet.getRange(startRow, 1, lastRow - startRow + 1, 6).getValues()
    .filter(r => String(r[3] || '').trim() || String(r[5] || '').trim())
    .forEach(r => { selected[String(r[0])] = r; });
  const pinnedIds = postsPins_().getDataRange().getValues().slice(1)
    .map(r => String(r[0] || '')).filter(Boolean);
  pinnedIds.forEach(id => {
    if (selected[id]) return;
    const match = sheet.getRange(2, 1, lastRow - 1, 1)
      .createTextFinder(id).matchEntireCell(true).findNext();
    if (match) selected[id] = sheet.getRange(match.getRow(), 1, 1, 6).getValues()[0];
  });
  return Object.keys(selected).map(id => {
    const r = selected[id];
    return {
      id: String(r[0]),
      username: String(r[1]),
      nama: String(r[2]),
      text: String(r[3] || ''),
      ts: Number(r[4]) || 0,
      imageUrl: String(r[5] || ''),
      kind: 'post'
    };
  });
}

function socialFeed(t) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };
  const profiles = socialUsers_();
  const feed = enrichSocialPosts_(socialPostRows_(100), username);
  const activeUsers = {};
  users_().getDataRange().getValues().slice(1).forEach(r => {
    if (String(r[3]).toUpperCase() == 'Y') {
      activeUsers[String(r[0]).toLowerCase()] = true;
    }
  });
  const chunks = {};
  data_().getDataRange().getValues().slice(1).forEach(r => {
    const key = String(r[0] || '').toLowerCase();
    if (profiles[key] && activeUsers[key]) (chunks[key] = chunks[key] || []).push(r);
  });
  const labels = {
    call: 'Mencatat panggilan', wa: 'Menghubungi calon nasabah',
    meet: 'Melakukan pertemuan', post: 'Mencatat aktivitas promosi',
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
      const profile = profiles[key];
      feed.push({
        id: 'activity-' + key + '-' + String(activity.id || ts),
        username: savedRows[0][0], nama: profile.nama,
        role: profile.role, profilePhoto: profile.photoUrl,
        text: labels[activity.type] || 'Mencatat aktivitas penjualan',
        ts: ts, kind: 'activity'
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
  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };
  return { rows: enrichSocialPosts_(socialPostRows_(100), username) };
}

function createPost(t, text, imageUrl) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };
  text = String(text || '').trim();
  imageUrl = String(imageUrl || '').trim();
  if (!text && !imageUrl) return { err: 'Isi postingan atau foto wajib diisi' };
  if (text.length > 500) return { err: 'Postingan maksimal 500 karakter' };
  if (imageUrl && (!/^https:\/\/.+/i.test(imageUrl) || imageUrl.length > 2000)) {
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
  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };
  postId = String(postId || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = posts_();
    const rows = sheet.getDataRange().getValues();
    const index = rows.findIndex((r, i) => i > 0 && String(r[0]) == postId);
    if (index < 1) return { err: 'Postingan tidak ditemukan' };
    if (String(rows[index][1]).toLowerCase() != String(username).toLowerCase()) {
      return { err: 'Anda hanya dapat menghapus postingan sendiri' };
    }
    sheet.deleteRow(index + 1);
    [
      { sheet: postsLikes_(), column: 1 },
      { sheet: postsPins_(), column: 1 },
      { sheet: postComments_(), column: 2 }
    ].forEach(meta => {
      const values = meta.sheet.getDataRange().getValues();
      for (let i = values.length - 1; i >= 1; i--) {
        if (String(values[i][meta.column - 1]) == postId) {
          meta.sheet.deleteRow(i + 1);
        }
      }
    });
  } finally {
    lock.releaseLock();
  }
  return { ok: 1 };
}

function togglePostLike(t, postId) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };
  postId = String(postId || '').trim();
  const post = posts_().getDataRange().getValues().slice(1)
    .some(r => String(r[0]) == postId);
  if (!post) return { err: 'Postingan tidak ditemukan' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = postsLikes_();
    const values = sheet.getDataRange().getValues();
    const index = values.findIndex((r, i) => i > 0 &&
      String(r[0]) == postId &&
      String(r[1]).toLowerCase() == String(username).toLowerCase());
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
  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };
  if (!['BM', 'DM'].includes(String(account[4]).toUpperCase())) {
    return { err: 'Hanya BM/DM yang dapat menyematkan postingan' };
  }
  postId = String(postId || '').trim();
  const exists = posts_().getDataRange().getValues().slice(1)
    .some(r => String(r[0]) == postId);
  if (!exists) return { err: 'Postingan tidak ditemukan' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = postsPins_();
    const values = sheet.getDataRange().getValues();
    const index = values.findIndex((r, i) => i > 0 && String(r[0]) == postId);
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
  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };
  postId = String(postId || '').trim();
  const exists = posts_().getDataRange().getValues().slice(1)
    .some(r => String(r[0]) == postId);
  if (!exists) return { err: 'Postingan tidak ditemukan' };
  const users = socialUsers_();
  const rows = postComments_().getDataRange().getValues().slice(1)
    .filter(r => String(r[1]) == postId)
    .slice(-100)
    .map(r => {
      const commenter = users[String(r[2]).toLowerCase()] || {};
      return {
        id: String(r[0]), username: String(r[2]),
        nama: commenter.nama || String(r[2]), photoUrl: commenter.photoUrl || '',
        text: String(r[3]), ts: Number(r[4]) || 0
      };
    });
  return { rows: rows };
}

function addPostComment(t, postId, text) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };
  postId = String(postId || '').trim();
  text = String(text || '').trim();
  if (!text) return { err: 'Komentar tidak boleh kosong' };
  if (text.length > 500) return { err: 'Komentar maksimal 500 karakter' };
  const exists = posts_().getDataRange().getValues().slice(1)
    .some(r => String(r[0]) == postId);
  if (!exists) return { err: 'Postingan tidak ditemukan' };
  postComments_().appendRow([Utilities.getUuid(), postId, username, text, Date.now()]);
  return { ok: 1 };
}
```

Pada `loadState(t)`, tambahkan `photoUrl: String(account[5] || '')` ke object sukses yang dikembalikan; gunakan nama variabel akun aktual yang sudah dipakai fungsi tersebut (`u`/`userRow` jika bukan `account`). Jalankan `setup()` supaya kolom F di `Users` tersedia. Feed menampilkan role BM/DM di samping nama, avatar profil, foto postingan, tombol suka/komentar, dan sematan.

Di `purgeSocialData_(username)`, setelah mendapatkan `postRows`, simpan ID postingan akun itu sebelum baris dihapus:

```javascript
const ownedPostIds = {};
postRows.slice(1).forEach(r => {
  if (String(r[1]).toLowerCase() == key) ownedPostIds[String(r[0])] = true;
});
```

Di akhir fungsi tersebut, sebelum `}`, bersihkan interaksi postingan agar penghapusan akun tidak meninggalkan suka, pin, atau komentar yatim:

```javascript
[
  { sheet: postsLikes_(), postColumn: 1, userColumn: 2 },
  { sheet: postsPins_(), postColumn: 1, userColumn: 2 },
  { sheet: postComments_(), postColumn: 2, userColumn: 3 }
].forEach(meta => {
  const values = meta.sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    const belongsToDeletedUser =
      String(values[i][meta.userColumn - 1]).toLowerCase() == key;
    const belongsToDeletedPost = !!ownedPostIds[String(values[i][meta.postColumn - 1])];
    if (belongsToDeletedUser || belongsToDeletedPost) meta.sheet.deleteRow(i + 1);
  }
});
```

#### Hapus grup chat dan chat global

Tambahkan endpoint berikut. Hapus grup dibatasi ke pembuat grup atau BM/DM, dan riwayat grup ikut dihapus. Chat global memakai `Messages.to = 'global'`; semua akun aktif dapat membaca dan membalas di satu ruang yang sama.

```javascript
function deleteChatGroup(t, groupId) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };
  groupId = String(groupId || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const groups = chatGroups_();
    const values = groups.getDataRange().getValues();
    const index = values.findIndex((r, i) => i > 0 && String(r[0]) == groupId);
    if (index < 1) return { err: 'Grup tidak ditemukan' };
    const owner = String(values[index][2] || '');
    if (owner.toLowerCase() != String(username).toLowerCase() &&
        !['BM', 'DM'].includes(String(account[4]).toUpperCase())) {
      return { err: 'Hanya pembuat grup atau BM/DM yang dapat menghapus grup' };
    }
    const messages = messages_();
    const messageRows = messages.getDataRange().getValues();
    for (let i = messageRows.length - 1; i >= 1; i--) {
      if (String(messageRows[i][2]) == 'group:' + groupId) messages.deleteRow(i + 1);
    }
    groups.deleteRow(index + 1);
  } finally {
    lock.releaseLock();
  }
  return { ok: 1 };
}

function globalChatHistory(t) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };
  const profiles = socialUsers_();
  const rows = messages_().getDataRange().getValues().slice(1)
    .filter(r => String(r[2]) == 'global')
    .sort((a, b) => Number(a[4]) - Number(b[4]))
    .slice(-200)
    .map(r => {
      const sender = profiles[String(r[1]).toLowerCase()] || {};
      return {
        id: String(r[0]), from: String(r[1]),
        nama: sender.nama || String(r[1]), role: sender.role || '',
        photoUrl: sender.photoUrl || '', text: String(r[3]),
        ts: Number(r[4]) || 0
      };
    });
  return { rows: rows };
}

function sendGlobalChat(t, unusedTarget, text) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };
  text = String(text || '').trim();
  if (!text) return { err: 'Pesan tidak boleh kosong' };
  if (text.length > 1000) return { err: 'Pesan maksimal 1000 karakter' };
  messages_().appendRow([Utilities.getUuid(), username, 'global', text, Date.now()]);
  return { ok: 1 };
}
```

Pada `chatGroups(t)`, tambahkan properti ini ke object grup:

```javascript
owner: String(r[2] || ''),
```

Pada `chatContacts(t)`, tambahkan properti ini pada object kontak:

```javascript
photoUrl: String(r[5] || ''),
```

Di fungsi `chatHistory` dan `groupHistory`, sebelum membuat hasil pesan siapkan lookup profil:

```javascript
const profiles = socialUsers_();
```

Lalu pada `.map(r => ({ ... }))` di `chatHistory`, tambahkan properti pengirim:

```javascript
nama: (profiles[String(r[1]).toLowerCase()] || {}).nama || String(r[1]),
role: (profiles[String(r[1]).toLowerCase()] || {}).role || '',
photoUrl: (profiles[String(r[1]).toLowerCase()] || {}).photoUrl || '',
```

Di `groupHistory`, properti `nama` sudah dibuat dari lookup yang ada; pertahankan dan tambahkan hanya `role` serta `photoUrl` dengan ekspresi yang sama. Jangan menghilangkan properti pesan yang sudah dikembalikan. Pada `chatNotificationSummary(t)`, ganti fungsi lama dengan implementasi berikut; kontak sekaligus membawa foto profil dan summary global:

```javascript
function chatNotificationSummary(t) {
  const username = user_(t);
  const account = username && findUser_(username);
  if (!account || String(account[3]).toUpperCase() != 'Y') return { err: 'AUTH' };

  const userKey = String(username).toLowerCase();
  const profiles = socialUsers_();
  const contacts = {};
  users_().getDataRange().getValues().slice(1).forEach(r => {
    const peer = String(r[0] || '');
    const profile = profiles[peer.toLowerCase()] || {};
    if (String(r[3]).toUpperCase() != 'Y' || peer.toLowerCase() == userKey) return;
    contacts[peer.toLowerCase()] = {
      username: peer, nama: profile.nama || peer,
      role: profile.role || 'BDO', photoUrl: profile.photoUrl || '',
      lastText: '', lastTs: 0, lastFrom: ''
    };
  });

  const groups = {};
  chatGroups_().getDataRange().getValues().slice(1).forEach(r => {
    const id = String(r[0]);
    const members = groupMembers_(r);
    if (members.indexOf(userKey) < 0) return;
    groups[id] = {
      id: id, name: String(r[1]), owner: String(r[2] || ''),
      memberCount: members.length, lastText: '', lastTs: 0, lastFrom: ''
    };
  });
  const global = { id: 'global', lastText: '', lastTs: 0, lastFrom: '' };

  messages_().getDataRange().getValues().slice(1).forEach(r => {
    const from = String(r[1] || ''), to = String(r[2] || '');
    const ts = Number(r[4]) || 0;
    if (to == 'global') {
      if (ts >= global.lastTs) {
        global.lastText = String(r[3] || '');
        global.lastTs = ts;
        global.lastFrom = from;
      }
      return;
    }
    if (to.indexOf('group:') == 0) {
      const group = groups[to.slice(6)];
      if (group && ts >= group.lastTs) {
        group.lastText = String(r[3] || '');
        group.lastTs = ts;
        group.lastFrom = from;
      }
      return;
    }
    let peer = '';
    if (from.toLowerCase() == userKey) peer = to.toLowerCase();
    else if (to.toLowerCase() == userKey) peer = from.toLowerCase();
    const contact = contacts[peer];
    if (contact && ts >= contact.lastTs) {
      contact.lastText = String(r[3] || '');
      contact.lastTs = ts;
      contact.lastFrom = from;
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
```

### Notifikasi pesan baru

`index.html` menggunakan Notification API bawaan browser, tanpa Firebase, Supabase, atau layanan push eksternal. Di halaman Chat, pengguna harus menekan **Aktifkan notifikasi** dan menyetujui izin browser. Workspace memeriksa ringkasan chat personal/grup/global setiap 5 detik selama halamannya masih terbuka, termasuk saat tab di background; browser dapat memperlambat atau menangguhkan polling tab background. Notifikasi dan badge hanya dibuat untuk pesan masuk dari anggota lain, bukan pesan sendiri. Gunakan implementasi `chatNotificationSummary(t)` pada bagian **Hapus grup chat dan chat global** di atas; fungsi tersebut juga mengembalikan ringkasan global.

Notifikasi browser dapat muncul saat tab workspace berada di background, tetapi tidak dapat dijamin setelah tab/browser ditutup atau sistem menghentikan halaman. Itu memerlukan Web Push dengan service worker dan server pengirim push. Gunakan hosting HTTPS.

## 7. Diagnostik jika login atau chat mendapat respons kosong

Frontend mengharapkan setiap request Apps Script mengembalikan JSON dengan properti `result`. Jika melihat pesan **“Apps Script tidak mengembalikan hasil”**, periksa hal berikut:

1. Fungsi `loadState(t)` harus selalu mengembalikan `{ err: 'AUTH' }` untuk token tidak valid, atau object `{ state, name, user, role }` untuk sesi aktif. Jangan menghapus `return` pada hasil sesi.
2. Fungsi `doPost(e)` harus membungkus hasil fungsi API sebagai `{ result: f.apply(null, req.args || []) }`.
3. Pastikan fungsi lama dan baru terdaftar pada `api_()`: `loadState`, `socialFeed`, `recentPosts`, `createPost`, `deletePost`, `togglePostLike`, `togglePostPin`, `postComments`, `addPostComment`, `updateProfilePhoto`, `chatContacts`, `chatNotificationSummary`, `chatHistory`, `chatSync`, `sendChat`, `chatGroups`, `createChatGroup`, `deleteChatGroup`, `groupHistory`, `sendGroupChat`, `globalChatHistory`, dan `sendGlobalChat`.
4. Jalankan `setup()` agar sheet `Posts`, `PostLikes`, `PostPins`, `PostComments`, `Messages`, `ChatGroups`, dan `ChatSummary` tersedia, kolom foto tersedia di `Posts` dan `Users`, serta ringkasan chat lama diindeks sekali. Menyimpan kode saja tidak memperbarui deployment Web App.
5. Keluar dari aplikasi, login kembali, lalu buka Feed/Chat. Token yang lama dapat tidak berlaku setelah akun dihapus atau sesi berakhir.

## 8. Feed dan chat terasa lambat

Frontend menampilkan postingan/pesan secara langsung sambil menyimpan ke Apps Script. Feed mengambil postingan terbaru dari `recentPosts` setiap 10 detik saat tab Feed terbuka; endpoint ini hanya membaca maksimal 100 baris terakhir dari sheet `Posts`, tidak memindai sheet `Data`. Tambahkan `recentPosts` pada `api_()` dan tempel fungsi di atas sebelum deploy versi baru.

Saat halaman Chat terlihat, `chatSync` memeriksa pesan baru setiap 1,5 detik dan hanya membaca baris `Messages` yang bertambah setelah cursor terakhir; request tidak ditumpuk jika request sebelumnya belum selesai. Ringkasan percakapan dibaca dari sheet `ChatSummary`, bukan dengan memindai semua pesan setiap polling. `chatHistory`, `groupHistory`, dan `globalChatHistory` memuat 50 pesan terbaru, lalu tombol **Muat pesan sebelumnya** melakukan pagination berdasarkan nomor baris sheet. Tidak ada pemangkasan otomatis riwayat; saat grup/akun dihapus, isi pesan terkait dibersihkan tanpa menggeser baris agar cursor aktif tetap akurat. Jalankan `setup()` sekali setelah menyalin backend baru agar index `ChatSummary` dibuat dari pesan lama.

Polling Apps Script bukan push real-time: interval aktual juga dipengaruhi waktu eksekusi Google, jaringan, dan throttling browser. Tab yang ditutup tidak menerima pesan atau notifikasi. Untuk foto, endpoint `uploadImg` menyimpan gambar di Drive dengan akses tautan, mengembalikan URL `uc?export=view`, dan `updateProfilePhoto` menyimpannya di kolom F `Users`; deploy backend baru lalu jalankan ulang simpan foto bila tautan lama gagal.

Endpoint sesi dapat diuji tanpa kredensial dengan memanggil `loadState` menggunakan token kosong: respons yang diharapkan adalah `{ "result": { "err": "AUTH" } }`, bukan respons tanpa `result`.

**Catatan riwayat Git:** data lama pernah ada di commit `388de7a`. Menghapusnya dari versi terbaru tidak menghapus data dari riwayat commit. Jika repository dapat diakses publik, data di commit lama masih dapat dilihat melalui history; penghapusan dari history memerlukan perubahan riwayat Git dan koordinasi sebelum force-push.
