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
    updateProfile: updateProfile,
    deleteUser: deleteUser
  };
}
```

Jangan menghapus endpoint lain yang sudah ada.

## 3. Tambahkan fungsi backend

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

    users_().deleteRow(target.idx);
  } finally {
    lock.releaseLock();
  }

  return { ok: 1 };
}
```

Setelah menempel perubahan, simpan project Apps Script dan **deploy versi baru** pada deployment Web App yang digunakan aplikasi. Jika tidak, frontend tetap memanggil versi API lama yang belum mengenali endpoint baru.
