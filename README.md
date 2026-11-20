# Rapot Digital ACS v13 — GitHub Pages (v13-pages)

Frontend v13 yang diadaptasi supaya bisa di-host di **GitHub Pages**, dengan backend
tetap di **Google Apps Script** (Kode.gs + Google Sheets). Isi folder ini:

| File | Isi |
|---|---|
| `index.html` | Frontend v13 hasil adaptasi (2.652 baris). Login tetap ada — **bukan** versi bypass dari artifact preview. |
| `Kode.gs` | Backend v12.1 persis + tambahan `doPost` di akhir file. `doGet` dan semua fungsi lama **tidak diubah**. |
| `README.md` | File ini. |

Repo: `badmanlabs-stack/Rapot-acs` (public, branch `main`).
URL setelah Pages aktif: `https://badmanlabs-stack.github.io/Rapot-acs/`

## Arsitektur split

```
GitHub Pages (statis)                Apps Script (backend + database)
┌─────────────────────┐              ┌──────────────────────────────────┐
│ index.html          │  POST JSON   │ Kode.gs (doPost baru)            │
│  - panggilBackend()  │ ───────────► │  - whitelist 36 fungsi frontend  │
│  - panggilServer()   │  {fn, args}  │  - semua fungsi v12.1 (tidak     │
│  - 8 halaman v13    │ ◄─────────── │    berubah)                      │
│  - tema dinamis     │ {ok, result} │ Google Sheets (database)         │
└─────────────────────┘              └──────────────────────────────────┘
```

## Yang berubah dari v13

### `index.html` — hanya transport server yang diganti
1. **Baru: `panggilBackend(namaFungsi, argsArray, signal)`** — `fetch` POST ke URL
   `/exec` dengan body JSON `{fn, args}` dan header `Content-Type: text/plain`
   (supaya request tergolong "simple request" dan tidak kena preflight CORS).
   Mengembalikan Promise: sukses → `data.result`; gagal transport / HTTP ≠ 200 /
   bukan JSON / `{ok:false}` → `reject(Error)`.
2. **`panggilServer` ditulis ulang bagian transportnya saja** — tanda tangan
   `panggilServer(namaFungsi, argsArray, opsi)` dan 13 call site-nya **tidak berubah**.
   Perilaku lama dipertahankan: timeout 90 detik (kini via `AbortController`),
   auto-retry 1x untuk read, loading overlay, dan dialog "Coba Lagi" manual untuk
   write/delete.
3. **27 panggilan langsung `google.script.run` diubah** menjadi
   `panggilBackend("namaFungsi", [args]).then(handler, gagalServer)` — isi success
   handler dan failure handler (`gagalServer`) dipertahankan apa adanya. Panggilan
   ini memang sengaja tidak memakai loading overlay / timeout di v13, jadi
   perilakunya juga dipertahankan.
4. **Login tidak diubah.** `prosesLogin` tetap lewat `panggilServer`, sesi tetap
   disimpan di `localStorage` (`ACS_SESSION`), wali tetap hanya bisa melihat
   anaknya sendiri (ditegakkan server via `bolehAksesAtlet`).
5. **Tidak ada lagi referensi `google.script.*`** di file ini (terverifikasi 0).
   Satu-satunya URL Google yang tersisa adalah CDN font (Google Fonts) dan
   backend `/exec` — keduanya memang harus begitu.

### `Kode.gs` — hanya `doPost` yang ditambah
- File ini = salinan persis `v13/Kode.gs` (md5 `c03a133cf85047fbc9d7452ab954237f`)
  **ditambah satu blok di akhir**: `FUNGSI_DIIZINKAN` (whitelist 36 fungsi) +
  `doPost(e)`.
- `doPost` membaca `e.postData.contents` sebagai JSON `{fn, args}`, menolak
  fungsi di luar whitelist, menjalankan fungsi yang diminta, dan mengembalikan
  `ContentService` JSON `{ok:true, result}` atau `{ok:false, error}`.
- Error aplikasi dari backend (mis. `{error:"SESI_HABIS"}`) tetap dikembalikan
  sebagai `result` normal — bukan sebagai `{ok:false}` — supaya `tanganiResError()`
  di frontend tetap bekerja seperti sebelumnya.
- Helper internal (`validasiToken`, `bolehAksesAtlet`, `denganLock`, dsb.)
  **tidak** diekspos lewat `doPost`.

### Yang TIDAK berubah
- Semua logika UI, 8 halaman, tema dinamis, validasi waktu, anti-duplikat,
  dan aturan hapus (konfirmasi dulu) — sama persis dengan v13.
- Tidak ada data mock/contoh di `index.html` (terverifikasi 0).
- File asli `v13/index.html` dan `v13/Kode.gs` tidak disentuh.

## Langkah deploy (untuk Ndaru)

### 1. Push ke GitHub
Push isi folder ini ke repo `badmanlabs-stack/Rapot-acs`, branch `main`, di root:
- `index.html`
- `README.md`
- `Kode.gs` (disimpan di repo sebagai sumber deploy Apps Script; GitHub Pages
  mengabaikannya karena bukan HTML — yang disajikan Pages hanya `index.html`).

### 2. Aktifkan GitHub Pages
Repo → **Settings → Pages → Build and deployment → Source: Deploy from a branch**
→ Branch: `main`, Folder: `/(root)` → **Save**. Tunggu ~1 menit, lalu buka
`https://badmanlabs-stack.github.io/Rapot-acs/`.

### 3. Deploy backend baru di Apps Script
1. Buka project Apps Script **RAPOT ACS**.
2. Ganti seluruh isi `Kode.gs` dengan file `Kode.gs` dari folder ini.
   (File `index.html` di project Apps Script **tidak perlu** diganti — `doGet`
   tetap menyajikan halaman lama di URL `/exec` sebagai fallback; frontend baru
   tidak memakainya.)
3. **Deploy → Manage deployments → edit deployment live → Version: New version
   → Deploy.** URL `/exec` **tetap sama** — frontend Pages sudah menunjuk ke URL itu.
4. Pastikan pengaturan deployment: **Execute as: Me** dan **Who has access:
   Anyone (Siapa saja)**. Kalau aksesnya "Anyone with Google account", browser
   akan memblokir panggilan dari Pages (CORS + login Google).

### 4. APK
APK lama (`id.acs.rapotdigital`) menunjuk ke URL `/exec`, yang setelah langkah 3
tetap menyajikan UI lama v12.1 (bukan v13). Supaya APK membuka UI baru,
**build ulang APK dengan URL `https://badmanlabs-stack.github.io/Rapot-acs/`**.

## Tes setelah deploy (di HP + laptop)
1. Buka URL Pages → halaman login muncul (bukan langsung masuk — login tetap ada).
2. Login pelatih → dashboard tampil (57 atlet), buka rapot salah satu atlet.
3. Login wali → hanya rapot anaknya yang terbuka; tombol pelatih tidak ada.
4. Coba simpan waktu 50m **12 detik** → harus **ditolak** (di bawah minimum 15 detik).
5. Coba hapus satu baris (presensi/SPP/waktu) → harus minta konfirmasi dulu.
6. Buka DevTools → Console (di laptop, F12) → tidak boleh ada error merah saat
   login dan pindah halaman. Kalau muncul error `CORS` / `Failed to fetch`:
   cek langkah 3 poin 4 (akses deployment harus "Anyone") dan pastikan version
   baru sudah di-deploy (bukan version lama tanpa `doPost`).
7. Kalau muncul alert **"Backend tidak mengembalikan JSON (apakah Kode.gs dengan
   doPost sudah di-deploy sebagai version baru?)"** → artinya deployment live
   masih Kode.gs lama. Ulangi langkah 3 (Deploy → Manage deployments → New version).

Daftar tes HP lengkap 12 poin ada di `v13/CATATAN-PERUBAHAN-v13.md` §5.

## Rollback
- Kalau Pages bermasalah, URL `/exec` lama tetap bisa dibuka dan menyajikan UI
  v12.1 seperti sebelumnya (karena `doGet` tidak diubah). Data di Sheets tidak
  tersentuh oleh pergantian hosting ini.
- Kalau `doPost` bermasalah, kembalikan `Kode.gs` di Apps Script ke versi
  sebelumnya (Deploy → Manage deployments → pilih version lama) — frontend
  Pages akan menampilkan error koneksi, sementara `/exec` lama tetap jalan.

## Yang belum bisa diverifikasi dari VM (harus dites setelah deploy)
- **CORS fetch dari Pages ke `/exec` di browser sungguhan** — VM bisa mencapai
  `/exec` (GET 200, POST 200), tapi header CORS hanya bisa dibuktikan browser.
  Lihat tes poin 6 di atas.
- **Login + semua 36 fungsi terhadap data asli** — VM tidak login ke akun asli.
- **Kirim WA (Fonnte), scan QR presensi, import Excel/PDF** — sama seperti v13,
  belum dites live; jalurnya tidak berubah oleh adaptasi ini.
- **Tanggal dari backend** — backend sudah memformat tanggal menjadi string
  di sisi server (`Utilities.formatDate`), jadi transport JSON aman; tetap
  verifikasi tampilan tanggal di halaman rapot & riwayat presensi dengan data asli.
