/**
 * =====================================================================
 *  ABSEN PKM: ABSENSI + REKAP + KETERANGAN
 *  SMP Labschool Ciracas
 * =====================================================================
 *  CARA PASANG DARI SHEET KOSONG
 *   1. Buat Google Sheet baru (kosong).
 *   2. Extensions > Apps Script.
 *   3. Tempel file ini ke Code.gs.
 *   4. Buat 2 file HTML (tanpa .html):  Index  dan  Rekap
 *   5. Pilih fungsi  setupTemplate  > Run > izinkan akses.
 *      -> Tab "Absensi" & "Keterangan" dibuat dan diisi data dummy.
 *      -> Folder Drive "Foto Absen PKM" dibuat otomatis.
 *   6. Deploy > New deployment > Web app
 *        Execute as : Me
 *        Who has access : Anyone
 *
 *  ALAMAT
 *   Absen guru : URL web app
 *   Rekap TU   : URL web app + ?page=rekap      (PIN XXXXXX)
 *
 *  UJI COBA ABSEN: gunakan nama Mahasiswa 7 atau Mahasiswa 8 (atau pegawai lain
 *  mulai urutan ke-7). Enam pegawai pertama (Mahasiswa 1 s.d. Mahasiswa 6) sudah
 *  tercatat absen datang hari ini oleh setupTemplate.
 *
 *  Menjalankan setupTemplate lagi akan MENGHAPUS isi kedua tab dan
 *  mengisi ulang data dummy. Jangan dijalankan di sheet data asli.
 *
 *  PENTING: daftar nama di DAFTAR_PEGAWAI (file ini) dan PEGAWAI
 *  (Index.html) harus SAMA PERSIS, termasuk huruf besar/kecilnya.
 * =====================================================================
 */

// ======================= KONFIGURASI =======================
const TZ               = 'Asia/Jakarta';   // dipakai langsung, tidak bergantung setelan proyek
const NAMA_SHEET_ABSEN = 'Absensi';        // di aplikasi asli: "Sheet1"
const NAMA_SHEET_KET   = 'Keterangan';
const PIN_REKAP = 'YOUR_SECRET_PIN';

// ---- Absen (kamera + lokasi) ----
// Kosongkan FOLDER_DRIVE_ID agar folder foto dibuat otomatis oleh setupTemplate.
const FOLDER_DRIVE_ID       = '';
const LATITUDE_SEKOLAH      = -6.000000; // Ganti koordinat menjadi Letak GPS sesuai keinginan Anda
const LONGITUDE_SEKOLAH     = 106.0000000; // Ganti koordinat menjadi Letak GPS sesuai keinginan Anda
const MAX_RADIUS_METER      = 20000000;  // UJI COBA: bisa dari mana saja. Produksi: 300
const BATAS_UMUR_FOTO_DETIK = 300;       // foto lebih tua dari 5 menit ditolak

// Sebuah tanggal dianggap HARI KERJA bila Senin–Jumat DAN yang hadir
// minimal sebanyak ini. Tanggal yang sepi dianggap libur.
const AMBANG_HARI_KERJA = 0.3;   // 30% pegawai
const AMBANG_MIN_ORANG  = 3;

// Jenis keterangan dan cara menghitungnya.
//   dikecualikan : hari itu dikeluarkan dari pembagi persentase
//   dihitungHadir: hari itu dihitung hadir (tapi tetap ditandai)
const JENIS_KET = {
  'Izin':       { dikecualikan: true },
  'Sakit':      { dikecualikan: true },
  'Dinas Luar': { dikecualikan: true },
  'Cuti':       { dikecualikan: true },
  'Lupa Absen': { dihitungHadir: true }
};

// Data dummy — ganti dengan DAFTAR_PEGAWAI asli saat digabung
const DAFTAR_PEGAWAI = [
  { nama: 'Mahasiswa 1',     kelompok: 'PKM' },
  { nama: 'Mahasiswa 2',     kelompok: 'PKM' },
  { nama: 'Mahasiswa 3',     kelompok: 'PKM' },
  { nama: 'Mahasiswa 4',     kelompok: 'PKM' },
  { nama: 'Mahasiswa 5',     kelompok: 'PKM' },
  { nama: 'Mahasiswa 6',     kelompok: 'PKM' },
  { nama: 'Mahasiswa 7',     kelompok: 'PKM' },
  { nama: 'Mahasiswa 8',     kelompok: 'PKM' },
  { nama: 'Mahasiswa 9',     kelompok: 'PKM' },
  { nama: 'Mahasiswa 10',    kelompok: 'PKM' }
];
// ===========================================================


// ======================= HALAMAN =======================
function doGet(e) {
  const halaman = (e && e.parameter && e.parameter.page) || 'absen';

  if (halaman === 'rekap') {
    return HtmlService.createTemplateFromFile('Rekap')
      .evaluate()
      .setTitle('Rekap Absen PKM')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Absen PKM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}


// ======================= SETUP + DATA DUMMY =======================
function setupTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---- Tab Absensi (struktur sama dengan aplikasi asli) ----
  const shA = ss.getSheetByName(NAMA_SHEET_ABSEN) || ss.insertSheet(NAMA_SHEET_ABSEN);
  shA.clear();
  const headA = ['Waktu Absen', 'Nama Guru', 'Status', 'Koordinat GPS',
                 'Link Foto Drive', 'Foto Wajah', 'Waktu Foto Diambil', 'Jeda (detik)'];
  shA.getRange(1, 1, 1, headA.length).setValues([headA]);

  // ---- Tab Keterangan ----
  const shK = ss.getSheetByName(NAMA_SHEET_KET) || ss.insertSheet(NAMA_SHEET_KET);
  shK.clear();
  const headK = ['Tanggal', 'Nama', 'Jenis', 'Catatan', 'Dicatat Pada'];
  shK.getRange(1, 1, 1, headK.length).setValues([headK]);
  shK.getRange('A:A').setNumberFormat('@');   // tanggal disimpan sebagai teks yyyy-MM-dd

  // ---- Buat data dummy ----
  const dummy = _buatDataDummy();

  if (dummy.absen.length) {
    shA.getRange(2, 1, dummy.absen.length, headA.length).setValues(dummy.absen);
    shA.getRange(2, 1, dummy.absen.length, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }
  if (dummy.ket.length) {
    shK.getRange(2, 1, dummy.ket.length, headK.length).setValues(dummy.ket);
    shK.getRange(2, 5, dummy.ket.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  }

  // ---- Rapikan tampilan ----
  [shA, shK].forEach(sh => {
    const lebar = sh.getLastColumn();
    sh.getRange(1, 1, 1, lebar)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#12447E');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, lebar);
  });

  // Hapus tab bawaan yang masih kosong (mis. "Sheet1")
  ss.getSheets().forEach(sh => {
    const n = sh.getName();
    if (n !== NAMA_SHEET_ABSEN && n !== NAMA_SHEET_KET &&
        sh.getLastRow() === 0 && ss.getSheets().length > 2) {
      ss.deleteSheet(sh);
    }
  });

  const folder = _folderFoto();

  SpreadsheetApp.flush();
  Logger.log('Selesai: ' + dummy.absen.length + ' baris absensi, ' +
             dummy.ket.length + ' keterangan dummy.');
  Logger.log('Folder foto: ' + folder.getName() + ' (' + folder.getUrl() + ')');
}

/** Folder foto: pakai FOLDER_DRIVE_ID bila diisi, selain itu buat sekali lalu diingat. */
function _folderFoto() {
  if (FOLDER_DRIVE_ID) return DriveApp.getFolderById(FOLDER_DRIVE_ID);

  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_FOTO_ID');
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* folder terhapus -> buat ulang */ }
  }
  const baru = DriveApp.createFolder('Foto Absen PKM');
  props.setProperty('FOLDER_FOTO_ID', baru.getId());
  return baru;
}

/** Membangkitkan ± 6 minggu absensi dengan pola yang realistis. */
function _buatDataDummy() {
  // Pengacak dengan benih tetap -> hasil selalu sama setiap setup
  let benih = 20260921;
  const acak = () => { benih = (benih * 1103515245 + 12345) % 2147483648; return benih / 2147483648; };
  const dua = n => String(n).padStart(2, '0');

  // Profil tiap pegawai: peluang hadir & peluang lupa absen pulang
  const profil = [
    [0.99, 0.02], [0.95, 0.05], [0.92, 0.30], [0.88, 0.05],
    [0.75, 0.10], [0.93, 0.10], [0.97, 0.03], [0.86, 0.15]
  ];

  // 30 hari kerja terakhir, sampai KEMARIN
  const hariIni = _kunci(new Date());
  const tanggal = [];
  let k = _geserKey(hariIni, -1);
  while (tanggal.length < 30) {
    const h = _hariKey(k);
    if (h !== 0 && h !== 6) tanggal.unshift(k);
    k = _geserKey(k, -1);
  }
  const idxLibur = 12;   // satu hari kerja dibuat "libur" (hanya 2 orang yang datang)

  const absen = [];
  const bolong = [];     // [nama, tanggal] hari yang dilewati -> bahan keterangan

  tanggal.forEach((tgl, ti) => {
    DAFTAR_PEGAWAI.forEach((p, pi) => {
      const [pHadir, pLupa] = profil[pi % profil.length];

      if (ti === idxLibur && pi >= 2) return;               // hari libur
      if (acak() > pHadir) { bolong.push([p.nama, tgl]); return; }

      const jamD = Utilities.parseDate(
        tgl + ' 06:' + dua(10 + Math.floor(acak() * 50)) + ':' + dua(Math.floor(acak() * 60)),
        TZ, 'yyyy-MM-dd HH:mm:ss');
      absen.push([jamD, p.nama, 'Hadir Datang', '-6.3159, 106.8730', '', '', jamD, '4 dtk']);

      if (acak() > pLupa) {
        const jamP = Utilities.parseDate(
          tgl + ' 16:' + dua(Math.floor(acak() * 45)) + ':' + dua(Math.floor(acak() * 60)),
          TZ, 'yyyy-MM-dd HH:mm:ss');
        absen.push([jamP, p.nama, 'Hadir Pulang', '-6.3159, 106.8730', '', '', jamP, '3 dtk']);
      }
    });
  });

  // Hari ini (bila hari kerja): 6 pegawai pertama (Mahasiswa 1 s.d. 6) sudah absen
  // datang. Pegawai urutan ke-7 dst (Mahasiswa 7, 8, ...) dibiarkan kosong
  // untuk uji coba absen.
  const hH = _hariKey(hariIni);
  if (hH !== 0 && hH !== 6) {
    DAFTAR_PEGAWAI.slice(0, 6).forEach(p => {
      const jamD = Utilities.parseDate(
        hariIni + ' 06:' + dua(10 + Math.floor(acak() * 50)) + ':' + dua(Math.floor(acak() * 60)),
        TZ, 'yyyy-MM-dd HH:mm:ss');
      absen.push([jamD, p.nama, 'Hadir Datang', '-6.3159, 106.8730', '', '', jamD, '4 dtk']);
    });
  }

  absen.sort((a, b) => a[0] - b[0]);   // urut waktu, seperti data asli

  // Sebagian hari bolong diberi keterangan; sisanya dibiarkan jadi "tanpa keterangan"
  const jenisList = ['Sakit', 'Izin', 'Lupa Absen', 'Dinas Luar', 'Cuti'];
  const catatan = {
    'Sakit': 'Demam, ada surat dokter', 'Izin': 'Keperluan keluarga',
    'Lupa Absen': 'Hadir, lupa absen', 'Dinas Luar': 'Workshop Kurikulum Merdeka',
    'Cuti': 'Cuti tahunan'
  };
  const ket = [];
  bolong.forEach((b, i) => {
    if (i % 6 === 5) return;   // 1 dari 6 sengaja dibiarkan tanpa keterangan
    const j = jenisList[i % jenisList.length];
    ket.push([b[1], b[0], j, catatan[j], new Date()]);
  });

  return { absen: absen, ket: ket };
}


// ======================= AKSES (PIN + TOKEN) =======================
function cekPin(pin) {
  if (String(pin) !== String(PIN_REKAP)) return { ok: false };
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('sesi_' + token, '1', 21600);   // 6 jam
  return { ok: true, token: token };
}

function _wajibToken(token) {
  if (!token || CacheService.getScriptCache().get('sesi_' + token) !== '1') {
    throw new Error('Sesi berakhir. Muat ulang halaman dan masukkan PIN lagi.');
  }
}


// ======================= UTILITAS TANGGAL =======================
// Semua perhitungan memakai kunci teks "yyyy-MM-dd" agar kebal zona waktu.
function _kunci(d) { return Utilities.formatDate(new Date(d), TZ, 'yyyy-MM-dd'); }

function _geserKey(key, n) {
  const p = key.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)).toISOString().slice(0, 10);
}

function _hariKey(key) {           // 0 = Minggu ... 6 = Sabtu
  const p = key.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay();
}

function _labelKey(key) {
  const HARI  = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const BULAN = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  const p = key.split('-').map(Number);
  return HARI[_hariKey(key)] + ', ' + p[2] + ' ' + BULAN[p[1] - 1] + ' ' + p[0];
}

function _validKey(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s)); }

function _ambang() {
  return Math.max(AMBANG_MIN_ORANG, Math.ceil(DAFTAR_PEGAWAI.length * AMBANG_HARI_KERJA));
}


// ======================= PEMBACA DATA =======================
function _sheet(nama) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nama);
  if (!sh) throw new Error('Tab "' + nama + '" tidak ditemukan. Jalankan setupTemplate dulu.');
  return sh;
}

/** Absensi dalam rentang: { nama: { tanggal: {datang, pulang} } } + jumlah orang per tanggal */
function _bacaAbsensi(mulai, selesai) {
  const sh = _sheet(NAMA_SHEET_ABSEN);
  const last = sh.getLastRow();
  const perOrang = {}, orangPerTgl = {};
  if (last < 2) return { perOrang, orangPerTgl };

  sh.getRange(2, 1, last - 1, 3).getValues().forEach(r => {
    if (!r[0] || !r[1]) return;
    const k = _kunci(r[0]);
    if (k < mulai || k > selesai) return;
    const nama = String(r[1]).trim(), status = String(r[2]);

    (orangPerTgl[k] = orangPerTgl[k] || {})[nama] = true;
    const slot = ((perOrang[nama] = perOrang[nama] || {})[k] =
                  perOrang[nama][k] || { datang: null, pulang: null });
    const t = new Date(r[0]);
    if (status.indexOf('Datang') !== -1 && (!slot.datang || t < slot.datang)) slot.datang = t;
    if (status.indexOf('Pulang') !== -1 && (!slot.pulang || t > slot.pulang)) slot.pulang = t;
  });
  return { perOrang, orangPerTgl };
}

/** Keterangan dalam rentang: { "nama|tanggal": {jenis, catatan} } */
function _bacaKeterangan(mulai, selesai) {
  const sh = _sheet(NAMA_SHEET_KET);
  const last = sh.getLastRow();
  const hasil = {};
  if (last < 2) return hasil;

  sh.getRange(2, 1, last - 1, 4).getValues().forEach(r => {
    const k = (r[0] instanceof Date) ? _kunci(r[0]) : String(r[0]).trim();
    if (!_validKey(k) || k < mulai || k > selesai) return;
    hasil[String(r[1]).trim() + '|' + k] = { jenis: String(r[2]), catatan: String(r[3] || '') };
  });
  return hasil;
}

/** Inti perhitungan, dipakai rekap maupun detail. */
function _hitung(mulai, selesai) {
  if (!_validKey(mulai) || !_validKey(selesai)) throw new Error('Format tanggal tidak valid.');

  const hariIni = _kunci(new Date());
  const { perOrang, orangPerTgl } = _bacaAbsensi(mulai, selesai);
  const ket = _bacaKeterangan(mulai, selesai);

  const ambang = _ambang();
  const hariKerja = Object.keys(orangPerTgl).filter(k => {
    const h = _hariKey(k);
    return k <= hariIni && h !== 0 && h !== 6 && Object.keys(orangPerTgl[k]).length >= ambang;
  }).sort();

  return { perOrang, ket, hariKerja, setKerja: new Set(hariKerja) };
}


// ======================= API UNTUK HALAMAN =======================
function ambilRekap(mulai, selesai, token) {
  _wajibToken(token);
  const H = _hitung(mulai, selesai);

  const baris = DAFTAR_PEGAWAI.map(p => {
    const milik = H.perOrang[p.nama] || {};
    const c = { hadir: 0, lupa: 0, izin: 0, sakit: 0, dinas: 0, cuti: 0, alpa: 0,
                tanpaPulang: 0, lembur: 0 };

    H.hariKerja.forEach(k => {
      const a = milik[k];
      if (a) {
        c.hadir++;
        if (a.datang && !a.pulang) c.tanpaPulang++;
        return;
      }
      const kt = H.ket[p.nama + '|' + k];
      if (!kt)                         c.alpa++;
      else if (kt.jenis === 'Lupa Absen') c.lupa++;
      else if (kt.jenis === 'Izin')       c.izin++;
      else if (kt.jenis === 'Sakit')      c.sakit++;
      else if (kt.jenis === 'Dinas Luar') c.dinas++;
      else if (kt.jenis === 'Cuti')       c.cuti++;
      else                                c.alpa++;   // jenis tak dikenal -> anggap tanpa keterangan
    });

    Object.keys(milik).forEach(k => { if (!H.setKerja.has(k)) c.lembur++; });

    const pembagi = H.hariKerja.length - (c.izin + c.sakit + c.dinas + c.cuti);
    c.persen = pembagi > 0 ? Math.round(((c.hadir + c.lupa) / pembagi) * 100) : null;

    return Object.assign({ nama: p.nama, kelompok: p.kelompok }, c);
  });

  return { mulai, selesai, hariKerja: H.hariKerja.length, baris };
}

function ambilDetail(nama, mulai, selesai, token) {
  _wajibToken(token);
  const H = _hitung(mulai, selesai);
  const milik = H.perOrang[nama] || {};
  const jam = d => d ? Utilities.formatDate(d, TZ, 'HH:mm') : '';

  // Tampilkan hari kerja + hari di luar hari kerja yang ia hadiri (lembur)
  const semua = Array.from(new Set(H.hariKerja.concat(Object.keys(milik)))).sort().reverse();

  const hari = semua.map(k => {
    const a = milik[k];
    const kerja = H.setKerja.has(k);
    const kt = H.ket[nama + '|' + k] || null;
    let status;
    if (a)       status = (a.datang && a.pulang) ? 'lengkap' : (a.datang ? 'tanpa-pulang' : 'tanpa-datang');
    else if (kt) status = 'keterangan';
    else         status = 'alpa';

    return {
      tanggal: k, label: _labelKey(k), hariKerja: kerja, status,
      jamDatang: a ? jam(a.datang) : '', jamPulang: a ? jam(a.pulang) : '',
      jenis: kt ? kt.jenis : '', catatan: kt ? kt.catatan : ''
    };
  });

  return { nama, hari };
}

function simpanKeterangan(token, data) {
  _wajibToken(token);
  data = data || {};
  const nama = String(data.nama || '').trim();
  const tgl = String(data.tanggal || '').trim();
  const jenis = String(data.jenis || '').trim();
  const catatan = String(data.catatan || '').trim().slice(0, 200);

  if (!DAFTAR_PEGAWAI.some(p => p.nama === nama)) throw new Error('Nama pegawai tidak dikenal.');
  if (!_validKey(tgl))                         throw new Error('Tanggal tidak valid.');
  if (tgl > _kunci(new Date()))                throw new Error('Tidak bisa memberi keterangan untuk tanggal yang belum tiba.');
  if (!JENIS_KET[jenis])                       throw new Error('Pilih jenis keterangan.');

  const kunci = LockService.getScriptLock();
  kunci.waitLock(20000);
  try {
    const sh = _sheet(NAMA_SHEET_KET);
    const last = sh.getLastRow();
    let baris = -1;
    if (last >= 2) {
      const v = sh.getRange(2, 1, last - 1, 2).getValues();
      for (let i = 0; i < v.length; i++) {
        const k = (v[i][0] instanceof Date) ? _kunci(v[i][0]) : String(v[i][0]).trim();
        if (k === tgl && String(v[i][1]).trim() === nama) { baris = i + 2; break; }
      }
    }
    const isi = [[tgl, nama, jenis, catatan, new Date()]];
    if (baris > 0) sh.getRange(baris, 1, 1, 5).setValues(isi);   // perbarui
    else           sh.appendRow(isi[0]);                          // baru
    SpreadsheetApp.flush();
  } finally {
    kunci.releaseLock();
  }
  return { ok: true, pesan: jenis + ' tersimpan untuk ' + _labelKey(tgl) + '.' };
}

function hapusKeterangan(token, nama, tanggal) {
  _wajibToken(token);
  const kunci = LockService.getScriptLock();
  kunci.waitLock(20000);
  try {
    const sh = _sheet(NAMA_SHEET_KET);
    const last = sh.getLastRow();
    if (last < 2) return { ok: true, pesan: 'Tidak ada yang dihapus.' };
    const v = sh.getRange(2, 1, last - 1, 2).getValues();
    for (let i = v.length - 1; i >= 0; i--) {
      const k = (v[i][0] instanceof Date) ? _kunci(v[i][0]) : String(v[i][0]).trim();
      if (k === String(tanggal) && String(v[i][1]).trim() === String(nama).trim()) {
        sh.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { ok: true, pesan: 'Keterangan ' + _labelKey(k) + ' dihapus.' };
      }
    }
    return { ok: true, pesan: 'Keterangan tidak ditemukan.' };
  } finally {
    kunci.releaseLock();
  }
}


// =====================================================================
//                         HALAMAN ABSEN GURU
// =====================================================================

function hitungJarak(lat1, lon1, lat2, lon2) {
  const R = 6371e3, rad = x => x * Math.PI / 180;
  const dLat = rad(lat2 - lat1), dLon = rad(lon2 - lon1);
  const a = Math.sin(dLat / 2) ** 2 +
            Math.cos(rad(lat1)) * Math.cos(rad(lat2)) * Math.sin(dLon / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function simpanAbsen(data) {
  // ---------- Validasi isian ----------
  if (!data || !data.nama || !data.status) {
    return { success: false, message: 'Data tidak lengkap. Silakan coba lagi.' };
  }
  if (!DAFTAR_PEGAWAI.some(p => p.nama === data.nama)) {
    return { success: false, message: 'Nama tidak terdaftar. Pilih ulang nama Anda.' };
  }
  if (['Hadir Datang', 'Hadir Pulang'].indexOf(data.status) === -1) {
    return { success: false, message: 'Status kehadiran tidak dikenal.' };
  }
  if (typeof data.lat !== 'number' || typeof data.lng !== 'number' || isNaN(data.lat) || isNaN(data.lng)) {
    return { success: false, message: 'Lokasi GPS tidak terdeteksi. Aktifkan GPS lalu coba lagi.' };
  }

  const jarak = hitungJarak(data.lat, data.lng, LATITUDE_SEKOLAH, LONGITUDE_SEKOLAH);
  if (jarak > MAX_RADIUS_METER) {
    return { success: false, message: 'Lokasi ditolak. Jarak Anda ' + Math.round(jarak) + ' meter dari sekolah.' };
  }

  if (!data.image || String(data.image).indexOf('base64,') === -1) {
    return { success: false, message: 'Foto tidak terekam. Ambil foto terlebih dahulu.' };
  }
  if (!data.waktuFoto) {
    return { success: false, message: 'Foto tidak valid. Ambil foto langsung dari kamera.' };
  }

  // Umur foto dari selisih dua jam di HP yang sama -> kebal jam HP yang meleset
  const waktuKirim = Number(data.waktuKirim) || Date.now();
  const selisihDetik = (waktuKirim - Number(data.waktuFoto)) / 1000;
  if (selisihDetik < -60) {
    return { success: false, message: 'Waktu foto tidak wajar. Ambil ulang foto.' };
  }
  if (selisihDetik > BATAS_UMUR_FOTO_DETIK) {
    return { success: false, message: 'Foto ditolak. Foto diambil ' + Math.round(selisihDetik / 60) +
                                      ' menit lalu, harap memotret langsung saat absen.' };
  }

  // ---------- Tahap 1: unggah foto (di luar kunci, karena lambat) ----------
  let file;
  try {
    const img = String(data.image);
    const jenis = img.substring(5, img.indexOf(';'));
    const bytes = Utilities.base64Decode(img.substr(img.indexOf('base64,') + 7));
    const blob = Utilities.newBlob(bytes, jenis, 'Absen_' + data.nama + '_' + Date.now() + '.jpg');
    file = _folderFoto().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    Logger.log('Gagal unggah foto: ' + err);
    return { success: false, message: 'Gagal menyimpan foto. Periksa koneksi lalu coba lagi.' };
  }

  // ---------- Tahap 2: tulis ke sheet (dikunci agar tidak saling menimpa) ----------
  const kunci = LockService.getScriptLock();
  try {
    kunci.waitLock(30000);
  } catch (err) {
    try { file.setTrashed(true); } catch (e) {}
    return { success: false, message: 'Server sedang sibuk (banyak yang absen bersamaan). Tunggu 10 detik lalu coba lagi.' };
  }

  let barisBaru = 0;
  const sekarang = new Date();
  try {
    const sh = _sheet(NAMA_SHEET_ABSEN);
    const hariIni = _kunci(sekarang);
    const last = sh.getLastRow();

    // Cegah data ganda (150 baris terakhir cukup untuk satu hari)
    if (last > 1) {
      const awal = Math.max(2, last - 149);
      const cek = sh.getRange(awal, 1, last - awal + 1, 7).getValues();
      for (let i = cek.length - 1; i >= 0; i--) {
        const r = cek[i];
        if (!r[0] || String(r[1]).trim() !== data.nama) continue;

        // (a) Foto yang sama sudah tersimpan (kiriman ulang setelah koneksi putus)
        if (r[6] && Math.abs(new Date(r[6]).getTime() - Number(data.waktuFoto)) < 1500) {
          try { file.setTrashed(true); } catch (e) {}
          return { success: true, message: 'Absen Anda sudah tersimpan pukul ' +
                                           Utilities.formatDate(new Date(r[0]), TZ, 'HH:mm:ss') + '.' };
        }
        // (b) Status yang sama sudah dicatat hari ini
        if (String(r[2]).trim() === data.status && _kunci(r[0]) === hariIni) {
          try { file.setTrashed(true); } catch (e) {}
          return { success: false, message: 'Anda sudah absen "' + data.status + '" hari ini pukul ' +
                                            Utilities.formatDate(new Date(r[0]), TZ, 'HH:mm') + '.' };
        }
      }
    }

    sh.appendRow([
      sekarang,
      data.nama,
      data.status,
      data.lat + ', ' + data.lng,
      file.getUrl(),
      '=IMAGE("https://lh3.googleusercontent.com/d/' + file.getId() + '")',
      new Date(Number(data.waktuFoto)),
      Math.round(selisihDetik) + ' dtk'
    ]);
    SpreadsheetApp.flush();
    barisBaru = sh.getLastRow();
  } catch (err) {
    Logger.log('Gagal tulis sheet: ' + err);
    try { file.setTrashed(true); } catch (e) {}
    return { success: false, message: 'Gagal menyimpan data absen. Silakan coba lagi.' };
  } finally {
    kunci.releaseLock();
  }

  // Perapian tampilan di luar kunci
  try { _sheet(NAMA_SHEET_ABSEN).setRowHeight(barisBaru, 90); } catch (e) {}

  return { success: true, message: 'Berhasil! Absen ' + data.status + ' pukul ' +
                                   Utilities.formatDate(sekarang, TZ, 'HH:mm:ss') + '.' };
}

/**
 * Status absen Senin–Jumat untuk satu orang.
 * offset 0 = minggu ini, -1 = minggu lalu, dst.
 * Status: lengkap | tanpa-pulang | tanpa-datang | kosong | sepi | nanti
 */
function statusMingguan(nama, offset) {
  try {
    const geser = Number(offset) || 0;
    const hariIni = _kunci(new Date());
    const h = _hariKey(hariIni);
    const senin = _geserKey(hariIni, -((h + 6) % 7) + geser * 7);   // Minggu ikut minggu sebelumnya
    const tgl = [0, 1, 2, 3, 4].map(i => _geserKey(senin, i));

    // Baca secukupnya dari bawah, bukan seluruh sheet
    const sh = _sheet(NAMA_SHEET_ABSEN);
    const last = sh.getLastRow();
    const perTgl = {};
    tgl.forEach(k => perTgl[k] = { orang: {}, datang: null, pulang: null });

    if (last > 1) {
      const maks = Math.max(20, DAFTAR_PEGAWAI.length * 2) * 7 * (Math.abs(geser) + 2);
      const awal = Math.max(2, last - maks + 1);
      sh.getRange(awal, 1, last - awal + 1, 3).getValues().forEach(r => {
        if (!r[0] || !r[1]) return;
        const k = _kunci(r[0]);
        if (!perTgl[k]) return;
        const n = String(r[1]).trim();
        perTgl[k].orang[n] = true;
        if (n !== String(nama).trim()) return;
        const t = new Date(r[0]), s = String(r[2]);
        if (s.indexOf('Datang') !== -1 && (!perTgl[k].datang || t < perTgl[k].datang)) perTgl[k].datang = t;
        if (s.indexOf('Pulang') !== -1 && (!perTgl[k].pulang || t > perTgl[k].pulang)) perTgl[k].pulang = t;
      });
    }

    const ambang = _ambang();
    const NAMA = ['Sen', 'Sel', 'Rab', 'Kam', 'Jum'];
    const BULAN = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
    const jam = d => d ? Utilities.formatDate(d, TZ, 'HH:mm') : '';
    const pendek = k => Number(k.slice(8)) + ' ' + BULAN[Number(k.slice(5, 7)) - 1];

    const hari = tgl.map((k, i) => {
      const x = perTgl[k];
      const jml = Object.keys(x.orang).length;
      let status;
      if (k > hariIni)              status = 'nanti';
      else if (x.datang && x.pulang) status = 'lengkap';
      else if (x.datang)            status = 'tanpa-pulang';
      else if (x.pulang)            status = 'tanpa-datang';
      else if (jml < ambang)        status = 'sepi';
      else                          status = 'kosong';
      return { tanggal: k, hari: NAMA[i], tgl: Number(k.slice(8)), hariIni: k === hariIni,
               status: status, jamDatang: jam(x.datang), jamPulang: jam(x.pulang), totalHadir: jml };
    });

    return { ok: true, offset: geser,
             label: pendek(tgl[0]) + ' – ' + pendek(tgl[4]) + ' ' + tgl[4].slice(0, 4),
             hari: hari };
  } catch (err) {
    Logger.log('statusMingguan: ' + err);
    return { ok: false, pesan: String(err.message || err) };
  }
}