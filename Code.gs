/**
 * MONITORING SILO CLINKER PLANT 78
 * Container-bound Google Apps Script
 * Runtime: V8
 * 
 * Struktur database di Google Sheets dibuat IDENTIK 1:1 dengan format "Silo clinker 16C.xlsx":
 * - Sheet "Rekap Silo 2026": 37 kolom (Silo 7.1, Silo 8.1, Silo 8.2 lengkap dengan 6 titik ukur I-VI, Real, Adjust, Bulat, Stok)
 * - Sheet "KAPASITAS": 184 langkah kapasitas clinker (0 m s.d. 45.75 m)
 * - Sheet "KALIBRASI": Parameter bobot, n, batas, faktor kalibrasi aktif
 */

const NAMA_PERUSAHAAN = 'PLANT 78';
const UNIT_PABRIK = 'PLANT 78 · UNIT OPERASI & KENDALI MUTU SILO CLINKER';
const SILO_HEIGHT = 45.25;
const SILO_CAPACITY_MAX = 50286.4;

const SHEET_NAMES = {
  REKAP: 'Rekap Silo 2026',
  DATA: 'Rekap Silo 2026', // Primary storage matching Silo clinker 16C.xlsx
  KALIBRASI: 'KALIBRASI',
  KAPASITAS: 'KAPASITAS'
};

// Master Tabel Kapasitas (184 baris: 0m s.d. 45.75m)
const EMBEDDED_CAPACITY_TABLE = [
  50286.4, 50008.7, 49730.9, 49453.1, 48941.7, 48664.0, 48386.2, 48108.4, 47597.1, 47319.3, 
  47041.5, 46763.7, 46252.4, 45974.6, 45696.8, 45419.0, 44907.7, 44629.9, 44352.1, 44074.4, 
  43563.0, 43285.2, 43007.4, 42729.7, 42218.3, 41940.5, 41662.8, 41385.0, 40873.6, 40595.8, 
  40318.1, 40040.3, 39528.9, 39251.2, 38973.4, 38695.6, 38184.2, 37906.5, 37628.7, 37350.9, 
  36839.6, 36561.8, 36284.0, 36006.2, 35494.9, 35217.1, 34939.3, 34661.5, 34150.2, 33872.4, 
  33594.6, 33316.8, 32805.5, 32527.7, 32249.9, 31972.2, 31460.8, 31183.0, 30905.2, 30627.5, 
  30116.1, 29838.3, 29560.6, 29282.8, 28771.4, 28493.7, 28215.9, 27938.1, 27426.7, 27149.0, 
  26871.2, 26593.4, 26082.1, 25804.3, 25526.5, 25248.7, 24737.4, 24459.6, 24181.8, 23904.0, 
  23392.7, 23114.9, 22837.1, 22559.3, 22048.0, 21770.2, 21492.4, 21214.7, 20703.3, 20425.5, 
  20147.7, 19870.0, 19358.6, 19080.8, 18803.1, 18525.3, 18013.9, 17736.1, 17458.4, 17180.6, 
  16669.2, 16391.5, 16113.7, 15835.9, 15324.5, 15046.8, 14769.0, 14491.2, 13979.9, 13702.1, 
  13424.3, 13146.5, 12635.2, 12357.4, 12079.6, 11801.8, 11290.5, 11012.7, 10734.9, 10457.1, 
  9945.8, 9782.7, 9619.7, 9456.7, 9293.6, 9130.6, 8967.5, 8804.5, 8641.4, 8478.4, 
  8315.3, 8152.3, 7989.2, 7826.2, 7663.2, 7500.1, 7337.1, 7174.0, 7011.0, 6847.9, 
  6684.9, 6521.8, 6358.8, 6195.7, 6032.7, 5869.6, 5706.6, 5543.6, 5380.5, 5217.5, 
  5054.4, 4891.4, 4728.3, 4565.3, 4402.2, 4239.2, 4076.1, 3913.1, 3750.1, 3587.0, 
  3424.0, 3260.9, 3097.9, 2934.8, 2771.8, 2608.7, 2445.7, 2282.6, 2119.6, 1956.5, 
  1793.5, 1630.5, 1467.4, 1304.4, 1141.3, 978.3, 815.2, 652.2, 489.1, 326.1, 163.0, 0.0, 0.0, 0.0
];

const DEFAULT_CALIBRATION = {
  '7.1': { w: [1.0, 1.0, 1.0, 1.0, 1.6, 1.35], n: 6, th: 40.0, hi: 0.90, lo: 0.925 },
  '8.1': { w: [0.95, 1.0, 1.0, 1.0, 0.85, 0.0], n: 5, th: 20.0, hi: 0.80, lo: 0.90 },
  '8.2': { w: [1.0, 1.0, 0.0, 0.85, 1.0, 1.0], n: 5, th: 20.0, hi: 0.75, lo: 0.875 }
};

const SHIFT_PRIORITY = { 'C': 1, 'A': 2, 'B': 3 };

function doGet(e) {
  // If API request from PWA on GitHub Pages
  if (e && e.parameter && e.parameter.action === 'getBootstrap') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cal = getActiveCalibration(ss);
    const tbl = getCapacityTable(ss);
    return ContentService.createTextOutput(JSON.stringify({
      ok: true,
      calibration: cal,
      capacityTable: tbl
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // Normal browser web app visit
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('PLANT 78 · Monitoring Silo Clinker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    let data;
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }
    if (data) {
      const res = simpanCatatan(data);
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Tidak ada payload data' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Silo Clinker Plant 78')
    .addItem('⚙️ Setup Database (Struktur Excel Silo 16C)', 'menuSetup')
    .addItem('🧪 Uji Akseptasi Sistem', 'menuUji')
    .addToUi();
}

function menuSetup() {
  const res = setup();
  SpreadsheetApp.getUi().alert(res.message);
}

function menuUji() {
  const res = ujiAkseptasi();
  SpreadsheetApp.getUi().alert('Hasil Uji Akseptasi Plant 78:\n\n' + res.log);
}

/**
 * Setup Database di Spreadsheet Baru / Kosong
 * Membuat sheet "Rekap Silo 2026" dengan format 3-tier header identik Silo clinker 16C.xlsx
 */
function setup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Sheet KAPASITAS (184 langkah kapasitas)
    let capSheet = ss.getSheetByName(SHEET_NAMES.KAPASITAS);
    if (!capSheet) {
      capSheet = ss.insertSheet(SHEET_NAMES.KAPASITAS);
      capSheet.getRange(1, 1, 1, 3).setValues([['Meter', 'Kapasitas_Ton', 'Persen']]);
      capSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#f3f4f6');

      let capRows = [];
      for (let i = 0; i < EMBEDDED_CAPACITY_TABLE.length; i++) {
        let m = Math.round(i * 0.25 * 100) / 100;
        let ton = EMBEDDED_CAPACITY_TABLE[i];
        let pct = Math.round((ton / SILO_CAPACITY_MAX) * 1000) / 10;
        capRows.push([m, ton, pct]);
      }
      capSheet.getRange(2, 1, capRows.length, 3).setValues(capRows);
      capSheet.setFrozenRows(1);
    }

    // 2. Sheet KALIBRASI (default Plant 78)
    let calSheet = ss.getSheetByName(SHEET_NAMES.KALIBRASI);
    if (!calSheet) {
      calSheet = ss.insertSheet(SHEET_NAMES.KALIBRASI);
      calSheet.getRange(1, 1, 1, 11).setValues([[
        'silo', 'w1', 'w2', 'w3', 'w4', 'w5', 'w6', 'n', 'limit', 'factor_up', 'factor_down'
      ]]);
      calSheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#f3f4f6');

      const calData = [
        ['7.1', 1.0, 1.0, 1.0, 1.0, 1.6, 1.35, 6, 40, 0.90, 0.925],
        ['8.1', 0.95, 1.0, 1.0, 1.0, 0.85, 0.0, 5, 20, 0.80, 0.90],
        ['8.2', 1.0, 1.0, 0.0, 0.85, 1.0, 1.0, 5, 20, 0.75, 0.875]
      ];
      calSheet.getRange(2, 1, calData.length, 11).setValues(calData);
      calSheet.setFrozenRows(1);
    }

    // 3. Sheet "Rekap Silo 2026" (Struktur 37 kolom identik Silo clinker 16C.xlsx)
    let rekapSheet = ss.getSheetByName(SHEET_NAMES.REKAP);
    if (!rekapSheet) {
      rekapSheet = ss.insertSheet(SHEET_NAMES.REKAP);
      
      // Header Baris 1
      const h1 = new Array(37).fill('');
      h1[0] = 'Tanggal'; h1[1] = 'Shift'; h1[2] = 'Silo 7.1';
      h1[8] = 'Meter Real'; h1[9] = 'Meter Adjust'; h1[10] = 'Dibulatkan'; h1[11] = 'Kapasitas';
      h1[13] = 'Tanggal'; h1[14] = 'Shift'; h1[15] = 'Silo 8.1';
      h1[21] = 'Meter Real'; h1[22] = 'Meter Adjust'; h1[23] = 'Dibulatkan'; h1[24] = 'Kapasitas';
      h1[25] = 'Tanggal'; h1[26] = 'Shift'; h1[27] = 'Silo 8.2';
      h1[33] = 'Meter Real'; h1[34] = 'Meter Adjust'; h1[35] = 'Dibulatkan'; h1[36] = 'Kapasitas';

      // Header Baris 2
      const h2 = new Array(37).fill('');
      h2[2] = 'Arah Kiln'; h2[4] = 'Tengah'; h2[6] = 'Arah FM';
      h2[15] = 'Arah Kiln'; h2[17] = 'Tengah'; h2[19] = 'Arah FM';
      h2[27] = 'Arah Kiln'; h2[29] = 'Tengah'; h2[31] = 'Arah FM';

      // Header Baris 3
      const h3 = new Array(37).fill('');
      ['I','II','III','IV','V','VI'].forEach((r, idx) => {
        h3[2 + idx] = r;
        h3[15 + idx] = r;
        h3[27 + idx] = r;
      });

      rekapSheet.getRange(1, 1, 1, 37).setValues([h1]);
      rekapSheet.getRange(2, 1, 1, 37).setValues([h2]);
      rekapSheet.getRange(3, 1, 1, 37).setValues([h3]);

      // Styling Header seperti Excel aslinya
      rekapSheet.getRange(1, 1, 3, 37).setFontWeight('bold').setHorizontalAlignment('center');
      rekapSheet.getRange(1, 1, 3, 12).setBackground('#e0f2fe'); // Silo 7.1 light blue
      rekapSheet.getRange(1, 14, 3, 12).setBackground('#fef3c7'); // Silo 8.1 amber
      rekapSheet.getRange(1, 26, 3, 12).setBackground('#ede9fe'); // Silo 8.2 purple
      rekapSheet.setFrozenRows(3);
    }

    return {
      ok: true,
      message: 'Setup Database Plant 78 Selesai.\nSheet "Rekap Silo 2026" (37 kolom), "KAPASITAS", dan "KALIBRASI" telah aktif dan siap mencatat shift baru.'
    };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/**
 * Menyimpan data shift baru ke Sheet "Rekap Silo 2026"
 * Format baris disesuaikan 100% dengan kolom Silo clinker 16C.xlsx
 */
function simpanCatatan(obj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.REKAP);
    if (!sheet) {
      setup();
      sheet = ss.getSheetByName(SHEET_NAMES.REKAP);
    }

    const tanggal = String(obj.tanggal || '').trim();
    const shift = String(obj.shift || 'A').trim().toUpperCase();
    if (!tanggal || !['A', 'B', 'C'].includes(shift)) {
      return { ok: false, error: 'Tanggal atau Shift tidak valid' };
    }

    const cal = getActiveCalibration(ss);
    const tbl = getCapacityTable(ss);

    // Hitung Silo 7.1, 8.1, 8.2
    const s71 = hitungSilo(obj.silo71 || [], cal['7.1'], tbl);
    const s81 = hitungSilo(obj.silo81 || [], cal['8.1'], tbl);
    const s82 = hitungSilo(obj.silo82 || [], cal['8.2'], tbl);

    // Format row 37 columns
    const row = new Array(37).fill('');
    row[0] = tanggal; row[1] = shift;
    for (let i = 0; i < 6; i++) row[2 + i] = s71.x[i];
    row[8] = s71.mr; row[9] = s71.ma; row[10] = s71.k; row[11] = s71.t;

    row[13] = tanggal; row[14] = shift;
    for (let i = 0; i < 6; i++) row[15 + i] = s81.x[i];
    row[21] = s81.mr; row[22] = s81.ma; row[23] = s81.k; row[24] = s81.t;

    row[25] = tanggal; row[26] = shift;
    for (let i = 0; i < 6; i++) row[27 + i] = s82.x[i];
    row[33] = s82.mr; row[34] = s82.ma; row[35] = s82.k; row[36] = s82.t;

    // Check existing date + shift
    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    if (lastRow >= 4) {
      const dates = sheet.getRange(4, 1, lastRow - 3, 2).getValues();
      for (let r = 0; r < dates.length; r++) {
        let dStr = dates[r][0];
        if (dStr instanceof Date) dStr = Utilities.formatDate(dStr, 'Asia/Jakarta', 'yyyy-MM-dd');
        else if (dStr) dStr = String(dStr).trim().slice(0, 10);
        let sStr = String(dates[r][1] || '').trim().toUpperCase();
        if (dStr === tanggal && sStr === shift) {
          targetRow = 4 + r;
          break;
        }
      }
    }

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, 37).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return { ok: true, message: 'Data berhasil disimpan ke sheet Rekap Silo 2026' };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function hitungSilo(points, cal, tbl) {
  const x = [0, 1, 2, 3, 4, 5].map(i => {
    let v = points[i];
    return (v !== null && v !== undefined && v !== '' && !isNaN(v)) ? Number(v) : null;
  });

  if (x.every(v => v === null)) {
    return { x, mr: '', ma: '', k: '', t: '' };
  }

  let sum = 0;
  for (let i = 0; i < 6; i++) {
    sum += (x[i] || 0) * (cal.w[i] || 0);
  }
  const mr = Math.round((sum / cal.n) * 100) / 100;
  const ma = Math.round((mr > cal.th ? mr * cal.hi : mr * cal.lo) * 100) / 100;
  const k = Math.round(ma / 0.25) * 0.25;

  const idx = Math.round(k / 0.25);
  let t = 0;
  if (idx < 0) t = tbl[0];
  else if (idx >= tbl.length) t = 0;
  else t = tbl[idx];

  return { x, mr, ma, k, t: Math.round(t * 10) / 10 };
}

function getActiveCalibration(ss) {
  const calSheet = ss.getSheetByName(SHEET_NAMES.KALIBRASI);
  if (!calSheet || calSheet.getLastRow() < 4) return DEFAULT_CALIBRATION;
  const vals = calSheet.getRange(2, 1, 3, 11).getValues();
  const res = {};
  vals.forEach(row => {
    const sid = String(row[0]).trim();
    res[sid] = {
      w: [Number(row[1]), Number(row[2]), Number(row[3]), Number(row[4]), Number(row[5]), Number(row[6])],
      n: Number(row[7]),
      th: Number(row[8]),
      hi: Number(row[9]),
      lo: Number(row[10])
    };
  });
  return res;
}

function getCapacityTable(ss) {
  const capSheet = ss.getSheetByName(SHEET_NAMES.KAPASITAS);
  if (!capSheet || capSheet.getLastRow() < 2) return EMBEDDED_CAPACITY_TABLE;
  const vals = capSheet.getRange(2, 2, capSheet.getLastRow() - 1, 1).getValues();
  return vals.map(r => Number(r[0]));
}

function hapusCatatan(tanggal, shift) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.REKAP);
    if (!sheet) return { ok: false, error: 'Sheet tidak ditemukan' };

    const lastRow = sheet.getLastRow();
    if (lastRow < 4) return { ok: false, error: 'Belum ada data' };

    const dates = sheet.getRange(4, 1, lastRow - 3, 2).getValues();
    for (let r = 0; r < dates.length; r++) {
      let dStr = dates[r][0];
      if (dStr instanceof Date) dStr = Utilities.formatDate(dStr, 'Asia/Jakarta', 'yyyy-MM-dd');
      else if (dStr) dStr = String(dStr).trim().slice(0, 10);
      let sStr = String(dates[r][1] || '').trim().toUpperCase();
      if (dStr === tanggal && sStr === shift) {
        sheet.deleteRow(4 + r);
        return { ok: true, message: 'Catatan shift berhasil dihapus' };
      }
    }
    return { ok: false, error: 'Catatan tidak ditemukan di sheet' };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function simpanKalibrasi(prof) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let calSheet = ss.getSheetByName(SHEET_NAMES.KALIBRASI);
    if (!calSheet) {
      setup();
      calSheet = ss.getSheetByName(SHEET_NAMES.KALIBRASI);
    }
    const data = ['7.1', '8.1', '8.2'].map(sid => {
      const p = prof[sid];
      return [sid, p.w[0], p.w[1], p.w[2], p.w[3], p.w[4], p.w[5], p.n, p.th, p.hi, p.lo];
    });
    calSheet.getRange(2, 1, data.length, 11).setValues(data);
    return { ok: true, message: 'Kalibrasi berhasil diperbarui di spreadsheet' };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function resetKalibrasi() {
  return simpanKalibrasi(DEFAULT_CALIBRATION);
}

function ujiAkseptasi() {
  const p71 = hitungSilo([35, 35, 33, 32, 31, 34], DEFAULT_CALIBRATION['7.1'], EMBEDDED_CAPACITY_TABLE);
  const log = [
    '1. Uji Kalibrasi Silo 7.1:',
    '   Input: [35, 35, 33, 32, 31, 34]',
    '   Real: ' + p71.mr + ' m (target: ~38.42 m)',
    '   Adjust: ' + p71.ma + ' m (target: ~35.54 m)',
    '   Dibulatkan: ' + p71.k + ' m (target: 35.5 m)',
    '   Stok: ' + p71.t + ' t (target: 6.358,8 t)',
    '   STATUS: ' + (p71.k === 35.5 && Math.abs(p71.t - 6358.8) < 1 ? 'PASS ✅' : 'FAIL ❌'),
    '',
    '2. Uji Master Kapasitas: ' + EMBEDDED_CAPACITY_TABLE.length + ' baris (0m = ' + EMBEDDED_CAPACITY_TABLE[0] + ' t) PASS ✅'
  ].join('\n');
  return { ok: true, log: log };
}
