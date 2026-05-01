/**
 * ================================================================
 * QC DASHBOARD API — Google Apps Script Web App
 * Mr.BOB Kampung Inggris — Creative Marketing Team
 * 
 * CARA DEPLOY:
 * 1. Buka Google Sheets → Extensions → Apps Script
 * 2. Paste seluruh kode ini
 * 3. Klik Deploy → New Deployment → Web App
 * 4. Execute as: Me | Who has access: Anyone
 * 5. Copy URL deployment /exec → paste ke dashboard HTML
 *    di variabel APPS_SCRIPT_URL
 * ================================================================
 */

const SPREADSHEET_ID = '1nOy456zmQ7DOdem5SYl6jCEWkoqZFaZydKuhw41PNvo';

// ================================================================
// NAMA TAB PER ANGGOTA (HARUS SAMA PERSIS dengan nama tab di Sheets)
// ================================================================
const MEMBER_TABS = {
  'Sulton': 'Sulton',
  'Dewi':   'Dewi',
  'Zakki':  'Zakki',
  'Alin':   'Alin',
  'Cindy':  'Cindy',
  'Sonia':  'Sonia',
  'Ilham':  'Ilham',
};

// ================================================================
// KOLOM YANG ADA DI SETIAP TAB (SESUAI STRUKTUR SPREADSHEET)
// ================================================================
// Berdasarkan screenshot:
// Col A=No, B=Tanggal, C=Link Content, D=Format Content,
// E=Jenis Revisi (dropdown: No Revisi / Revisi Minor / Revisi Major / Request)
// F=Total revisi, G onwards = checkbox bagian revisi (Hook, Font, Warna, dst)
// Last col = Dealdone (checkbox)

const COL_TANGGAL     = 1;  // B (index 1, 0-based)
const COL_JENIS_REVISI = 4; // E (index 4, 0-based)
const COL_TOTAL_REVISI = 5; // F (index 5, 0-based)
// Kolom checkbox revisi mulai dari index 6 (G) sampai sebelum Keterangan/Dealdone

// ================================================================
// MAIN HANDLER
// ================================================================
function doGet(e) {
  const params = e.parameter || {};
  const dateFrom = params.from || '';
  const dateTo   = params.to   || '';
  
  try {
    const result = getAllMembersData(dateFrom, dateTo);
    return jsonOutput({
      status: 'ok',
      data: result,
      timestamp: new Date().toISOString(),
      timezone: Session.getScriptTimeZone(),
    });
  } catch(err) {
    return jsonOutput({ status: 'error', message: err.toString() });
  }
}

// ================================================================
// BACA DATA SEMUA ANGGOTA
// ================================================================
function getAllMembersData(dateFrom, dateTo) {
  const ss = getSpreadsheet();
  const result = {};
  
  for (const [memberName, tabName] of Object.entries(MEMBER_TABS)) {
    try {
      const sheet = ss.getSheetByName(tabName);
      if (!sheet) {
        result[memberName] = emptyMember();
        continue;
      }
      result[memberName] = parseMemberSheet(sheet, memberName, dateFrom, dateTo);
    } catch(err) {
      Logger.log(`Error parsing ${memberName}: ${err}`);
      result[memberName] = emptyMember();
    }
  }
  
  return result;
}

// ================================================================
// PARSE SATU TAB ANGGOTA
// ================================================================
function parseMemberSheet(sheet, memberName, dateFrom, dateTo) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 6 || lastCol < 5) return emptyMember();
  
  // Baca semua data mulai baris 6 (row index 6, setelah header)
  // Header ada di baris 4-5 biasanya
  const startRow = 6;
  const numRows  = lastRow - startRow + 1;
  
  if (numRows <= 0) return emptyMember();
  
  const range = sheet.getRange(startRow, 1, numRows, lastCol);
  const values = range.getValues();
  
  // Parse date filter
  const fromDate = dateFrom ? parseDate(dateFrom) : null;
  const toDate   = dateTo   ? parseDate(dateTo)   : null;
  
  let projects = 0;
  let noRevisi = 0;
  let minor    = 0;
  let major    = 0;
  
  // Kumpulkan isu (checkbox yang dicentang)
  // Key = nama kolom, value = count berapa kali dicentang
  const issueCounts = {};
  
  // Ambil header nama kolom dari baris 5 (index 4 di sheet = baris ke-5)
  const headerRow = sheet.getRange(5, 1, 1, lastCol).getValues()[0];
  
  // Kolom revisi checkbox: dari kolom G (index 6) sampai sebelum Keterangan
  // Cari kolom "Keterangan" dan "Dealdone"
  let keteranganColIdx = -1;
  let dealdoneColIdx   = -1;
  for (let c = 0; c < headerRow.length; c++) {
    const h = String(headerRow[c]).toLowerCase().trim();
    if (h === 'keterangan') keteranganColIdx = c;
    if (h === 'dealdone')   dealdoneColIdx   = c;
  }
  
  // Kolom checkbox revisi: dari index 6 sampai keteranganColIdx-1 (atau dealdoneColIdx-1)
  const checkboxEndCol = keteranganColIdx > 0 ? keteranganColIdx : (dealdoneColIdx > 0 ? dealdoneColIdx : lastCol - 1);
  const checkboxStartCol = 6; // G
  
  // Nama kolom checkbox
  const checkboxColNames = [];
  for (let c = checkboxStartCol; c < checkboxEndCol && c < headerRow.length; c++) {
    const name = String(headerRow[c]).trim();
    if (name && name !== '') checkboxColNames.push({ idx: c, name });
  }
  
  for (const row of values) {
    // Skip baris kosong
    const tanggal    = row[COL_TANGGAL];
    const jenisRevisi = String(row[COL_JENIS_REVISI] || '').trim();
    
    // Cek apakah baris ini ada isinya (ada tanggal atau jenis revisi)
    if (!tanggal && !jenisRevisi) continue;
    if (!jenisRevisi) continue;
    
    // Filter tanggal
    if (tanggal && (fromDate || toDate)) {
      const rowDate = tanggal instanceof Date ? tanggal : parseDate(String(tanggal));
      if (rowDate) {
        if (fromDate && rowDate < fromDate) continue;
        if (toDate   && rowDate > toDate)   continue;
      }
    }
    
    // Normalisasi jenis revisi
    const jenis = jenisRevisi.toLowerCase();
    
    // Skip baris yang belum ada jenis revisinya (masih kosong/placeholder)
    if (!jenis || jenis === '' || jenis === 'undefined') continue;
    
    // Hitung project (setiap baris dengan jenis revisi = 1 project)
    projects++;
    
    if (jenis.includes('no revisi') || jenis === 'no revisi') {
      noRevisi++;
    } else if (jenis.includes('minor')) {
      minor++;
      // Hitung checkbox yang dicentang
      countCheckedBoxes(row, checkboxColNames, issueCounts);
    } else if (jenis.includes('major')) {
      major++;
      countCheckedBoxes(row, checkboxColNames, issueCounts);
    }
    // "Request" dihitung sebagai project tapi tidak masuk revisi count
    // (opsional, bisa disesuaikan)
  }
  
  // Buat array issues
  const issues = [];
  for (const [name, count] of Object.entries(issueCounts)) {
    if (count > 0 && name) {
      const revisedProjects = minor + major;
      const pct = revisedProjects > 0 ? Math.round(count / revisedProjects * 100) : 0;
      issues.push({ name, count, pct });
    }
  }
  issues.sort((a, b) => b.count - a.count);
  
  return { projects, noRevisi, minor, major, issues };
}

// ================================================================
// COUNT CHECKBOX YANG DICENTANG
// ================================================================
function countCheckedBoxes(row, checkboxColNames, issueCounts) {
  for (const { idx, name } of checkboxColNames) {
    const val = row[idx];
    // Google Sheets checkbox: TRUE/true/"TRUE"
    if (val === true || val === 'TRUE' || val === 'true' || val === 1) {
      issueCounts[name] = (issueCounts[name] || 0) + 1;
    }
  }
}

// ================================================================
// HELPERS
// ================================================================
function emptyMember() {
  return { projects: 0, noRevisi: 0, minor: 0, major: 0, issues: [] };
}

function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== 'PASTE_SPREADSHEET_ID_HERE') {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function jsonOutput(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseDate(str) {
  if (!str) return null;
  if (str instanceof Date) return str;
  
  // Format: YYYY-MM-DD
  const parts = String(str).split('-');
  if (parts.length === 3) {
    return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  }
  
  // Format: DD/MM/YYYY
  const parts2 = String(str).split('/');
  if (parts2.length === 3) {
    return new Date(parseInt(parts2[2]), parseInt(parts2[1]) - 1, parseInt(parts2[0]));
  }
  
  return new Date(str);
}

// ================================================================
// TEST FUNCTION (jalankan manual untuk debug)
// ================================================================
function testGetData() {
  const today = new Date();
  const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
  const from = Utilities.formatDate(firstDay, 'Asia/Jakarta', 'yyyy-MM-dd');
  const to   = Utilities.formatDate(today,    'Asia/Jakarta', 'yyyy-MM-dd');
  
  const result = getAllMembersData(from, to);
  Logger.log(JSON.stringify(result, null, 2));
}
