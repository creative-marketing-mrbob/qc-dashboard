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

const MEMBER_TAB_ALIASES = {
  'Alin': ['Alin', 'LINDA MONICA HADI KUSUMA', 'Linda Monica Hadi Kusuma', 'Linda Monica', 'Linda', 'ALIN'],
};

// ================================================================
// KOLOM YANG ADA DI SETIAP TAB (SESUAI STRUKTUR SPREADSHEET)
// ================================================================
// Script mencari kolom berdasarkan header, jadi aman walaupun ada kolom
// tambahan seperti Kategori sebelum Link Content.
// Last col = Dealdone (checkbox)

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
      const sheet = findMemberSheet(ss, memberName, tabName);
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
  
  const headerRowIndex = findHeaderRow(sheet, lastCol);
  const headerRow = sheet.getRange(headerRowIndex, 1, 1, lastCol).getValues()[0];
  const columns = getColumnMap(headerRow);
  if (columns.tanggal < 0 || columns.jenisRevisi < 0) return emptyMember();

  // Baca semua data mulai setelah header.
  const startRow = headerRowIndex + 1;
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
  
  // Kolom revisi checkbox: setelah Total revisi sampai sebelum Keterangan/Dealdone.
  const checkboxEndCol = columns.keterangan > 0 ? columns.keterangan : (columns.dealdone > 0 ? columns.dealdone : lastCol);
  const checkboxStartCol = columns.totalRevisi >= 0 ? columns.totalRevisi + 1 : columns.jenisRevisi + 1;
  
  // Nama kolom checkbox
  const checkboxColNames = [];
  for (let c = checkboxStartCol; c < checkboxEndCol && c < headerRow.length; c++) {
    const name = String(headerRow[c]).trim();
    if (name && name !== '') checkboxColNames.push({ idx: c, name });
  }
  
  for (const row of values) {
    // Skip baris kosong
    const tanggal    = row[columns.tanggal];
    const jenisRevisi = String(row[columns.jenisRevisi] || '').trim();
    
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

function findMemberSheet(ss, memberName, defaultTabName) {
  const candidates = [defaultTabName].concat(MEMBER_TAB_ALIASES[memberName] || []);
  const seen = {};
  for (const candidate of candidates) {
    const tabName = String(candidate || '').trim();
    if (!tabName || seen[tabName]) continue;
    seen[tabName] = true;
    const sheet = ss.getSheetByName(tabName);
    if (sheet) return sheet;
  }
  return null;
}

function findHeaderRow(sheet, lastCol) {
  const maxRows = Math.min(sheet.getLastRow(), 10);
  const rows = sheet.getRange(1, 1, maxRows, lastCol).getValues();
  for (let r = 0; r < rows.length; r++) {
    const normalized = rows[r].map(normalizeHeader);
    if (normalized.indexOf('tanggal') >= 0 && normalized.indexOf('jenis revisi') >= 0) {
      return r + 1;
    }
  }
  return 5;
}

function getColumnMap(headerRow) {
  const normalized = headerRow.map(normalizeHeader);
  return {
    tanggal: normalized.indexOf('tanggal'),
    jenisRevisi: normalized.indexOf('jenis revisi'),
    totalRevisi: normalized.indexOf('total revisi'),
    keterangan: normalized.indexOf('keterangan'),
    dealdone: normalized.indexOf('dealdone'),
  };
}

function normalizeHeader(value) {
  return String(value || '').toLowerCase().trim().replace(/\s+/g, ' ');
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
