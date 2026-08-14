/**
 * =========================================================================
 * BACKEND API GOOGLE APPS SCRIPT
 * SISTEM ABSENSI WEBINAR PERDALIN - GDM 2026
 * =========================================================================
 * 
 * Petunjuk Penyebaran (Deployment):
 * 1. Letakkan kode ini di editor Google Apps Script spreadsheet Anda.
 * 2. Deploy sebagai Web App.
 * 3. Set "Execute as" ke "Me" (email Anda).
 * 4. Set "Who has access" ke "Anyone".
 * 5. Salin URL Web App yang dihasilkan, lalu tempelkan di file `config.js` Anda.
 */

const CONFIG = {
  PESERTA_SHEET_NAME: "PESERTA", // Membaca data cepat dari PESERTA (4 kolom hasil cermin)
  RAW_SHEET_NAME: "RAW", // Menulis pendaftaran OTS baru ke RAW
  ATTENDANCE_SHEET_NAME: "Kehadiran",
  SPREADSHEET_ID: "1fEzzdCOt4Gof-ZUmgW6Px_EjdJ1jkJKH0-8eUsEES9o", // ID Spreadsheet pendaftaran & absensi Anda
  ADMIN_PASSWORD: "#GDMPERDALIN26" // Password default untuk masuk ke dashboard admin
};

const QUIZ_CONFIG = {
  SCORE_SHEET_NAME: "Quiz_Scores",
  START_TIME: "2026-08-15T11:00:00+07:00" // Jam 11 Siang WIB pada 15 Agustus 2026
};

/**
 * Handle GET Requests
 * Mengembalikan informasi status API jika diakses langsung melalui browser
 */
function doGet(e) {
  const info = {
    status: "success",
    message: "API Absensi Webinar PERDALIN - GDM 2026 Aktif. Harap kirimkan POST request untuk berinteraksi dengan API ini."
  };
  return ContentService.createTextOutput(JSON.stringify(info))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handle POST Requests (Pintu Gerbang API Utama)
 * Menangani komunikasi CORS dengan aman menggunakan Content-Type text/plain
 */
function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) {
      return errorResponse("Payload data tidak ditemukan.");
    }
    
    // Parse data JSON yang dikirim dari klien
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    let result;
    
    if (action === "check") {
      result = checkParticipant(postData.email);
    } else if (action === "record") {
      result = recordAttendance(postData.email);
    } else if (action === "registerOTS") {
      result = registerOTS(postData.email, postData.name, postData.phone, postData.institution);
    } else if (action === "getDashboard") {
      if (postData.password !== CONFIG.ADMIN_PASSWORD) {
        result = { status: "error", message: "Password admin salah atau akses ditolak." };
      } else {
        result = getDashboardData();
      }
    } else if (action === "resetData") {
      result = resetSpreadsheetData();
    } else if (action === "getQuizStatus") {
      result = getQuizStatus();
    } else if (action === "submitQuiz") {
      result = submitQuiz(postData.emailOrId, postData.score, postData.timeTaken);
    } else if (action === "getLeaderboard") {
      result = getLeaderboard();
    } else if (action === "compareSheets") {
      result = compareSheetsData();
    } else if (action === "syncSheets") {
      result = syncSheetsData(postData.additions);
    } else {
      result = { status: "error", message: "Aksi '" + action + "' tidak dikenali oleh API." };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return errorResponse(error.toString());
  }
}

function errorResponse(message) {
  const result = { status: "error", message: message };
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Mendapatkan objek Spreadsheet
 */
function getSpreadsheet() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } else {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

/**
 * Mendapatkan objek Sheet Pendaftaran (RAW)
 */
function getRegistrationSheet(ss) {
  const sheet = ss.getSheetByName(CONFIG.RAW_SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet pendaftaran '" + CONFIG.RAW_SHEET_NAME + "' tidak ditemukan.");
  }
  return sheet;
}

function getPesertaSheet(ss) {
  const sheet = ss.getSheetByName(CONFIG.PESERTA_SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet pembacaan cepat '" + CONFIG.PESERTA_SHEET_NAME + "' tidak ditemukan. Silakan buat sheet '" + CONFIG.PESERTA_SHEET_NAME + "' terlebih dahulu.");
  }
  return sheet;
}

/**
 * Mendapatkan objek Sheet Absensi (Kehadiran)
 */
function getAttendanceSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.ATTENDANCE_SHEET_NAME);
    
    // Inisialisasi Header
    const headers = ["ID", "Email", "Nama", "Status Kehadiran", "Waktu Absen"];
    sheet.appendRow(headers);
    
    // Styling Header
    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#f1f5f9");
    headerRange.setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    
    // Atur Lebar Kolom
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 250);
    sheet.setColumnWidth(4, 150);
    sheet.setColumnWidth(5, 200);
  } else {
    // Migrasi kolom ID jika belum ada di sheet Kehadiran yang sudah ada
    const lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      const firstHeader = String(sheet.getRange(1, 1).getValue()).toLowerCase().trim();
      if (firstHeader !== "id") {
        sheet.insertColumnBefore(1);
        sheet.getRange(1, 1).setValue("ID");
        sheet.setColumnWidth(1, 120);
        sheet.getRange(1, 1).setFontWeight("bold");
        sheet.getRange(1, 1).setBackground("#f1f5f9");
        sheet.getRange(1, 1).setHorizontalAlignment("center");
      }
    }
  }
  return sheet;
}

/**
 * Mendapatkan objek Sheet Skor Kuis (Quiz_Scores)
 */
function getQuizScoresSheet(ss) {
  let sheet = ss.getSheetByName(QUIZ_CONFIG.SCORE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(QUIZ_CONFIG.SCORE_SHEET_NAME);
    
    // Inisialisasi Header
    const headers = ["ID Peserta", "Email", "Nama", "Skor Kuis", "Waktu Jawab (Detik)", "Waktu Kirim"];
    sheet.appendRow(headers);
    
    // Styling Header
    const headerRange = sheet.getRange(1, 1, 1, 6);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#e2e8f0");
    headerRange.setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    
    // Atur Lebar Kolom
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 250);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 200);
  }
  return sheet;
}

/**
 * Menemukan indeks kolom di sheet 'RAW' dan membuat kolom pendukung jika belum ada
 */
function getRegistrationColumnIndicesAndPrepare(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    throw new Error("Sheet pendaftaran '" + CONFIG.RAW_SHEET_NAME + "' kosong.");
  }
  
  // Baca seluruh data sheet pendaftaran dalam 1 panggilan API saja!
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const regLastRow = allData.length;
  
  let idCol = -1;
  let emailCol = -1;
  let nameCol = -1;
  let phoneCol = -1;
  let instCol = -1;
  let statusCol = -1;
  
  // Prioritas 1: Cari kolom khusus IDP atau ID Peserta terlebih dahulu
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i]).toLowerCase().trim();
    if (header === "idp" || header === "id peserta" || header === "id_peserta" || header === "registration id") {
      idCol = i + 1;
      break;
    }
  }
  
  // Prioritas 2: Jika tidak ada IDP, baru cari kolom ID generik atau No
  if (idCol === -1) {
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i]).toLowerCase().trim();
      if (header === "id" || header === "no" || header === "no." || header === "nomor" || header.includes("registration")) {
        idCol = i + 1;
        break;
      }
    }
  }
  
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i]).toLowerCase().trim();
    if (header.includes("email")) {
      emailCol = i + 1;
    } else if (header === "nama" || header === "name" || header === "nama lengkap" || header === "fullname" || header === "nama peserta" || header === "nama_peserta" || header === "peserta") {
      nameCol = i + 1;
    } else if (header.includes("tel") || header.includes("hp") || header.includes("phone") || header.includes("wa")) {
      phoneCol = i + 1;
    } else if (header.includes("instansi") || header.includes("organisasi") || header.includes("rs") || header.includes("institusi")) {
      instCol = i + 1;
    } else if (header.includes("tipe") || header.includes("kategori") || header.includes("status pendaftaran") || header.includes("registrasi")) {
      statusCol = i + 1;
    }
  }
  
  let currentLast = lastCol;
  let needsHeaderUpdate = false;
  
  // Jika kolom ID belum ada, tambahkan di paling kanan
  if (idCol === -1) {
    currentLast++;
    sheet.getRange(1, currentLast).setValue("ID");
    idCol = currentLast;
    needsHeaderUpdate = true;
  }
  
  if (emailCol === -1) throw new Error("Kolom 'Email' tidak ditemukan di baris pertama sheet 'RAW'.");
  if (nameCol === -1) throw new Error("Kolom 'Nama' tidak ditemukan di baris pertama sheet 'RAW'.");
  
  if (phoneCol === -1) {
    currentLast++;
    sheet.getRange(1, currentLast).setValue("WhatsApp");
    phoneCol = currentLast;
    needsHeaderUpdate = true;
  }
  if (instCol === -1) {
    currentLast++;
    sheet.getRange(1, currentLast).setValue("Instansi");
    instCol = currentLast;
    needsHeaderUpdate = true;
  }
  if (statusCol === -1) {
    currentLast++;
    sheet.getRange(1, currentLast).setValue("Tipe Registrasi");
    statusCol = currentLast;
    needsHeaderUpdate = true;
  }
  
  // Jika ada perubahan kolom header, flush perubahan dan baca ulang allData agar allData sinkron
  if (needsHeaderUpdate) {
    SpreadsheetApp.flush();
    // Kembalikan kolom-kolom baru
    return { idCol, emailCol, nameCol, phoneCol, instCol, statusCol };
  }
  
  // Backfill ID otomatis untuk data lama yang kosong
  if (regLastRow >= 2) {
    let regCount = 0;
    let otsCount = 0;
    let needsUpdate = false;
    
    // Tahap 1: Hitung angka maksimum REG dan OTS yang sudah ada
    for (let i = 1; i < regLastRow; i++) {
      const currentId = String(allData[i][idCol - 1] || "").trim().toUpperCase();
      if (currentId.indexOf("REG") === 0) {
        const num = parseInt(currentId.substring(3), 10);
        if (!isNaN(num) && num > regCount) regCount = num;
      } else if (currentId.indexOf("OTS") === 0) {
        const num = parseInt(currentId.substring(3), 10);
        if (!isNaN(num) && num > otsCount) otsCount = num;
      }
    }
    
    // Tahap 2: Isi ID kosong dengan format berurutan (HANYA jika baris tersebut berisi data nama/email valid)
    const idValues = [];
    for (let i = 1; i < regLastRow; i++) {
      let currentId = String(allData[i][idCol - 1] || "").trim();
      const emailVal = emailCol !== -1 ? String(allData[i][emailCol - 1] || "").trim() : "";
      const nameVal = nameCol !== -1 ? String(allData[i][nameCol - 1] || "").trim() : "";
      const statusType = statusCol !== -1 ? String(allData[i][statusCol - 1] || "").trim().toUpperCase() : "";
      
      if ((currentId === "" || currentId === "0" || currentId === "0.0") && (emailVal || nameVal) && nameVal !== "0" && nameVal !== "0.0") {
        if (statusType.includes("OTS")) {
          otsCount++;
          currentId = "OTS" + ("000" + otsCount).slice(-3);
        } else {
          regCount++;
          currentId = "REG" + ("000" + regCount).slice(-3);
        }
        needsUpdate = true;
      }
      idValues.push([currentId]);
    }
    
    if (needsUpdate) {
      sheet.getRange(2, idCol, regLastRow - 1, 1).setValues(idValues);
      SpreadsheetApp.flush();
    }
  }
  
  return { idCol, emailCol, nameCol, phoneCol, instCol, statusCol };
}

/**
 * Memeriksa status pendaftaran di sheet 'RAW' dan status kehadiran di sheet 'Kehadiran'
 */
function checkParticipant(identifier) {
  if (!identifier) {
    return { status: "error", message: "ID Peserta atau Email tidak boleh kosong." };
  }
  
  identifier = identifier.trim().toLowerCase();
  const ss = getSpreadsheet();
  
  // 1. Validasi Pendaftaran di sheet 'PESERTA' (baca cepat)
  const pesSheet = getPesertaSheet(ss);
  const regLastRow = pesSheet.getLastRow();
  
  if (regLastRow < 2) {
    return { status: "not_found", message: "Database pendaftaran kosong." };
  }
  
  const regData = pesSheet.getRange(2, 1, regLastRow - 1, 4).getValues(); // 4 kolom: ID, Nama, Email, Tipe
  let name = "";
  let id = "";
  let email = "";
  let isRegistered = false;
  
  for (let i = 0; i < regData.length; i++) {
    const row = regData[i];
    const rowId = String(row[0] || "").trim().toLowerCase();
    const rowName = String(row[1] || "").trim();
    const rowEmail = String(row[2] || "").trim().toLowerCase();
    
    if ((rowId === identifier || rowEmail === identifier) && rowName && rowName !== "0" && rowName !== "0.0") {
      name = rowName;
      id = String(row[0] || "").trim();
      email = String(row[2] || "").trim();
      isRegistered = true;
      break;
    }
  }
  
  if (!isRegistered) {
    return { status: "not_found", message: "ID Peserta atau Email Anda tidak terdaftar." };
  }
  
  // 2. Cek status absensi di sheet 'Kehadiran'
  const attSheet = getAttendanceSheet(ss);
  const attLastRow = attSheet.getLastRow();
  
  if (attLastRow >= 2) {
    const attData = attSheet.getRange(2, 1, attLastRow - 1, 5).getValues(); // 5 kolom: ID, Email, Nama, Status Kehadiran, Waktu Absen
    for (let i = 0; i < attData.length; i++) {
      const row = attData[i];
      const attId = String(row[0]).trim();
      
      if (attId === id) {
        const timestamp = row[4];
        let formattedTime = "";
        if (timestamp instanceof Date) {
          formattedTime = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd MMMM yyyy HH:mm:ss");
        } else if (timestamp) {
          formattedTime = String(timestamp);
        } else {
          formattedTime = "Baru-baru ini";
        }
        
        return {
          status: "already_present",
          name: name,
          time: formattedTime
        };
      }
    }
  }
  
  return {
    status: "found",
    name: name
  };
}

/**
 * Mencatat kehadiran peserta reguler ke sheet 'Kehadiran'
 */
function recordAttendance(identifier) {
  if (!identifier) {
    return { status: "error", message: "ID Peserta atau Email tidak boleh kosong." };
  }
  
  identifier = identifier.trim().toLowerCase();
  const ss = getSpreadsheet();
  
  // 1. Ambil Nama dan ID dari data pendaftaran 'PESERTA'
  const pesSheet = getPesertaSheet(ss);
  const regLastRow = pesSheet.getLastRow();
  
  let name = "";
  let id = "";
  let email = "";
  let isRegistered = false;
  
  if (regLastRow >= 2) {
    const regData = pesSheet.getRange(2, 1, regLastRow - 1, 4).getValues(); // 4 kolom: ID, Nama, Email, Tipe
    for (let i = 0; i < regData.length; i++) {
      const row = regData[i];
      const rowId = String(row[0] || "").trim().toLowerCase();
      const rowName = String(row[1] || "").trim();
      const rowEmail = String(row[2] || "").trim().toLowerCase();
      
      if ((rowId === identifier || rowEmail === identifier) && rowName && rowName !== "0" && rowName !== "0.0") {
        name = rowName;
        id = String(row[0] || "").trim();
        email = String(row[2] || "").trim();
        isRegistered = true;
        break;
      }
    }
  }
  
  if (!isRegistered) {
    return { status: "error", message: "ID Peserta atau Email pendaftaran tidak ditemukan." };
  }
  
  // 2. Cek double-submit di sheet 'Kehadiran' berdasarkan ID
  const attSheet = getAttendanceSheet(ss);
  const attLastRow = attSheet.getLastRow();
  
  if (attLastRow >= 2) {
    const attData = attSheet.getRange(2, 1, attLastRow - 1, 1).getValues(); // Kolom 1 adalah ID
    for (let i = 0; i < attData.length; i++) {
      const attId = String(attData[i][0]).trim();
      if (attId === id) {
        return { status: "error", message: "Kehadiran Anda sudah tercatat sebelumnya." };
      }
    }
  }
  
  // 3. Tambahkan baris kehadiran baru (ID, Email, Nama, Status Kehadiran, Waktu Absen)
  const now = new Date();
  attSheet.appendRow([id, email, name, "HADIR", now]);
  SpreadsheetApp.flush();
  
  const formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd MMMM yyyy HH:mm:ss");
  
  return {
    status: "success",
    name: name,
    time: formattedTime
  };
}

/**
 * Mendaftarkan peserta On-The-Spot (OTS) baru ke sheet 'RAW' dan langsung absen di 'Kehadiran'
 */
function registerOTS(email, name, phone, institution) {
  if (!email || !name) {
    return { status: "error", message: "Email dan Nama wajib diisi." };
  }
  
  email = email.trim().toLowerCase();
  name = name.trim();
  phone = phone ? phone.trim() : "-";
  institution = institution ? institution.trim() : "-";
  
  const ss = getSpreadsheet();
  
  // 1. Cek apakah email sudah terdaftar di sheet 'RAW'
  const regSheet = getRegistrationSheet(ss);
  const regCols = getRegistrationColumnIndicesAndPrepare(regSheet);
  const regLastRow = regSheet.getLastRow();
  
  if (regLastRow >= 2) {
    const regData = regSheet.getRange(2, 1, regLastRow - 1, regSheet.getLastColumn()).getValues();
    for (let i = 0; i < regData.length; i++) {
      const row = regData[i];
      const rowEmail = String(row[regCols.emailCol - 1]).trim().toLowerCase();
      if (rowEmail === email) {
        return { status: "error", message: "Email ini sudah terdaftar sebagai peserta reguler. Silakan masukkan email Anda di form utama." };
      }
    }
  }
  
  // 2. Generate Next OTS ID (OTS001, OTS002, dst)
  const otsId = generateNextOtsId(regSheet, regCols, regLastRow);
  
  // 3. Tambah data ke sheet 'RAW' pada kolom yang sesuai
  const nextRow = regSheet.getLastRow() + 1;
  regSheet.getRange(nextRow, regCols.idCol).setValue(otsId);
  regSheet.getRange(nextRow, regCols.emailCol).setValue(email);
  regSheet.getRange(nextRow, regCols.nameCol).setValue(name);
  regSheet.getRange(nextRow, regCols.phoneCol).setValue(phone);
  regSheet.getRange(nextRow, regCols.instCol).setValue(institution);
  regSheet.getRange(nextRow, regCols.statusCol).setValue("OTS");
  
  // 4. Catat di sheet 'Kehadiran' (ID, Email, Nama, Status Kehadiran, Waktu Absen)
  const attSheet = getAttendanceSheet(ss);
  const attLastRow = attSheet.getLastRow();
  
  if (attLastRow >= 2) {
    const attData = attSheet.getRange(2, 1, attLastRow - 1, 2).getValues(); // Baca 2 kolom pertama (ID, Email)
    for (let i = 0; i < attData.length; i++) {
      const attId = String(attData[i][0]).trim();
      const attEmail = String(attData[i][1]).trim().toLowerCase();
      if (attId === otsId || attEmail === email) {
        return { status: "error", message: "Email sudah terdaftar absen." };
      }
    }
  }
  
  const now = new Date();
  attSheet.appendRow([otsId, email, name, "HADIR (OTS)", now]);
  SpreadsheetApp.flush();
  
  const formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd MMMM yyyy HH:mm:ss");
  
  return {
    status: "success",
    name: name,
    time: formattedTime
  };
}

/**
 * Menghimpun data kehadiran dan data tidak hadir untuk Dashboard & Admin panel
 */
function getDashboardData() {
  const tStart = new Date().getTime();
  const ss = getSpreadsheet();
  const tSS = new Date().getTime();
  
  // 1. Ambil data pendaftaran (PESERTA - 4 kolom)
  const pesSheet = getPesertaSheet(ss);
  const allRegData = pesSheet.getDataRange().getValues();
  const regLastRow = allRegData.length;
  const tRegInfo = new Date().getTime();
  
  const registrants = [];
  if (regLastRow >= 2) {
    for (let i = 1; i < regLastRow; i++) {
      const row = allRegData[i];
      const id = String(row[0] || "").trim();
      const name = String(row[1] || "").trim();
      const email = String(row[2] || "").trim().toLowerCase();
      const statusType = String(row[3] || "REGULER").trim();
      
      // Cukup ada ID ATAU Email yang terisi, serta nama valid, maka dianggap peserta sah
      if ((id || email) && name && name !== "0" && name !== "0.0") {
        registrants.push({
          id: id,
          email: email,
          name: name,
          type: statusType
        });
      }
    }
  }
  const tRegData = new Date().getTime();
  
  // 2. Ambil data kehadiran (Kehadiran)
  const attSheet = getAttendanceSheet(ss);
  const attLastRow = attSheet.getLastRow();
  
  // Sinkronisasi & Koreksi otomatis ID di sheet Kehadiran agar 100% cocok dengan ID dari RAW (misal: GDM001, PPI001)
  if (attLastRow >= 2) {
    const attRange = attSheet.getRange(2, 1, attLastRow - 1, 2); // Kolom 1 (ID), Kolom 2 (Email)
    const attValues = attRange.getValues();
    let attNeedsUpdate = false;
    
    for (let i = 0; i < attValues.length; i++) {
      const attId = String(attValues[i][0]).trim();
      const attEmail = String(attValues[i][1]).trim().toLowerCase();
      
      if (attEmail) {
        const matchedReg = registrants.find(r => r.email === attEmail);
        if (matchedReg && attId !== matchedReg.id) {
          attValues[i][0] = matchedReg.id;
          attNeedsUpdate = true;
        }
      }
    }
    
    if (attNeedsUpdate) {
      attSheet.getRange(2, 1, attLastRow - 1, 2).setValues(attValues);
      SpreadsheetApp.flush();
    }
  }
  
  // Baca seluruh data kehadiran dalam 1 panggilan API saja!
  const allAttData = attSheet.getDataRange().getValues();
  const attRows = allAttData.length;
  
  const attendanceMap = {};
  if (attRows >= 2) {
    for (let i = 1; i < attRows; i++) {
      const row = allAttData[i];
      const id = String(row[0]).trim();
      const email = String(row[1]).trim().toLowerCase();
      const name = String(row[2]).trim();
      const status = String(row[3]).trim();
      const timestamp = row[4];
      
      let formattedTime = "";
      if (timestamp instanceof Date) {
        formattedTime = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd MMM yyyy HH:mm:ss");
      } else if (timestamp) {
        formattedTime = String(timestamp);
      }
      
      if (id && name && name !== "0" && name !== "0.0") {
        attendanceMap[id] = {
          email: email,
          name: name,
          status: status,
          time: formattedTime
        };
      }
    }
  }
  const tAttData = new Date().getTime();
  
  // 3. Susun data Sudah Hadir (presentList) dan Belum Hadir (absentList)
  const presentList = [];
  const absentList = [];
  let totalOTS = 0;
  
  registrants.forEach(reg => {
    if (attendanceMap[reg.id]) {
      presentList.push({
        id: reg.id,
        email: reg.email,
        name: reg.name,
        type: reg.type,
        time: attendanceMap[reg.id].time,
        status: attendanceMap[reg.id].status
      });
      if (reg.type === "OTS") totalOTS++;
    } else {
      absentList.push({
        id: reg.id,
        email: reg.email,
        name: reg.name,
        type: reg.type
      });
    }
  });
  
  // Tambahan cadangan jika ada di sheet Kehadiran tapi belum di RAW (misalnya pendaftar OTS)
  Object.keys(attendanceMap).forEach(attId => {
    const foundInRegistrants = registrants.some(r => r.id === attId);
    if (!foundInRegistrants) {
      const attInfo = attendanceMap[attId];
      presentList.push({
        id: attId,
        email: attInfo.email,
        name: attInfo.name,
        type: attInfo.status.includes("OTS") ? "OTS" : "REGULER",
        time: attInfo.time,
        status: attInfo.status
      });
      if (attInfo.status.includes("OTS")) totalOTS++;
    }
  });
  
  presentList.sort((a, b) => new Date(b.time) - new Date(a.time));
  absentList.sort((a, b) => a.name.localeCompare(b.name));
  
  const totalRegistered = registrants.length;
  const totalPresent = presentList.length;
  const pctPresent = totalRegistered > 0 ? Math.round((totalPresent / totalRegistered) * 100) : 0;
  
  const tEnd = new Date().getTime();
  
  return {
    status: "success",
    presentList: presentList,
    absentList: absentList,
    stats: {
      totalRegistered: totalRegistered,
      totalPresent: totalPresent,
      totalOTS: totalOTS,
      pctPresent: pctPresent
    },
    timings: {
      initSpreadsheet: tSS - tStart,
      loadRegInfo: tRegInfo - tSS,
      loadRegData: tRegData - tRegInfo,
      loadAttData: tAttData - tRegData,
      processData: tEnd - tAttData,
      totalExecution: tEnd - tStart
    }
  };
}

/**
 * Membuat ID OTS otomatis berikutnya (misalnya OTS001, OTS002, dst)
 */
function generateNextOtsId(regSheet, regCols, regLastRow) {
  let maxOtsNumber = 0;
  if (regLastRow >= 2) {
    const ids = regSheet.getRange(2, regCols.idCol, regLastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      const currentId = String(ids[i][0]).trim();
      if (currentId.toUpperCase().indexOf("OTS") === 0) {
        const numStr = currentId.substring(3);
        const num = parseInt(numStr, 10);
        if (!isNaN(num) && num > maxOtsNumber) {
          maxOtsNumber = num;
        }
      }
    }
  }
  const nextOtsNumber = maxOtsNumber + 1;
  const formattedNumber = ("000" + nextOtsNumber).slice(-3); // Minimal 3 digit, misal OTS001
  return "OTS" + formattedNumber;
}

function resetSpreadsheetData() {
  const ss = getSpreadsheet();
  
  // 1. Bersihkan sheet RAW (hapus kolom 14 ke kanan secara hati-hati, karena kolom 13 adalah Tipe Registrasi)
  const regSheet = getRegistrationSheet(ss);
  
  // Tambahkan kolom "Tipe Registrasi" di kolom 13 secara bersih jika belum ada di 12 kolom pertama
  const lastColCheck = Math.min(regSheet.getLastColumn(), 12);
  let headers = [];
  if (lastColCheck > 0) {
    headers = regSheet.getRange(1, 1, 1, lastColCheck).getValues()[0];
  }
  let statusColExists = false;
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i]).toLowerCase().trim();
    if (header.includes("tipe") || header.includes("kategori") || header.includes("status pendaftaran") || header.includes("registrasi")) {
      statusColExists = true;
    }
  }
  if (!statusColExists) {
    regSheet.getRange(1, 13).setValue("Tipe Registrasi");
  }

  const regLastCol = regSheet.getLastColumn();
  if (regLastCol > 13) {
    regSheet.deleteColumns(14, regLastCol - 13);
  }
  
  // 2. Bersihkan sheet Kehadiran (kosongkan semua baris setelah header)
  const attSheet = getAttendanceSheet(ss);
  const attLastRow = attSheet.getLastRow();
  if (attLastRow >= 2) {
    attSheet.deleteRows(2, attLastRow - 1);
  }

  // 3. Buat dan Inisialisasi sheet PESERTA dengan rumus cermin agar performa baca super cepat
  let pesSheet = ss.getSheetByName(CONFIG.PESERTA_SHEET_NAME);
  if (!pesSheet) {
    pesSheet = ss.insertSheet(CONFIG.PESERTA_SHEET_NAME);
  }
  pesSheet.clear();
  // Gunakan formula QUERY untuk menyalin data secara dinamis dari RAW (hanya kolom ID, Nama, Email, Kategori)
  // Ini menghindari pembacaan data kolom yang kotor dan mempercepat load time
  pesSheet.getRange(1, 1).setValue('=QUERY(RAW!A:M; "SELECT A, C, F, M WHERE A IS NOT NULL"; 1)');
  
  SpreadsheetApp.flush();
  
  // 4. Bersihkan sheet skor kuis jika ada
  let quizSheet = ss.getSheetByName(QUIZ_CONFIG.SCORE_SHEET_NAME);
  if (quizSheet) {
    const quizLastRow = quizSheet.getLastRow();
    if (quizLastRow >= 2) {
      quizSheet.deleteRows(2, quizLastRow - 1);
    }
  }
  
  return {
    status: "success",
    message: "Spreadsheet, data kehadiran, kuis, dan sheet cermin baca cepat PESERTA berhasil di-reset dengan bersih!"
  };
}

/**
 * Memeriksa status waktu mulai kuis berdasarkan server time
 */
function getQuizStatus() {
  const now = new Date();
  const startTime = new Date(QUIZ_CONFIG.START_TIME);
  const isOpen = now.getTime() >= startTime.getTime();
  
  return {
    status: "success",
    open: isOpen,
    serverTime: Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX"),
    startTime: QUIZ_CONFIG.START_TIME
  };
}

/**
 * Menyimpan skor kuis peserta setelah memvalidasi ID dan satu kali percobaan
 */
function submitQuiz(emailOrId, score, timeTaken) {
  if (!emailOrId) {
    return { status: "error", message: "ID Peserta atau Email tidak boleh kosong." };
  }
  if (score === undefined || timeTaken === undefined) {
    return { status: "error", message: "Skor dan Waktu Jawab wajib dikirimkan." };
  }
  
  emailOrId = emailOrId.trim().toLowerCase();
  score = parseInt(score, 10);
  timeTaken = parseFloat(timeTaken);
  
  if (isNaN(score) || isNaN(timeTaken)) {
    return { status: "error", message: "Skor atau Waktu Jawab tidak valid." };
  }
  
  const ss = getSpreadsheet();
  
  // 1. Validasi pendaftaran peserta di sheet cermin PESERTA (baca cepat)
  const pesSheet = getPesertaSheet(ss);
  const regLastRow = pesSheet.getLastRow();
  if (regLastRow < 2) {
    return { status: "error", message: "Database pendaftaran kosong." };
  }
  
  const regData = pesSheet.getRange(2, 1, regLastRow - 1, 4).getValues(); // ID, Nama, Email, Kategori
  let id = "";
  let name = "";
  let email = "";
  let isRegistered = false;
  
  for (let i = 0; i < regData.length; i++) {
    const row = regData[i];
    const rowId = String(row[0] || "").trim().toLowerCase();
    const rowName = String(row[1] || "").trim();
    const rowEmail = String(row[2] || "").trim().toLowerCase();
    
    if ((rowId === emailOrId || rowEmail === emailOrId) && rowName && rowName !== "0" && rowName !== "0.0") {
      id = String(row[0] || "").trim();
      name = rowName;
      email = String(row[2] || "").trim();
      isRegistered = true;
      break;
    }
  }
  
  if (!isRegistered) {
    return { status: "error", message: "ID Peserta atau Email Anda tidak terdaftar sebagai peserta webinar." };
  }
  
  // 2. Akses/Inisialisasi sheet Quiz_Scores dan cek double-attempt
  const quizSheet = getQuizScoresSheet(ss);
  const quizLastRow = quizSheet.getLastRow();
  
  if (quizLastRow >= 2) {
    const quizData = quizSheet.getRange(2, 1, quizLastRow - 1, 2).getValues(); // Baca ID (kolom 1) dan Email (kolom 2)
    for (let i = 0; i < quizData.length; i++) {
      const qId = String(quizData[i][0]).trim().toLowerCase();
      const qEmail = String(quizData[i][1]).trim().toLowerCase();
      if (qId === id.toLowerCase() || qEmail === email.toLowerCase()) {
        return { status: "error", message: "Anda sudah berpartisipasi dalam kuis ini! Hanya diperbolehkan 1x percobaan." };
      }
    }
  }
  
  // 3. Rekam skor kuis secara aman
  const now = new Date();
  quizSheet.appendRow([id, email, name, score, timeTaken, now]);
  SpreadsheetApp.flush();
  
  return {
    status: "success",
    name: name,
    score: score,
    message: "Skor kuis Anda berhasil disimpan!"
  };
}

/**
 * Mengambil papan peringkat Top 20 kuis (urut skor desc, waktu jawab asc, submit asc)
 */
function getLeaderboard() {
  const ss = getSpreadsheet();
  const quizSheet = getQuizScoresSheet(ss);
  const quizLastRow = quizSheet.getLastRow();
  
  if (quizLastRow < 2) {
    return { status: "success", rankings: [] };
  }
  
  const allScores = quizSheet.getRange(2, 1, quizLastRow - 1, 6).getValues(); // ID, Email, Nama, Score, TimeTaken, Timestamp
  const players = [];
  
  for (let i = 0; i < allScores.length; i++) {
    const row = allScores[i];
    const id = String(row[0] || "").trim();
    const email = String(row[1] || "").trim();
    const name = String(row[2] || "").trim();
    const score = parseInt(row[3], 10);
    const timeTaken = parseFloat(row[4]);
    const timestamp = row[5];
    
    if (id && name) {
      players.push({
        id: id,
        email: email,
        name: name,
        score: isNaN(score) ? 0 : score,
        timeTaken: isNaN(timeTaken) ? 999 : timeTaken,
        timestamp: timestamp instanceof Date ? timestamp.getTime() : new Date().getTime()
      });
    }
  }
  
  // Urutkan: 1. Skor Tertinggi (desc) -> 2. Waktu Tercepat (asc) -> 3. Waktu Submit Terawal (asc)
  players.sort((a, b) => {
    if (b.score !== a.score) {
      return b.score - a.score;
    }
    if (a.timeTaken !== b.timeTaken) {
      return a.timeTaken - b.timeTaken;
    }
    return a.timestamp - b.timestamp;
  });
  
  // Ambil Top 20 untuk layar proyektor
  const rankings = [];
  const limit = Math.min(players.length, 20);
  for (let i = 0; i < limit; i++) {
    const p = players[i];
    rankings.push({
      rank: i + 1,
      id: p.id,
      name: p.name,
      score: p.score,
      timeTaken: p.timeTaken
    });
  }
  
  return {
    status: "success",
    rankings: rankings
  };
}

/**
 * Mengambil data sheet RAW dan sheet DATA DARI PERDALIN untuk perbandingan
 */
/**
 * Fungsi Mandiri untuk Sinkronisasi Data dari sheet 'COCOK DATA' ke sheet 'RAW'
 * Buka Apps Script, pilih fungsi ini di dropdown atas, lalu klik RUN.
 */
function syncCocokData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.RAW_SHEET_NAME);
  const cocokSheet = ss.getSheetByName("COCOK DATA");
  
  if (!rawSheet) {
    Logger.log("Error: Sheet 'RAW' tidak ditemukan.");
    return;
  }
  if (!cocokSheet) {
    Logger.log("Error: Sheet 'COCOK DATA' tidak ditemukan.");
    return;
  }
  
  const rawRows = rawSheet.getLastRow();
  const rawCols = rawSheet.getLastColumn();
  const rawData = rawRows > 0 ? rawSheet.getRange(1, 1, rawRows, rawCols).getValues() : [];
  
  const cocokRows = cocokSheet.getLastRow();
  const cocokCols = cocokSheet.getLastColumn();
  const cocokData = cocokRows > 0 ? cocokSheet.getRange(1, 1, cocokRows, cocokCols).getValues() : [];
  
  if (cocokData.length <= 1) {
    Logger.log("Info: Sheet 'COCOK DATA' kosong atau hanya berisi header.");
    return;
  }
  
  // Indeks data RAW berdasarkan ID (Kolom A) dan Email (Kolom F)
  const rawMap = new Map();
  const rawEmails = new Set();
  
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    const id = String(row[0] || "").trim().toLowerCase();
    const email = String(row[5] || "").trim().toLowerCase();
    if (id) rawMap.set(id, i + 1);
    if (email) rawEmails.add(email);
  }
  
  const additions = [];
  const perdalinIds = new Set();
  
  // Headers COCOK DATA:
  // 0: NO, 1: KODE, 2: Nama Peserta, 3: NIK, 4: Email, 5: No. Whatsapp, 6: Asal Instansi, 7: Profesi
  for (let i = 1; i < cocokData.length; i++) {
    const row = cocokData[i];
    const id = String(row[1] || "").trim().toLowerCase(); // KODE di index 1
    const email = String(row[4] || "").trim().toLowerCase(); // Email di index 4
    const name = String(row[2] || "").trim(); // Nama di index 2
    
    if (!id) continue;
    
    // Cek duplikat di dalam COCOK DATA sendiri
    if (perdalinIds.has(id)) {
      continue;
    }
    perdalinIds.add(id);
    
    // Cek apakah ID atau Email sudah ada di RAW
    const inRawById = rawMap.has(id);
    const inRawByEmail = email ? rawEmails.has(email) : false;
    
    if (!inRawById && !inRawByEmail && name && name !== "0" && name !== "0.0") {
      // Petakan baris COCOK DATA ini ke dalam format kolom RAW (13 kolom)
      const rawFormattedRow = [
        row[1], // ID Peserta (KODE)
        row[3], // NIK
        row[2], // Nama Peserta
        row[6], // Asal Instansi
        row[5], // No Handphone (No. Whatsapp)
        row[4], // Email akun LMS
        "",     // Cabang
        "",     // Nama PS
        "",     // Divisi
        row[7], // Profesi
        "",     // Jabatan
        "",     // Produk
        "Peserta" // Tipe Registrasi
      ];
      additions.push(rawFormattedRow);
    }
  }
  
  Logger.log("=== MEMULAI SINKRONISASI COCOK DATA ===");
  Logger.log("Jumlah data baru yang ditemukan: " + additions.length);
  
  if (additions.length > 0) {
    for (let i = 0; i < additions.length; i++) {
      rawSheet.appendRow(additions[i]);
      Logger.log("Menambahkan: [" + additions[i][0] + "] " + additions[i][2]);
    }
    SpreadsheetApp.flush();
    
    // Re-sync sheet cermin PESERTA dengan memicu penulisan ulang formula QUERY
    resetSpreadsheetData();
    Logger.log("✓ Sinkronisasi berhasil! " + additions.length + " data ditambahkan dan sheet PESERTA diperbarui.");
  } else {
    Logger.log("✓ Semua data di sheet 'COCOK DATA' sudah sesuai dan ada di sheet 'RAW'. Tidak ada data yang perlu ditambahkan.");
  }
}
