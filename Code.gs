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
  REGISTRATION_SHEET_NAME: "RAW",
  ATTENDANCE_SHEET_NAME: "Kehadiran",
  SPREADSHEET_ID: "", // Kosongkan jika script di-bind langsung ke Spreadsheet
  ADMIN_PASSWORD: "#GDMPERDALIN26" // Password default untuk masuk ke dashboard admin
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
  const sheet = ss.getSheetByName(CONFIG.REGISTRATION_SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet pendaftaran '" + CONFIG.REGISTRATION_SHEET_NAME + "' tidak ditemukan.");
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
    const headers = ["Email", "Nama", "Status Kehadiran", "Waktu Absen"];
    sheet.appendRow(headers);
    
    // Styling Header
    const headerRange = sheet.getRange(1, 1, 1, 4);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#f1f5f9");
    headerRange.setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    
    // Atur Lebar Kolom
    sheet.setColumnWidth(1, 250);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 200);
  }
  return sheet;
}

/**
 * Menemukan indeks kolom di sheet 'RAW' dan membuat kolom pendukung jika belum ada
 */
function getRegistrationColumnIndicesAndPrepare(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    throw new Error("Sheet pendaftaran '" + CONFIG.REGISTRATION_SHEET_NAME + "' kosong.");
  }
  
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let emailCol = -1;
  let nameCol = -1;
  let phoneCol = -1;
  let instCol = -1;
  let statusCol = -1;
  
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i]).toLowerCase().trim();
    if (header.includes("email")) {
      emailCol = i + 1;
    } else if (header.includes("nama") || header.includes("name") || header.includes("fullname") || header.includes("peserta")) {
      nameCol = i + 1;
    } else if (header.includes("tel") || header.includes("hp") || header.includes("phone") || header.includes("wa")) {
      phoneCol = i + 1;
    } else if (header.includes("instansi") || header.includes("organisasi") || header.includes("rs") || header.includes("institusi")) {
      instCol = i + 1;
    } else if (header.includes("tipe") || header.includes("kategori") || header.includes("status pendaftaran") || header.includes("registrasi")) {
      statusCol = i + 1;
    }
  }
  
  if (emailCol === -1) throw new Error("Kolom 'Email' tidak ditemukan di baris pertama sheet 'RAW'.");
  if (nameCol === -1) throw new Error("Kolom 'Nama' tidak ditemukan di baris pertama sheet 'RAW'.");
  
  let currentLast = lastCol;
  
  if (phoneCol === -1) {
    currentLast++;
    sheet.getRange(1, currentLast).setValue("WhatsApp");
    phoneCol = currentLast;
  }
  if (instCol === -1) {
    currentLast++;
    sheet.getRange(1, currentLast).setValue("Instansi");
    instCol = currentLast;
  }
  if (statusCol === -1) {
    currentLast++;
    sheet.getRange(1, currentLast).setValue("Tipe Registrasi");
    statusCol = currentLast;
  }
  
  return { emailCol, nameCol, phoneCol, instCol, statusCol };
}

/**
 * Memeriksa status pendaftaran di sheet 'RAW' dan status kehadiran di sheet 'Kehadiran'
 */
function checkParticipant(email) {
  if (!email) {
    return { status: "error", message: "Email tidak boleh kosong." };
  }
  
  email = email.trim().toLowerCase();
  const ss = getSpreadsheet();
  
  // 1. Validasi Pendaftaran di sheet 'RAW'
  const regSheet = getRegistrationSheet(ss);
  const regCols = getRegistrationColumnIndicesAndPrepare(regSheet);
  const regLastRow = regSheet.getLastRow();
  
  if (regLastRow < 2) {
    return { status: "not_found", message: "Database pendaftaran kosong." };
  }
  
  const regData = regSheet.getRange(2, 1, regLastRow - 1, regSheet.getLastColumn()).getValues();
  let name = "";
  let isRegistered = false;
  
  for (let i = 0; i < regData.length; i++) {
    const row = regData[i];
    const rowEmail = String(row[regCols.emailCol - 1]).trim().toLowerCase();
    if (rowEmail === email) {
      name = row[regCols.nameCol - 1];
      isRegistered = true;
      break;
    }
  }
  
  if (!isRegistered) {
    return { status: "not_found", message: "Email Anda belum terdaftar." };
  }
  
  // 2. Cek status absensi di sheet 'Kehadiran'
  const attSheet = getAttendanceSheet(ss);
  const attLastRow = attSheet.getLastRow();
  
  if (attLastRow >= 2) {
    const attData = attSheet.getRange(2, 1, attLastRow - 1, 4).getValues();
    for (let i = 0; i < attData.length; i++) {
      const row = attData[i];
      const attEmail = String(row[0]).trim().toLowerCase();
      
      if (attEmail === email) {
        const timestamp = row[3];
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
function recordAttendance(email) {
  if (!email) {
    return { status: "error", message: "Email tidak boleh kosong." };
  }
  
  email = email.trim().toLowerCase();
  const ss = getSpreadsheet();
  
  // 1. Ambil Nama dari data pendaftaran 'RAW'
  const regSheet = getRegistrationSheet(ss);
  const regCols = getRegistrationColumnIndicesAndPrepare(regSheet);
  const regLastRow = regSheet.getLastRow();
  
  let name = "";
  let isRegistered = false;
  
  if (regLastRow >= 2) {
    const regData = regSheet.getRange(2, 1, regLastRow - 1, regSheet.getLastColumn()).getValues();
    for (let i = 0; i < regData.length; i++) {
      const row = regData[i];
      const rowEmail = String(row[regCols.emailCol - 1]).trim().toLowerCase();
      if (rowEmail === email) {
        name = row[regCols.nameCol - 1];
        isRegistered = true;
        break;
      }
    }
  }
  
  if (!isRegistered) {
    return { status: "error", message: "Email pendaftaran tidak ditemukan." };
  }
  
  // 2. Cek double-submit di sheet 'Kehadiran'
  const attSheet = getAttendanceSheet(ss);
  const attLastRow = attSheet.getLastRow();
  
  if (attLastRow >= 2) {
    const attData = attSheet.getRange(2, 1, attLastRow - 1, 1).getValues();
    for (let i = 0; i < attData.length; i++) {
      const attEmail = String(attData[i][0]).trim().toLowerCase();
      if (attEmail === email) {
        return { status: "error", message: "Kehadiran Anda sudah tercatat sebelumnya." };
      }
    }
  }
  
  // 3. Tambahkan baris kehadiran baru
  const now = new Date();
  attSheet.appendRow([email, name, "HADIR", now]);
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
  
  // 2. Tambah data ke sheet 'RAW' pada kolom yang sesuai
  const nextRow = regSheet.getLastRow() + 1;
  regSheet.getRange(nextRow, regCols.emailCol).setValue(email);
  regSheet.getRange(nextRow, regCols.nameCol).setValue(name);
  regSheet.getRange(nextRow, regCols.phoneCol).setValue(phone);
  regSheet.getRange(nextRow, regCols.instCol).setValue(institution);
  regSheet.getRange(nextRow, regCols.statusCol).setValue("OTS");
  
  // 3. Catat di sheet 'Kehadiran'
  const attSheet = getAttendanceSheet(ss);
  
  const attLastRow = attSheet.getLastRow();
  if (attLastRow >= 2) {
    const attData = attSheet.getRange(2, 1, attLastRow - 1, 1).getValues();
    for (let i = 0; i < attData.length; i++) {
      const attEmail = String(attData[i][0]).trim().toLowerCase();
      if (attEmail === email) {
        return { status: "error", message: "Email sudah terdaftar absen." };
      }
    }
  }
  
  const now = new Date();
  attSheet.appendRow([email, name, "HADIR (OTS)", now]);
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
  const ss = getSpreadsheet();
  
  // 1. Ambil data pendaftaran (RAW)
  const regSheet = getRegistrationSheet(ss);
  const regCols = getRegistrationColumnIndicesAndPrepare(regSheet);
  const regLastRow = regSheet.getLastRow();
  
  const registrants = [];
  if (regLastRow >= 2) {
    const regData = regSheet.getRange(2, 1, regLastRow - 1, regSheet.getLastColumn()).getValues();
    for (let i = 0; i < regData.length; i++) {
      const row = regData[i];
      registrants.push({
        email: String(row[regCols.emailCol - 1]).trim().toLowerCase(),
        name: String(row[regCols.nameCol - 1]).trim(),
        type: String(row[regCols.statusCol - 1] || "REGULER").trim()
      });
    }
  }
  
  // 2. Ambil data kehadiran (Kehadiran)
  const attSheet = getAttendanceSheet(ss);
  const attLastRow = attSheet.getLastRow();
  
  const attendanceMap = {};
  if (attLastRow >= 2) {
    const attData = attSheet.getRange(2, 1, attLastRow - 1, 4).getValues();
    for (let i = 0; i < attData.length; i++) {
      const row = attData[i];
      const email = String(row[0]).trim().toLowerCase();
      const name = String(row[1]).trim();
      const status = String(row[2]).trim();
      const timestamp = row[3];
      
      let formattedTime = "";
      if (timestamp instanceof Date) {
        formattedTime = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd MMM yyyy HH:mm:ss");
      } else if (timestamp) {
        formattedTime = String(timestamp);
      }
      
      attendanceMap[email] = {
        name: name,
        status: status,
        time: formattedTime
      };
    }
  }
  
  // 3. Susun data Sudah Hadir (presentList) dan Belum Hadir (absentList)
  const presentList = [];
  const absentList = [];
  let totalOTS = 0;
  
  registrants.forEach(reg => {
    if (attendanceMap[reg.email]) {
      presentList.push({
        email: reg.email,
        name: reg.name,
        type: reg.type,
        time: attendanceMap[reg.email].time,
        status: attendanceMap[reg.email].status
      });
      if (reg.type === "OTS") totalOTS++;
    } else {
      absentList.push({
        email: reg.email,
        name: reg.name,
        type: reg.type
      });
    }
  });
  
  // Tambahan cadangan jika ada di sheet Kehadiran tapi belum di RAW (misalnya pendaftar OTS)
  Object.keys(attendanceMap).forEach(attEmail => {
    const foundInRegistrants = registrants.some(r => r.email === attEmail);
    if (!foundInRegistrants) {
      const attInfo = attendanceMap[attEmail];
      presentList.push({
        email: attEmail,
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
  
  return {
    status: "success",
    presentList: presentList,
    absentList: absentList,
    stats: {
      totalRegistered: totalRegistered,
      totalPresent: totalPresent,
      totalOTS: totalOTS,
      pctPresent: pctPresent
    }
  };
}
