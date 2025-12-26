// ============================================================================
// BACKEND: PORTAL IKLAN LKIM (VERSI MUKTAMAD - CLEANED)
// ============================================================================

var SPREADSHEET_ID = ''; // Masukkan ID Spreadsheet anda di sini jika perlu
var PASSWORD_SALT = "LKIM_SECURE_SALT_v1_#99283!";

function getDb() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Portal Iklan LKIM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================================
// 1. FUNGSI UTILITI & KESELAMATAN
// ============================================================================

function sanitizeInput(str) {
  if (!str) return "";
  var stringVal = String(str);
  return stringVal.replace(/<[^>]*>?/gm, "").trim();
}

function cleanString(str) {
  return str ? sanitizeInput(str).toLowerCase() : "";
}

function hashPassword(password) {
  if (!password) return "";
  // Gabungkan Password + Salt sebelum hash
  var saltedPayload = String(password).trim() + PASSWORD_SALT; 
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, saltedPayload);
  var txtHash = '';
  for (var i = 0; i < rawHash.length; i++) {
    var hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length == 1) txtHash += '0';
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

function uploadToDrive(base64Data, fileName, mimeType) {
  try {
    var splitBase64 = base64Data.split(',');
    var data = Utilities.base64Decode(splitBase64[1]);
    var blob = Utilities.newBlob(data, mimeType, fileName);
    var file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://drive.google.com/uc?export=view&id=" + file.getId();
  } catch (e) {
    throw new Error("Gagal muat naik ke Drive: " + e.toString());
  }
}

// HELPER: FORMAT DATE (DD/MM/YYYY)
function formatDateDMY(dateVal) {
  if (!dateVal) return "-";
  var d = new Date(dateVal);
  // Jika tarikh tak valid, pulangkan string asal
  if (isNaN(d.getTime())) return String(dateVal);
  // Format ke dd/MM/yyyy (Cth: 25/12/2025)
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

// HELPER: FORMAT NO TELEFON (FORCE STRING)
function formatPhoneNumber(phone) {
  if (!phone) return "";
  // 1. Bersihkan nombor
  var str = phone.toString().replace(/[^0-9]/g, '');
  // 2. Standardkan awalan
  if (str.startsWith('0')) {
    str = '60' + str.substring(1);
  } else if (!str.startsWith('60')) {
    str = '60' + str;
  }

  // 3. Format dengan Dash
  var prefix = "+60";
  var networkCode = str.substring(2, 4);
  var restOfNumber = str.substring(4);

  var finalNumber = prefix + networkCode + "-" + restOfNumber;
  // 4. PENTING: Tambah ' di depan supaya Sheet baca sebagai Text
  return "'" + finalNumber;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// FUNGSI NOTIFIKASI TUGASAN (DASHBOARD)
function getPendingTaskCount(role) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Applications');
  if (!sheet) return 0;
  
  var data = sheet.getDataRange().getValues();
  var count = 0;

  // Loop data (mula baris 2)
  for (var i = 1; i < data.length; i++) {
    var status = data[i][15]; // Lajur P (Index 15) adalah Status

    if (role === 'staff' && status === 'Dalam Proses') {
      count++;
    } 
    else if (role === 'approver' && status === 'Disokong') {
      count++;
    } 
    else if (role === 'admin' && (status === 'Dalam Proses' || status === 'Disokong')) {
      count++;
    }
  }
  return count;
}

function isAdmin(email) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  var searchEmail = cleanString(email);
  
  for (var i = 1; i < data.length; i++) {
    // Lajur A (Indeks 0) = Email, Lajur E (Indeks 4) = Role
    if (cleanString(data[i][0]) === searchEmail && data[i][4] === 'admin') {
      return true;
    }
  }
  return false;
}

// ============================================================================
// 2. API UTAMA DASHBOARD ADMIN
// ============================================================================

function getAdminData() {
  return {
    apps: getAllApplications(),
    assets: getAssets(),
    banners: getBanners(),
    users: getUsers(),
    waiting: getWaitingList(),
    logs: getAuditLogs(),
    stats: getDashboardStats()
  };
}

function deleteItem(type, rowIndex) {
  var ss = getDb();
  var sheetName = '';
  if (type === 'asset') sheetName = 'Assets';
  else if (type === 'user') sheetName = 'Users';
  else if (type === 'banner') sheetName = 'Banners';
  else if (type === 'waiting') sheetName = 'WaitingList';

  var sheet = ss.getSheetByName(sheetName);
  if (sheet && rowIndex > 0) {
    try {
      sheet.deleteRow(rowIndex);
      return { success: true, message: "Item berjaya dipadam." };
    } catch (e) {
      return { success: false, message: "Ralat memadam: " + e.message };
    }
  }

  var userEmail = Session.getActiveUser().getEmail();
  if (!isAdmin(userEmail)) throw new Error("Akses Ditolak");

  return { success: false, message: "Sheet tidak dijumpai." };
}

// ============================================================================
// 3. API ASET (ASSETS)
// ============================================================================

function getAssets() {
  var ss = getDb();
  var sheetAssets = ss.getSheetByName('Assets');
  var sheetApps = ss.getSheetByName('Applications');
  if (!sheetAssets) return [];

  var dataAssets = sheetAssets.getDataRange().getDisplayValues();
  var dataApps = sheetApps ? sheetApps.getDataRange().getDisplayValues() : [];

  // LANGKAH 1: KIRA BERAPA RAMAI DAH "BERJAYA" (BOOKED)
  var bookedAssets = [];
  // Mula dari i=1 (Skip Header)
  for (var i = 1; i < dataApps.length; i++) {
    var statusApp = dataApps[i][15]; // Lajur P (Status)
    var namaAsetApp = dataApps[i][10]; // Lajur K (Lokasi/Nama Aset)

    if (statusApp === 'Berjaya' || statusApp === 'Lulus') {
      if (namaAsetApp) {
        bookedAssets.push(String(namaAsetApp).toLowerCase().trim());
      }
    }
  }

  var assets = [];
  for (var i = 1; i < dataAssets.length; i++) {
    if (dataAssets[i] && dataAssets[i][0]) {

      var statusManual = dataAssets[i][3];
      var namaAset = dataAssets[i][0];
      var namaAsetKecil = String(namaAset).toLowerCase().trim();

      var kekosonganRaw = dataAssets[i].length > 2 ? dataAssets[i][2] : "0";
      var currentCapacity = 0;
      var totalCapacityLabel = "";
      var parsedString = String(kekosonganRaw).trim();

      if (parsedString.includes('/')) {
        var parts = parsedString.split('/');
        currentCapacity = parseInt(parts[0]);
        if (isNaN(currentCapacity)) currentCapacity = 0;
        totalCapacityLabel = parts[1] || '';
      } else {
        currentCapacity = parseInt(parsedString);
        if (isNaN(currentCapacity)) currentCapacity = 0;
      }

      var totalBooked = 0;
      for (var k = 0; k < bookedAssets.length; k++) {
        if (bookedAssets[k].includes(namaAsetKecil)) totalBooked++;
      }

      var realVacancy = Math.max(0, currentCapacity - totalBooked);
      if (statusManual === 'Ditutup' || statusManual === 'Penyelenggaraan') {
        realVacancy = 0;
      }

      var displayKekosongan = "";
      if (statusManual === 'Ditutup' || statusManual === 'Penyelenggaraan') {
        displayKekosongan = "0 / " + (totalCapacityLabel || currentCapacity);
      } else {
        displayKekosongan = totalCapacityLabel ? (realVacancy + " / " + totalCapacityLabel) : realVacancy.toString();
      }

      var finalStatus = "";
      if (statusManual === 'Ditutup') {
        finalStatus = "Ditutup";
      } else if (statusManual === 'Penyelenggaraan') {
        finalStatus = "Penyelenggaraan";
      } else if (statusManual === 'Penuh') {
        finalStatus = "Penuh (Manual)";
      } else if (realVacancy === 0) {
        finalStatus = "Penuh";
      } else {
        finalStatus = "Aktif";
      }

      assets.push({
        id: i + 1,
        nama: namaAset,
        kategori: dataAssets[i][1],
        kekosongan: displayKekosongan,
        rawKekosongan: kekosonganRaw,
        realVacancy: realVacancy,
        status: finalStatus,
        lokasi: dataAssets[i][4] || "",
        deskripsi: dataAssets[i][5] || "",
        saiz: dataAssets[i][6] || "-",
        info: dataAssets[i][7] || "-",
        harga: dataAssets[i][8] || "0"
      });
    }
  }
  return assets;
}

function addAsset(form) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Assets');
  if (!sheet) {
    sheet = ss.insertSheet('Assets');
    sheet.appendRow(['Nama', 'Kategori', 'Kekosongan', 'Status', 'Lokasi', 'Deskripsi', 'Saiz', 'Info', 'Harga']);
  }

  sheet.appendRow([form.nama, form.kategori, form.kekosongan, form.status, form.lokasi, form.deskripsi || "-", form.saiz || "-", form.info || "-", form.harga || "0"]);

  // Kita guna 'Admin' umum sebab frontend tak hantar emel admin dalam form asset
  logAudit("ADMIN", "admin", "TAMBAH_ASET", "Menambah aset baru: " + form.nama);
  return { success: true, message: "Aset berjaya ditambah." };
}

function updateAsset(form) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Assets');
  var rowIndex = form.id;
  if (rowIndex > sheet.getLastRow()) return { success: false, message: "Aset tidak dijumpai." };

  sheet.getRange(rowIndex, 1, 1, 9).setValues([[form.nama, form.kategori, form.kekosongan, form.status, form.lokasi, form.deskripsi, form.saiz, form.info, form.harga]]);
  return { success: true, message: "Aset berjaya dikemaskini." };
}

// ============================================================================
// 4. API PERMOHONAN (APPLICATIONS)
// ============================================================================

function submitApplication(data) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Applications');

  if (!sheet) {
    sheet = ss.insertSheet('Applications');
    sheet.appendRow(['AppID', 'Date', 'Email', 'Nama', 'IC', 'Phone', 'Syarikat', 'Alamat', 'Pekerjaan', 'Pendapatan', 'Lokasi', 'Kategori', 'TarikhMula', 'Tempoh', 'Tujuan', 'Status', 'HargaFinal', 'UlasanStaff', 'UlasanAdmin', 'ProcessedBy', 'HargaAsal', 'UserAction', 'UserTime', 'UserNote']);
  }

  var id = 'APP-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
  try {
    var rowData = [
      String(id), // 0: ID
      new Date(), // 1: Date
      String(data.email), // 2: Email
      String(data.nama), // 3: Nama
      String(data.noIc), // 4: IC
      formatPhoneNumber(data.noTelefon), // 5: Phone
      String(data.syarikat || "-"), // 6: Syarikat
      String(data.alamat), // 7: Alamat
      String(data.pekerjaan), // 8: Pekerjaan
      String(data.pendapatan), // 9: Pendapatan
      String(data.lokasi), // 10: Lokasi
      String(data.jenisAset), // 11: Jenis
      String(data.tarikhMula), // 12: Tarikh Mula
      String(data.tempoh), // 13: Tempoh
      String(data.tujuan), // 14: Tujuan
      'Dalam Proses', // 15: Status
      '', // 16: HargaFinal
      '', // 17: UlasanStaff
      '', // 18: UlasanAdmin
      '', // 19: ProcessedBy
      String(data.hargaAsal || "0") // 20: HargaAsal
    ];

    sheet.appendRow(rowData);
    // Kod Baru
    var identitiUser = data.nama + " (" + data.email + ")";
    logAudit(identitiUser, "user", "MOHON_ASET", "Memohon aset: " + data.lokasi);
    removeFromWaitingList(data.email, data.lokasi);
    return { success: true, message: "Permohonan berjaya dihantar! ID: " + id };
  } catch (e) {
    return { success: false, message: "Ralat Backend: " + e.toString() };
  }
}

function getAllApplications() {
  var ss = getDb();
  var sheet = ss.getSheetByName('Applications');
  if (!sheet) return [];

  var data = sheet.getDataRange().getDisplayValues();
  var apps = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      apps.push({
        rowIndex: i + 1,
        appId: data[i][0],
        dateTime: data[i][1],
        email: data[i][2],
        nama: data[i][3],
        ic: data[i][4],
        phone: data[i][5],
        syarikat: data[i][6],
        alamat: data[i][7],
        pekerjaan: data[i][8],
        pendapatan: data[i][9],
        lokasi: data[i][10],
        jenis: data[i][11],
        tarikhMula: formatDateDMY(data[i][12]),
        tempoh: data[i][13],
        tujuan: data[i][14],
        status: data[i][15],
        hargaFinal: data[i][16] || "",
        ulasanStaff: data[i][17] || "",
        ulasanAdmin: data[i][18] || "",
        processedBy: data[i][19] || "",
        hargaAsal: data[i][20] || "0",
        userAction: data[i][21] || "",
        userTime: data[i][22] || "",
        userNote: data[i][23] || "",
        signLink: data[i][25] || ""
      });
    }
  }
  return apps.reverse();
}

// FUNGSI PROSES PERMOHONAN (LENGKAP DENGAN LABEL ADMIN/APPROVER)
function processApplication(data) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Applications');

  // 1. Dapatkan Row Index yang betul
  var rowIndex = parseInt(data.rowIndex);
  if (isNaN(rowIndex) || rowIndex < 2) {
    return { success: false, message: "Ralat: Baris data tidak dijumpai." };
  }

  // 2. Tentukan Status Baru berdasarkan butang yang ditekan
  var newStatus = "";
  if (data.action === 'sokong') newStatus = "Disokong";
  else if (data.action === 'lulus') newStatus = "Menunggu Persetujuan";
  else if (data.action === 'tolak') newStatus = "Ditolak";

  var timestamp = new Date();

  // A. Lajur P (16) -> Simpan Status
  sheet.getRange(rowIndex, 16).setValue(newStatus);

  // B. Lajur Q (17) -> Simpan Harga Final (Jika ada)
  if (data.harga) {
    sheet.getRange(rowIndex, 17).setValue(data.harga);
  }

  // C. Lajur R (18) -> Simpan Ulasan Staff (Jika ada)
  if (data.ulasanStaff) {
    sheet.getRange(rowIndex, 18).setValue(data.ulasanStaff);
  }

  // D. Lajur S (19) -> Simpan Ulasan Admin/Approver (DENGAN LABEL)
  if (data.ulasanAdmin) {
    var label = "";
    if (data.userRole === 'admin') {
      label = "[SUPER ADMIN] ";
    } else {
      label = "[APPROVER] ";
    }
    var finalComment = label + data.ulasanAdmin;
    sheet.getRange(rowIndex, 19).setValue(finalComment);
  }

  // E. Lajur T (20) -> Log Siapa Yang Proses
  var processorInfo = data.staffName + " [" + (data.userRole || "Staff") + "] (" + timestamp.toString() + ")";
  sheet.getRange(rowIndex, 20).setValue(processorInfo);

  // --- TAMBAHAN: PANGGIL FUNGSI EMEL (Walaupun ia commented dalam function tu) ---
  // Dapatkan Emel Pemohon (Lajur C - Index 2 dalam array data[i] tapi kita tak ada data[i] sini)
  // Kita ambil dari sheet direct
  var emailPemohon = sheet.getRange(rowIndex, 3).getValue(); 
  
  var subject = "Status Permohonan: " + newStatus;
  var body = "<p>Permohonan anda telah dikemaskini kepada status: <strong>" + newStatus + "</strong>.</p><p>Sila log masuk ke portal untuk maklumat lanjut.</p>";
  
  sendEmailNotification(emailPemohon, subject, body);
  // -----------------------------------------------------------------------------

  // Dapatkan emel staff/admin dari sheet Users (atau guna nama je sementara)
  var identitiStaff = data.staffName + " (" + data.staffEmail + ")";
  logAudit(identitiStaff, data.userRole, "PROSES_PERMOHONAN", "Status ditukar ke: " + newStatus + " (Row: " + data.rowIndex + ")");

  return { success: true, message: "Keputusan berjaya dikemaskini kepada: " + newStatus };
}

function getUserHistory(email) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Applications');
  if (!sheet) return [];

  var data = sheet.getDataRange().getDisplayValues();
  var history = [];
  var search = cleanString(email);

  for (var i = data.length - 1; i > 0; i--) {
    // Logic untuk memastikan email match
    var rowEmail = cleanString(data[i][2]);
    if (rowEmail !== search && cleanString(data[i][1]) === search) rowEmail = search; // Fallback logic

    if (rowEmail === search) {
      history.push({
        appId: data[i][0],
        dateTime: data[i][1],
        nama: data[i][3],
        ic: data[i][4],
        noTelefon: data[i][5],
        syarikat: data[i][6],
        alamat: data[i][7],
        pekerjaan: data[i][8],
        pendapatan: data[i][9],
        lokasi: data[i][10],
        jenisAset: data[i][11],
        tarikhMula: formatDateDMY(data[i][12]),
        tempoh: data[i][13],
        tujuan: data[i][14],
        status: data[i][15],
        hargaLulus: data[i][16],
        ulasanStaff: data[i][17],
        ulasanAdmin: data[i][18],
        hargaAsal: data[i][20] || "0",
        userAction: data[i][21] || "",
        userTime: data[i][22] || "",
        userNote: data[i][23] || "",
        signLink: data[i][25] || ""
      });
    }
  }
  return history;
}

// FUNGSI USER: TERIMA/TOLAK (DIKEMASKINI: JANA PDF LENGKAP)
function submitUserDecision(data) {
  var ss = getDb();
  var appSheet = ss.getSheetByName('Applications');
  var appData = appSheet.getDataRange().getDisplayValues();

  var rowIndex = -1;
  var rowData = null;

  // Cari row berdasarkan App ID
  for (var i = 0; i < appData.length; i++) {
    if (appData[i][0] == data.appId) {
      rowIndex = i + 1;
      rowData = appData[i]; // Simpan data row untuk kegunaan template
      break;
    }
  }

  if (rowIndex === -1) return { success: false, message: "Ralat: Permohonan tidak dijumpai." };

  var newStatus = "";
  var userActionText = "";
  var finalNote = data.note || "-";

  // --- LOGIK BARU: JANA AGREEMENT PDF JIKA TERIMA ---
  if (data.action === 'terima') {
    
    // 1. Validasi: Mesti ada tandatangan
    if (!data.signature) {
      return { success: false, message: "Sila turunkan tandatangan digital dahulu." };
    }

    try {
      // A. Ambil butiran dari sheet untuk dimasukkan ke dalam PDF
      var info = {
        nama: rowData[3],
        ic: rowData[4],
        alamat: rowData[7],
        lokasi: rowData[10],
        tempoh: rowData[13],
        harga: rowData[16] || rowData[20], // Harga Lulus atau Harga Asal
        tarikh: formatDateDMY(new Date())
      };

      // B. Sediakan Template HTML Perjanjian (Boleh ubah ayat ikut kesesuaian)
      var htmlAgreement = `
        <div style="font-family: Arial, sans-serif; padding: 40px; line-height: 1.6; color: #333;">
          <div style="text-align: center; border-bottom: 2px solid #000; padding-bottom: 20px; margin-bottom: 30px;">
            <h2 style="margin: 0;">LEMBAGA KEMAJUAN IKAN MALAYSIA</h2>
            <p style="margin: 0; font-size: 12px;">PERJANJIAN PENYEWAAN ASET</p>
          </div>

          <p>PERJANJIAN INI dibuat pada <strong>${info.tarikh}</strong> antara <strong>LEMBAGA KEMAJUAN IKAN MALAYSIA (LKIM)</strong> dan:</p>
          
          <table style="width: 100%; margin-bottom: 20px;">
            <tr><td width="150"><strong>Nama Penyewa</strong></td><td>: ${info.nama}</td></tr>
            <tr><td><strong>No. Kad Pengenalan</strong></td><td>: ${info.ic}</td></tr>
            <tr><td><strong>Alamat</strong></td><td>: ${info.alamat}</td></tr>
          </table>

          <p>Penyewa dengan ini bersetuju untuk menyewa aset berikut tertakluk kepada syarat-syarat yang ditetapkan:</p>

          <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
            <tr style="background-color: #f0f0f0;">
              <td style="padding: 10px; border: 1px solid #ddd;"><strong>Lokasi Aset</strong></td>
              <td style="padding: 10px; border: 1px solid #ddd;">${info.lokasi}</td>
            </tr>
            <tr>
              <td style="padding: 10px; border: 1px solid #ddd;"><strong>Tempoh Sewaan</strong></td>
              <td style="padding: 10px; border: 1px solid #ddd;">${info.tempoh}</td>
            </tr>
            <tr style="background-color: #f0f0f0;">
              <td style="padding: 10px; border: 1px solid #ddd;"><strong>Kadar Sewaan</strong></td>
              <td style="padding: 10px; border: 1px solid #ddd;">RM ${info.harga} / Bulan</td>
            </tr>
          </table>

          <h3>PENGAKUAN PENYEWA:</h3>
          <p>Saya dengan ini mengaku bahawa segala maklumat yang diberikan adalah benar dan saya bersetuju mematuhi segala syarat penyewaan yang ditetapkan oleh LKIM.</p>
          
          <br><br>
          
          <div style="margin-top: 20px;">
            <p><strong>Tandatangan Penyewa:</strong></p>
            <img src="${data.signature}" style="width: 200px; height: auto; border-bottom: 1px solid #000;" />
            <p>Tarikh: ${info.tarikh}</p>
          </div>

          <br><br><br>
          <div style="font-size: 10px; color: #888; text-align: center; border-top: 1px solid #eee; padding-top: 10px;">
            Dokumen ini dijana secara digital melalui Sistem e-Sewaan Aset LKIM. ID Rujukan: ${data.appId}
          </div>
        </div>
      `;

      // C. Tukar HTML kepada PDF Blob
      var blob = Utilities.newBlob(htmlAgreement, MimeType.HTML).getAs(MimeType.PDF);
      blob.setName("Perjanjian_Sewaan_" + data.appId + ".pdf");

      // D. Simpan ke Google Drive
      var file = DriveApp.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var fileUrl = file.getUrl();

      // 3. Simpan Link PDF Lengkap ke Lajur Z (Column 26)
      appSheet.getRange(rowIndex, 26).setValue(fileUrl);
      
      newStatus = "Berjaya";
      userActionText = "BERSETUJU (PERJANJIAN LENGKAP DIREKOD)";

    } catch (e) {
      return { success: false, message: "Gagal menjana perjanjian PDF: " + e.toString() };
    }

  } else {
    // Jika Tolak
    newStatus = "Dibatalkan";
    userActionText = "MENOLAK";
  }

  // Kemaskini data lain dalam Sheet
  var timestamp = new Date().toString();
  var startDate = new Date(appData[rowIndex-1][12]); // Column M (Index 12)
  var duration = appData[rowIndex-1][13];            // Column N (Index 13)
  var endDate = calculateEndDate(startDate, duration);

  appSheet.getRange(rowIndex, 16).setValue(newStatus);      // Status
  appSheet.getRange(rowIndex, 22).setValue(userActionText); // User Action
  appSheet.getRange(rowIndex, 23).setValue(timestamp);      // Time
  appSheet.getRange(rowIndex, 24).setValue(finalNote);      // Note
  appSheet.getRange(rowIndex, 25).setValue(endDate);        // End Date

  return { success: true, message: "Tahniah! Perjanjian telah dijana dan disimpan. Status kini: " + newStatus };
}

// ============================================================================
// 5. API PENGGUNA (USERS)
// ============================================================================

function getUserDetails(email) {
  try {
    var ss = getDb();
    var sheet = ss.getSheetByName('Users');
    var users = sheet.getDataRange().getValues();
    var searchEmail = cleanString(email);
    for (var i = 1; i < users.length; i++) {
      if (cleanString(users[i][0]) == searchEmail) {
        return {
          success: true,
          user: {
            email: users[i][0],
            nama: users[i][2] || "Pengguna",
            noTelefon: users[i][3] || "-",
            role: users[i][4] || "user",
            syarikat: users[i][5] || ""
          }
        };
      }
    }
    return { success: false, message: "Pengguna tidak dijumpai" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function updateUserHeartbeat(email) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  var targetEmail = cleanString(email);
  for (var i = 1; i < data.length; i++) {
    if (cleanString(data[i][0]) == targetEmail) {
      sheet.getRange(i + 1, 7).setValue(new Date());
      return { success: true };
    }
  }
  return { success: false };
}

function changeUserPassword(email, oldPass, newPass) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  var targetEmail = cleanString(email);
  var oldPassHash = hashPassword(oldPass);

  for (var i = 1; i < data.length; i++) {
    if (cleanString(data[i][0]) == targetEmail) {
      if (String(data[i][1]) == oldPassHash) {
        sheet.getRange(i + 1, 2).setValue(hashPassword(newPass));
        return { success: true, message: "Kata laluan berjaya ditukar." };
      } else {
        return { success: false, message: "Kata laluan lama salah." };
      }
    }
  }
  return { success: false, message: "Pengguna tidak dijumpai." };
}

function getUsers() {
  var ss = getDb();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var users = [];
  var now = new Date().getTime();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      var lastActive = data[i][6];
      var statusOnline = "Offline";
      var isOnline = false;

      if (lastActive instanceof Date) {
        var diff = now - lastActive.getTime();
        var minutes = Math.floor(diff / 60000);

        if (minutes < 5) {
          statusOnline = "Online";
          isOnline = true;
        } else if (minutes < 60) {
          statusOnline = minutes + " minit lepas";
        } else if (minutes < 1440) {
          statusOnline = Math.floor(minutes / 60) + " jam lepas";
        } else {
          statusOnline = Math.floor(minutes / 1440) + " hari lepas";
        }
      } else {
        statusOnline = "Belum pernah login";
      }

      users.push({
        rowIndex: i + 1,
        email: data[i][0],
        nama: data[i][2],
        noTelefon: data[i][3],
        role: data[i][4],
        syarikat: data[i][5],
        lastSeen: statusOnline,
        isOnline: isOnline,
        status: data[i][7] || "Aktif" // Ambil status dari Kolum H
      });
    }
  }
  return users;
}

function adminToggleStatus(email, newStatus) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  var targetEmail = cleanString(email);

  for (var i = 1; i < data.length; i++) {
    if (cleanString(data[i][0]) == targetEmail) {
      // Kolum H (Index 8 dalam notation 1-based sheet)
      sheet.getRange(i + 1, 8).setValue(newStatus); 
      return { success: true, message: "Status pengguna berjaya ditukar kepada " + newStatus };
    }
  }
  return { success: false, message: "Pengguna tidak dijumpai." };
}

function loginUser(email, password) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Users');
  var users = sheet.getDataRange().getValues();
  var searchEmail = cleanString(email);
  var searchPass = hashPassword(password);
  
  for (var i = 1; i < users.length; i++) {
    if (cleanString(users[i][0]) == searchEmail && String(users[i][1]) == searchPass) {
      
      // --- TAMBAHAN BARU: SEMAK STATUS AKAUN ---
      // Anggap Kolum H (Index 7) adalah status. Jika kosong, anggap 'Aktif'.
      var statusAkaun = users[i][7] ? String(users[i][7]) : "Aktif";
      
      if (statusAkaun === "Nyahaktif") {
        return { success: false, message: "Akaun anda telah dinyahaktifkan. Sila hubungi pentadbir." };
      }
      // -----------------------------------------

      sheet.getRange(i + 1, 7).setValue(new Date()); // Update Last Login

      var identitiUser = users[i][2] + " (" + email + ")"; // users[i][2] ialah Nama
      logAudit(identitiUser, users[i][4] || "user", "LOGIN", "Pengguna berjaya log masuk.");

      return {
        success: true,
        user: {
          email: users[i][0],
          nama: users[i][2],
          noTelefon: users[i][3],
          role: users[i][4] || "user",
          syarikat: users[i][5] || ""
        }
      };
    }
  }
  return { success: false, message: "Emel atau kata laluan salah." };
}

function registerUser(data) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Users');
  var users = sheet.getDataRange().getValues();
  var newEmail = cleanString(data.email);
  for (var i = 1; i < users.length; i++) {
    if (cleanString(users[i][0]) == newEmail) return { success: false, message: "Emel telah didaftarkan." };
  }
  sheet.appendRow([data.email, hashPassword(data.password), data.nama, formatPhoneNumber(data.noTelefon), 'user', data.syarikat || '', new Date()]);
  return { success: true, message: "Pendaftaran berjaya." };
}

function adminAddUser(form) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (cleanString(data[i][0]) == cleanString(form.email)) return { success: false, message: "Emel sudah wujud." };
  }
  sheet.appendRow([form.email, hashPassword(form.password), form.nama, formatPhoneNumber(form.noTelefon), form.role, form.syarikat, new Date()]);
  return { success: true, message: "Pengguna ditambah." };
}

function adminUpdateUser(form) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (cleanString(data[i][0]) == cleanString(form.email)) {
      sheet.getRange(i + 1, 3).setValue(form.nama);
      sheet.getRange(i + 1, 4).setValue(formatPhoneNumber(form.noTelefon));
      sheet.getRange(i + 1, 5).setValue(form.role);
      sheet.getRange(i + 1, 6).setValue(form.syarikat);
      if (form.password) sheet.getRange(i + 1, 2).setValue(hashPassword(form.password));
      return { success: true, message: "Pengguna dikemaskini." };
    }
  }
  return { success: false, message: "Pengguna tidak dijumpai." };
}

function adminResetPass(email) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var newPass = Math.floor(100000 + Math.random() * 900000).toString();
    sheet.getRange(i + 1, 2).setValue(hashPassword(newPass));
    return { success: true, message: "Kata laluan direset ke: " + newPass };
  }
}

// ============================================================================
// 6. API BANNER & PUBLIC
// ============================================================================

function getBanners() {
  var ss = getDb();
  var sheet = ss.getSheetByName('Banners');
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  var banners = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      banners.push({ rowIndex: i + 1, image: data[i][0], title: data[i][1], subtitle: data[i][2], expiry: data[i][3] });
    }
  }
  return banners;
}

function addBanner(form) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Banners');
  if (!sheet) {
    sheet = ss.insertSheet('Banners');
    sheet.appendRow(['ImageURL', 'Title', 'Subtitle', 'Expiry']);
  }
  var imageUrl = form.url || "";
  if (form.imageFile) {
    try {
      imageUrl = uploadToDrive(form.imageFile.data, form.imageFile.name, form.imageFile.type);
    } catch (e) {
      return { success: false, message: "Upload gambar gagal." };
    }
  }
  sheet.appendRow([imageUrl, form.title, form.subtitle, form.expiry]);
  return { success: true, message: "Banner berjaya ditambah." };
}

function updateBanner(form) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Banners');
  
  var rowIndex = parseInt(form.rowIndex);

  if (!sheet || isNaN(rowIndex) || rowIndex < 2) {
    return { success: false, message: "Banner tidak dijumpai atau ID tidak sah." };
  }

  var imageUrl = form.image;
  
  if (form.imageFile && form.imageFile.data) {
    try {
      imageUrl = uploadToDrive(form.imageFile.data, form.imageFile.name, form.imageFile.type);
    } catch (e) {
      return { success: false, message: "Gagal muat naik gambar baru: " + e.toString() };
    }
  }

  try {
    sheet.getRange(rowIndex, 1).setValue(imageUrl);
    sheet.getRange(rowIndex, 2).setValue(form.title);
    sheet.getRange(rowIndex, 3).setValue(form.subtitle);
    sheet.getRange(rowIndex, 4).setValue(form.expiry);

    return { success: true, message: "Banner berjaya dikemaskini." };
  } catch(e) {
    return { success: false, message: "Ralat backend: " + e.toString() };
  }
}

function getPublicData() {
  var banners = [];
  try {
    var all = getBanners();
    var today = new Date();
    all.forEach(b => {
      var exp = b.expiry ? new Date(b.expiry) : null;
      if (!exp || exp >= today) banners.push(b);
    });
  } catch (e) {}
  return { assets: getAssets(), banners: banners, settings: getSystemSettings() };
}

// ============================================================================
// 7. FUNGSI TAMBAHAN (EMEL, PDF, STATISTIK)
// ============================================================================

// A. FUNGSI PENGHANTARAN EMEL (DIBIARKAN SEBAGAI KOMEN)
function sendEmailNotification(to, subject, htmlBody) {
  // --- Buka komen di bawah (buang //) untuk aktifkan emel ---
  
  // var emailTemplate = `
  //   <div style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">
      
  //     <div style="background-color: #1e3a8a; padding: 20px; text-align: center; border-bottom: 4px solid #facc15;">
  //       <h2 style="color: #ffffff; margin: 0; font-size: 20px; text-transform: uppercase; letter-spacing: 1px;">Sistem e-Sewaan Aset</h2>
  //       <p style="color: #facc15; margin: 5px 0 0; font-size: 10px; font-weight: bold; letter-spacing: 2px;">LEMBAGA KEMAJUAN IKAN MALAYSIA</p>
  //     </div>

  //     <div style="padding: 30px; background-color: #ffffff; color: #334155; line-height: 1.6;">
  //       ${htmlBody}
        
  //       <br><br>
  //       <p style="font-size: 12px; color: #64748b; border-top: 1px solid #f1f5f9; padding-top: 20px;">
  //         Sila log masuk ke portal untuk maklumat lanjut: <br>
  //         <a href="${ScriptApp.getService().getUrl()}" style="color: #1e3a8a; font-weight: bold; text-decoration: none;">Klik Sini untuk ke Portal e-Sewaan</a>
  //       </p>
  //     </div>

  //     <div style="background-color: #f8fafc; padding: 15px; text-align: center; font-size: 11px; color: #94a3b8;">
  //       &copy; ${new Date().getFullYear()} Lembaga Kemajuan Ikan Malaysia (LKIM). <br>
  //       Ini adalah emel janaan komputer, tandatangan tidak diperlukan.
  //     </div>
  //   </div>
  // `;

  // try {
  //   MailApp.sendEmail({
  //     to: to,
  //     subject: "[e-Sewaan LKIM] " + subject,
  //     htmlBody: emailTemplate,
  //     name: "Portal Aset LKIM" // Nama pengirim yang akan keluar di inbox
  //   });
  //   Logger.log("Emel berjaya dihantar ke: " + to);
  // } catch (e) {
  //   Logger.log("Gagal hantar emel: " + e.toString());
  // }
}

// B. API STATISTIK DASHBOARD
function getDashboardStats() {
  var ss = getDb();
  var sheet = ss.getSheetByName('Applications');
  if (!sheet) return null;
  
  var data = sheet.getDataRange().getValues();
  var stats = {
    total: 0,
    lulus: 0,
    ditolak: 0,
    dalamProses: 0,
    kutipan: 0.00,
    // Data Baru untuk Carta Terperinci
    revenueByCategory: {},
    topAssets: {} 
  };

  // Mula dari row 1 (skip header)
  for (var i = 1; i < data.length; i++) {
    stats.total++;
    var status = data[i][15]; // Col P: Status
    var harga = parseFloat(String(data[i][16] || data[i][20]).replace(/[^0-9.]/g, '')) || 0; // Harga
    var lokasiAset = data[i][10]; // Col K: Nama Aset/Lokasi
    var kategori = data[i][11];   // Col L: Kategori

    // 1. Kira Status Utama
    if (status === 'Berjaya' || status === 'Lulus') {
      stats.lulus++;
      stats.kutipan += harga;

      // 2. Kira Kutipan Ikut Kategori (Hanya yang Lulus/Berjaya)
      if (kategori) {
        var katKey = String(kategori).trim();
        if (!stats.revenueByCategory[katKey]) stats.revenueByCategory[katKey] = 0;
        stats.revenueByCategory[katKey] += harga;
      }

    } else if (status === 'Ditolak' || status === 'Dibatalkan') {
      stats.ditolak++;
    } else if (status === 'Dalam Proses' || status === 'Disokong' || status === 'Menunggu Persetujuan') {
      stats.dalamProses++;
    }

    // 3. Kira Aset Paling Popular (Berdasarkan Jumlah Permohonan Masuk, tak kira status)
    // Ini membantu admin tahu aset mana yang 'laku keras' atau viral
    if (lokasiAset) {
      var asetName = String(lokasiAset).split('-')[0].trim(); // Ambil nama depan saja (buang negeri)
      if (!stats.topAssets[asetName]) stats.topAssets[asetName] = 0;
      stats.topAssets[asetName]++;
    }
  }
  
  // --- FORMAT DATA UNTUK FRONTEND ---

  // Format Revenue Array
  stats.revenueArray = Object.keys(stats.revenueByCategory).map(function(key) {
    return { name: key, nilai: stats.revenueByCategory[key] };
  }).sort(function(a, b) { return b.nilai - a.nilai }); // Susun paling tinggi ke rendah

  // Format Top Assets Array (Ambil Top 5 Sahaja)
  stats.topAssetsArray = Object.keys(stats.topAssets).map(function(key) {
    return { name: key, jumlah: stats.topAssets[key] };
  }).sort(function(a, b) { return b.jumlah - a.jumlah }).slice(0, 5);

  return stats;
}

// C. FUNGSI JANA SURAT TAWARAN (PDF)
function generateOfferLetter(appId) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Applications');
  var data = sheet.getDataRange().getValues();
  var appData = null;

  // Cari Data
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == appId) {
      appData = {
        nama: data[i][3],
        ic: data[i][4],
        alamat: data[i][7],
        lokasi: data[i][10],
        jenis: data[i][11],
        tempoh: data[i][13],
        harga: data[i][16] || data[i][20], // Harga Lulus atau Asal
        tarikhSurat: formatDateDMY(new Date())
      };
      break;
    }
  }

  if (!appData) return { success: false, message: "Data tidak dijumpai" };

  // HTML Template Surat (Kini SAMA dengan Draf)
  var htmlContent = `
    <div style="font-family: Arial, sans-serif; padding: 40px; line-height: 1.6;">
      <div style="text-align: center; border-bottom: 2px solid #000; padding-bottom: 20px; margin-bottom: 30px;">
        <h2 style="margin: 0;">LEMBAGA KEMAJUAN IKAN MALAYSIA</h2>
        <p style="margin: 0; font-size: 12px;">Aras 5, Menara LKIM, Jalan Topaz 7/4, Seksyen 7, 47120 Puchong, Selangor</p>
      </div>

      <p style="text-align: right;"><strong>Tarikh:</strong> ${appData.tarikhSurat}</p>
      
      <p><strong>KEPADA:</strong><br>
      ${appData.nama}<br>
      ${appData.ic}<br>
      ${appData.alamat}</p>

      <br>
      <h3>TAWARAN PENYEWAAN ASET LKIM (${appData.jenis.toUpperCase()})</h3>
      
      <p>Sukacita dimaklumkan bahawa permohonan tuan/puan untuk menyewa aset berikut telah <strong>DILULUSKAN</strong>:</p>

      <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
        <tr>
            <td style="padding: 8px; border: 1px solid #ddd; width: 40%; font-weight: bold;">Lokasi Aset</td>
            <td style="padding: 8px; border: 1px solid #ddd;">${appData.lokasi}</td>
        </tr>
        <tr>
            <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold;">Tempoh Sewaan</td>
            <td style="padding: 8px; border: 1px solid #ddd;">${appData.tempoh}</td>
        </tr>
        <tr>
            <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold;">Kadar Sewaan Bulanan</td>
            <td style="padding: 8px; border: 1px solid #ddd;">RM ${appData.harga}</td>
        </tr>
      </table>

      <p>Syarat-syarat tambahan:</p>
      <ul>
        <li>Penyewa hendaklah menjelaskan bayaran deposit sebelum kunci diserahkan.</li>
        <li>Sebarang ubah suai struktur perlu mendapat kebenaran bertulis LKIM.</li>
        <li>LKIM berhak menamatkan penyewaan sekiranya terdapat pelanggaran syarat.</li>
      </ul>

      <p>Sila hadir ke pejabat LKIM dalam tempoh 14 hari bekerja untuk urusan perjanjian sewaan dan pembayaran deposit.</p>
      
      <br><br>
      <div style="margin-top: 50px;">
        <p>Sekian, terima kasih.</p>
        <p><strong>"MALAYSIA MADANI"</strong></p>
        <br>
        <p><i>Dokumen ini adalah cetakan komputer dan tidak memerlukan tandatangan.</i></p>
      </div>
    </div>
  `;

  // Tukar HTML ke PDF Blob
  var blob = Utilities.newBlob(htmlContent, "text/html", "Surat_Tawaran_" + appId + ".html");
  var pdf = blob.getAs("application/pdf");
  
  // Tukar ke Base64 supaya user boleh download terus
  var base64 = Utilities.base64Encode(pdf.getBytes());
  
  // Return format Data URI supaya browser boleh baca
  return { 
      success: true, 
      data: "data:application/pdf;base64," + base64, 
      filename: "Surat_Tawaran_" + appId + ".pdf" 
  };
}

function getOfferLetterContent(appId) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Applications');
  var data = sheet.getDataRange().getValues();
  var appData = null;

  // 1. Cari Data Permohonan
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == appId) {
      appData = {
        nama: data[i][3],
        ic: data[i][4],
        alamat: data[i][7],
        lokasi: data[i][10],
        jenis: data[i][11],
        tempoh: data[i][13],
        harga: data[i][16] || data[i][20], // Harga Lulus (Q) atau Asal (U)
        tarikhSurat: formatDateDMY(new Date())
      };
      break;
    }
  }

  if (!appData) return { success: false, message: "Data permohonan tidak dijumpai." };

  // 2. HTML Template (SAMA SEPERTI generateOfferLetter)
  // Nota: Pastikan template ini SAMA dengan yang ada dalam generateOfferLetter supaya konsisten.
  var htmlContent = `
    <div style="font-family: Arial, sans-serif; padding: 40px; line-height: 1.6; max-width: 800px; margin: 0 auto;">
      <div style="text-align: center; border-bottom: 2px solid #000; padding-bottom: 20px; margin-bottom: 30px;">
        <h2 style="margin: 0;">LEMBAGA KEMAJUAN IKAN MALAYSIA</h2>
        <p style="margin: 0; font-size: 12px;">Aras 5, Menara LKIM, Jalan Topaz 7/4, Seksyen 7, 47120 Puchong, Selangor</p>
      </div>

      <p style="text-align: right;"><strong>Tarikh:</strong> ${appData.tarikhSurat}</p>
      
      <p><strong>KEPADA:</strong><br>
      ${appData.nama}<br>
      ${appData.ic}<br>
      ${appData.alamat}</p>

      <br>
      <h3>TAWARAN PENYEWAAN ASET LKIM (${appData.jenis.toUpperCase()})</h3>
      
      <p>Sukacita dimaklumkan bahawa permohonan tuan/puan untuk menyewa aset berikut telah <strong>DILULUSKAN</strong>:</p>

      <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd; width: 40%; font-weight: bold;">Lokasi Aset</td>
          <td style="padding: 8px; border: 1px solid #ddd;">${appData.lokasi}</td>
        </tr>
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold;">Tempoh Sewaan</td>
          <td style="padding: 8px; border: 1px solid #ddd;">${appData.tempoh}</td>
        </tr>
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold;">Kadar Sewaan Bulanan</td>
          <td style="padding: 8px; border: 1px solid #ddd;">RM ${appData.harga}</td>
        </tr>
      </table>

      <p>Syarat-syarat tambahan:</p>
      <ul>
        <li>Penyewa hendaklah menjelaskan bayaran deposit sebelum kunci diserahkan.</li>
        <li>Sebarang ubah suai struktur perlu mendapat kebenaran bertulis LKIM.</li>
        <li>LKIM berhak menamatkan penyewaan sekiranya terdapat pelanggaran syarat.</li>
      </ul>

      <p>Sila hadir ke pejabat LKIM dalam tempoh 14 hari bekerja untuk urusan perjanjian sewaan dan pembayaran deposit.</p>
      
      <br><br>
      <div style="margin-top: 50px;">
        <p>Sekian, terima kasih.</p>
        <p><strong>"MALAYSIA MADANI"</strong></p>
        <br>
        <p><i>(Ini adalah draf paparan komputer untuk semakan awal)</i></p>
      </div>
    </div>
  `;

  return { success: true, html: htmlContent };
}

// ============================================================================
// 8. AUTOMATION: SEMAKAN TAMAT TEMPOH SEWAAN
// ============================================================================

function checkExpiredRentals() {
  var ss = getDb();
  var sheet = ss.getSheetByName('Applications');
  var data = sheet.getDataRange().getValues();
  var today = new Date();
  var changesCount = 0;

  // Mula dari row 1 (skip header)
  for (var i = 1; i < data.length; i++) {
    var status = data[i][15]; // Lajur P (Status)
    
    // Hanya semak yang status 'Berjaya' (Sedang menyewa)
    if (status === 'Berjaya') {
      var tarikhMula = new Date(data[i][12]); // Lajur M (Tarikh Mula)
      var tempohStr = String(data[i][13]).toLowerCase(); // Lajur N (Tempoh) - cth: "6 Bulan"
      
      if (isValidDate(tarikhMula)) {
        var tarikhTamat = calculateEndDate(tarikhMula, tempohStr);
        
        // Semak jika hari ini dah melepasi tarikh tamat
        if (tarikhTamat && today > tarikhTamat) {
          // 1. Tukar Status kepada 'Tamat Tempoh'
          sheet.getRange(i + 1, 16).setValue('Tamat Tempoh'); 
          
          // 2. (Optional) Log Nota Sistem
          var oldNotes = data[i][23] || "";
          sheet.getRange(i + 1, 24).setValue(oldNotes + " [Sistem: Sewaan tamat pada " + formatDateDMY(tarikhTamat) + "]");
          
          changesCount++;
        }
      }
    }
  }
  
  if (changesCount > 0) {
    Logger.log("Sebanyak " + changesCount + " sewaan telah ditamatkan secara automatik.");
  }
}

// Helper: Kira Tarikh Tamat berdasarkan string "1 Tahun" atau "6 Bulan"
function calculateEndDate(startDate, durationStr) {
  var d = new Date(startDate);
  if (isNaN(d.getTime())) return null; // Return null jika tarikh mula rosak
  
  var str = String(durationStr).toLowerCase();
  var amount = parseInt(str.replace(/[^0-9]/g, '')) || 0; 
  
  // Logik yang lebih fleksibel
  if (str.includes('tahun') || str.includes('year')) {
    d.setFullYear(d.getFullYear() + amount);
  } else if (str.includes('hari') || str.includes('day')) {
    d.setDate(d.getDate() + amount);
  } else {
    // Default ke bulan jika perkataan 'bulan' ada ATAU tiada unit dinyatakan
    d.setMonth(d.getMonth() + (amount || 0)); 
  }
  return d;
}

// Helper: Validasi Date
function isValidDate(d) {
  return d instanceof Date && !isNaN(d);
}

// ============================================================================
// 9. FUNGSI SAMBUNG SEWA (RENEWAL)
// ============================================================================

function submitRenewalApplication(data) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Applications');
  var rows = sheet.getDataRange().getValues();
  var oldData = null;

  // 1. Cari Data Lama
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.oldAppId) { // Column A: AppID
      oldData = rows[i];
      break;
    }
  }

  if (!oldData) return { success: false, message: "Rekod lama tidak dijumpai." };

  // 2. Jana ID Baru
  var newId = 'APP-R-' + Math.floor(Math.random() * 1000000000); // 'R' tanda Renewal

  // 3. Siapkan Data Baru (Copy Old Data + New Duration)
  // Structure: [ID, Date, Email, Name, IC, Phone, Syarikat, Alamat, Pekerjaan, Pendapatan, Lokasi, Jenis, TarikhMula(Baru), Tempoh(Baru), Tujuan, Status...]
  
  var newRow = [
    newId,                        // 0: ID Baru
    new Date(),                   // 1: Tarikh Mohon Sekarang
    oldData[2],                   // 2: Email (Sama)
    oldData[3],                   // 3: Nama (Sama)
    oldData[4],                   // 4: IC (Sama)
    oldData[5],                   // 5: Phone (Sama)
    oldData[6],                   // 6: Syarikat (Sama)
    oldData[7],                   // 7: Alamat (Sama)
    oldData[8],                   // 8: Pekerjaan (Sama)
    oldData[9],                   // 9: Pendapatan (Sama)
    oldData[10],                  // 10: Lokasi/Aset (Sama)
    oldData[11],                  // 11: Jenis (Sama)
    data.newStartDate,            // 12: Tarikh Mula (Input User)
    data.newDuration,             // 13: Tempoh Baru (Input User)
    "PEMBAHARUAN SEWAAN: " + oldData[0], // 14: Tujuan (Auto-Label)
    'Dalam Proses',               // 15: Status Reset
    '', '', '', '',               // 16-19: Kosongkan field admin
    oldData[16] || oldData[20]    // 20: Harga Asal (Ambil harga lulus terakhir)
  ];

  // 4. Masukkan ke Sheet
  sheet.appendRow(newRow);

  return { success: true, message: "Permohonan pembaharuan berjaya dihantar! ID: " + newId };
}

// ============================================================================
// 10. API ALERT & MONITORING
// ============================================================================

// A. Kira permohonan tertunggak > 3 Hari
function getOverdueStats() {
  var ss = getDb();
  var sheet = ss.getSheetByName('Applications');
  var data = sheet.getDataRange().getValues();
  var overdueCount = 0;
  var today = new Date();
  var threeDaysAgo = new Date();
  threeDaysAgo.setDate(today.getDate() - 3);

  // Status sasaran untuk dipantau
  var targetStatuses = ["Dalam Proses", "Disokong", "Lulus", "Menunggu Persetujuan"];

  for (var i = 1; i < data.length; i++) {
    var appDate = new Date(data[i][1]); // Lajur B: Tarikh Mohon
    var status = data[i][15]; // Lajur P: Status

    // Semak jika status termasuk dalam sasaran DAN tarikh mohon lebih tua dari 3 hari lepas
    if (targetStatuses.includes(status) && appDate < threeDaysAgo) {
      overdueCount++;
    }
  }
  return overdueCount;
}

// ============================================================================
// 11. API SENARAI MENUNGGU (WAITING LIST)
// ============================================================================

function joinWaitingList(data) {
  var ss = getDb();
  var sheet = ss.getSheetByName('WaitingList');
  
  // Jika sheet belum wujud, cipta baru
  if (!sheet) {
    sheet = ss.insertSheet('WaitingList');
    sheet.appendRow(['Tarikh', 'Nama Aset', 'Lokasi', 'Nama Pemohon', 'Emel', 'No Telefon', 'Nota']);
  }

  try {
    sheet.appendRow([
      new Date(),
      data.assetName,
      data.location,
      data.name,
      data.email,
      "'" + data.phone, // Tambah ' supaya sheet baca sebagai text
      data.note || "-"
    ]);
    return { success: true, message: "Anda telah berjaya dimasukkan ke Senarai Menunggu!" };
  } catch (e) {
    return { success: false, message: "Ralat: " + e.toString() };
  }
}

function getWaitingList() {
  var ss = getDb();
  var sheet = ss.getSheetByName('WaitingList');
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  var list = [];
  
  // Mula dari row 1 (skip header)
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      list.push({
        rowIndex: i + 1,
        date: formatDateDMY(data[i][0]),
        asset: data[i][1],
        location: data[i][2],
        name: data[i][3],
        email: data[i][4],
        phone: data[i][5],
        note: data[i][6] || "-"
      });
    }
  }
  return list.reverse(); // Tunjuk yang paling baru dahulu
}

function removeFromWaitingList(email, namaAsetGabungan) {
  var ss = getDb();
  var sheet = ss.getSheetByName('WaitingList');
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  var targetEmail = cleanString(email);
  
  // Kita loop dari bawah ke atas supaya bila delete row, index tak lari
  for (var i = data.length - 1; i > 0; i--) {
    var rowEmail = cleanString(data[i][4]); // Col E: Email
    var rowAsset = cleanString(data[i][1]); // Col B: Nama Aset (Waiting List)
    
    // Logic: Jika Email SAMA dan Nama Aset ada kaitan
    // (Guna .includes() sebab dalam Application nama aset mungkin digabung dengan lokasi)
    if (rowEmail === targetEmail && cleanString(namaAsetGabungan).includes(rowAsset)) {
       sheet.deleteRow(i + 1);
    }
  }
}

// ============================================================================
// 12. MODUL PENYELENGGARAAN (MAINTENANCE)
// ============================================================================

// A. HANTAR ADUAN BARU
function submitMaintenanceReport(data) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Maintenance');
  
  if (!sheet) {
    sheet = ss.insertSheet('Maintenance');
    sheet.appendRow(['TicketID', 'Tarikh', 'Nama Aset', 'Lokasi', 'Pengadu', 'Emel', 'No. Telefon', 'Isu Kerosakan', 'Status', 'Nota Admin', 'Tarikh Kemaskini']);
  }

  var ticketId = 'TKT-' + Math.floor(Math.random() * 100000);
  
  try {
    sheet.appendRow([
      ticketId,
      new Date(),
      data.assetName,
      data.location,
      data.reporterName,
      data.email,
      "'" + data.phone,
      data.issue,
      'Baru',           // Status Awal
      '-',              // Nota Admin
      '-'               // Tarikh Kemaskini
    ]);
    return { success: true, message: "Aduan berjaya dihantar. No Tiket: " + ticketId };
  } catch (e) {
    return { success: false, message: "Ralat: " + e.toString() };
  }
}

// B. DAPATKAN SENARAI ADUAN (USER & ADMIN)
function getMaintenanceTickets(email, role) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Maintenance');
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getDisplayValues();
  var tickets = [];
  var isStaff = (role === 'admin' || role === 'staff' || role === 'approver');
  var targetEmail = cleanString(email);

  // Mula row 1 (skip header)
  for (var i = 1; i < data.length; i++) {
    var rowEmail = cleanString(data[i][5]); // Col F: Emel Pengadu

    // Jika Admin/Staff -> Ambil SEMUA. Jika User biasa -> Ambil EMEL DIA SAHAJA.
    if (isStaff || rowEmail === targetEmail) {
       if (data[i][0]) {
         tickets.push({
           rowIndex: i + 1,
           id: data[i][0],
           date: formatDateDMY(data[i][1]),
           asset: data[i][2],
           location: data[i][3],
           reporter: data[i][4],
           email: data[i][5],
           phone: data[i][6],
           issue: data[i][7],
           status: data[i][8],
           adminNote: data[i][9],
           updated: data[i][10]
         });
       }
    }
  }
  return tickets.reverse();
}

// C. UPDATE STATUS ADUAN (ADMIN SAHAJA)
function updateTicketStatus(rowIndex, newStatus, adminNote) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Maintenance');
  
  if (sheet && rowIndex > 1) {
     sheet.getRange(rowIndex, 9).setValue(newStatus); // Col I: Status
     sheet.getRange(rowIndex, 10).setValue(adminNote); // Col J: Nota Admin
     sheet.getRange(rowIndex, 11).setValue(new Date()); // Col K: Tarikh Kemaskini
     return { success: true, message: "Status tiket berjaya dikemaskini." };
  }
  return { success: false, message: "Tiket tidak dijumpai." };
}

// ============================================================================
// 13. MODUL E-AGREEMENT (TANDATANGAN DIGITAL)
// ============================================================================

function saveSignedAgreement(appId, base64Pdf) {
  try {
    var ss = getDb();
    var sheet = ss.getSheetByName('Applications');
    var data = sheet.getDataRange().getDisplayValues();
    var rowIndex = -1;

    // 1. Cari row permohonan berdasarkan ID
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == appId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: "ID Permohonan tidak dijumpai." };

    var contentType = 'application/pdf';
    var cleanBase64 = base64Pdf.split(',')[1];
    var blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), contentType, "Perjanjian_" + appId + ".pdf");
    var file = DriveApp.createFile(blob);
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = file.getUrl();
    
    sheet.getRange(rowIndex, 26).setValue(fileUrl);
    
    currentNoteCell.setValue(newNote);

    return { success: true, message: "Perjanjian berjaya ditandatangani dan disimpan!", url: fileUrl };
  } catch (e) {
    return { success: false, message: "Ralat Server: " + e.toString() };
  }
}

// ============================================================================
// 14. API KALENDAR KEKOSONGAN (BARU)
// ============================================================================

function getAssetOccupancy(assetName) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Applications');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var bookedDates = [];
  var targetAsset = cleanString(assetName);

  // Mula dari baris 2 (skip header)
  for (var i = 1; i < data.length; i++) {
    var rowLokasi = cleanString(data[i][10]); // Col K: Lokasi/Nama Aset
    var status = data[i][15]; // Col P: Status

    // Hanya ambil yang status AKTIF (Berjaya, Lulus, Menunggu Persetujuan)
    // Dan nama aset sepadan
    if ((status === 'Berjaya' || status === 'Lulus' || status === 'Menunggu Persetujuan') && 
        rowLokasi.includes(targetAsset)) {
      
      var startDate = new Date(data[i][12]); // Col M: Tarikh Mula
      var tempohStr = String(data[i][13]);   // Col N: Tempoh
      
      if (isValidDate(startDate)) {
        var endDate = calculateEndDate(startDate, tempohStr);
        // Simpan range tarikh
        bookedDates.push({
          start: startDate.getTime(),
          end: endDate.getTime()
        });
      }
    }
  }
  return bookedDates;
}

// ============================================================================
// 15. MODUL JEJAK AUDIT (AUDIT TRAIL)
// ============================================================================

// Fungsi Rekod Log (Dipanggil oleh fungsi lain)
function logAudit(email, role, action, details) {
  try {
    var ss = getDb();
    var sheet = ss.getSheetByName('Logs');
    
    // Jika sheet tiada, cipta baru
    if (!sheet) {
      sheet = ss.insertSheet('Logs');
      sheet.appendRow(['Timestamp', 'User Email', 'Role', 'Action', 'Details']);
      sheet.setColumnWidth(1, 150); // Lebarkan kolum tarikh
      sheet.setColumnWidth(5, 300); // Lebarkan kolum details
    }
    
    // Masukkan data log
    sheet.appendRow([new Date(), email, role, action, details]);
    
  } catch (e) {
    // Jangan hentikan sistem jika log gagal, cuma rekod di konsol
    console.error("Gagal rekod audit: " + e.toString()); 
  }
}

// API untuk Admin baca Logs
function getAuditLogs() {
  var ss = getDb();
  var sheet = ss.getSheetByName('Logs');
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getDisplayValues();
  var logs = [];
  
  // Baca dari bawah (terkini dulu), hadkan 500 rekod terakhir untuk laju
  var limit = 500;
  var count = 0;
  
  for (var i = data.length - 1; i > 0; i--) {
    logs.push({
      timestamp: data[i][0],
      email: data[i][1],
      role: data[i][2],
      action: data[i][3],
      details: data[i][4]
    });
    count++;
    if (count >= limit) break;
  }
  return logs;
}

// ============================================================================
// 16. MODUL TETAPAN SISTEM (SYSTEM SETTINGS) - BARU
// ============================================================================

function getSystemSettings() {
  var ss = getDb();
  var sheet = ss.getSheetByName('Settings');
  
  // Default Settings jika sheet belum wujud
  var defaults = {
    title: "PORTAL SEWAAN ASET",
    subtitle: "Lembaga Kemajuan Ikan Malaysia",
    logo: "https://upload.wikimedia.org/wikipedia/ms/2/28/Logo_LKIM.png", // Logo asal
    phone: "03-8888 1234",
    email: "aset@lkim.gov.my",
    address: "Aras 5, Menara LKIM, Jalan Topaz 7/4, Seksyen 7, 47120 Puchong, Selangor"
  };

  if (!sheet) {
    return defaults;
  }

  // Kita guna format Key-Value di Column A dan B
  // Row 1: Title, Row 2: Subtitle, dsb...
  // Tapi lebih mudah guna 1 Row data (Header Row 1, Data Row 2)
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return defaults;

  var row = data[1]; // Row 2 (Index 1)
  return {
    title: row[0] || defaults.title,
    subtitle: row[1] || defaults.subtitle,
    logo: row[2] || defaults.logo,
    phone: row[3] || defaults.phone,
    email: row[4] || defaults.email,
    address: row[5] || defaults.address
  };
}

function saveSystemSettings(form) {
  var ss = getDb();
  var sheet = ss.getSheetByName('Settings');
  
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    sheet.appendRow(['Title', 'Subtitle', 'LogoURL', 'Phone', 'Email', 'Address']); // Header
    sheet.appendRow(['', '', '', '', '', '']); // Placeholder Row
  }

  var logoUrl = form.currentLogo;
  // Jika ada file baru diupload
  if (form.logoFile && form.logoFile.data) {
     try {
        logoUrl = uploadToDrive(form.logoFile.data, form.logoFile.name, form.logoFile.type);
     } catch(e) {
        return { success: false, message: "Gagal muat naik logo: " + e.toString() };
     }
  }

  // Simpan di Row 2 (Overwrite)
  sheet.getRange(2, 1, 1, 6).setValues([[
    form.title, 
    form.subtitle, 
    logoUrl, 
    form.phone, 
    form.email, 
    form.address
  ]]);

  return { success: true, message: "Tetapan sistem berjaya disimpan!" };
}

// ============================================================================
// 17. API LAPORAN INVENTORI ASET (BARU)
// ============================================================================

function getAssetInventoryStats() {
  var ss = getDb();
  var sheet = ss.getSheetByName('Assets');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var stats = {};
  
  // Mula dari row 1 (skip header)
  for (var i = 1; i < data.length; i++) {
    var kategori = data[i][1]; // Col B: Kategori
    var status = data[i][3];   // Col D: Status

    // Bersihkan nama kategori (trim spaces)
    if (kategori) {
      kategori = String(kategori).trim();
      
      if (!stats[kategori]) {
        stats[kategori] = { nama: kategori, total: 0, aktif: 0, tidakAktif: 0 };
      }

      stats[kategori].total++;
      
      // Logik status: 'Aktif' dikira aktif, selebihnya (Penyelenggaraan/Penuh/Rosak) dikira Tidak Aktif
      // Anda boleh ubah logik ini jika 'Penuh' patut dikira Aktif.
      if (status === 'Aktif') {
        stats[kategori].aktif++;
      } else {
        stats[kategori].tidakAktif++;
      }
    }
  }

  // Tukar object kepada array untuk dihantar ke frontend
  var reportArray = Object.keys(stats).map(function(key) {
    return stats[key];
  });

  // Susun ikut abjad kategori
  reportArray.sort(function(a, b) {
    return a.nama.localeCompare(b.nama);
  });

  return reportArray;
}

// ============================================================================
// 18. MODUL LUPA KATA LALUAN (RESET PASSWORD)
// ============================================================================

function requestPasswordReset(email) {
  var ss = getDb();
  var userSheet = ss.getSheetByName('Users');
  var tokenSheet = ss.getSheetByName('ResetTokens');
  
  // 1. Pastikan Sheet ResetTokens wujud
  if (!tokenSheet) tokenSheet = ss.insertSheet('ResetTokens');

  var cleanEmail = cleanString(email);
  
  // 1. RATE LIMITING CHECK
  var tokens = tokenSheet.getDataRange().getValues();
  var now = new Date().getTime();
  var twoMinutes = 2 * 60 * 1000;

  // Semak dari bawah (rekod terkini)
  for (var i = tokens.length - 1; i > 0; i--) {
    var rowEmail = cleanString(tokens[i][0]);
    var timestamp = tokens[i][3]; // Kita perlu tambah col timestamp created jika belum ada. 
    // Tapi kita boleh guna expiry masa (Column C - Index 2). 
    // Expiry = Created + 15 min. So, Created = Expiry - 15 min.
    
    if (rowEmail === cleanEmail) {
      var expiryTime = parseFloat(tokens[i][2]);
      var createdTime = expiryTime - (15 * 60 * 1000); 
      
      // Jika request terakhir dibuat kurang dari 2 minit lepas
      if ((now - createdTime) < twoMinutes) {
        return { success: false, message: "Sila tunggu 2 minit sebelum meminta OTP baru." };
      }
      break; // Jumpa latest record, stop checking
    }
  }

// 2. Semak User Wujud
  var userFound = false;
  var userData = userSheet.getDataRange().getValues();
  for (var i = 1; i < userData.length; i++) {
    if (cleanString(userData[i][0]) == cleanEmail) {
      userFound = true;
      break;
    }
  }

  if (!userFound) {
    return { success: false, message: "Emel tidak dijumpai dalam sistem." };
  }

  // 3. Jana OTP & Tarikh Luput (15 Minit dari sekarang)
  var otp = Math.floor(100000 + Math.random() * 900000).toString();
  var expiry = new Date().getTime() + (15 * 60 * 1000); // 15 minit

  // 4. Simpan ke Database (ResetTokens)
  tokenSheet.appendRow([cleanEmail, otp, expiry]);

  // 5. Hantar Emel OTP
  var subject = "Kod Reset Kata Laluan Portal LKIM";
  var body = "<p>Kod OTP anda ialah: <strong>" + otp + "</strong></p><p>Kod ini sah selama 15 minit.</p>";
  
  // Pastikan fungsi sendEmailNotification aktif atau guna MailApp terus
  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: body,
      name: "Sistem e-Sewaan LKIM"
    });
    return { success: true, message: "Kod OTP telah dihantar ke emel anda." };
  } catch (e) {
    console.log("OTP Error for " + email + ": " + e.toString());
    return { success: false, message: "Gagal hantar emel. Sila hubungi Admin." };
  }
}

function verifyAndResetPassword(email, otp, newPassword) {
  var ss = getDb();
  var tokenSheet = ss.getSheetByName('ResetTokens');
  var userSheet = ss.getSheetByName('Users');
  
  if (!tokenSheet) return { success: false, message: "Sistem ralat: Token sheet missing." };

  var cleanEmail = cleanString(email);
  var tokens = tokenSheet.getDataRange().getValues();
  var now = new Date().getTime();
  var validTokenIndex = -1;

  // 1. Cari Token yang Sah (Email sama, OTP sama, Belum Expire)
  // Kita cari dari bawah (terkini) ke atas
  for (var i = tokens.length - 1; i > 0; i--) {
    var rowEmail = cleanString(tokens[i][0]);
    var rowOtp = String(tokens[i][1]).trim();
    var rowExpiry = parseFloat(tokens[i][2]);

    if (rowEmail === cleanEmail && rowOtp === String(otp).trim()) {
      if (now < rowExpiry) {
        validTokenIndex = i + 1; // Jumpa!
        break;
      } else {
        return { success: false, message: "Kod OTP telah tamat tempoh." };
      }
    }
  }

  if (validTokenIndex === -1) {
    return { success: false, message: "Kod OTP tidak sah atau salah." };
  }

  // 2. Tukar Password Pengguna
  var users = userSheet.getDataRange().getValues();
  var userIndex = -1;

  for (var j = 1; j < users.length; j++) {
    if (cleanString(users[j][0]) === cleanEmail) {
      userIndex = j + 1;
      break;
    }
  }

  if (userIndex !== -1) {
    // Hash password baru sebelum simpan
    userSheet.getRange(userIndex, 2).setValue(hashPassword(newPassword));
    
    // 3. Padam token yang dah guna (Pilihan: Atau biarkan saja)
    tokenSheet.deleteRow(validTokenIndex);

    return { success: true, message: "Kata laluan berjaya ditukar. Sila log masuk." };
  }

  return { success: false, message: "Pengguna tidak dijumpai." };
}

// ============================================================================
// 19. MODUL JANA PDF LAPORAN INVENTORI
// ============================================================================

function getInventoryReportHTML() {
  var ss = getDb(); //
  var settings = getSystemSettings(); //
  
  // Dapatkan data statistik
  var stats = getAssetInventoryStats(); //
  var totalAssets = 0;
  var totalActive = 0;
  var totalInactive = 0;

  // Kira total
  stats.forEach(function(item){
    totalAssets += item.total;
    totalActive += item.aktif;
    totalInactive += item.tidakAktif;
  });

  var reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  // HTML Putih Bersih (Sama macam tadi, tapi tiada penukaran PDF)
  var html = `
    <!DOCTYPE html>
    <html>
      <head>
        <title>Laporan Inventori Aset</title>
        <style>
          body { font-family: Arial, sans-serif; font-size: 12px; color: #000; padding: 40px; background: white; }
          .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #000; padding-bottom: 10px; }
          .header h1 { margin: 0; font-size: 18px; text-transform: uppercase; }
          .header p { margin: 5px 0 0; font-size: 10px; }
          .info { margin-bottom: 20px; text-align: right; font-size: 10px; }
          table { width: 100%; border-collapse: collapse; margin-top: 10px; }
          th, td { border: 1px solid #000; padding: 8px; text-align: center; }
          th { background-color: #f0f0f0; font-weight: bold; }
          .text-left { text-align: left; }
          .footer-row { background-color: #ddd; font-weight: bold; }
          
          /* Butang Print (Hanya nampak di skrin, hilang bila print) */
          .no-print { text-align: center; margin-bottom: 20px; }
          .btn-print { background: #1e293b; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; font-weight: bold; }
          @media print {
            .no-print { display: none; }
            body { padding: 0; }
          }
        </style>
      </head>
      <body>
        <div class="no-print">
           <button onclick="window.print()" class="btn-print">🖨️ CETAK SEKARANG</button>
           <p style="font-size:10px; color:#666; margin-top:5px;">(Format ini telah dioptimumkan untuk kertas A4)</button>
        </div>

        <div class="header">
          <h1>${settings.title.toUpperCase()}</h1>
          <p>${settings.subtitle}</p>
          <p>${settings.address}</p>
        </div>

        <div class="info">
          <strong>LAPORAN RINGKASAN INVENTORI ASET</strong><br>
          Dijana pada: ${reportDate}
        </div>

        <div style="text-align: center; margin-bottom: 20px; border: 1px solid #ccc; padding: 15px;">
           <div style="display:inline-block; width: 30%; border-right: 1px solid #ccc;">
              <span style="font-size: 20px; font-weight: bold;">${totalAssets}</span><br>
              <span style="font-size: 9px;">JUMLAH ASET</span>
           </div>
           <div style="display:inline-block; width: 30%; border-right: 1px solid #ccc;">
              <span style="font-size: 20px; font-weight: bold;">${totalActive}</span><br>
              <span style="font-size: 9px;">AKTIF</span>
           </div>
           <div style="display:inline-block; width: 30%;">
              <span style="font-size: 20px; font-weight: bold;">${totalInactive}</span><br>
              <span style="font-size: 9px;">TIDAK AKTIF</span>
           </div>
        </div>

        <table>
          <thead>
            <tr>
              <th width="5%">BIL</th>
              <th width="45%" class="text-left">KATEGORI ASET</th>
              <th width="15%">JUMLAH UNIT</th>
              <th width="15%">AKTIF</th>
              <th width="20%">ROSAK / PENUH</th>
            </tr>
          </thead>
          <tbody>
  `;

  for (var i = 0; i < stats.length; i++) {
    html += `
      <tr>
        <td>${i + 1}</td>
        <td class="text-left"><strong>${stats[i].nama}</strong></td>
        <td>${stats[i].total}</td>
        <td>${stats[i].aktif}</td>
        <td>${stats[i].tidakAktif}</td>
      </tr>
    `;
  }

  html += `
            <tr class="footer-row">
              <td colspan="2" class="text-left">JUMLAH BESAR</td>
              <td>${totalAssets}</td>
              <td>${totalActive}</td>
              <td>${totalInactive}</td>
            </tr>
          </tbody>
        </table>

        <br><br>
        <div style="font-size: 9px; text-align: center; color: #555; margin-top: 50px;">
           <i>Dokumen ini dijana oleh Sistem e-Sewaan Aset LKIM.</i>
        </div>
        
        <script>
           // Auto Print bila laman dibuka
           window.onload = function() { setTimeout(function(){ window.print(); }, 500); }
        </script>
      </body>
    </html>
  `;

  return { success: true, html: html };
}
