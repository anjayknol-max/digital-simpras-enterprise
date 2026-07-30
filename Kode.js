const namaSheetData = "Sheet1";

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('DIGITAL SIMPRAS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function checkLogin(username, password) {
  if (username === "sarpras" && password === "admin123") {
    return { success: true, role: "sarpras", message: "Login Sarpras berhasil" };
  } else if (username === "asrama" && password === "admin123") {
    return { success: true, role: "asrama", message: "Login Asrama berhasil" };
  } else {
    return { success: false, message: "Username atau Password salah!" };
  }
}

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(namaSheetData);
    if (!sheet || sheet.getLastRow() < 2) {
      return { rows: [], unitKerjaList: [], tahunList: [] };
    }

    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(2, 1, lastRow - 1, 10);
    const dataMentah = range.getValues();
    const dataTampilan = range.getDisplayValues();

    const unitKerjaSet = new Set();
    const tahunSet = new Set();

    const rows = dataTampilan.map((baris, i) => {
      let tahunDitemukan = "";
      const nilaiTanggal = dataMentah[i][0];

      if (nilaiTanggal instanceof Date && !isNaN(nilaiTanggal)) {
        tahunDitemukan = nilaiTanggal.getFullYear().toString();
      } else {
        const cocokTahun = baris[0].match(/\d{4}/);
        tahunDitemukan = cocokTahun ? cocokTahun[0] : "";
      }

      if (tahunDitemukan) tahunSet.add(tahunDitemukan);
      const unit = baris[3] ? baris[3].toString().trim() : "";
      if (unit) unitKerjaSet.add(unit);

      const status = baris[9] && baris[9].toString().trim() !== "" ? baris[9].toString().trim() : "Proses";
      baris[9] = status;

      baris.push(tahunDitemukan);
      return baris;
    });

    const unitKerjaList = Array.from(unitKerjaSet).sort();
    const tahunList = Array.from(tahunSet).sort((a, b) => b - a);

    return { rows, unitKerjaList, tahunList };
  } catch (e) {
    return { error: e.message };
  }
}

function simpanLaporanBaru(formData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(namaSheetData);
    if (!sheet) {
      sheet = ss.insertSheet(namaSheetData);
      sheet.appendRow(["Timestamp", "Nama", "HP", "Unit Kerja", "Barang", "Deskripsi", "Lokasi", "", "", "Status"]);
    }

    const date = new Date();
    const months = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
    const formattedDate = `${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()}`;

    sheet.appendRow([formattedDate, formData.nama, formData.hp, formData.unit, formData.barang, formData.deskripsi, formData.lokasi, "", "", "Proses"]);
    return { success: true, message: "Laporan berhasil disimpan!" };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateStatusLaporan(arrayIndex, newStatus) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheetData);
    if (!sheet) return { success: false, message: "Sheet data tidak ditemukan." };
    sheet.getRange(Number(arrayIndex) + 2, 10).setValue(newStatus);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function simpanLabel(labelData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Sheet3");
    if (!sheet) {
      sheet = ss.insertSheet("Sheet3");
      sheet.appendRow(["Timestamp", "Barang", "Part", "Tanggal", "Teknisi"]);
    }
    sheet.appendRow([new Date(), labelData.barang, labelData.part, labelData.tanggal, labelData.teknisi]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getAsramaData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Asrama");
  if (!sheet) {
    sheet = ss.insertSheet("Asrama");
    sheet.appendRow(["No Kamar", "Bed 1", "Bed 2", "Bed 3", "Bed 4"]);
    for (let i = 1; i <= 16; i++) {
      sheet.appendRow([i, "", "", "", ""]);
    }
    sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#2d3436").setFontColor("white");
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(5, 200);
  }

  const data = sheet.getDataRange().getValues();
  return data.slice(1);
}

function simpanDataAsrama(kamarNo, bed1, bed2, bed3, bed4) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Asrama");
    if (!sheet) return { success: false, message: "Sheet Asrama tidak ditemukan." };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == kamarNo) {
        sheet.getRange(i + 1, 2, 1, 4).setValues([[bed1, bed2, bed3, bed4]]);
        return { success: true };
      }
    }
    return { success: false, message: "Kamar tidak ditemukan di database." };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function bersihkanTimestampGForm(e) {
  Utilities.sleep(2000);
  try {
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    const rangeTanggal = sheet.getRange(row, 1);
    const nilaiAsli = rangeTanggal.getValue();
    if (nilaiAsli instanceof Date) {
      const daftarBulan = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
      const tglSaja = `${nilaiAsli.getDate()} ${daftarBulan[nilaiAsli.getMonth()]} ${nilaiAsli.getFullYear()}`;
      rangeTanggal.setValue(tglSaja);
      sheet.getRange(row, 10).setValue("Proses");
    }
  } catch (err) {
    console.error("Gagal karena: " + err.message);
  }
}