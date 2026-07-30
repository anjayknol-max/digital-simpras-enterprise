function checkLogin(username, password) {
  if (username === 'sarpras' && password === ADMIN_CREDENTIALS.sarpras) {
    return { success: true, role: 'sarpras', message: 'Login Sarpras berhasil' };
  } else if (username === 'asrama' && password === ADMIN_CREDENTIALS.asrama) {
    return { success: true, role: 'asrama', message: 'Login Asrama berhasil' };
  }
  return { success: false, message: 'Username atau Password salah!' };
}

function getDashboardData() {
  try {
    const sheet = getSheetByName(SHEET_DATA_NAME);
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
      const tahunDitemukan = getYearFromValue(dataMentah[i][0]);
      if (tahunDitemukan) tahunSet.add(tahunDitemukan);

      const unit = baris[3] ? baris[3].toString().trim() : '';
      if (unit) unitKerjaSet.add(unit);

      baris[9] = normalizeStatus(baris[9]);
      baris.push(tahunDitemukan);
      return baris;
    });

    return {
      rows,
      unitKerjaList: Array.from(unitKerjaSet).sort(),
      tahunList: Array.from(tahunSet).sort((a, b) => Number(b) - Number(a))
    };
  } catch (e) {
    return { error: e.message };
  }
}

function simpanLaporanBaru(formData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = ensureSheet(SHEET_DATA_NAME, ['Timestamp','Nama','HP','Unit Kerja','Barang','Deskripsi','Lokasi','','','Status']);
    const date = new Date();
    const formattedDate = formatDateIndo(date);

    sheet.appendRow([formattedDate, formData.nama, formData.hp, formData.unit, formData.barang, formData.deskripsi, formData.lokasi, '', '', 'Proses']);
    return { success: true, message: 'Laporan berhasil disimpan!' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateStatusLaporan(arrayIndex, newStatus) {
  try {
    const sheet = getSheetByName(SHEET_DATA_NAME);
    if (!sheet) return { success: false, message: 'Sheet data tidak ditemukan.' };
    sheet.getRange(Number(arrayIndex) + 2, 10).setValue(newStatus);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function simpanLabel(labelData) {
  try {
    const sheet = ensureSheet(SHEET_LABEL_NAME, ['Timestamp','Barang','Part','Tanggal','Teknisi']);
    sheet.appendRow([new Date(), labelData.barang, labelData.part, labelData.tanggal, labelData.teknisi]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getAsramaData() {
  const sheet = ensureSheet(SHEET_ASRAMA_NAME, ['No Kamar','Bed 1','Bed 2','Bed 3','Bed 4']);
  if (sheet.getLastRow() === 1) {
    for (let i = 1; i <= 16; i++) {
      sheet.appendRow([i, '', '', '', '']);
    }
    sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#2d3436').setFontColor('white');
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
    const sheet = getSheetByName(SHEET_ASRAMA_NAME);
    if (!sheet) return { success: false, message: 'Sheet Asrama tidak ditemukan.' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == kamarNo) {
        sheet.getRange(i + 1, 2, 1, 4).setValues([[bed1, bed2, bed3, bed4]]);
        return { success: true };
      }
    }
    return { success: false, message: 'Kamar tidak ditemukan di database.' };
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
      rangeTanggal.setValue(formatDateIndo(nilaiAsli));
      sheet.getRange(row, 10).setValue('Proses');
    }
  } catch (err) {
    console.error('Gagal karena: ' + err.message);
  }
}
