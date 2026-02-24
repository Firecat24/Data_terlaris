const NAMA_SHEET_DASBOR = "Analisis";
const NAMA_SHEET_HELPER = "_Helper";
const BARIS_AWAL = 3;
const KOLOM_BULAN = 1;   // Kolom A
const KOLOM_NAMA = 2;    // Kolom B
const KOLOM_SATUAN = 3;  // Kolom C
const KOLOM_OUTPUT_START = 4; // Kolom D (QTY 2023)

// Konfigurasi kolom di sheet bulanan (JANUARI, FEBRUARI, dll.)
const DATA_RANGES = [
  { nama: "B3:B", satuan: "C3:C", qty: "D3:D" }, // Tabel 2023
  { nama: "F3:F", satuan: "G3:G", qty: "H3:H" }, // Tabel 2024
  { nama: "J3:J", satuan: "K3:K", qty: "L3:L" }  // Tabel 2025
];

/**
 * Fungsi utama yang berjalan setiap kali ada editan.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (sheet.getName() !== NAMA_SHEET_DASBOR || row < BARIS_AWAL) {
    return;
  }

  const bulanCell = sheet.getRange(row, KOLOM_BULAN);
  const namaCell = sheet.getRange(row, KOLOM_NAMA);
  const satuanCell = sheet.getRange(row, KOLOM_SATUAN);
  const outputRange = sheet.getRange(row, KOLOM_OUTPUT_START, 1, DATA_RANGES.length); 

  const bulanTerpilih = bulanCell.getValue();

  // KASUS 1: Pengguna mengubah BULAN (Kolom A)
  if (col === KOLOM_BULAN) {
    namaCell.clearContent().clearDataValidations();
    satuanCell.clearContent().clearDataValidations();
    outputRange.clearContent();
    if (bulanTerpilih) {
      buatDropdownNama(ss, bulanTerpilih, namaCell);
    }
  } 
  // KASUS 2: Pengguna mengubah NAMA (Kolom B)
  else if (col === KOLOM_NAMA) {
    satuanCell.clearContent().clearDataValidations();
    outputRange.clearContent();
    const namaTerpilih = namaCell.getValue();
    if (namaTerpilih) {
      buatDropdownSatuan(ss, bulanTerpilih, namaTerpilih, satuanCell);
    }
  }
  // KASUS 3: Pengguna mengubah SATUAN (Kolom C)
  else if (col === KOLOM_SATUAN) {
    const namaTerpilih = namaCell.getValue();
    const satuanTerpilih = satuanCell.getValue();
    if (bulanTerpilih && namaTerpilih && satuanTerpilih) {
      ambilDataQty(ss, bulanTerpilih, namaTerpilih, satuanTerpilih, outputRange);
    } else {
      outputRange.clearContent();
    }
  }
}

/**
 * ====================================================================
 * HELPER 1: Membuat dropdown "Nama" (VERSI PERBAIKAN 500+ ITEM)
 * ====================================================================
 */
function buatDropdownNama(ss, namaSheet, targetCell) {
  try {
    const sheetBulan = ss.getSheetByName(namaSheet);
    if (!sheetBulan) return; 

    // 1. Ambil semua nama (sama seperti sebelumnya)
    let daftarNama = [];
    DATA_RANGES.forEach(rangeInfo => {
      const values = sheetBulan.getRange(rangeInfo.nama).getValues();
      values.forEach(row => {
        if (row[0]) daftarNama.push(row[0]);
      });
    });
    const uniqueNama = [...new Set(daftarNama)].filter(Boolean);
    
    if (uniqueNama.length === 0) return; // Keluar jika tidak ada nama

    // 2. Dapatkan atau buat sheet helper
    let helperSheet = ss.getSheetByName(NAMA_SHEET_HELPER);
    if (!helperSheet) {
      helperSheet = ss.insertSheet(NAMA_SHEET_HELPER);
      helperSheet.hideSheet();
    }
    const helperRange = helperSheet.getRange(1, 1, helperSheet.getMaxRows());
    helperRange.clearContent();
    const dataUntukSheet = uniqueNama.map(item => [item]);
    const targetHelperRange = helperSheet.getRange(1, 1, dataUntukSheet.length, 1);
    targetHelperRange.setValues(dataUntukSheet);

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(targetHelperRange) 
      .setAllowInvalid(true)
      .build();
      
    targetCell.setDataValidation(rule);
    
  } catch (err) {
    Logger.log(`Error di buatDropdownNama: ${err} - ${err.stack}`);
  }
}