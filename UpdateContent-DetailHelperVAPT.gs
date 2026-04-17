/**
 * Fungsi Update Content Helper VAPT
 * Fitur: Sumber "Content - Helper" (tetap), Target dinamis, Ninja Copy A1:K11, dan Auto Rata Kiri baris 10+.
 */
function jalankanUpdateContentHelperVAPT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // 1. Konfigurasi Fix
  const sourceSheetName = "Content - Helper"; // Sesuai dengan nama tab sumber di Master
  const targetUpdateText = "Content - Helper"; 

  const listSheet = ss.getSheetByName("Report List");
  const sourceSheet = ss.getSheetByName(sourceSheetName);

  if (!sourceSheet || !listSheet) {
    ui.alert(`Error: Tab '${sourceSheetName}' atau 'Report List' tidak ditemukan!`);
    return;
  }

  // 2. Dinamisasi Kolom
  const lastRowList = listSheet.getLastRow();
  const lastColList = listSheet.getLastColumn();

  if (lastRowList < 2) {
    ui.alert("Daftar di 'Report List' kosong.");
    return;
  }

  const headers = listSheet.getRange(1, 1, 1, lastColList).getValues()[0].map(h => h.toString().trim().toLowerCase());
  
  let colIdIdx = -1, colTabIdx = -1, colUpdateIdx = -1, colRunIdx = -1;

  // Pencarian kata kunci kolom
  headers.forEach((h, i) => {
    if (h.includes("id") || h.includes("link")) colIdIdx = i; 
    if (h.includes("tab")) colTabIdx = i;                     
    if (h.includes("update")) colUpdateIdx = i;               
    if (h.includes("run")) colRunIdx = i;                     
  });

  if (colIdIdx === -1 || colTabIdx === -1 || colUpdateIdx === -1) {
    ui.alert("Error: Kolom ID/Link, Tab Name, atau Update tidak ditemukan di baris pertama (Header).");
    return;
  }

  const listData = listSheet.getRange(2, 1, lastRowList - 1, lastColList).getValues();

  let logStatus = [];
  let successCount = 0;
  let skipCount = 0;

  // 3. Looping Eksekusi
  listData.forEach((row, index) => {
    const rowNum = index + 2;
    
    const reportId = row[colIdIdx] ? row[colIdIdx].toString().trim() : ""; 
    const tabName = row[colTabIdx] ? row[colTabIdx].toString().trim() : "";  
    const rawUpdateValue = row[colUpdateIdx] ? row[colUpdateIdx].toString() : "";
    const updateType = rawUpdateValue.replace(/\u00A0/g, ' ').trim(); 

    let isRun = false;
    if (colRunIdx !== -1) {
      isRun = row[colRunIdx] === true || (row[colRunIdx] && row[colRunIdx].toString().trim().toUpperCase() === "TRUE");
    }

    // FILTER: Centang TRUE dan Update harus sama dengan "Content - Helper"
    if (!isRun || updateType.toLowerCase() !== targetUpdateText.toLowerCase()) {
      skipCount++;
      return; 
    }

    if (!reportId || !tabName) {
      logStatus.push(`❌ Baris ${rowNum}: ID Spreadsheet target atau Nama Tab kosong.`);
      return;
    }

    let tempSheet = null;
    let targetSs = null;

    try {
      targetSs = SpreadsheetApp.openById(reportId);
      let targetSheet = targetSs.getSheetByName(tabName); // Targetnya baru menggunakan tabName (misal: "helper")

      if (!targetSheet) {
        logStatus.push(`❌ Baris ${rowNum}: Gagal. Tab TARGET '${tabName}' tidak ditemukan di file target.`);
        return;
      }

      // ==============================================================
      // PROSES COPY PASTE (Ninja Copy untuk 100% Presisi Format)
      // ==============================================================
      
      tempSheet = sourceSheet.copyTo(targetSs);
      tempSheet.setName("Temp_Helper_" + Math.random());
      
      const sourceRange = tempSheet.getRange("A1:K11");
      const targetRange = targetSheet.getRange("A1:K11");

      // Paste semua data & format dari Master A1:K11
      sourceRange.copyTo(targetRange);
      
      // Sinkronisasi Ukuran Sel
      for (let i = 1; i <= 11; i++) {
        targetSheet.setColumnWidth(i, tempSheet.getColumnWidth(i));
        targetSheet.setRowHeight(i, tempSheet.getRowHeight(i));
      }

      // ==============================================================
      // AUTO RATA KIRI UNTUK BARIS 10 KE BAWAH
      // ==============================================================
      let lastDataRow = targetSheet.getLastRow();
      // Cek apakah data (hasil formula) sampai baris 10 atau lebih
      if (lastDataRow >= 10) {
        // Ambil range dari baris 10 sampai baris terakhir, dan paksa rata kiri (Left Alignment)
        targetSheet.getRange(10, 1, lastDataRow - 9, 11).setHorizontalAlignment("left");
      }

      successCount++;
      logStatus.push(`✅ Baris ${rowNum}: BERHASIL disalin ke tab '${tabName}' (${targetSs.getName()}).`);

    } catch (e) {
      logStatus.push(`⚠️ Baris ${rowNum}: ERROR - ${e.message}`);
    } finally {
      // Hapus tab bayangan
      if (targetSs && tempSheet) {
        try { targetSs.deleteSheet(tempSheet); } catch(err) {} 
      }
    }
  });

  // 4. Log Pop-up
  const summary = `<b>Proses Update Content Helper Selesai.</b><br>Berhasil: ${successCount}<br>Dilewati: ${skipCount}<br><br><b>Detail Tracing:</b><br>`;
  const htmlOutput = HtmlService.createHtmlOutput(
    `<div style="font-family: sans-serif; font-size: 13px;">${summary}<pre style="background:#f4f4f4;padding:10px;height:250px;overflow:auto;white-space:pre-wrap;">${logStatus.join("\n")}</pre></div>`
  ).setWidth(650).setHeight(480);
  ui.showModalDialog(htmlOutput, 'Execution Log - Update Helper');
}