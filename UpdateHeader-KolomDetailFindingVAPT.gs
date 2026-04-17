/**
 * Fungsi Standarisasi Kolom - Detail Finding VAPT
 * Fitur: Smart Sorting, Auto-Delete Kosong, Mark Red, Auto-Width, Safe Auto-Filter, 
 * dan Freeze View (Hanya Header Row, Nol-kan Freeze Kolom).
 */
function jalankanUpdateKolomDetailFindingVAPT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const detailSheetName = "Kolom - Detail Finding VAPT"; 
  const listSheet = ss.getSheetByName("Report List");
  const detailSheet = ss.getSheetByName(detailSheetName);

  if (!detailSheet || !listSheet) {
    ui.alert(`Error: Tab '${detailSheetName}' atau 'Report List' tidak ditemukan!`);
    return;
  }

  // Pengaturan Lebar Kolom Standar (Bisa Anda ubah angkanya)
  const DEFAULT_COL_WIDTH = 120; 

  // 1. Ambil Standar Kolom (A: Huruf, B: Nama, C: Note)
  const lastRowDetail = detailSheet.getLastRow();
  const standardConfig = detailSheet.getRange(2, 1, lastRowDetail - 1, 3).getValues();
  
  let standardCols = [];
  standardConfig.forEach((config) => {
    let name = config[1] ? config[1].toString().trim() : "";
    let note = config[2] ? config[2].toString().trim() : "";
    if (name !== "") {
      standardCols.push({ name: name, note: note });
    }
  });

  // 2. Ambil List Report dan Dinamisasi Kolom
  const lastRowList = listSheet.getLastRow();
  const lastColList = listSheet.getLastColumn();

  if (lastRowList < 2) {
    ui.alert("Daftar di 'Report List' kosong.");
    return;
  }

  // Baca header dan cari index kolom secara dinamis
  const headers = listSheet.getRange(1, 1, 1, lastColList).getValues()[0].map(h => h.toString().trim().toLowerCase());
  
  let colIdIdx = -1, colTabIdx = -1, colHeaderRowIdx = -1, colUpdateIdx = -1, colRunIdx = -1;

  headers.forEach((h, i) => {
    if (h.includes("id") || h.includes("link")) colIdIdx = i; 
    if (h.includes("tab")) colTabIdx = i;                     
    if (h.includes("header row")) colHeaderRowIdx = i;        
    if (h.includes("update")) colUpdateIdx = i;               
    if (h.includes("run")) colRunIdx = i;                     
  });

  if (colIdIdx === -1 || colTabIdx === -1 || colUpdateIdx === -1) {
    ui.alert("Error: Kolom ID/Link, Tab Name, atau Update tidak ditemukan di baris pertama 'Report List'.");
    return;
  }
  
  const listRange = listSheet.getRange(2, 1, lastRowList - 1, lastColList);
  const listData = listRange.getValues();
  const richTextData = listRange.getRichTextValues();

  let logStatus = [];
  let successCount = 0;
  let skipCount = 0;

  listData.forEach((row, index) => {
    const rowNum = index + 2;
    
    // SMART ID EXTRACTOR (Support Smart Chips & URL)
    let rawReportId = row[colIdIdx] ? row[colIdIdx].toString().trim() : ""; 
    let linkUrl = richTextData[index][colIdIdx].getLinkUrl();
    let reportId = linkUrl ? linkUrl : rawReportId;
    let matchId = reportId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (matchId) reportId = matchId[1];

    const tabName = row[colTabIdx] ? row[colTabIdx].toString().trim() : "";  
    const headerRow = (colHeaderRowIdx !== -1 && row[colHeaderRowIdx]) ? parseInt(row[colHeaderRowIdx]) || 10 : 10; 
    const updateType = row[colUpdateIdx] ? row[colUpdateIdx].toString().trim() : ""; 

    // BACA KOLOM RUN
    let isRun = row[colRunIdx] === true || (row[colRunIdx] && row[colRunIdx].toString().trim().toUpperCase() === "TRUE");

    if (!isRun || updateType.toLowerCase() !== detailSheetName.toLowerCase()) {
      skipCount++; return; 
    }

    try {
      const targetSs = SpreadsheetApp.openById(reportId);
      const targetSheet = targetSs.getSheetByName(tabName);
      if (!targetSheet) { logStatus.push(`❌ Baris ${rowNum}: Tab '${tabName}' tidak ditemukan.`); return; }

      // =========================================================
      // PROSES PENGURUTAN KOLOM STANDAR (MOVE)
      // =========================================================
      let lastCol = targetSheet.getLastColumn();
      let headersTarget = targetSheet.getRange(headerRow, 1, 1, Math.max(lastCol, 1)).getValues()[0].map(h => h ? h.toString().trim() : "");

      for (let i = 0; i < standardCols.length; i++) {
        let stdName = standardCols[i].name;
        let stdNote = standardCols[i].note;
        let currentPointer = i + 1;
        
        let foundIdx = -1;
        // Cari kolom standar di seluruh sheet
        for (let j = 0; j < headersTarget.length; j++) {
          if (headersTarget[j].toLowerCase() === stdName.toLowerCase()) {
            foundIdx = j;
            break;
          }
        }

        if (foundIdx !== -1) {
          // Jika ditemukan di urutan salah, pindahkan
          if ((foundIdx + 1) !== currentPointer) {
            targetSheet.moveColumns(targetSheet.getRange(1, foundIdx + 1), currentPointer);
            let movedName = headersTarget.splice(foundIdx, 1)[0];
            headersTarget.splice(i, 0, movedName);
          }
          targetSheet.getRange(headerRow, currentPointer).setNote(stdNote).setBackground(null);
        } else {
          // Jika tidak ada, buat baru & beri lebar standar
          targetSheet.insertColumnBefore(currentPointer);
          targetSheet.getRange(headerRow, currentPointer).setValue(stdName).setNote(stdNote).setBackground(null);
          targetSheet.setColumnWidth(currentPointer, DEFAULT_COL_WIDTH);
          headersTarget.splice(i, 0, stdName);
        }
      }

      // =========================================================
      // PEMBERSIHAN & PENANDAAN KOLOM ASING (SETELAH NOTES / AE)
      // =========================================================
      let targetLastRow = targetSheet.getLastRow();
      let totalColsNow = targetSheet.getLastColumn();
      let standardCount = standardCols.length;

      // Loop mundur dari kolom terakhir sampai tepat setelah kolom standar terakhir
      for (let c = totalColsNow; c > standardCount; c--) {
        let isAlienEmpty = true;
        if (targetLastRow > headerRow) {
          let colData = targetSheet.getRange(headerRow + 1, c, targetLastRow - headerRow, 1).getValues();
          for (let r = 0; r < colData.length; r++) {
            if (colData[r][0] !== "" && colData[r][0] !== null) {
              isAlienEmpty = false; 
              break;
            }
          }
        }

        if (isAlienEmpty) {
          targetSheet.deleteColumn(c); // Hapus jika kosong
        } else {
          // BERI WARNA MERAH jika ada isinya tapi nama tidak sesuai standar
          targetSheet.getRange(headerRow, c).setBackground("#FF0000"); 
          // Pastikan kolom alien ini juga tidak terlalu kecil
          if (targetSheet.getColumnWidth(c) < 50) {
            targetSheet.setColumnWidth(c, DEFAULT_COL_WIDTH);
          }
        }
      }

      // =========================================================
      // SAFE UNMERGE & AUTO-APPLY FILTER
      // =========================================================
      if (targetSheet.getFilter()) targetSheet.getFilter().remove();

      let finalLastCol = targetSheet.getLastColumn();
      let finalLastRow = targetSheet.getLastRow();

      if (finalLastCol > 0) {
        let numRowsToFilter = finalLastRow >= headerRow ? (finalLastRow - headerRow + 1) : 1;
        let filterRange = targetSheet.getRange(headerRow, 1, numRowsToFilter, finalLastCol);
        filterRange.getMergedRanges().forEach(m => m.breakApart()); 
        filterRange.createFilter();
      }

      // =========================================================
      // FREEZE VIEW (Hanya Baris Header, Hapus Freeze Kolom)
      // =========================================================
      targetSheet.setFrozenRows(headerRow);
      targetSheet.setFrozenColumns(0);

      successCount++;
      logStatus.push(`✅ Baris ${rowNum}: BERHASIL (${targetSs.getName()}). Kolom asing ditandai Merah & Freeze Header aktif.`);

    } catch (e) {
      logStatus.push(`⚠️ Baris ${rowNum}: ERROR - ${e.message}`);
    }
  });

  const summary = `<b>Proses Selesai.</b><br>Berhasil: ${successCount}<br>Dilewati: ${skipCount}<br><br><b>Detail Tracing:</b><br>`;
  ui.showModalDialog(HtmlService.createHtmlOutput(
    `<div style="font-family: sans-serif; font-size: 13px;">${summary}<pre style="background: #f9f9f9; padding: 10px; border: 1px solid #ddd; height: 250px; overflow-y: auto; white-space: pre-wrap;">${logStatus.join("\n")}</pre></div>`
  ).setWidth(650).setHeight(480), 'Execution Log');
}