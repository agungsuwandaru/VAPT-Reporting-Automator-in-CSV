/**
 * Fungsi Standarisasi Kolom - Khusus PoC VAPT
 * Fitur: Ninja Format Mirroring (Dari "Content - PoC VAPT" Baris 1), Smart ID, Filter Run, 
 * Auto-Border (Tanpa Freeze Panes).
 */
function jalankanUpdateKolomPoCVAPT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const detailSheetName = "Kolom - PoC VAPT";   // Tab sumber untuk daftar nama & note
  const formatSheetName = "Content - PoC VAPT"; // Tab sumber untuk gaya format header
  const targetUpdateText = "Kolom - PoC VAPT"; 

  const listSheet = ss.getSheetByName("Report List");
  const detailSheet = ss.getSheetByName(detailSheetName);
  const formatSheet = ss.getSheetByName(formatSheetName);

  if (!detailSheet || !listSheet || !formatSheet) {
    ui.alert(`Error: Pastikan tab '${detailSheetName}', '${formatSheetName}', dan 'Report List' ada!`);
    return;
  }

  // 1. Ambil Standar Kolom dari Master (List Vertikal)
  const lastRowDetail = detailSheet.getLastRow();
  if (lastRowDetail < 2) return;
  
  const standardConfig = detailSheet.getRange(2, 1, lastRowDetail - 1, 3).getValues();
  let standardNames = [];
  let standardNotes = [];
  
  standardConfig.forEach((config) => {
    let name = config[1] ? config[1].toString().trim() : "";
    let note = config[2] ? config[2].toString().trim() : "";
    if (name !== "") {
      standardNames.push(name);
      standardNotes.push(note);
    }
  });

  // 2. Ambil List Report & Dinamisasi Kolom
  const lastRowList = listSheet.getLastRow();
  const lastColList = listSheet.getLastColumn();
  if (lastRowList < 2) return;

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
    ui.alert("Error: Kolom ID/Link, Tab, atau Update tidak ditemukan di baris pertama 'Report List'.");
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
    
    // SMART ID EXTRACTOR (Support Smart Chips)
    let rawReportId = row[colIdIdx] ? row[colIdIdx].toString().trim() : ""; 
    let linkUrl = richTextData[index][colIdIdx].getLinkUrl();
    let reportId = linkUrl ? linkUrl : rawReportId;
    let matchId = reportId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (matchId) reportId = matchId[1];

    const tabName = row[colTabIdx] ? row[colTabIdx].toString().trim() : "";  
    const headerRow = (colHeaderRowIdx !== -1 && row[colHeaderRowIdx]) ? parseInt(row[colHeaderRowIdx]) || 1 : 1; 
    const updateType = row[colUpdateIdx] ? row[colUpdateIdx].toString().trim() : ""; 
    const isRun = row[colRunIdx] === true || (row[colRunIdx] && row[colRunIdx].toString().toUpperCase() === "TRUE");

    if (!isRun || updateType.toLowerCase() !== targetUpdateText.toLowerCase()) {
      skipCount++; return; 
    }

    if (!tabName) {
      logStatus.push(`❌ Baris ${rowNum}: Gagal. Nama Tab target kosong.`);
      return;
    }

    let tempSheet = null;
    let targetSs = null;

    try {
      targetSs = SpreadsheetApp.openById(reportId);
      const targetSheet = targetSs.getSheetByName(tabName);
      if (!targetSheet) { logStatus.push(`❌ Baris ${rowNum}: Gagal. Tab '${tabName}' tidak ada.`); return; }

      // --- LOGIKA PENENTUAN POSISI KOLOM ---
      let lastCol = targetSheet.getLastColumn();
      let existingHeaders = [];
      if (lastCol > 0) {
        existingHeaders = targetSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => h ? h.toString().trim() : "");
      }

      let isMatchAtStart = existingHeaders.length >= standardNames.length && standardNames.every((val, i) => val === existingHeaders[i]);
      let isMatchAtEnd = false;
      if (!isMatchAtStart && existingHeaders.length >= standardNames.length) {
        let offset = existingHeaders.length - standardNames.length;
        isMatchAtEnd = standardNames.every((val, i) => val === existingHeaders[offset + i]);
      }

      let targetStartCol = 1;
      if (isMatchAtStart) {
        targetStartCol = 1;
      } else if (isMatchAtEnd) {
        targetStartCol = existingHeaders.length - standardNames.length + 1;
      } else {
        if (lastCol > 0) targetSheet.getRange(headerRow, 1, 1, lastCol).setBackground("#FF0000"); 
        targetStartCol = lastCol + 1;
        for (let i = 0; i < standardNames.length; i++) {
          targetSheet.getRange(headerRow, targetStartCol + i).setValue(standardNames[i]);
        }
      }

      // ==============================================================
      // AUTO-BORDER SELURUH AREA DATA DULUAN 
      // ==============================================================
      let fRow = targetSheet.getLastRow();
      if (fRow < headerRow) fRow = headerRow;
      let maxCol = Math.max(lastCol, targetStartCol + standardNames.length - 1);
      
      if (maxCol > 0) {
        targetSheet.getRange(headerRow, 1, fRow - headerRow + 1, maxCol)
                   .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      }

      // ==============================================================
      // TRIK NINJA: COPY TAB "Content - PoC VAPT" KE TARGET
      // ==============================================================
      // Kali ini kita salin formatSheet ("Content - PoC VAPT") agar dapat header indahnya!
      tempSheet = formatSheet.copyTo(targetSs);
      tempSheet.setName("Temp_Ninja_" + Math.random());
      tempSheet.hideSheet();

      // Ambil header dari tab format (Content - PoC VAPT) baris 1 untuk pencocokan nama
      let templateLastCol = tempSheet.getLastColumn();
      let templateHeaders = [];
      if (templateLastCol > 0) {
        templateHeaders = tempSheet.getRange(1, 1, 1, templateLastCol).getValues()[0].map(h => h ? h.toString().trim().toLowerCase() : "");
      }

      for (let i = 0; i < standardNames.length; i++) {
        let targetCol = targetStartCol + i;
        let targetCell = targetSheet.getRange(headerRow, targetCol);
        
        let headerNameLower = standardNames[i].toLowerCase();
        let sourceColIdx = templateHeaders.indexOf(headerNameLower) + 1;
        
        if (sourceColIdx > 0) {
          // JIKA KETEMU DI BARIS 1: Copy format dan lebar kolom dari cell yang tepat
          let sourceCell = tempSheet.getRange(1, sourceColIdx);
          sourceCell.copyTo(targetCell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
          targetSheet.setColumnWidth(targetCol, tempSheet.getColumnWidth(sourceColIdx));
        } else {
          // JIKA TIDAK KETEMU (Nama Beda Sedikit): Asumsikan posisinya urut dari awal (A, B, C...)
          let fallbackIdx = i + 1;
          let sourceCell = tempSheet.getRange(1, fallbackIdx);
          sourceCell.copyTo(targetCell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
          targetSheet.setColumnWidth(targetCol, tempSheet.getColumnWidth(fallbackIdx));
        }
        
        // Copy Note (Komentar) dari Kolom - PoC VAPT
        targetCell.setNote(standardNotes[i]);
      }

      successCount++;
      logStatus.push(`✅ Baris ${rowNum}: BERHASIL (${targetSs.getName()}). Format Ninja diterapkan.`);

    } catch (e) {
      logStatus.push(`⚠️ Baris ${rowNum}: ERROR - ${e.message}`);
    } finally {
      // PENTING: Hapus Tab Bayangan
      if (targetSs && tempSheet) {
        try { targetSs.deleteSheet(tempSheet); } catch(err) {} 
      }
    }
  });

  // Tampilkan Log
  const summary = `<b>Update PoC Selesai.</b><br>Berhasil: ${successCount}<br>Dilewati: ${skipCount}<br><br><b>Detail:</b><br>`;
  ui.showModalDialog(HtmlService.createHtmlOutput(
    `<div style="font-family:sans-serif;font-size:12px;">${summary}<pre style="background:#f4f4f4;padding:10px;height:250px;overflow:auto;">${logStatus.join("\n")}</pre></div>`
  ).setWidth(650).setHeight(480), 'Execution Log');
}