/**
 * Fungsi Update Content Detail Finding VAPT
 * Fitur: PELINDUNG HEADER, Ninja Copy (Format & Border), Smart ID, Dynamic Indexing, 
 * Filter Run, Auto-Border, Freeze Panes, Auto-Clear.
 */
function jalankanUpdateContentDetailFindingVAPT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const sourceSheetName = "Content - Detail Finding VAPT"; 
  const standardColSheetName = "Kolom - Detail Finding VAPT"; // Sumber standar kolom
  const targetUpdateText = "Content - Detail Finding VAPT"; 

  const listSheet = ss.getSheetByName("Report List");
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const standardColSheet = ss.getSheetByName(standardColSheetName);

  if (!sourceSheet || !listSheet || !standardColSheet) {
    ui.alert(`Error: Tab '${sourceSheetName}', '${standardColSheetName}', atau 'Report List' tidak ditemukan!`);
    return;
  }

  // --- 1. AMBIL STANDAR KOLOM UNTUK PELINDUNG ---
  const lastRowStd = standardColSheet.getLastRow();
  if (lastRowStd < 2) {
    ui.alert(`Tab '${standardColSheetName}' kosong.`);
    return;
  }
  // Ambil data nama kolom dari Kolom B (index 2)
  const stdNamesRaw = standardColSheet.getRange(2, 2, lastRowStd - 1, 1).getValues();
  const standardNames = stdNamesRaw.map(r => r[0] ? r[0].toString().trim() : "").filter(String);

  // 2. Cek Ketersediaan Template Content
  const lastColDetail = sourceSheet.getLastColumn();
  if (lastColDetail === 0) {
    ui.alert(`Tab '${sourceSheetName}' kosong.`);
    return;
  }

  // 3. Ambil Data Report List & Dinamisasi Kolom
  const lastRowList = listSheet.getLastRow();
  const lastColList = listSheet.getLastColumn();
  
  if (lastRowList < 2 || lastColList < 1) {
    ui.alert("Daftar di 'Report List' kosong.");
    return;
  }

  // Baca header dan ubah jadi huruf kecil semua untuk pencocokan dinamis
  const headers = listSheet.getRange(1, 1, 1, lastColList).getValues()[0].map(h => h.toString().trim().toLowerCase());
  
  let colIdIdx = -1, colTabIdx = -1, colHeaderRowIdx = -1, colUpdateIdx = -1, colRunIdx = -1;

  // Pencarian pintar hanya dengan kata kunci
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

  // AMBIL DATA & RICH TEXT (Untuk membaca link di balik Smart Chip)
  const listRange = listSheet.getRange(2, 1, lastRowList - 1, lastColList);
  const listData = listRange.getValues();
  const richTextData = listRange.getRichTextValues();

  let logStatus = [];
  let successCount = 0;
  let skipCount = 0;

  listData.forEach((row, index) => {
    const rowNum = index + 2;
    
    // =========================================================
    // FITUR: SMART ID EXTRACTOR (Aman untuk Smart Chip)
    // =========================================================
    let rawReportId = row[colIdIdx] ? row[colIdIdx].toString().trim() : ""; 
    let linkUrl = richTextData[index][colIdIdx].getLinkUrl();
    
    let reportId = linkUrl ? linkUrl : rawReportId;
    let matchId = reportId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (matchId) {
      reportId = matchId[1];
    }
    // =========================================================

    const tabName = row[colTabIdx] ? row[colTabIdx].toString().trim() : "";  
    const headerRow = (colHeaderRowIdx !== -1 && row[colHeaderRowIdx]) ? parseInt(row[colHeaderRowIdx]) || 10 : 10; 
    
    const rawUpdateValue = row[colUpdateIdx] ? row[colUpdateIdx].toString() : "";
    const updateType = rawUpdateValue.replace(/\u00A0/g, ' ').trim(); 

    // BACA KOLOM RUN
    let isRun = false;
    if (colRunIdx !== -1) {
      isRun = row[colRunIdx] === true || (row[colRunIdx] && row[colRunIdx].toString().trim().toUpperCase() === "TRUE");
    }

    // FILTER 1: Cek apakah kolom Run dicentang (TRUE)
    if (!isRun) {
      skipCount++;
      return; 
    }

    // FILTER 2: Cek Teks Kolom Update
    if (updateType.toLowerCase().indexOf(targetUpdateText.toLowerCase()) === -1) {
      skipCount++;
      return; 
    }

    if (!reportId) {
      logStatus.push(`❌ Baris ${rowNum}: ID/Link Spreadsheet kosong atau tidak valid.`);
      return;
    }

    let tempSheet = null; // Variabel untuk Tab Bayangan
    let targetSs = null;

    try {
      targetSs = SpreadsheetApp.openById(reportId);
      const targetSheet = targetSs.getSheetByName(tabName);

      if (!targetSheet) {
        logStatus.push(`❌ Baris ${rowNum}: Tab '${tabName}' tidak ditemukan.`);
        return;
      }

      // ==============================================================
      // FITUR PELINDUNG: CEK KESESUAIAN NAMA DAN URUTAN HEADER
      // ==============================================================
      let lastColTarget = targetSheet.getLastColumn();
      
      if (lastColTarget < standardNames.length) {
        logStatus.push(`⚠️ Baris ${rowNum}: DILOMPATI (${targetSs.getName()}) - Jumlah kolom kurang. Harap Update Kolom dulu!`);
        return; // Lewati baris ini
      }

      let targetHeaders = targetSheet.getRange(headerRow, 1, 1, standardNames.length).getValues()[0].map(h => h ? h.toString().trim() : "");
      let isStandard = true;

      for (let i = 0; i < standardNames.length; i++) {
        if (targetHeaders[i].toLowerCase() !== standardNames[i].toLowerCase()) {
          isStandard = false;
          break;
        }
      }

      if (!isStandard) {
        logStatus.push(`⚠️ Baris ${rowNum}: DILOMPATI (${targetSs.getName()}) - Urutan/Nama Header belum standar. Harap Update Kolom dulu!`);
        return; // Lewati baris ini (mencegah kerusakan data)
      }
      // ==============================================================

      // --- TENTUKAN RANGE TARGET ---
      const startRow = headerRow + 1; // Mulai di bawah header
      const targetLastRow = targetSheet.getLastRow();
      let endRow = startRow; // Default: 1 baris
      
      if (targetLastRow >= startRow) {
        const findingNameValues = targetSheet.getRange(startRow, 9, targetLastRow - startRow + 1, 1).getValues();
        let lastFilledRelIndex = -1;
        for (let i = findingNameValues.length - 1; i >= 0; i--) {
          if (findingNameValues[i][0].toString().trim() !== "") {
            lastFilledRelIndex = i;
            break;
          }
        }
        if (lastFilledRelIndex !== -1) endRow = startRow + lastFilledRelIndex;
      }

      // Auto-Insert baris jika sheet kurang panjang
      if (targetSheet.getMaxRows() < endRow) {
        targetSheet.insertRowsAfter(targetSheet.getMaxRows(), endRow - targetSheet.getMaxRows());
      }

      const numRows = endRow - startRow + 1;
      const targetRangeAll = targetSheet.getRange(startRow, 1, numRows, lastColDetail);
      
      // Bersihkan Validasi & Format Lama pada Content
      targetRangeAll.clearDataValidations(); 
      targetRangeAll.setBackground(null);    
      targetRangeAll.setFontColor(null);     

      // ==============================================================
      // 1. EKSEKUSI AUTO-BORDER DULUAN (Sebagai garis dasar hitam)
      // ==============================================================
      let targetMaxCol = targetSheet.getLastColumn();
      let lastHeaderCol = 0;
      
      if (targetMaxCol > 0) {
        let headerVals = targetSheet.getRange(headerRow, 1, 1, targetMaxCol).getValues()[0];
        for (let i = headerVals.length - 1; i >= 0; i--) {
          if (headerVals[i] !== null && headerVals[i].toString().trim() !== "") {
            lastHeaderCol = i + 1;
            break;
          }
        }
      }

      if (lastHeaderCol > 0 && endRow >= headerRow) {
        let borderRange = targetSheet.getRange(headerRow, 1, endRow - headerRow + 1, lastHeaderCol);
        borderRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      }

      // ==============================================================
      // 2. TRIK NINJA: COPY SHEET SEMENTARA UNTUK MEMPERTAHANKAN FORMAT
      // ==============================================================
      tempSheet = sourceSheet.copyTo(targetSs); // Copy ke file target
      tempSheet.setName("Temp_Template_" + Math.random()); // Nama acak sementara
      tempSheet.hideSheet(); // Sembunyikan agar rapi

      // --------------------------------------------------------------
      // FITUR: COPY FORMAT HEADER (Baris 1 Master -> Header Target)
      // (Dieksekusi setelah border hitam agar format khusus template menang)
      // --------------------------------------------------------------
      const sourceHeaderRange = tempSheet.getRange(1, 1, 1, lastColDetail);
      const targetHeaderRange = targetSheet.getRange(headerRow, 1, 1, lastColDetail);
      sourceHeaderRange.copyTo(targetHeaderRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

      // --------------------------------------------------------------
      // COPY FORMAT CONTENT (Baris 2 Master -> Content Rows Target)
      // --------------------------------------------------------------
      const tempRange = tempSheet.getRange(2, 1, 1, lastColDetail);
      const templateFormulas = tempRange.getFormulas()[0];

      // A. Copy Data Validation (Lokal) -> Menjaga Warna Dropdown 100%
      tempRange.copyTo(targetRangeAll, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
      
      // B. Copy Format Content (Lokal)
      tempRange.copyTo(targetRangeAll, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

      // C. Terapkan Formula (Mencegah timpa manual text)
      for (let colIdx = 0; colIdx < lastColDetail; colIdx++) {
        if (templateFormulas[colIdx] !== "") {
          let sourceFormulaCell = tempSheet.getRange(2, colIdx + 1);
          let targetColRange = targetSheet.getRange(startRow, colIdx + 1, numRows, 1);
          sourceFormulaCell.copyTo(targetColRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
        }
      }

      // ==============================================================
      // FREEZE PANES (Kolom & Baris)
      // ==============================================================
      targetSheet.setFrozenColumns(9);       // Freeze sampai Kolom I
      targetSheet.setFrozenRows(headerRow);  // Freeze baris Header ke atas

      successCount++;
      logStatus.push(`✅ Baris ${rowNum}: BERHASIL (${targetSs.getName()}) - Header & Content di-update.`);

    } catch (e) {
      logStatus.push(`⚠️ Baris ${rowNum}: ERROR - ${e.message}`);
    } finally {
      // PENTING: Selalu Hapus Tab Bayangan meski terjadi Error
      if (targetSs && tempSheet) {
        try { targetSs.deleteSheet(tempSheet); } catch(err) {} 
      }
    }
  });

  // Tampilkan Log
  const summary = `<b>Proses Content Update Selesai.</b><br>Berhasil: ${successCount}<br>Dilewati: ${skipCount}<br><br><b>Detail Tracing:</b><br>`;
  const htmlOutput = HtmlService.createHtmlOutput(
    `<div style="font-family: sans-serif; font-size: 13px;">${summary}<pre style="background:#f4f4f4;padding:10px;height:250px;overflow:auto;white-space:pre-wrap;">${logStatus.join("\n")}</pre></div>`
  ).setWidth(650).setHeight(480);
  ui.showModalDialog(htmlOutput, 'Execution Log - Content Detail Finding');
}