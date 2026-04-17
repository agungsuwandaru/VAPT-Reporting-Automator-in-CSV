/**
 * Fungsi Update Content PoC VAPT
 * Fitur: SUPER DEBUGGER, Auto-Shrink Width + Extra Padding Filter, Ninja Copy Format, 
 * Alignment (Left-Top), dan Fail-Safe Header.
 */
function jalankanUpdateContentPoCVAPT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const sourceSheetName = "Content - PoC VAPT"; 
  const standardColSheetName = "Kolom - PoC VAPT"; 
  const targetUpdateText = "Content - PoC VAPT"; 

  const listSheet = ss.getSheetByName("Report List");
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const standardColSheet = ss.getSheetByName(standardColSheetName);

  if (!sourceSheet || !listSheet || !standardColSheet) {
    ui.alert(`Error: Tab '${sourceSheetName}', '${standardColSheetName}', atau 'Report List' tidak ditemukan!`);
    return;
  }

  // Helper untuk membersihkan spasi bandel (Non-breaking space)
  const cleanStr = (str) => str.toString().replace(/\u00A0/g, ' ').trim().toLowerCase();

  // 1. Ambil Standar Kolom (Kolom B Master)
  const lastRowStd = standardColSheet.getLastRow();
  const stdNamesRaw = standardColSheet.getRange(2, 2, lastRowStd - 1, 1).getValues();
  const standardNames = stdNamesRaw.map(r => r[0] ? r[0].toString().trim() : "").filter(String);

  // 2. Ambil Data Report List
  const lastRowList = listSheet.getLastRow();
  const lastColList = listSheet.getLastColumn();
  if (lastRowList < 2) return;

  const headers = listSheet.getRange(1, 1, 1, lastColList).getValues()[0].map(h => cleanStr(h));
  let colIdIdx = -1, colTabIdx = -1, colHeaderRowIdx = -1, colUpdateIdx = -1, colRunIdx = -1;

  headers.forEach((h, i) => {
    if (h.includes("id") || h.includes("link")) colIdIdx = i; 
    if (h.includes("tab")) colTabIdx = i;                     
    if (h.includes("header row")) colHeaderRowIdx = i;        
    if (h.includes("update")) colUpdateIdx = i;               
    if (h.includes("run")) colRunIdx = i;                     
  });

  const listRange = listSheet.getRange(2, 1, lastRowList - 1, lastColList);
  const listData = listRange.getValues();
  const richTextData = listRange.getRichTextValues();

  let logStatus = [];
  let successCount = 0;
  let skipCount = 0;

  listData.forEach((row, index) => {
    const rowNum = index + 2;
    const isRun = row[colRunIdx] === true || (row[colRunIdx] && row[colRunIdx].toString().toUpperCase() === "TRUE");
    const rawUpdateValue = row[colUpdateIdx] ? row[colUpdateIdx].toString().trim() : "";

    if (!isRun || rawUpdateValue !== targetUpdateText) {
      skipCount++; return; 
    }

    // SMART ID EXTRACTOR
    let rawReportId = row[colIdIdx] ? row[colIdIdx].toString().trim() : ""; 
    let linkUrl = richTextData[index][colIdIdx].getLinkUrl();
    let reportId = linkUrl ? linkUrl : rawReportId;
    let matchId = reportId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (matchId) reportId = matchId[1];

    const tabName = row[colTabIdx] ? row[colTabIdx].toString().trim() : "";  
    const headerRow = (colHeaderRowIdx !== -1 && row[colHeaderRowIdx]) ? parseInt(row[colHeaderRowIdx]) || 1 : 1; 

    let tempSheet = null;
    let targetSs = null;

    try {
      targetSs = SpreadsheetApp.openById(reportId);
      const targetSheet = targetSs.getSheetByName(tabName);
      if (!targetSheet) { logStatus.push(`❌ Baris ${rowNum}: Tab '${tabName}' tidak ditemukan.`); return; }

      // --- SUPER DEBUGGER: PROTEKSI & TRACING HEADER ---
      let lastColTarget = targetSheet.getLastColumn();
      if (lastColTarget < standardNames.length) {
        logStatus.push(`⚠️ Baris ${rowNum}: LOMPAT (${targetSs.getName()}) - Jumlah kolom kurang.`);
        return;
      }

      let targetHeadersValues = targetSheet.getRange(headerRow, 1, 1, standardNames.length).getValues()[0];
      let mismatches = [];

      for (let i = 0; i < standardNames.length; i++) {
        let rawTarget = targetHeadersValues[i] || "";
        if (cleanStr(rawTarget) !== cleanStr(standardNames[i])) {
          mismatches.push(`Kolom ${i+1} -> Harusnya: "${standardNames[i]}" | Terbaca: "${rawTarget}"`);
        }
      }

      if (mismatches.length > 0) {
        logStatus.push(`⚠️ Baris ${rowNum}: LOMPAT (${targetSs.getName()}) - Header tidak standar:\n    ❌ ` + mismatches.join(`\n    ❌ `));
        return;
      }

      // --- LANJUT PROSES UPDATE CONTENT ---
      const startRow = headerRow + 1;
      const targetLastRow = targetSheet.getLastRow();
      let endRow = startRow;
      
      if (targetLastRow >= startRow) {
        const checkRange = targetSheet.getRange(startRow, 1, targetLastRow - startRow + 1, lastColTarget).getValues();
        for (let i = checkRange.length - 1; i >= 0; i--) {
          if (checkRange[i].join("").trim() !== "") {
            endRow = startRow + i;
            break;
          }
        }
      }

      const numRows = endRow - startRow + 1;
      const lastColSource = sourceSheet.getLastColumn();
      const targetRangeAll = targetSheet.getRange(startRow, 1, numRows, lastColSource);
      
      targetRangeAll.clearDataValidations().setBackground(null).setFontColor(null);

      tempSheet = sourceSheet.copyTo(targetSs);
      tempSheet.setName("Temp_PoC_" + Math.random());
      tempSheet.hideSheet();

      // Mirror Format Header & Content
      tempSheet.getRange(1, 1, 1, lastColSource).copyTo(targetSheet.getRange(headerRow, 1, 1, lastColSource), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      
      const tempRange = tempSheet.getRange(2, 1, 1, lastColSource);
      tempRange.copyTo(targetRangeAll, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
      tempRange.copyTo(targetRangeAll, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

      // Formulas
      const templateFormulas = tempRange.getFormulas()[0];
      for (let c = 0; c < lastColSource; c++) {
        if (templateFormulas[c] !== "") {
          tempSheet.getRange(2, c + 1).copyTo(targetSheet.getRange(startRow, c + 1, numRows, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
        }
      }

      targetRangeAll.setHorizontalAlignment("left").setVerticalAlignment("top");
      targetSheet.getRange(headerRow, 1, 1, lastColSource).setHorizontalAlignment("left").setVerticalAlignment("top");

      // ==============================================================
      // FITUR BARU: DYNAMIC WIDTH DENGAN EXTRA PADDING
      // ==============================================================
      let standardHeadersLower = standardNames.map(n => n.toLowerCase());
      standardHeadersLower.forEach((hName, i) => {
        let colIdx = i + 1;
        if (hName.includes("description") || hName.includes("evidence")) {
          let hasContent = false;
          if (targetLastRow >= startRow) {
            let colValues = targetSheet.getRange(startRow, colIdx, numRows, 1).getValues();
            hasContent = colValues.some(v => v[0] && v[0].toString().trim() !== "");
          }

          if (hasContent) {
            targetSheet.setColumnWidth(colIdx, 450);
          } else {
            // Trik Ninja: Resize otomatis sesuai teks, LALU tambah 40 pixel untuk tombol Filter
            targetSheet.autoResizeColumn(colIdx);
            let autoWidth = targetSheet.getColumnWidth(colIdx);
            targetSheet.setColumnWidth(colIdx, autoWidth + 40);
          }
        }
      });

      // Border & Freeze
      targetSheet.getRange(headerRow, 1, numRows + 1, lastColSource).setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      targetSheet.setFrozenRows(headerRow);
      targetSheet.setFrozenColumns(4); 

      successCount++;
      logStatus.push(`✅ Baris ${rowNum}: BERHASIL (${targetSs.getName()})`);

    } catch (e) {
      logStatus.push(`⚠️ Baris ${rowNum}: ERROR - ${e.message}`);
    } finally {
      if (targetSs && tempSheet) {
        try { targetSs.deleteSheet(tempSheet); } catch(err) {} 
      }
    }
  });

  const summary = `<b>Update Content PoC Selesai.</b><br>Berhasil: ${successCount}<br>Dilewati: ${skipCount}<br><br><b>Detail Tracing:</b><br>`;
  ui.showModalDialog(HtmlService.createHtmlOutput(
    `<div style="font-family:sans-serif;font-size:13px;">${summary}<pre style="background:#f4f4f4;padding:10px;height:250px;overflow:auto;white-space:pre-wrap;">${logStatus.join("\n")}</pre></div>`
  ).setWidth(650).setHeight(480), 'Execution Log - Content PoC');
}