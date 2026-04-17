/**
 * Fungsi khusus Google Sheets yang otomatis jalan saat file dibuka
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Membuat Menu Utama "⚡ VAPT TOOLS"
  ui.createMenu('⚡ VAPT TOOLS')
      .addItem('Update Header - Detail Finding VAPT', 'jalankanUpdateKolomDetailFindingVAPT')
      .addItem('Update Header - PoC VAPT', "jalankanUpdateKolomPoCVAPT")
      .addItem('Update Content - Detail Finding VAPT',"jalankanUpdateContentDetailFindingVAPT")
      .addItem('Update Content - PoC VAPT',"jalankanUpdateContentPoCVAPT")
      .addItem('Update Content - Helper VAPT',"jalankanUpdateContentHelperVAPT")
      .addItem('Update Header - Dashboard BGN',"jalankanUpdateKolomDashboardBGN")
      // Jika nanti ada fungsi lain, tinggal tambah addItem lagi di bawahnya:
      // .addSeparator() // Garis pembatas
      // .addItem('Sync Data Vulnerability', 'namaFungsiLain')
      .addToUi();
}
