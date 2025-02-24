function onFormSubmit(e) {
  if (!e) {
    Logger.log("Tidak ada objek event.");
    return;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RENCANA");
  if (!sheet) {
    Logger.log("Sheet 'RENCANA' tidak ditemukan.");
    return;
  }
  
  // Dapatkan baris terakhir yang baru ditambahkan
  var lastRow = sheet.getLastRow();
  
  // Tandai kolom J (status) pada baris terbaru dengan "PENDING"
  sheet.getRange(lastRow, 11).setValue("PENDING"); // Kolom K = 11
  
  Logger.log("Baris baru ditandai sebagai PENDING di sheet RENCANA.");
  
  // Validasi untuk membaca worksheet 'Publikasi' pada sel B3
  var publikasiSheet = ss.getSheetByName("Publikasi");
  if (!publikasiSheet) {
    Logger.log("Sheet 'Publikasi' tidak ditemukan.");
    return;
  }
  
  // Baca nilai dari sel B3 di sheet 'Publikasi'
  var statusPublikasi = publikasiSheet.getRange("B3").getValue();
  
  // Periksa apakah status adalah "Aktif" atau "Tidak Aktif"
  if (statusPublikasi === "Tidak Aktif") {
    Logger.log("Status Publikasi: Tidak Aktif. Proses queue dibatalkan.");
    return; // Hentikan eksekusi jika status adalah "Tidak Aktif"
  } else if (statusPublikasi === "Aktif") {
    Logger.log("Status Publikasi: Aktif. Melanjutkan ke processQueue.");
    // Panggil processQueue untuk memulai pemrosesan
    processQueue();
  } else {
    Logger.log("Status Publikasi tidak dikenali: " + statusPublikasi);
    return; // Hentikan eksekusi jika status tidak dikenali
  }
}