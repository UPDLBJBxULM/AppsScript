function GenPDFHAR() {
  var startTime = Date.now(); // Menyimpan waktu mulai eksekusi
  
  // Menampilkan alert bahwa proses dimulai
  SpreadsheetApp.getUi().alert('Sedang diproses!');
  
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Mendapatkan spreadsheet aktif
  var sheet = ss.getActiveSheet(); // Mendapatkan sheet yang aktif
  
  // Mengakses sheet HAR untuk mengambil data E6
  var generateSheet = ss.getSheetByName('HAR'); // Ganti dengan nama sheet yang sesuai
  
  // Cek apakah sheet ditemukan
  if (!generateSheet) {
    Logger.log('Sheet HAR tidak ditemukan.');
    SpreadsheetApp.getUi().alert('Sheet HAR tidak ditemukan.');
    return; // Hentikan eksekusi script jika sheet tidak ada
  }

  // Mendapatkan data periode dari E6
  var periodeData = generateSheet.getRange('E6').getValue();
  
  // Log data untuk melihat nilai sebelum perubahan
  Logger.log('Data E6: ' + periodeData);
  
  // Ekstrak tahun dan bulan dari periodeData
  var monthYearMatch = periodeData.match(/(\d{4}) (\w+)/);
  if (monthYearMatch) {
    var year = monthYearMatch[1]; // Ambil tahun
    var monthName = monthYearMatch[2]; // Ambil nama bulan
    var month = new Date(Date.parse(monthName + " 1, 2021")).getMonth() + 1; // Mengonversi nama bulan ke angka
    month = month.toString().padStart(2, '0'); // Menambahkan nol di depan jika bulan < 10
  } else {
    Logger.log('Format data E6 tidak sesuai.');
    SpreadsheetApp.getUi().alert('Format data E6 tidak sesuai.');
    return; // Hentikan eksekusi jika format tidak sesuai
  }
  
  // Format nama PDF sesuai dengan data yang ditemukan
  var pdfName = 'HAR_' + year + month + '_Infrastruktur dan Jaringan' + '.pdf';
  
  Logger.log('Nama file PDF: ' + pdfName);
  
  Logger.log('Mulai menyembunyikan baris 1 hingga 2');
  // Menyembunyikan baris 1 hingga 2
  sheet.hideRows(1, 2);
  
  Logger.log('Mulai menyembunyikan gambar dan shape');
  // Menyembunyikan semua gambar dan shape
  var shapes = sheet.getDrawings();
  shapes.forEach(function(shape) {
    shape.remove(); // Menghapus shape
  });

  Logger.log('Mendapatkan folder untuk menyimpan PDF');
  var folderId= '1_vVHzvtTJeNL_48sbFpP2PV7gZ9G0FXq';
  var folder = DriveApp.getFolderById(folderId);
  
  Logger.log('Mendapatkan URL untuk mengkonversi sheet ke PDF');
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?';
  
  // Menentukan parameter ekspor ke PDF
  var params = {
    format: 'pdf', // Format file PDF
    exportFormat: 'pdf', // Ekspor format PDF
    gid: sheet.getSheetId(), // ID sheet
    size: 'A4', // Ukuran kertas A4
    portrait: false, // Orientasi landscape
    fitw: true, // Menyesuaikan lebar
    fitt: true, // Menyesuaikan tinggi
    top_margin: 0.5, // Margin atas
    bottom_margin: 0.5, // Margin bawah
    left_margin: 0.5, // Margin kiri
    right_margin: 0.5, // Margin kanan
    sheetnames: false, // Tidak menampilkan nama sheet
    printtitle: false, // Tidak menampilkan judul
  };
  
  Logger.log('Mengunduh PDF dari URL');
  // Membuat URL dengan parameter untuk mendownload PDF
  var response = UrlFetchApp.fetch(url + Object.keys(params).map(function(key) {
    return key + '=' + encodeURIComponent(params[key]);
  }).join('&'));
  
  Logger.log('Mendapatkan file PDF sebagai blob');
  // Mendapatkan file PDF sebagai blob
  var pdfBlob = response.getBlob().setName(pdfName);
  
  Logger.log('Menyimpan PDF ke folder Google Drive');
  // Menyimpan PDF ke folder yang telah ditentukan
  var file = folder.createFile(pdfBlob);
  
  Logger.log('Menampilkan kembali baris yang disembunyikan');
  // Menampilkan kembali baris yang disembunyikan
  sheet.showRows(1, 2);
  
  Logger.log('Menambahkan link ke sel B1');
  // Menambahkan link ke sel B1
  sheet.getRange('B1').setValue(file.getUrl()); // Menyisipkan URL file di B1
  
  Logger.log('Proses selesai, file PDF berhasil disimpan');
  // Menampilkan pesan saat proses selesai
  SpreadsheetApp.getUi().alert('Proses selesai!');
  
  // Menghitung durasi eksekusi
  var endTime = Date.now(); // Menyimpan waktu akhir eksekusi
  var duration = (endTime - startTime) / 1000; // Durasi dalam detik
  
  // Menampilkan log durasi eksekusi
  Logger.log('Proses selesai dalam ' + duration + ' detik.');
  // Menyimpan log akhir
  Logger.log('File PDF berhasil disimpan di folder Google Drive: ' + pdfName);
}