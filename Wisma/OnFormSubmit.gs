function onFormSubmit(e) {

  if (!e || !e.range) return Logger.log('e tidak ada');

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  Logger.log("Nama Sheet: " + sheetName);

  if (sheetName == 'ROOMPLOT') {
    onFormSubmit_roomPlot(e);
  } else if (sheetName == 'GUESTCI') {
    onFormSubmit_checkIn(e);
  } else if (sheetName == 'GUESTCO') {
    onFormSubmit_checkOut(e);
  } else {
    Logger.log("Form tidak dikenali.");
  };
}