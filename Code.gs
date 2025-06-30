function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const data = JSON.parse(e.postData.contents);
  const rows = sheet.getDataRange().getValues();
  const now = new Date();

  // Semak jika kod sudah digunakan
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().toUpperCase() === data.kod.toUpperCase()) {
      if (rows[i][5].toString().toLowerCase() === 'digunakan') {
        return ContentService.createTextOutput(JSON.stringify({ mesej: "Kod telah digunakan!" }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      // Simpan data pengguna ke baris tersebut
      sheet.getRange(i + 1, 2).setValue(data.nama);
      sheet.getRange(i + 1, 3).setValue(data.telefon);
      sheet.getRange(i + 1, 4).setValue(data.organisasi);
      sheet.getRange(i + 1, 5).setValue(now);
      sheet.getRange(i + 1, 6).setValue("Digunakan");

      return ContentService.createTextOutput(JSON.stringify({ mesej: "Pendaftaran berjaya!" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ mesej: "Kod tidak sah atau belum dimasukkan oleh admin." }))
    .setMimeType(ContentService.MimeType.JSON);
}