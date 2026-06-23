/**
 * Google Apps Script — synchronizacja tereny.geojson z Google Drive
 *
 * 1. Utwórz plik tereny.geojson na Google Drive (lub prześlij istniejący).
 * 2. Skopiuj ID pliku z URL: drive.google.com/file/d/FILE_ID/view
 * 3. Wklej FILE_ID poniżej.
 * 4. Rozszerzenia → Apps Script → wklej ten kod → Wdróż → Aplikacja internetowa
 *    (wykonuj jako: Ja, dostęp: Każdy).
 * 5. Skopiuj URL wdrożenia do CONFIG.DRIVE_SYNC_URL w admin.html
 */

var FILE_ID = '1mI6SMecuweOvA1xED0jQKwYcIwM11wgn';

function doGet() {
  var file = DriveApp.getFileById(FILE_ID);
  return ContentService
    .createTextOutput(file.getBlob().getDataAsString())
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var file = DriveApp.getFileById(FILE_ID);
  file.setContent(e.postData.contents);
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, updated: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}
