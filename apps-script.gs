// Google Apps Script untuk borang tempahan halatuju.testpaper.com
// Pasang dalam Sheet tempahan: Extensions > Apps Script > tampal kod ini > Save
// Deploy > New deployment > Type: Web app > Execute as: Me > Who has access: Anyone > Deploy
// Salin URL Web App dan tampal pada SHEET_URL dalam index.html.
function doPost(e) {
  var p = e.parameter;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Temu Janji') || ss.insertSheet('Temu Janji');
  if (sh.getLastRow() === 0) sh.appendRow(['Masa','Nama Ibu/Bapa','WhatsApp','Emel','Nama Anak','Tahap','Bil. Anak','Sekolah Sasaran','Khidmat','Tarikh','Slot Masa','Catatan','Lokasi','utm_source','utm_medium','utm_campaign','fbclid','Kod Agent','Status']);
  sh.appendRow([new Date(), p.nama, "'" + p.wa, p.emel, p.anak, p.tahap, "'" + p.pax, p.sasaran, p.servis, "'" + p.tarikh, "'" + p.masa, p.catatan, p.lokasi, p.utm_source, p.utm_medium, p.utm_campaign, p.fbclid, p.agent, 'Belum disahkan']);
  return ContentService.createTextOutput('ok');
}
