// Code.gs penuh (project Apps Script TestpaperKelas2027)
// - Borang pendaftaran kelas (JSON + secret) -> tab "Pendaftaran"
// - Borang tempahan halatuju.testpaper.com (jenis=temujanji) -> tab "Temu Janji"
// - GET ?action=slots -> bilangan tempahan setiap slot (untuk tutup slot penuh)
// Selepas edit: Deploy > Manage deployments > edit > Version: New version > Deploy
const SHEET_ID = "1FsuXEtCoxeoUb-iZJkAhoUIx9t9UKHV_pSyeDhWIBmA";
const SHEET_NAME = "Pendaftaran";
const WEBHOOK_SECRET = "testpaperKelas2027";
const TJ_SHEET = "Temu Janji";
const KAPASITI = 5; // perunding setiap slot

function doPost(e) {
  if (e && e.parameter && e.parameter.jenis === "temujanji") {
    return handleTemuJanji(e);
  }
  try {
    const data = JSON.parse(e.postData.contents || "{}");
    if (data.secret !== WEBHOOK_SECRET) {
      return response({ ok: false, error: "Unauthorized" });
    }
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    sheet.appendRow([
      new Date(),
      data.parentName || "",
      data.phone || "",
      data.studentName || "",
      data.studentLevel || "",
      data.className || "",
      data.consent ? "Ya" : "Tidak",
      data.email || "",
    ]);
    return response({ ok: true });
  } catch (error) {
    return response({ ok: false, error: error.message });
  }
}

function handleTemuJanji(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const p = e.parameter;
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(TJ_SHEET) || ss.insertSheet(TJ_SHEET);
    if (sh.getLastRow() === 0) {
      sh.appendRow(["Masa", "Nama Ibu/Bapa", "WhatsApp", "Emel", "Nama Anak", "Tahap",
        "Bil. Anak", "Sekolah Sasaran", "Khidmat", "Tarikh", "Slot Masa", "Catatan",
        "Lokasi", "utm_source", "utm_medium", "utm_campaign", "fbclid", "Kod Agent", "Status"]);
    }
    const full = (slotCounts()[(p.tarikh || "") + " " + (p.masa || "")] || 0) >= KAPASITI;
    sh.appendRow([
      new Date(), p.nama || "", "'" + (p.wa || ""), p.emel || "", p.anak || "", p.tahap || "",
      "'" + (p.pax || ""), p.sasaran || "", p.servis || "", "'" + (p.tarikh || ""), "'" + (p.masa || ""),
      p.catatan || "", p.lokasi || "", p.utm_source || "", p.utm_medium || "", p.utm_campaign || "",
      p.fbclid || "", p.agent || "", full ? "SLOT PENUH - hubungi" : "Belum disahkan",
    ]);
    return response({ ok: true, full: full });
  } catch (error) {
    return response({ ok: false, error: error.message });
  } finally {
    lock.releaseLock();
  }
}

// Kira tempahan setiap "YYYY-MM-DD HH:MM" (tak kira baris berstatus Batal)
function slotCounts() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TJ_SHEET);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 19).getDisplayValues();
  rows.forEach(function (r) {
    if (/batal/i.test(r[18])) return;
    const k = String(r[9]).replace(/^'/, "") + " " + String(r[10]).replace(/^'/, "");
    out[k] = (out[k] || 0) + 1;
  });
  return out;
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "slots") {
    return response({ ok: true, slots: slotCounts() });
  }
  return response({ ok: true, message: "Testpaper registration webhook aktif." });
}

function response(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// Uji dalam editor: pilih testTempahan > Run
function testTempahan() {
  const r = doPost({ parameter: {
    jenis: "temujanji", nama: "TEST - boleh padam", wa: "0123456789", emel: "test@test.com",
    anak: "Anak Test", tahap: "Tahun 6", pax: "1", servis: "Saringan RIASEC",
    tarikh: "2026-11-28", masa: "09:00", lokasi: "Booth Testpaper Lot No 4",
  }});
  Logger.log(r.getContent());
}
