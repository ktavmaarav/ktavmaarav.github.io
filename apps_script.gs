// Landing-page form receiver (Google Apps Script web app).
// Each submission: file (or a share link for big files) -> Drive folder, row -> sheet in the same folder, mail -> NOTIFY.
// Deploy: script.google.com > New project > paste > Deploy > New deployment > Web app,
//         Execute as: Me, Who has access: Anyone. Copy the /exec URL into ENDPOINT in the page.

const FOLDER_NAME = 'פניות – עמוד חינם';
const MAX_BYTES = 20 * 1024 * 1024;
const NOTIFY = 'benzimraaa@gmail.com';

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const link = String(d.link || '').slice(0, 2000);
    if ((!d.data && !link) || !d.contact) return reply_({ ok: false, error: 'missing' });
    if (link && !/^https?:\/\/\S+$/.test(link)) return reply_({ ok: false, error: 'bad_link' });
    const bytes = d.data ? Utilities.base64Decode(d.data) : null;
    if (bytes && bytes.length > MAX_BYTES) return reply_({ ok: false, error: 'too_big' });

    const folder = folder_();
    const stamp = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd HH-mm');
    const name = stamp + ' ' + String(d.fileName || 'page').slice(0, 120);
    const file = bytes && folder.createFile(Utilities.newBlob(bytes, d.mimeType || 'application/octet-stream', name));

    const contact = String(d.contact).slice(0, 200);
    const message = String(d.message || '').slice(0, 5000);
    sheet_(folder).appendRow([new Date(), contact, message, d.fileName || '', file ? file.getUrl() : '', link]);

    MailApp.sendEmail({
      to: NOTIFY,
      subject: 'פנייה חדשה מדף הנחיתה: ' + contact,
      body: 'פרטי קשר: ' + contact + '\n\nהודעה:\n' + (message || '(אין)') + (file ? '\n\nקובץ: ' + file.getUrl() : '') + (link ? '\n\nקישור שיתוף: ' + link : '')
    });
    return reply_({ ok: true });
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  }
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function sheet_(folder) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) {
    const sh = SpreadsheetApp.openById(id).getSheets()[0];
    if (!sh.getRange(1, 6).getValue()) sh.getRange(1, 6).setValue('קישור שיתוף');
    return sh;
  }
  const ss = SpreadsheetApp.create('פניות – עמוד חינם (רשימה)');
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  const sh = ss.getSheets()[0];
  sh.appendRow(['תאריך', 'טלפון / מייל', 'הודעה', 'שם הקובץ', 'קישור לקובץ', 'קישור שיתוף']);
  props.setProperty('SHEET_ID', ss.getId());
  return sh;
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
