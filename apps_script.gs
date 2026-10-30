// Upload endpoint for the "send a page" form on https://ktavmaarav.github.io/
// Deploy: script.google.com -> New project -> paste this -> Deploy -> New deployment
//   type: Web app | Execute as: Me | Who has access: Anyone
// Then put the /exec URL into ENDPOINT in index.html.

const FOLDER_NAME = "פניות מהאתר - מפענח צפונות";
const NOTIFY = "a533155683@gmail.com";
const MAX_BYTES = 20 * 1024 * 1024;

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const contact = String(req.contact || "").slice(0, 200);
    const message = String(req.message || "").slice(0, 5000);
    if (!contact || !req.data) return reply_({ ok: false, error: "missing fields" });

    const bytes = Utilities.base64Decode(req.data);
    if (bytes.length > MAX_BYTES) return reply_({ ok: false, error: "too large" });

    const stamp = Utilities.formatDate(new Date(), "Asia/Jerusalem", "yyyy-MM-dd HH-mm");
    const name = stamp + " - " + contact.replace(/[\\\/:*?"<>|]/g, "_") + " - " + (req.fileName || "page");
    const blob = Utilities.newBlob(bytes, req.mimeType || "application/octet-stream", name);
    const file = folder_().createFile(blob);
    file.setDescription("קשר: " + contact + (message ? "\nהערה: " + message : ""));

    MailApp.sendEmail({
      to: NOTIFY,
      subject: "עמוד חדש מהאתר: " + contact,
      body: "קשר: " + contact + "\n" +
            (message ? "הערה: " + message + "\n" : "") +
            "קובץ: " + (req.fileName || "") + "\n" + file.getUrl()
    });
    return reply_({ ok: true });
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return reply_({ ok: true, service: "upload" });
}
