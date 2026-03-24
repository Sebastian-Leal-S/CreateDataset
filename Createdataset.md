/**
 * SIGA Access Mail Processor — Google Apps Script
 *
 * Searches Gmail threads with subject "Acceso a SIGA".
 * For each thread:
 *   1. Scans the first message for image attachments and saves them to Drive.
 *   2. Reads the next message in the thread and extracts the first 4-digit sequence.
 *
 * Optimizations:
 *   - Fetches exactly the last MAX_THREADS threads in a single API call.
 *   - Batch loads all messages with GmailApp.getMessagesForThreads() — one API call total.
 *   - Drive folder is resolved once and reused across all threads.
 *   - Plain-text body is preferred over HTML to avoid regex overhead on markup.
 *
 * Permissions required: Gmail, Drive.
 * Run processNewSIGAThreads() manually or via a time-based trigger.
 */

const CONFIG = {
  SUBJECT: "Acceso a SIGA",
  FOLDER_NAME: "SIGA Access Images",
  MAX_THREADS: 100,
  IMAGE_TYPES: new Set(["image/jpeg", "image/png", "image/gif", "image/webp", "image/bmp"]),
  FOUR_DIGIT_REGEX: /\d{4}/,
};

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

function processNewSIGAThreads() {
  const folder = getOrCreateFolder(CONFIG.FOLDER_NAME);
  const threads = GmailApp.search(`subject:"${CONFIG.SUBJECT}"`, 0, CONFIG.MAX_THREADS);
  const allMessages = GmailApp.getMessagesForThreads(threads);
  const results = [];

  for (let i = 0; i < threads.length; i++) {
    const messages = allMessages[i];
    if (!messages.length) continue;

    const code = messages.length > 1 ? extractFirstFourDigits(messages[1]) : null;
    const imageSaved = saveImageAttachments(messages[0], folder, code);

    results.push({
      subject: messages[0].getSubject(),
      date: messages[0].getDate(),
      imageSaved,
      code,
    });

    Logger.log(`[OK] "${messages[0].getSubject()}" | image: ${imageSaved} | code: ${code}`);
  }

  Logger.log(`Processed ${results.length} thread(s).`);
  return results;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function getOrCreateFolder(name) {
  const iter = DriveApp.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
}

function saveImageAttachments(message, folder, code) {
  const attachments = message.getAttachments({ includeInlineImages: true });
  const prefix = code ?? "NO_CODE";
  let saved = false;

  for (const att of attachments) {
    if (!CONFIG.IMAGE_TYPES.has(att.getContentType().toLowerCase())) continue;
    const filename = `${prefix}_${att.getName()}`;
    folder.createFile(att.copyBlob().setName(filename));
    saved = true;
  }

  return saved;
}

function extractFirstFourDigits(message) {
  const body = message.getPlainBody() || stripHtml(message.getBody());
  const match = body.match(CONFIG.FOUR_DIGIT_REGEX);
  return match ? match[0] : null;
}

function stripHtml(html) {
  return html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
}
