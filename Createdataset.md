/**
 * SIGA Access Mail Processor — Google Apps Script
 *
 * Searches Gmail threads with subject "Acceso a SIGA".
 * For each thread:
 *   1. Scans the first message for image attachments and saves them to Drive.
 *   2. Reads the next message in the thread and extracts the first 4-digit sequence.
 *
 * Optimizations:
 *   - Batch processes threads with GmailApp.getMessagesForThreads() to minimize API calls.
 *   - Skips threads already processed using a persistent ScriptProperties cache.
 *   - Uses pagination to safely handle large inboxes (> 500 threads).
 *   - Drive folder is resolved once and reused across all threads.
 *   - Plain-text body is preferred over HTML to avoid regex overhead on markup.
 *
 * Permissions required: Gmail, Drive, Script Properties.
 * Run processNewSIGAThreads() manually or via a time-based trigger.
 */

const CONFIG = {
  SUBJECT: "Acceso a SIGA",
  FOLDER_NAME: "SIGA Access Images",
  BATCH_SIZE: 50,
  IMAGE_TYPES: new Set(["image/jpeg", "image/png", "image/gif", "image/webp", "image/bmp"]),
  FOUR_DIGIT_REGEX: /\d{4}/,
};

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

function processNewSIGAThreads() {
  const cache = PropertiesService.getScriptProperties();
  const folder = getOrCreateFolder(CONFIG.FOLDER_NAME);
  const results = [];
  let start = 0;

  while (true) {
    const threads = GmailApp.search(`subject:"${CONFIG.SUBJECT}"`, start, CONFIG.BATCH_SIZE);
    if (!threads.length) break;

    const allMessages = GmailApp.getMessagesForThreads(threads);

    for (let i = 0; i < threads.length; i++) {
      const threadId = threads[i].getId();
      if (cache.getProperty(threadId)) continue;

      const messages = allMessages[i];
      if (!messages.length) continue;

      const imageSaved = saveImageAttachments(messages[0], folder);
      const code = messages.length > 1 ? extractFirstFourDigits(messages[1]) : null;

      cache.setProperty(threadId, "done");

      results.push({
        subject: messages[0].getSubject(),
        date: messages[0].getDate(),
        imageSaved,
        code,
      });

      Logger.log(`[OK] "${messages[0].getSubject()}" | image: ${imageSaved} | code: ${code}`);
    }

    if (threads.length < CONFIG.BATCH_SIZE) break;
    start += CONFIG.BATCH_SIZE;
  }

  Logger.log(`Processed ${results.length} new thread(s).`);
  return results;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function getOrCreateFolder(name) {
  const iter = DriveApp.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
}

function saveImageAttachments(message, folder) {
  const attachments = message.getAttachments({ includeInlineImages: true });
  let saved = false;

  for (const att of attachments) {
    if (!CONFIG.IMAGE_TYPES.has(att.getContentType().toLowerCase())) continue;
    const filename = `${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss")}_${att.getName()}`;
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

// ---------------------------------------------------------------------------
// Utility: reset processed thread cache (run manually if needed)
// ---------------------------------------------------------------------------

function resetCache() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log("Cache cleared.");
}