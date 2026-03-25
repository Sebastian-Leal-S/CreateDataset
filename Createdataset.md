/*
  processNewSIGAThreads – Google Apps Script
  - Searches Gmail in pages of 500 (GmailApp.search limit) up to MAX_THREADS total.
  - First message of each thread: saves image attachments to Drive.
  - Second message (if present): extracts the first 4-digit code from plain body.
  - Valid code  → image saved as  "<CODE>_<filename>"  in the root folder.
  - No code or non-4-digit code → saved in the NO_CODE subfolder.
  - Code is always uppercased.
*/

const CONFIG = {
  SUBJECT: "Acceso a SIGA",
  FOLDER_NAME: "SIGA Access Images",
  MAX_THREADS: 2000,
  PAGE_SIZE: 500,
  NO_CODE_FOLDER: "NO_CODE",
  IMAGE_TYPES: new Set(["image/jpeg", "image/png"]),
  FOUR_DIGIT_REGEX: /^\d{4}$/,
};

function processNewSIGAThreads() {
  const rootFolder = getOrCreateFolder(CONFIG.FOLDER_NAME);
  const noCodeFolder = getOrCreateSubfolder(rootFolder, CONFIG.NO_CODE_FOLDER);

  const threads = fetchAllThreads();
  const allMessages = GmailApp.getMessagesForThreads(threads);
  const results = [];

  for (let i = 0; i < threads.length; i++) {
    const messages = allMessages[i];
    if (!messages.length) continue;

    const rawCode = messages.length > 1 ? extractFourDigitCode(messages[1]) : null;
    const code = rawCode ? rawCode.toUpperCase() : null;
    const isValidCode = code && CONFIG.FOUR_DIGIT_REGEX.test(code);

    const targetFolder = isValidCode ? rootFolder : noCodeFolder;
    const prefix = isValidCode ? code : "NO_CODE";
    const imageSaved = saveImageAttachments(messages[0], targetFolder, prefix);

    results.push({
      subject: messages[0].getSubject(),
      date: messages[0].getDate(),
      imageSaved,
      code,
    });

    Logger.log(`[OK] "${messages[0].getSubject()}" | image: ${imageSaved} | code: ${code} | valid: ${isValidCode}`);
  }

  Logger.log(`Processed ${results.length} thread(s).`);
  return results;
}

function fetchAllThreads() {
  const threads = [];
  let start = 0;

  while (start < CONFIG.MAX_THREADS) {
    const page = GmailApp.search(`subject:"${CONFIG.SUBJECT}"`, start, CONFIG.PAGE_SIZE);
    if (!page.length) break;
    threads.push(...page);
    start += page.length;
    if (page.length < CONFIG.PAGE_SIZE) break;
  }

  return threads;
}

function getOrCreateFolder(name) {
  const iter = DriveApp.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
}

function getOrCreateSubfolder(parent, name) {
  const iter = parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parent.createFolder(name);
}

function saveImageAttachments(message, folder, prefix) {
  const attachments = message.getAttachments({ includeInlineImages: true });
  let saved = false;

  for (const att of attachments) {
    if (!CONFIG.IMAGE_TYPES.has(att.getContentType().toLowerCase())) continue;
    const filename = `${prefix}_${att.getName()}`;
    folder.createFile(att.copyBlob().setName(filename));
    saved = true;
  }

  return saved;
}

function extractFourDigitCode(message) {
  const match = message.getPlainBody().match(/\d{4}/);
  return match ? match[0] : null;
}
