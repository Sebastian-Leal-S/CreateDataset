/*
  processNewSIGAThreads – Google Apps Script
  - Searches Gmail in pages of 500 (GmailApp.search limit) up to MAX_THREADS total.
  - Processes each page immediately to avoid the 500-thread limit on getMessagesForThreads.
  - Valid code (first 4 chars of body) → image saved in root folder as "<CODE>.<ext>".
  - No code → saved in the NO_CODE subfolder as "NO_CODE.<ext>".
  - Code is always uppercased.
*/

const CONFIG = {
  SUBJECT: "033 - Acceso a SIGA - captcha image",
  FOLDER_NAME: "captchas",
  MAX_THREADS: 2000,
  PAGE_SIZE: 500,
  NO_CODE_FOLDER: "NO_CODE",
  IMAGE_TYPES: new Set(["image/jpeg", "image/png"]),
};

function processNewSIGAThreads() {
  const rootFolder = getOrCreateFolder(CONFIG.FOLDER_NAME);
  const noCodeFolder = getOrCreateSubfolder(rootFolder, CONFIG.NO_CODE_FOLDER);

  let start = 0;
  let totalProcessed = 0;

  while (start < CONFIG.MAX_THREADS) {
    const threads = GmailApp.search(`subject:"${CONFIG.SUBJECT}"`, start, CONFIG.PAGE_SIZE);
    if (!threads.length) break;

    const allMessages = GmailApp.getMessagesForThreads(threads);

    for (let i = 0; i < threads.length; i++) {
      const messages = allMessages[i];
      if (!messages.length) continue;

      const rawCode = messages.length > 1 ? extractCode(messages[1]) : null;
      const code = rawCode ? rawCode.toUpperCase() : null;

      const targetFolder = code ? rootFolder : noCodeFolder;
      const label = code ?? "NO_CODE";
      const imageSaved = saveImageAttachments(messages[0], targetFolder, label, !!code);

      totalProcessed++;
    }

    Logger.log(`--- Segmento procesado: threads ${start + 1} a ${start + threads.length} | acumulado: ${totalProcessed} ---`);

    start += threads.length;
    if (threads.length < CONFIG.PAGE_SIZE) break;
  }

  Logger.log(`Proceso finalizado. Total de threads procesados: ${totalProcessed}`);
}

function getOrCreateFolder(name) {
  const iter = DriveApp.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
}

function getOrCreateSubfolder(parent, name) {
  const iter = parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parent.createFolder(name);
}

function saveImageAttachments(message, folder, code, hasValidCode) {
  const attachments = message.getAttachments({ includeInlineImages: true });
  let saved = false;

  for (const att of attachments) {
    if (!CONFIG.IMAGE_TYPES.has(att.getContentType().toLowerCase())) continue;
    const ext = att.getName().split(".").pop();
    const filename = `${code}.${ext}`;
    folder.createFile(att.copyBlob().setName(filename));

    if (hasValidCode) {
      Logger.log(`[OK] Imagen guardada como ${filename} en "${CONFIG.FOLDER_NAME}"`);
    } else {
      Logger.log(`[NO_CODE] Imagen guardada como ${filename} en "${CONFIG.FOLDER_NAME}/${CONFIG.NO_CODE_FOLDER}"`);
    }

    saved = true;
  }

  return saved;
}

function extractCode(message) {
  return message.getPlainBody().substring(0, 4);
}
