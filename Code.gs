/**
 * BuzzBlu Monthly Football Tournament 2026
 * Four-file Apps Script project: Code.gs + Index.html + Style.html + Script.html.
 *
 * One-time setup:
 * 1) Set SPREADSHEET_ID in Project Settings > Script Properties.
 * 2) Optionally set ORGANIZER_EMAILS to comma-separated authorized emails.
 * 3) Run setupTournament() and authorize Sheets/Drive access.
 */
const APP_CONFIG = Object.freeze({
  SPREADSHEET_PROPERTY: 'SPREADSHEET_ID',
  ORGANIZERS_PROPERTY: 'ORGANIZER_EMAILS',
  PHOTO_FOLDER_NAME: 'BuzzBlu Football Tournament 2026 - Player Photos',
  SHEET_NAME: 'Player Registrations',
  MAX_PHOTO_BYTES: 10 * 1024 * 1024,
  ALLOWED_MIME_TYPES: ['image/jpeg', 'image/png', 'image/webp'],
  HEADERS: [
    'Timestamp', 'Registration ID', 'Player Name', 'Playing Position',
    'Jersey Number', 'Jersey Size', 'Project Name', 'Photo Drive File ID',
    'Registration Status'
  ]
});

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.pageMode = (e && e.parameter && e.parameter.view === 'admin') ? 'admin' : 'register';
  return template.evaluate()
    .setTitle('BuzzBlu Monthly Football Tournament 2026 | Player Registration')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Includes Style.html or Script.html in the page template. */
function include(filename) {
  if (filename !== 'Style' && filename !== 'Script') {
    throw new Error('Invalid include requested.');
  }
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Run once from the Apps Script editor after setting SPREADSHEET_ID. */
function setupTournament() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty(APP_CONFIG.SPREADSHEET_PROPERTY);
  if (!spreadsheetId) {
    throw new Error('Set SPREADSHEET_ID in Project Settings > Script Properties first.');
  }
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(APP_CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(APP_CONFIG.SHEET_NAME);

  // Preserve all existing rows. Add headers only when the sheet is empty.
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, APP_CONFIG.HEADERS.length).setValues([APP_CONFIG.HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, APP_CONFIG.HEADERS.length)
      .setBackground('#0d1a29').setFontColor('#ffffff').setFontWeight('bold');
    sheet.autoResizeColumns(1, APP_CONFIG.HEADERS.length);
  } else {
    const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), APP_CONFIG.HEADERS.length))
      .getDisplayValues()[0];
    const missing = APP_CONFIG.HEADERS.filter(h => current.indexOf(h) === -1);
    if (missing.length) {
      throw new Error('The existing sheet has different headers. No data was changed. Missing headers: ' + missing.join(', ') + '. Use a new empty sheet or reconcile headers manually.');
    }
  }

  const folder = getOrCreatePhotoFolder_();
  return {
    success: true,
    spreadsheetName: ss.getName(),
    sheetName: APP_CONFIG.SHEET_NAME,
    photoFolderName: folder.getName(),
    message: 'Setup complete. Keep the spreadsheet and photo folder private.'
  };
}

/** Public registration endpoint. Returns success only after both Drive and Sheets writes succeed. */
function submitRegistration(formData) {
  const lock = LockService.getScriptLock();
  let photoFile = null;
  let appendedRow = null;
  let sheet = null;
  try {
    validateRegistration_(formData);
    lock.waitLock(25000);

    sheet = getRegistrationSheet_();
    const jerseyNumber = Number(formData.jerseyNumber);
    // Warn about duplicates to organizers in their sheet; duplicate enforcement is configurable.
    const enforceUnique = PropertiesService.getScriptProperties().getProperty('ENFORCE_UNIQUE_JERSEY') === 'true';
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const jerseyValues = sheet.getRange(2, 5, lastRow - 1, 1).getValues().flat();
      if (enforceUnique && jerseyValues.some(v => String(v) === String(jerseyNumber))) {
        return { success: false, field: 'jerseyNumber', message: 'This jersey number is already registered. Please choose another number.' };
      }
    }

    const registrationId = createRegistrationId_();
    photoFile = savePhoto_(formData.photo, registrationId, formData.playerName);
    const safeName = safeSheetText_(formData.playerName);
    const safeProject = safeSheetText_(formData.projectName);
    const row = [
      new Date(), registrationId, safeName, formData.position,
      jerseyNumber, formData.jerseySize, safeProject,
      photoFile.getId(), 'Pending Review'
    ];
    sheet.appendRow(row);
    appendedRow = sheet.getLastRow();
    SpreadsheetApp.flush();

    return { success: true, registrationId: registrationId, message: 'Registration received.' };
  } catch (err) {
    // Best-effort rollback if one side of the transaction failed.
    if (appendedRow && sheet) {
      try { sheet.deleteRow(appendedRow); } catch (ignored) {}
    }
    if (photoFile) {
      try { photoFile.setTrashed(true); } catch (ignored) {}
    }
    console.error(err && err.stack ? err.stack : err);
    return { success: false, message: publicErrorMessage_(err) };
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function validateRegistration_(data) {
  if (!data || typeof data !== 'object') throw new Error('Registration data was missing. Please refresh and try again.');
  const name = cleanText_(data.playerName, 100);
  const project = cleanText_(data.projectName, 100);
  if (!name) throw fieldError_('playerName', 'Please enter your full name.');
  if (!project) throw fieldError_('projectName', 'Please enter your project name.');
  const positions = ['Goalkeeper','Full Back','Center Back','Defensive Midfielder','Attacking Midfielder','Center Midfielder','Right Winger','Left Winger','Center Forward','Finisher'];
  const sizes = ['S','M','L','XL','XXL','XXXL'];
  if (positions.indexOf(String(data.position)) === -1) throw fieldError_('position', 'Please select a valid playing position.');
  if (sizes.indexOf(String(data.jerseySize)) === -1) throw fieldError_('jerseySize', 'Please select a valid jersey size.');
  const jersey = Number(data.jerseyNumber);
  if (!Number.isInteger(jersey) || jersey < 0 || jersey > 99) throw fieldError_('jerseyNumber', 'Jersey number must be an integer from 0 to 99.');
  if (data.privacyConsent !== true) throw fieldError_('privacyConsent', 'Please confirm the privacy notice.');
  if (!data.photo || typeof data.photo.dataUrl !== 'string' || typeof data.photo.mimeType !== 'string') {
    throw fieldError_('playerPhoto', 'Please upload a player photo.');
  }
  if (APP_CONFIG.ALLOWED_MIME_TYPES.indexOf(data.photo.mimeType) === -1) {
    throw fieldError_('playerPhoto', 'Only JPG, PNG, and WebP images are allowed.');
  }
  const match = data.photo.dataUrl.match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=\s]+)$/);
  if (!match || match[1] !== data.photo.mimeType) throw fieldError_('playerPhoto', 'The photo data is invalid. Please select it again.');
  const bytes = Utilities.base64Decode(match[2]);
  if (!bytes.length || bytes.length > APP_CONFIG.MAX_PHOTO_BYTES) throw fieldError_('playerPhoto', 'Photo must be smaller than or equal to 10 MB.');
  // Basic file-signature validation avoids trusting only the browser MIME type.
  if (!hasValidImageSignature_(bytes, data.photo.mimeType)) {
    throw fieldError_('playerPhoto', 'The uploaded file does not appear to be a valid image.');
  }
  data.playerName = name;
  data.projectName = project;
}

function hasValidImageSignature_(bytes, mimeType) {
  const b = bytes.map(v => v & 255);
  if (mimeType === 'image/jpeg') return b.length > 3 && b[0] === 0xFF && b[1] === 0xD8 && b[2] === 0xFF;
  if (mimeType === 'image/png') return b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47 && b[4] === 0x0D && b[5] === 0x0A && b[6] === 0x1A && b[7] === 0x0A;
  if (mimeType === 'image/webp') return b.length > 12 && String.fromCharCode.apply(null, b.slice(0, 4)) === 'RIFF' && String.fromCharCode.apply(null, b.slice(8, 12)) === 'WEBP';
  return false;
}

function savePhoto_(photo, registrationId, playerName) {
  const match = photo.dataUrl.match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=\s]+)$/);
  const bytes = Utilities.base64Decode(match[2]);
  const extension = photo.mimeType === 'image/jpeg' ? 'jpg' : (photo.mimeType === 'image/png' ? 'png' : 'webp');
  const safeName = String(playerName).replace(/[^a-zA-Z0-9_-]+/g, '-').replace(/^-+|-+$/g, '').substring(0, 45) || 'player';
  const blob = Utilities.newBlob(bytes, photo.mimeType, registrationId + '_' + safeName + '.' + extension);
  const folder = getOrCreatePhotoFolder_();
  return folder.createFile(blob);
}

function getOrCreatePhotoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('PHOTO_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (ignored) {}
  }
  const folders = DriveApp.getFoldersByName(APP_CONFIG.PHOTO_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(APP_CONFIG.PHOTO_FOLDER_NAME);
  props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
}

function getRegistrationSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(APP_CONFIG.SPREADSHEET_PROPERTY);
  if (!id) throw new Error('Spreadsheet is not configured. Ask the organizer to complete setup.');
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName(APP_CONFIG.SHEET_NAME);
  if (!sheet || sheet.getLastRow() === 0) throw new Error('Registration sheet is not initialized. Ask the organizer to run setupTournament().');
  return sheet;
}

function createRegistrationId_() {
  return 'BBFT-2026-' + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
}
function cleanText_(value, maxLength) {
  if (typeof value !== 'string') return '';
  return value.replace(/[\u0000-\u001F\u007F]/g, '').trim().substring(0, maxLength);
}
function safeSheetText_(value) {
  const s = String(value || '').trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function fieldError_(field, message) {
  const err = new Error(message);
  err.field = field;
  return err;
}
function publicErrorMessage_(err) {
  if (err && err.field) return err.message;
  return 'We could not save your registration. Please try again. If the problem continues, contact the tournament organizer.';
}

/* Organizer-only functions. Configure ORGANIZER_EMAILS as comma-separated addresses.
 * For deployments where Session.getActiveUser().getEmail() is blank, this intentionally
 * fails closed. Use an organization-controlled deployment where identity is available.
 */
function requireOrganizer_() {
  const allowed = (PropertiesService.getScriptProperties().getProperty(APP_CONFIG.ORGANIZERS_PROPERTY) || '')
    .split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
  const email = (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email || allowed.indexOf(email) === -1) {
    throw new Error('Organizer access denied. Sign in with an authorized Google account in a supported deployment.');
  }
  return email;
}

function getOrganizerDashboardData() {
  requireOrganizer_();
  const sheet = getRegistrationSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, total: 0, pending: 0, approved: 0, rejected: 0, players: [] };
  const values = sheet.getRange(2, 1, lastRow - 1, APP_CONFIG.HEADERS.length).getValues();
  const players = values.map((r, i) => ({
    timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    registrationId: String(r[1]), playerName: String(r[2]), position: String(r[3]),
    jerseyNumber: String(r[4]), jerseySize: String(r[5]), projectName: String(r[6]),
    photoFileId: String(r[7]), status: String(r[8]), rowNumber: i + 2
  }));
  return {
    success: true, total: players.length,
    pending: players.filter(p => p.status === 'Pending Review').length,
    approved: players.filter(p => p.status === 'Approved').length,
    rejected: players.filter(p => p.status === 'Rejected').length,
    players: players
  };
}

function updateRegistrationStatus(registrationId, status) {
  requireOrganizer_();
  const allowedStatuses = ['Pending Review', 'Approved', 'Rejected'];
  if (allowedStatuses.indexOf(status) === -1) throw new Error('Invalid registration status.');
  const sheet = getRegistrationSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No registrations found.');
  const ids = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues().flat();
  const index = ids.indexOf(String(registrationId));
  if (index === -1) throw new Error('Registration was not found.');
  sheet.getRange(index + 2, 9).setValue(status);
  return { success: true, status: status };
}

/** Returns a short-lived data URL only to an authenticated organizer. */
function getPrivatePlayerPhoto(fileId) {
  requireOrganizer_();
  const folder = getOrCreatePhotoFolder_();
  const file = DriveApp.getFileById(String(fileId));
  // Verify that the requested file belongs to the configured photo folder.
  const parents = file.getParents();
  let belongs = false;
  while (parents.hasNext()) if (parents.next().getId() === folder.getId()) belongs = true;
  if (!belongs) throw new Error('Photo access denied.');
  const blob = file.getBlob();
  if (blob.getBytes().length > 2 * 1024 * 1024) {
    throw new Error('This photo is too large to preview in the dashboard. Open it directly in Drive with organizer permissions.');
  }
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

/** Returns CSV text to the authorized organizer; the browser can download it locally. */
function exportRegistrationsCsv() {
  requireOrganizer_();
  const sheet = getRegistrationSheet_();
  const data = sheet.getDataRange().getDisplayValues();
  return data.map(row => row.map(cell => '"' + String(cell).replace(/"/g, '""') + '"').join(',')).join('\r\n');
}
