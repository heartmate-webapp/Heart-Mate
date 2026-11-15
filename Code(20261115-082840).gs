/**
 * ============================================================================
 * HEART MATE WEB APP — GOOGLE APPS SCRIPT API
 * ============================================================================
 * Version: 2.0.0
 * Runtime: Google Apps Script V8
 * Database: Google Sheets
 * File storage: Google Drive
 * Music search: YouTube Data API v3
 * Frontend: GitHub Pages (index.html / dashboard.html)
 *
 * IMPORTANT SETUP
 * 1) Paste this entire file into Code.gs in a Google Apps Script project.
 * 2) Run setupSystem() once and grant the requested permissions.
 *    - If this is a spreadsheet-bound script, it uses that spreadsheet.
 *    - If this is a standalone script, it creates the database spreadsheet.
 * 3) Enable YouTube Data API v3 in the Google Cloud project linked to this
 *    Apps Script project, then run saveYouTubeApiKey('YOUR_API_KEY') once.
 * 4) Deploy > New deployment > Web app:
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5) Use the deployed /exec URL in the frontend as API_URL.
 *
 * FRONTEND TRANSPORT
 * GitHub Pages and Apps Script are different origins. This API intentionally
 * supports JSONP through `callback=` so the frontend can call it without a
 * browser CORS preflight. The index/dashboard integration should use the
 * JSONP helper described in the API CONTRACT near the bottom of this file.
 * For file uploads, send base64 image data in small JSONP chunks.
 *
 * SECURITY NOTES
 * - Admin password and YouTube API key are stored in Script Properties.
 * - Admin actions require a short-lived session token.
 * - Login uses a one-time challenge proof; the raw password is not sent in the
 *   request URL.
 * - Public gift URLs are unguessable IDs, not recipient accounts/passwords.
 * - Anyone with a published gift URL can view that gift. Do not store private
 *   or highly sensitive information in a public gift.
 * - JSONP is suitable for this small personal project. Admin API calls still
 *   require a token and the frontend should keep that token in memory only.
 * - Public creation is intentionally open. Nonce, honeypot and throttles are
 *   best-effort abuse controls; use CAPTCHA / edge protection for high traffic.
 * ============================================================================
 */

/* -------------------------------------------------------------------------- */
/* 01. CONFIGURATION                                                          */
/* -------------------------------------------------------------------------- */

const APP = Object.freeze({
  NAME: 'Heart Mate',
  VERSION: '2.0.0',
  DEFAULT_ADMIN_PASSWORD: 'ASSD55290',
  SESSION_TTL_SECONDS: 21600, // 6 hours (CacheService maximum)
  LOGIN_CHALLENGE_TTL_SECONDS: 300,
  LOGIN_FAIL_TTL_SECONDS: 900,
  MAX_LOGIN_FAILURES: 10,
  YOUTUBE_CACHE_TTL_SECONDS: 1800,
  YOUTUBE_MAX_RESULTS: 10,
  MAX_IMAGE_BYTES: 1000000, // Compress images in the browser before upload.
  IMAGE_CHUNK_SIZE: 4000,  // Base64 characters per JSONP request; keep in sync with index.html.
  MAX_IMAGE_CHUNKS: 340,   // Enough for a 1 MB image encoded as Base64 at 4,000 chars/chunk.
  MAX_TEXT_LENGTH: 5000,
  SHEETS: Object.freeze({
    GIFTS: 'Gifts',
    CARDS: 'Cards',
    SONGS: 'Songs',
    PHOTOS: 'Photos',
    MEMORIES: 'Memories',
    SETTINGS: 'Settings'
  }),
  HEADERS: Object.freeze({
    Gifts: [
      'gift_id', 'gift_type', 'title', 'recipient_name', 'sender_name',
      'status', 'theme', 'message', 'payload_json', 'created_at',
      'updated_at', 'view_count', 'last_viewed_at'
    ],
    Cards: [
      'gift_id', 'design_id', 'greeting', 'message', 'closing',
      'created_at', 'updated_at'
    ],
    Songs: [
      'gift_id', 'youtube_video_id', 'video_title', 'channel_title',
      'thumbnail_url', 'watch_url', 'embed_url', 'created_at', 'updated_at'
    ],
    Photos: [
      'photo_id', 'gift_id', 'drive_file_id', 'image_url',
      'thumbnail_url', 'caption', 'photo_type', 'created_at'
    ],
    Memories: [
      'memory_id', 'gift_id', 'memory_title', 'memory_text',
      'memory_date', 'photo_url', 'created_at'
    ],
    Settings: ['setting_key', 'setting_value', 'updated_at']
  })
});

const PROP = Object.freeze({
  SPREADSHEET_ID: 'GIFT_MEMORY_SPREADSHEET_ID',
  ASSETS_FOLDER_ID: 'GIFT_MEMORY_ASSETS_FOLDER_ID',
  TEMP_FOLDER_ID: 'GIFT_MEMORY_TEMP_FOLDER_ID',
  ADMIN_PASSWORD: 'GIFT_MEMORY_ADMIN_PASSWORD',
  LOGIN_SALT: 'GIFT_MEMORY_LOGIN_SALT',
  YOUTUBE_API_KEY: 'GIFT_MEMORY_YOUTUBE_API_KEY',
  SITE_URL: 'GIFT_MEMORY_SITE_URL'
});

/* -------------------------------------------------------------------------- */
/* 02. FIRST-TIME SETUP                                                       */
/* -------------------------------------------------------------------------- */

/**
 * Run this function manually once from the Apps Script editor.
 * Creates the database tabs and their headers automatically, plus Drive
 * folders for image files and temporary upload chunks.
 */
function setupSystem() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    const savedId = props.getProperty(PROP.SPREADSHEET_ID);
    if (savedId) {
      spreadsheet = SpreadsheetApp.openById(savedId);
    } else {
      spreadsheet = SpreadsheetApp.create('Heart Mate — Database');
    }
  }

  // Reuse the default empty tab as Gifts so a new spreadsheet has no stray Sheet1.
  const defaultSheet = spreadsheet.getSheetByName('Sheet1');
  if (defaultSheet && spreadsheet.getSheets().length === 1 && defaultSheet.getLastRow() === 0) {
    defaultSheet.setName(APP.SHEETS.GIFTS);
  }

  props.setProperty(PROP.SPREADSHEET_ID, spreadsheet.getId());

  Object.keys(APP.HEADERS).forEach(function (sheetName) {
    ensureSheet_(spreadsheet, sheetName, APP.HEADERS[sheetName]);
  });

  ensureDriveFolder_(PROP.ASSETS_FOLDER_ID, 'HeartMate_Assets');
  ensureDriveFolder_(PROP.TEMP_FOLDER_ID, 'HeartMate_Temporary_Uploads');

  if (!props.getProperty(PROP.ADMIN_PASSWORD)) {
    props.setProperty(PROP.ADMIN_PASSWORD, APP.DEFAULT_ADMIN_PASSWORD);
  }
  if (!props.getProperty(PROP.LOGIN_SALT)) {
    props.setProperty(PROP.LOGIN_SALT, Utilities.getUuid());
  }

  // Keep the site URL in both Settings and Script Properties once configured.
  const storedSiteUrl = props.getProperty(PROP.SITE_URL);
  if (storedSiteUrl) {
    upsertSetting_('site_url', storedSiteUrl);
  }

  const result = {
    ok: true,
    message: 'ตั้งค่าระบบและสร้างหัวตารางเรียบร้อยแล้ว',
    spreadsheet_name: spreadsheet.getName(),
    spreadsheet_url: spreadsheet.getUrl(),
    sheets: Object.keys(APP.HEADERS),
    assets_folder_ready: true,
    temporary_folder_ready: true,
    admin_password_initialized: true
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Save the YouTube Data API v3 key in Script Properties.
 * Run manually from the editor, for example:
 * saveYouTubeApiKey('AIza...');
 * Do not put the key in index.html, dashboard.html, GitHub, or a Sheet cell.
 */
function saveYouTubeApiKey(apiKey) {
  const key = String(apiKey || '').trim();
  if (key.length < 20) {
    throw new Error('กรุณาระบุ YouTube Data API Key ที่ถูกต้อง');
  }
  PropertiesService.getScriptProperties().setProperty(PROP.YOUTUBE_API_KEY, key);
  return 'บันทึก YouTube API Key ลง Script Properties แล้ว';
}

/**
 * Optional: set the GitHub Pages base URL used to generate full gift links.
 * Example: saveSiteUrl('https://username.github.io/repository-name');
 */
function saveSiteUrl(siteUrl) {
  const normalized = normalizeSiteUrl_(siteUrl);
  PropertiesService.getScriptProperties().setProperty(PROP.SITE_URL, normalized);
  upsertSetting_('site_url', normalized);
  return 'บันทึก URL เว็บไซต์แล้ว: ' + normalized;
}

/** Returns the database spreadsheet URL for setup/maintenance. */
function getDatabaseUrl() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP.SPREADSHEET_ID);
  if (!id) throw new Error('ยังไม่ได้ตั้งค่าระบบ กรุณารัน setupSystem() ก่อน');
  return SpreadsheetApp.openById(id).getUrl();
}

function ensureSheet_(spreadsheet, sheetName, requiredHeaders) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const existing = sheet.getLastRow() > 0
    ? sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    : [];
  const existingHeaders = existing.map(function (value) { return String(value).trim(); });
  const knownHeaders = existingHeaders.filter(function (value) { return value !== ''; });
  const lastHeaderColumn = existingHeaders.reduce(function (last, value, index) {
    return value ? index + 1 : last;
  }, 0);

  if (knownHeaders.length === 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
  } else {
    const missing = requiredHeaders.filter(function (header) {
      return knownHeaders.indexOf(header) === -1;
    });
    if (missing.length) {
      sheet.getRange(1, lastHeaderColumn + 1, 1, missing.length).setValues([missing]);
    }
  }

  const finalColumns = sheet.getLastColumn();
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 30);
  sheet.getRange(1, 1, 1, finalColumns)
    .setBackground('#FCE4EC')
    .setFontColor('#6D4355')
    .setFontWeight('bold')
    .setWrap(true);
  if (!sheet.getFilter() && sheet.getLastRow() > 1) {
    sheet.getRange(1, 1, sheet.getLastRow(), finalColumns).createFilter();
  }
  sheet.autoResizeColumns(1, Math.min(finalColumns, 12));
  return sheet;
}

function ensureDriveFolder_(propertyKey, folderName) {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(propertyKey);
  if (existingId) {
    try {
      DriveApp.getFolderById(existingId).getName();
      return existingId;
    } catch (error) {
      // The folder was removed or access changed; create a replacement.
    }
  }
  const folder = DriveApp.createFolder(folderName);
  props.setProperty(propertyKey, folder.getId());
  return folder.getId();
}

/* -------------------------------------------------------------------------- */
/* 03. WEB APP ENTRY POINT AND JSONP RESPONSE                                 */
/* -------------------------------------------------------------------------- */

/** Main browser API entry point. All frontend requests use GET + JSONP. */
function doGet(event) {
  const params = event && event.parameter ? event.parameter : {};
  const action = String(params.action || 'ping').trim();
  let callback = '';

  try {
    callback = normalizeCallback_(params.callback || '');
    const result = routeAction_(action, params);
    return makeResponse_(result, callback);
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    const message = error && error.message ? error.message : 'เกิดข้อผิดพลาดภายในระบบ';
    return makeResponse_({ ok: false, error: message }, callback);
  }
}

function routeAction_(action, params) {
  switch (action) {
    case 'ping':
      return { ok: true, app: APP.NAME, version: APP.VERSION, timestamp: nowIso_() };

    case 'getPublicConfig':
      return getPublicConfig_();

    case 'getPublicNonce':
      return getPublicNonce_();

    case 'createPublicGift':
      return createPublicGift_(params);

    case 'getLoginChallenge':
      return getLoginChallenge_();

    case 'login':
      return login_(params);

    case 'logout':
      return logout_(params.token);

    case 'listGifts':
      return listGifts_(params);

    case 'getGift':
      return getGift_(params);

    case 'createGift':
      return createGift_(params);

    case 'updateGift':
      return updateGift_(params);

    case 'deleteGift':
      return deleteGift_(params);

    case 'searchYouTube':
      return searchYouTube_(params);

    case 'saveSettings':
      return saveSettings_(params);

    case 'startImageUpload':
      return startImageUpload_(params);

    case 'uploadImageChunk':
      return uploadImageChunk_(params);

    case 'finishImageUpload':
      return finishImageUpload_(params);

    case 'deletePhoto':
      return deletePhoto_(params);

    default:
      throw new Error('ไม่รู้จัก API action: ' + action);
  }
}

function normalizeCallback_(callback) {
  const value = String(callback || '').trim();
  // Allow a simple dotted callback path, e.g. GiftApi.callbacks.cb123.
  if (!value) return '';
  const valid = /^[A-Za-z_$][A-Za-z0-9_$]*(\.[A-Za-z_$][A-Za-z0-9_$]*){0,3}$/;
  if (!valid.test(value)) throw new Error('ชื่อ JSONP callback ไม่ถูกต้อง');
  return value;
}

function makeResponse_(data, callback) {
  let body = JSON.stringify(data);
  body = body.replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');

  if (callback) {
    return ContentService.createTextOutput('/**/' + callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function parsePayload_(params) {
  let raw = '';
  try {
    if (params.payload_gzip64) {
      // Preferred transport for the public creator: gzip UTF-8 JSON, then
      // Base64URL it so Thai characters do not expand inside the query string.
      let encoded = String(params.payload_gzip64).replace(/-/g, '+').replace(/_/g, '/');
      while (encoded.length % 4) encoded += '=';
      const compressedBytes = Utilities.base64Decode(encoded);
      raw = Utilities.ungzip(Utilities.newBlob(compressedBytes, 'application/gzip', 'payload.json.gz'))
        .getDataAsString('UTF-8');
    } else if (params.payload64) {
      // Uncompressed Base64URL fallback for older browsers.
      let encoded = String(params.payload64).replace(/-/g, '+').replace(/_/g, '/');
      while (encoded.length % 4) encoded += '=';
      const bytes = Utilities.base64Decode(encoded);
      raw = Utilities.newBlob(bytes).getDataAsString('UTF-8');
    } else if (params.payload) {
      raw = String(params.payload);
    } else {
      return {};
    }
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      throw new Error('payload ต้องเป็น JSON object');
    }
    return parsed;
  } catch (error) {
    throw new Error('ข้อมูล payload ไม่ถูกต้อง: ' + error.message);
  }
}

/* -------------------------------------------------------------------------- */
/* 04. AUTHENTICATION AND SESSION MANAGEMENT                                  */
/* -------------------------------------------------------------------------- */

function getLoginChallenge_() {
  const cache = CacheService.getScriptCache();
  const challengeId = Utilities.getUuid().replace(/-/g, '');
  const nonce = Utilities.getUuid() + Utilities.getUuid();
  const salt = PropertiesService.getScriptProperties().getProperty(PROP.LOGIN_SALT);
  if (!salt) throw new Error('ยังไม่ได้ตั้งค่าระบบ กรุณารัน setupSystem()');

  cache.put('LOGIN_CHALLENGE_' + challengeId, nonce, APP.LOGIN_CHALLENGE_TTL_SECONDS);
  return {
    ok: true,
    challenge_id: challengeId,
    nonce: nonce,
    salt: salt,
    algorithm: 'SHA-256(password + ":" + salt + ":" + nonce)',
    expires_in: APP.LOGIN_CHALLENGE_TTL_SECONDS
  };
}

/**
 * The frontend calculates SHA-256(password + ':' + salt + ':' + nonce) using
 * Web Crypto and sends only that proof. Each challenge can be used once.
 */
function login_(params) {
  const cache = CacheService.getScriptCache();
  const failuresKey = 'LOGIN_FAILURE_COUNT';
  const currentFailures = Number(cache.get(failuresKey) || 0);
  if (currentFailures >= APP.MAX_LOGIN_FAILURES) {
    throw new Error('พยายามเข้าสู่ระบบมากเกินไป กรุณารอประมาณ 15 นาที');
  }

  const challengeId = String(params.challenge_id || '').trim();
  const suppliedProof = String(params.proof || '').trim().toLowerCase();
  if (!challengeId || !/^[a-f0-9]{32}$/.test(challengeId)) {
    throw new Error('ข้อมูลยืนยันการเข้าสู่ระบบไม่ถูกต้อง');
  }

  const cacheKey = 'LOGIN_CHALLENGE_' + challengeId;
  const nonce = cache.get(cacheKey);
  cache.remove(cacheKey); // One-time challenge: prevent replay.
  if (!nonce) return loginFailed_('คำขอเข้าสู่ระบบหมดอายุ กรุณาลองใหม่');

  const props = PropertiesService.getScriptProperties();
  const password = props.getProperty(PROP.ADMIN_PASSWORD);
  const salt = props.getProperty(PROP.LOGIN_SALT);
  const expectedProof = sha256Hex_(password + ':' + salt + ':' + nonce);

  if (!constantTimeEquals_(expectedProof, suppliedProof)) {
    return loginFailed_('รหัสเข้าใช้งานไม่ถูกต้อง');
  }

  cache.remove(failuresKey);
  const token = Utilities.getUuid() + Utilities.getUuid();
  const sessionKey = 'ADMIN_SESSION_' + sha256Hex_(token);
  cache.put(sessionKey, JSON.stringify({ created_at: nowIso_(), role: 'admin' }), APP.SESSION_TTL_SECONDS);

  return {
    ok: true,
    token: token,
    token_type: 'Bearer',
    expires_in: APP.SESSION_TTL_SECONDS,
    message: 'เข้าสู่ระบบสำเร็จ'
  };
}

function loginFailed_(message) {
  const cache = CacheService.getScriptCache();
  const failuresKey = 'LOGIN_FAILURE_COUNT';
  const failures = Number(cache.get(failuresKey) || 0) + 1;
  cache.put(failuresKey, String(failures), APP.LOGIN_FAIL_TTL_SECONDS);
  throw new Error(message);
}

function requireAdmin_(token) {
  const value = String(token || '').trim();
  if (value.length < 20) throw new Error('กรุณาเข้าสู่ระบบก่อนใช้งาน');

  const key = 'ADMIN_SESSION_' + sha256Hex_(value);
  const session = CacheService.getScriptCache().get(key);
  if (!session) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  return JSON.parse(session);
}

function logout_(token) {
  const value = String(token || '').trim();
  if (value) CacheService.getScriptCache().remove('ADMIN_SESSION_' + sha256Hex_(value));
  return { ok: true, message: 'ออกจากระบบแล้ว' };
}


/* -------------------------------------------------------------------------- */
/* 05. DATABASE HELPERS                                                       */
/* -------------------------------------------------------------------------- */

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP.SPREADSHEET_ID);
  if (!id) throw new Error('ยังไม่ได้ตั้งค่าฐานข้อมูล กรุณารัน setupSystem()');
  return SpreadsheetApp.openById(id);
}

function getSheet_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error('ไม่พบชีต ' + sheetName + ' กรุณารัน setupSystem()');
  return sheet;
}

function getHeaders_(sheet) {
  if (sheet.getLastColumn() < 1) throw new Error('ชีต ' + sheet.getName() + ' ยังไม่มีหัวตาราง');
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
    .map(function (value) { return String(value).trim(); });
}

function appendObject_(sheetName, record) {
  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);
  const row = headers.map(function (header) {
    return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
  });
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function readObjects_(sheetName) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];

  const headers = getHeaders_(sheet);
  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  return values.map(function (row) {
    const item = {};
    headers.forEach(function (header, index) {
      item[header] = row[index] instanceof Date ? row[index].toISOString() : row[index];
    });
    return item;
  });
}

function findRowByValue_(sheetName, columnName, value) {
  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);
  const columnIndex = headers.indexOf(columnName);
  if (columnIndex === -1 || sheet.getLastRow() < 2) return -1;

  const values = sheet.getRange(2, columnIndex + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  for (let index = 0; index < values.length; index += 1) {
    if (String(values[index][0]) === String(value)) return index + 2;
  }
  return -1;
}

function getRecordById_(sheetName, idColumn, idValue) {
  const rowNumber = findRowByValue_(sheetName, idColumn, idValue);
  if (rowNumber === -1) return null;
  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);
  const values = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const record = {};
  headers.forEach(function (header, index) {
    record[header] = values[index] instanceof Date ? values[index].toISOString() : values[index];
  });
  return { row_number: rowNumber, record: record };
}

function updateRecordById_(sheetName, idColumn, idValue, updates) {
  const found = getRecordById_(sheetName, idColumn, idValue);
  if (!found) throw new Error('ไม่พบข้อมูลที่ต้องการแก้ไข');

  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);
  Object.keys(updates).forEach(function (key) {
    const index = headers.indexOf(key);
    if (index !== -1) sheet.getRange(found.row_number, index + 1).setValue(updates[key]);
  });
  return found.row_number;
}

function deleteRowsByValue_(sheetName, columnName, value) {
  const sheet = getSheet_(sheetName);
  const rowNumbers = [];
  const columnIndex = getHeaders_(sheet).indexOf(columnName);
  if (columnIndex === -1 || sheet.getLastRow() < 2) return;
  const values = sheet.getRange(2, columnIndex + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  values.forEach(function (row, index) {
    if (String(row[0]) === String(value)) rowNumbers.push(index + 2);
  });
  rowNumbers.reverse().forEach(function (rowNumber) { sheet.deleteRow(rowNumber); });
}

function withScriptLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

/* -------------------------------------------------------------------------- */
/* 06. GIFT CRUD                                                              */
/* -------------------------------------------------------------------------- */

function listGifts_(params) {
  requireAdmin_(params.token);
  const typeFilter = String(params.type || '').trim();
  const statusFilter = String(params.status || '').trim();
  const keyword = String(params.q || '').trim().toLowerCase();

  let gifts = readObjects_(APP.SHEETS.GIFTS).map(giftFromRow_);
  if (typeFilter) gifts = gifts.filter(function (gift) { return gift.gift_type === typeFilter; });
  if (statusFilter) gifts = gifts.filter(function (gift) { return gift.status === statusFilter; });
  if (keyword) {
    gifts = gifts.filter(function (gift) {
      const text = [gift.title, gift.recipient_name, gift.sender_name, gift.gift_id]
        .join(' ').toLowerCase();
      return text.indexOf(keyword) !== -1;
    });
  }
  gifts.sort(function (a, b) { return String(b.created_at).localeCompare(String(a.created_at)); });
  return { ok: true, gifts: gifts, count: gifts.length };
}

function getGift_(params) {
  const giftId = sanitizeText_(params.id || params.gift_id, 80);
  if (!giftId) throw new Error('กรุณาระบุรหัสของขวัญ');

  const found = getRecordById_(APP.SHEETS.GIFTS, 'gift_id', giftId);
  if (!found) throw new Error('ไม่พบของขวัญนี้');
  const gift = giftFromRow_(found.record);

  if (gift.status !== 'published') {
    requireAdmin_(params.token);
  }

  const relatedPhotos = readObjects_(APP.SHEETS.PHOTOS).filter(function (photo) {
    return String(photo.gift_id) === giftId;
  });
  const relatedMemories = readObjects_(APP.SHEETS.MEMORIES).filter(function (memory) {
    return String(memory.gift_id) === giftId;
  });
  const card = getRecordById_(APP.SHEETS.CARDS, 'gift_id', giftId);
  const song = getRecordById_(APP.SHEETS.SONGS, 'gift_id', giftId);

  gift.photos = relatedPhotos;
  gift.memories = relatedMemories;
  gift.card = card ? card.record : null;
  gift.song = song ? song.record : null;

  // Public view counter. Failures to count a view should not block the gift.
  try {
    withScriptLock_(function () {
      const latest = getRecordById_(APP.SHEETS.GIFTS, 'gift_id', giftId);
      if (latest) {
        const nextCount = Number(latest.record.view_count || 0) + 1;
        updateRecordById_(APP.SHEETS.GIFTS, 'gift_id', giftId, {
          view_count: nextCount,
          last_viewed_at: nowIso_()
        });
        gift.view_count = nextCount;
      }
    });
  } catch (error) {
    console.warn('Unable to increment view counter: ' + error.message);
  }

  return { ok: true, gift: gift };
}

/**
 * Issues a short-lived, one-time nonce for the public gift creator.
 * This is an anti-CSRF / replay barrier, not a CAPTCHA.
 */
function getPublicNonce_() {
  const nonce = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').substring(0, 16);
  CacheService.getScriptCache().put('PUBLIC_CREATE_NONCE_' + nonce, '1', 600);
  return { ok: true, nonce: nonce, expires_in: 600, min_form_time_ms: 1200 };
}

/**
 * Public creation endpoint. Visitors can create and publish their own gift
 * without an admin account. Public gifts are link-accessible by design.
 * Best-effort abuse controls: one-time nonce, honeypot, elapsed-time check,
 * per-browser throttle, and a global daily ceiling.
 */
function createPublicGift_(params) {
  const nonce = String(params.nonce || '').trim();
  if (!/^[a-f0-9]{48}$/.test(nonce)) throw new Error('คำขอสร้างของขวัญไม่ถูกต้อง กรุณาลองใหม่');
  const nonceKey = 'PUBLIC_CREATE_NONCE_' + nonce;
  const nonceCache = CacheService.getScriptCache();
  if (!nonceCache.get(nonceKey)) throw new Error('คำขอสร้างของขวัญหมดอายุ กรุณาลองใหม่');
  nonceCache.remove(nonceKey);

  // Hidden fields should remain empty. They discourage simple automated spam.
  if (String(params.website || '').trim() || String(params.company || '').trim()) {
    throw new Error('ไม่สามารถส่งคำขอนี้ได้');
  }
  const startedAt = Number(params.form_started_at || 0);
  if (startedAt && (Date.now() - startedAt < 1200 || Date.now() - startedAt > 86400000)) {
    throw new Error('กรุณาตรวจสอบข้อมูลก่อนส่งอีกครั้ง');
  }

  const clientId = normalizePublicClientId_(params.client_id);
  incrementCacheLimit_('PUBLIC_CREATE_' + clientId, 4, 3600, 'คุณสร้างของขวัญบ่อยเกินไป กรุณารอสักครู่แล้วลองใหม่');
  incrementDailyLimit_('PUBLIC_GIFT_CREATE', 300, 'ขณะนี้มีการสร้างของขวัญจำนวนมาก กรุณาลองใหม่ภายหลัง');

  const input = parsePayload_(params);
  input.status = 'published'; // Publicly-created gifts are published immediately.
  const payload = normalizeGiftPayload_(input);

  return withScriptLock_(function () {
    const giftId = createUniqueGiftId_();
    const timestamp = nowIso_();
    const record = {
      gift_id: giftId,
      gift_type: payload.gift_type,
      title: payload.title,
      recipient_name: payload.recipient_name,
      sender_name: payload.sender_name,
      status: 'published',
      theme: payload.theme,
      message: payload.message,
      payload_json: JSON.stringify(payload.content),
      created_at: timestamp,
      updated_at: timestamp,
      view_count: 0,
      last_viewed_at: ''
    };
    appendObject_(APP.SHEETS.GIFTS, record);
    syncGiftDetailSheets_(giftId, payload, timestamp);

    // This one-time-returned secret lets the creator attach their own photos
    // to any gift type without exposing an editable public endpoint to link viewers.
    const ownerUploadToken = Utilities.getUuid() + Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty(
      'PUBLIC_OWNER_UPLOAD_' + giftId,
      JSON.stringify({ hash: sha256Hex_(ownerUploadToken), created_at_ms: Date.now() })
    );

    const publicUrl = buildGiftUrl_(giftId);
    return {
      ok: true,
      gift_id: giftId,
      owner_upload_token: ownerUploadToken,
      relative_url: '?id=' + encodeURIComponent(giftId),
      public_url: publicUrl,
      gift: Object.assign({}, record, { content: payload.content, public_url: publicUrl }),
      message: 'สร้างของขวัญและเผยแพร่เรียบร้อยแล้ว'
    };
  });
}

function normalizePublicClientId_(value) {
  const id = String(value || '').trim();
  if (!/^[A-Za-z0-9_-]{16,64}$/.test(id)) throw new Error('ไม่พบรหัสเซสชันของเบราว์เซอร์ กรุณารีเฟรชหน้าเว็บ');
  return id;
}

function incrementCacheLimit_(key, limit, ttlSeconds, errorMessage) {
  return withScriptLock_(function () {
    const cache = CacheService.getScriptCache();
    const count = Number(cache.get(key) || 0);
    if (count >= limit) throw new Error(errorMessage);
    cache.put(key, String(count + 1), ttlSeconds);
    return count + 1;
  });
}

function incrementDailyLimit_(prefix, limit, errorMessage) {
  return withScriptLock_(function () {
    const dateKey = Utilities.formatDate(new Date(), 'Etc/UTC', 'yyyyMMdd');
    const key = prefix + '_' + dateKey;
    const props = PropertiesService.getScriptProperties();
    const count = Number(props.getProperty(key) || 0);
    if (count >= limit) throw new Error(errorMessage);
    props.setProperty(key, String(count + 1));
    return count + 1;
  });
}

function createGift_(params) {
  requireAdmin_(params.token);
  const payload = normalizeGiftPayload_(parsePayload_(params));
  return withScriptLock_(function () {
    const giftId = createUniqueGiftId_();
    const timestamp = nowIso_();
    const record = {
      gift_id: giftId,
      gift_type: payload.gift_type,
      title: payload.title,
      recipient_name: payload.recipient_name,
      sender_name: payload.sender_name,
      status: payload.status,
      theme: payload.theme,
      message: payload.message,
      payload_json: JSON.stringify(payload.content),
      created_at: timestamp,
      updated_at: timestamp,
      view_count: 0,
      last_viewed_at: ''
    };
    appendObject_(APP.SHEETS.GIFTS, record);
    syncGiftDetailSheets_(giftId, payload, timestamp);

    const publicUrl = buildGiftUrl_(giftId);
    return {
      ok: true,
      gift_id: giftId,
      relative_url: '?id=' + encodeURIComponent(giftId),
      public_url: publicUrl,
      gift: Object.assign({}, record, { content: payload.content, public_url: publicUrl }),
      message: 'สร้างของขวัญเรียบร้อยแล้ว'
    };
  });
}

function updateGift_(params) {
  requireAdmin_(params.token);
  const giftId = sanitizeText_(params.id || params.gift_id, 80);
  if (!giftId) throw new Error('กรุณาระบุรหัสของขวัญ');
  const patch = parsePayload_(params);

  return withScriptLock_(function () {
    const found = getRecordById_(APP.SHEETS.GIFTS, 'gift_id', giftId);
    if (!found) throw new Error('ไม่พบของขวัญที่ต้องการแก้ไข');
    const oldGift = giftFromRow_(found.record);
    const merged = Object.assign({}, oldGift, patch);
    merged.content = Object.assign({}, oldGift.content || {}, patch.content || {});
    const payload = normalizeGiftPayload_(merged);
    const timestamp = nowIso_();

    updateRecordById_(APP.SHEETS.GIFTS, 'gift_id', giftId, {
      gift_type: payload.gift_type,
      title: payload.title,
      recipient_name: payload.recipient_name,
      sender_name: payload.sender_name,
      status: payload.status,
      theme: payload.theme,
      message: payload.message,
      payload_json: JSON.stringify(payload.content),
      updated_at: timestamp
    });
    syncGiftDetailSheets_(giftId, payload, timestamp);

    return {
      ok: true,
      gift_id: giftId,
      public_url: buildGiftUrl_(giftId),
      message: 'บันทึกการแก้ไขเรียบร้อยแล้ว'
    };
  });
}

function deleteGift_(params) {
  requireAdmin_(params.token);
  const giftId = sanitizeText_(params.id || params.gift_id, 80);
  if (!giftId) throw new Error('กรุณาระบุรหัสของขวัญ');

  return withScriptLock_(function () {
    const found = getRecordById_(APP.SHEETS.GIFTS, 'gift_id', giftId);
    if (!found) throw new Error('ไม่พบของขวัญที่ต้องการลบ');

    const photos = readObjects_(APP.SHEETS.PHOTOS).filter(function (photo) {
      return String(photo.gift_id) === giftId;
    });
    photos.forEach(function (photo) {
      if (photo.drive_file_id) {
        try { DriveApp.getFileById(String(photo.drive_file_id)).setTrashed(true); }
        catch (error) { console.warn('Unable to trash Drive file: ' + error.message); }
      }
    });

    [APP.SHEETS.CARDS, APP.SHEETS.SONGS, APP.SHEETS.PHOTOS, APP.SHEETS.MEMORIES]
      .forEach(function (sheetName) { deleteRowsByValue_(sheetName, 'gift_id', giftId); });
    deleteRowsByValue_(APP.SHEETS.GIFTS, 'gift_id', giftId);
    PropertiesService.getScriptProperties().deleteProperty('PUBLIC_OWNER_UPLOAD_' + giftId);
    return { ok: true, message: 'ลบของขวัญและข้อมูลที่เกี่ยวข้องแล้ว' };
  });
}

function normalizeGiftPayload_(input) {
  const allowedTypes = ['flower', 'card', 'song', 'photobooth', 'memory'];
  const giftType = String(input.gift_type || input.type || '').trim().toLowerCase();
  if (allowedTypes.indexOf(giftType) === -1) {
    throw new Error('ประเภทของขวัญไม่ถูกต้อง ต้องเป็น flower, card, song, photobooth หรือ memory');
  }

  const status = String(input.status || 'published').toLowerCase();
  if (['draft', 'published', 'archived'].indexOf(status) === -1) {
    throw new Error('สถานะของขวัญไม่ถูกต้อง');
  }

  const content = input.content && typeof input.content === 'object' && !Array.isArray(input.content)
    ? input.content
    : (input.data && typeof input.data === 'object' && !Array.isArray(input.data) ? input.data : {});

  if (giftType === 'song') {
    const videoId = String(content.youtube_video_id || content.video_id || '').trim();
    if (!/^[A-Za-z0-9_-]{11}$/.test(videoId)) {
      throw new Error('กรุณาเลือกเพลง YouTube ที่มี Video ID ถูกต้องก่อนบันทึก');
    }
    content.youtube_video_id = videoId;
  }

  return {
    gift_type: giftType,
    title: sanitizeText_(input.title || defaultGiftTitle_(giftType), 160),
    recipient_name: sanitizeText_(input.recipient_name, 120),
    sender_name: sanitizeText_(input.sender_name, 120),
    status: status,
    theme: sanitizeText_(input.theme || 'pastel', 80),
    message: sanitizeText_(input.message, APP.MAX_TEXT_LENGTH),
    content: content
  };
}

function giftFromRow_(row) {
  let content = {};
  try { content = row.payload_json ? JSON.parse(String(row.payload_json)) : {}; }
  catch (error) { content = {}; }

  return {
    gift_id: String(row.gift_id || ''),
    gift_type: String(row.gift_type || ''),
    title: String(row.title || ''),
    recipient_name: String(row.recipient_name || ''),
    sender_name: String(row.sender_name || ''),
    status: String(row.status || 'draft'),
    theme: String(row.theme || 'pastel'),
    message: String(row.message || ''),
    content: content,
    created_at: String(row.created_at || ''),
    updated_at: String(row.updated_at || ''),
    view_count: Number(row.view_count || 0),
    last_viewed_at: String(row.last_viewed_at || '')
  };
}

function syncGiftDetailSheets_(giftId, payload, timestamp) {
  const content = payload.content || {};
  if (payload.gift_type !== 'card') deleteRowsByValue_(APP.SHEETS.CARDS, 'gift_id', giftId);
  if (payload.gift_type !== 'song') deleteRowsByValue_(APP.SHEETS.SONGS, 'gift_id', giftId);
  if (payload.gift_type !== 'memory') deleteRowsByValue_(APP.SHEETS.MEMORIES, 'gift_id', giftId);

  if (payload.gift_type === 'card') {
    upsertByGiftId_(APP.SHEETS.CARDS, giftId, {
      gift_id: giftId,
      design_id: sanitizeText_(content.design_id || content.design || 'classic', 80),
      greeting: sanitizeText_(content.greeting, 300),
      message: sanitizeText_(content.message || payload.message, APP.MAX_TEXT_LENGTH),
      closing: sanitizeText_(content.closing, 300),
      created_at: timestamp,
      updated_at: timestamp
    });
  }

  if (payload.gift_type === 'song') {
    const videoId = String(content.youtube_video_id || '');
    upsertByGiftId_(APP.SHEETS.SONGS, giftId, {
      gift_id: giftId,
      youtube_video_id: videoId,
      video_title: sanitizeText_(content.video_title || content.title, 300),
      channel_title: sanitizeText_(content.channel_title, 200),
      thumbnail_url: safeHttpUrl_(content.thumbnail_url),
      watch_url: 'https://www.youtube.com/watch?v=' + videoId,
      embed_url: 'https://www.youtube.com/embed/' + videoId,
      created_at: timestamp,
      updated_at: timestamp
    });
  }

  if (payload.gift_type === 'memory' && Array.isArray(content.entries)) {
    deleteRowsByValue_(APP.SHEETS.MEMORIES, 'gift_id', giftId);
    content.entries.slice(0, 100).forEach(function (entry) {
      appendObject_(APP.SHEETS.MEMORIES, {
        memory_id: createId_('mem'),
        gift_id: giftId,
        memory_title: sanitizeText_(entry.title, 200),
        memory_text: sanitizeText_(entry.text, 2000),
        memory_date: sanitizeText_(entry.date, 40),
        photo_url: safeHttpUrl_(entry.photo_url),
        created_at: timestamp
      });
    });
  }
}

function upsertByGiftId_(sheetName, giftId, record) {
  const existing = getRecordById_(sheetName, 'gift_id', giftId);
  if (existing) {
    updateRecordById_(sheetName, 'gift_id', giftId, record);
  } else {
    appendObject_(sheetName, record);
  }
}

function createUniqueGiftId_() {
  let id;
  do { id = 'gift-' + Utilities.getUuid().replace(/-/g, '').substring(0, 12).toLowerCase(); }
  while (getRecordById_(APP.SHEETS.GIFTS, 'gift_id', id));
  return id;
}

function defaultGiftTitle_(type) {
  const titles = {
    flower: 'ช่อดอกไม้ของขวัญ',
    card: 'การ์ดพิเศษ',
    song: 'เพลงที่อยากส่งให้',
    photobooth: 'โฟโต้บูธความทรงจำ',
    memory: 'กล่องความทรงจำ'
  };
  return titles[type] || 'ของขวัญพิเศษ';
}

function buildGiftUrl_(giftId) {
  const base = PropertiesService.getScriptProperties().getProperty(PROP.SITE_URL);
  if (!base) return '?id=' + encodeURIComponent(giftId);
  return base.replace(/\/+$/, '') + '/?id=' + encodeURIComponent(giftId);
}

/* -------------------------------------------------------------------------- */
/* 07. YOUTUBE SEARCH                                                         */
/* -------------------------------------------------------------------------- */

function searchYouTube_(params) {
  if (params.token) {
    requireAdmin_(params.token);
  } else {
    const publicClientId = normalizePublicClientId_(params.client_id);
    incrementCacheLimit_('PUBLIC_YT_CLIENT_' + publicClientId, 30, 3600, 'ค้นหาเพลงบ่อยเกินไป กรุณารอสักครู่');
  }
  const query = sanitizeText_(params.q || params.query, 100);
  if (query.length < 2) throw new Error('กรุณาพิมพ์ชื่อเพลงหรือศิลปินอย่างน้อย 2 ตัวอักษร');

  const maxResults = Math.max(1, Math.min(Number(params.max_results || APP.YOUTUBE_MAX_RESULTS), APP.YOUTUBE_MAX_RESULTS));
  const cache = CacheService.getScriptCache();
  const cacheKey = 'YT_SEARCH_' + sha256Hex_(query.toLowerCase() + '|' + maxResults);
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  if (!params.token) {
    incrementDailyLimit_('PUBLIC_YOUTUBE_SEARCH', 85, 'ระบบค้นหาเพลงใช้งานครบโควตาประจำวันแล้ว กรุณาลองใหม่ในวันถัดไป');
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty(PROP.YOUTUBE_API_KEY);
  if (!apiKey) throw new Error('ยังไม่ได้ตั้งค่า YouTube API Key กรุณารัน saveYouTubeApiKey() ก่อน');

  const paramsMap = {
    part: 'snippet',
    q: query,
    type: 'video',
    videoEmbeddable: 'true',
    maxResults: String(maxResults),
    order: 'relevance',
    regionCode: 'TH',
    relevanceLanguage: 'th',
    safeSearch: 'moderate',
    key: apiKey
  };
  const queryString = Object.keys(paramsMap).map(function (key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(paramsMap[key]);
  }).join('&');
  const endpoint = 'https://www.googleapis.com/youtube/v3/search?' + queryString;
  const response = UrlFetchApp.fetch(endpoint, { muteHttpExceptions: true });
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();
  let data;
  try { data = JSON.parse(responseText); }
  catch (error) { throw new Error('YouTube ส่งข้อมูลกลับมาไม่ถูกต้อง'); }

  if (statusCode < 200 || statusCode >= 300 || data.error) {
    const apiMessage = data.error && data.error.message ? data.error.message : 'ไม่สามารถค้นหาเพลงได้';
    throw new Error('YouTube API: ' + apiMessage);
  }

  const results = (data.items || []).map(function (item) {
    const snippet = item.snippet || {};
    const thumbnails = snippet.thumbnails || {};
    const videoId = item.id && item.id.videoId ? item.id.videoId : '';
    return {
      video_id: videoId,
      title: snippet.title || '',
      channel_title: snippet.channelTitle || '',
      description: snippet.description || '',
      published_at: snippet.publishedAt || '',
      thumbnail_url: (thumbnails.high && thumbnails.high.url) ||
        (thumbnails.medium && thumbnails.medium.url) ||
        (thumbnails.default && thumbnails.default.url) || '',
      watch_url: videoId ? 'https://www.youtube.com/watch?v=' + videoId : '',
      embed_url: videoId ? 'https://www.youtube.com/embed/' + videoId : ''
    };
  }).filter(function (item) { return /^[A-Za-z0-9_-]{11}$/.test(item.video_id); });

  const result = { ok: true, query: query, results: results, count: results.length };
  cache.put(cacheKey, JSON.stringify(result), APP.YOUTUBE_CACHE_TTL_SECONDS);
  return result;
}

/* -------------------------------------------------------------------------- */
/* 08. GOOGLE DRIVE IMAGE UPLOADS                                             */
/* -------------------------------------------------------------------------- */

/**
 * Starts one photo upload. The client compresses the image to <= 1 MB, encodes
 * it as base64, splits the base64 string into 12,000-character chunks, then
 * calls uploadImageChunk sequentially before calling finishImageUpload.
 *
 * Upload is allowed for an authenticated admin, the creator's short-lived
 * upload-owner secret, or (for photobooth gifts only) a public gift link.
 */
function startImageUpload_(params) {
  const payload = parsePayload_(params);
  const giftId = sanitizeText_(payload.gift_id, 80);
  if (!giftId) throw new Error('กรุณาระบุ gift_id สำหรับรูปภาพ');

  const isAdmin = Boolean(params.token);
  if (isAdmin) requireAdmin_(params.token);
  const giftFound = getRecordById_(APP.SHEETS.GIFTS, 'gift_id', giftId);
  if (!giftFound) throw new Error('ไม่พบของขวัญที่ต้องการแนบรูป');
  if (!isAdmin && String(giftFound.record.status) !== 'published') {
    throw new Error('ของขวัญนี้ยังไม่เปิดให้เพิ่มรูปภาพ');
  }

  let isOwnerUpload = false;
  if (!isAdmin && payload.owner_upload_token) {
    const ownerProperty = PropertiesService.getScriptProperties().getProperty('PUBLIC_OWNER_UPLOAD_' + giftId);
    if (ownerProperty) {
      try {
        const ownerRecord = JSON.parse(ownerProperty);
        const stillValid = Date.now() - Number(ownerRecord.created_at_ms || 0) <= 7 * 86400000;
        isOwnerUpload = stillValid && constantTimeEquals_(ownerRecord.hash, sha256Hex_(payload.owner_upload_token));
      } catch (error) {
        isOwnerUpload = false;
      }
    }
  }

  // Link viewers may add community snapshots only to Photo Booth gifts.
  // The original creator's secret allows attaching images to any gift type.
  if (!isAdmin && !isOwnerUpload && String(giftFound.record.gift_type) !== 'photobooth') {
    throw new Error('การเพิ่มรูปภาพประเภทนี้ต้องใช้ลิงก์สิทธิ์ของผู้สร้างของขวัญ');
  }
  if (!isAdmin && !isOwnerUpload) {
    const publicUploadKey = 'PUBLIC_UPLOADS_' + giftId;
    const currentUploads = Number(CacheService.getScriptCache().get(publicUploadKey) || 0);
    if (currentUploads >= 20) throw new Error('ของขวัญนี้มีการอัปโหลดรูปครบจำนวนชั่วคราวแล้ว');
    CacheService.getScriptCache().put(publicUploadKey, String(currentUploads + 1), 3600);
  }

  const mimeType = String(payload.mime_type || '').toLowerCase();
  if (['image/jpeg', 'image/png', 'image/webp'].indexOf(mimeType) === -1) {
    throw new Error('รองรับเฉพาะไฟล์ JPG, PNG และ WebP');
  }

  const expectedBytes = Number(payload.expected_bytes || 0);
  const totalChunks = Number(payload.total_chunks || 0);
  if (!Number.isFinite(expectedBytes) || expectedBytes < 1 || expectedBytes > APP.MAX_IMAGE_BYTES) {
    throw new Error('รูปภาพต้องมีขนาดไม่เกิน 1 MB หลังบีบอัด');
  }
  if (!Number.isInteger(totalChunks) || totalChunks < 1 || totalChunks > APP.MAX_IMAGE_CHUNKS) {
    throw new Error('จำนวนชิ้นข้อมูลรูปภาพไม่ถูกต้อง');
  }

  const uploadId = Utilities.getUuid().replace(/-/g, '');
  const uploadToken = Utilities.getUuid() + Utilities.getUuid();
  const tempRootId = PropertiesService.getScriptProperties().getProperty(PROP.TEMP_FOLDER_ID);
  if (!tempRootId) throw new Error('ไม่พบโฟลเดอร์อัปโหลดชั่วคราว กรุณารัน setupSystem()');
  const tempFolder = DriveApp.getFolderById(tempRootId).createFolder(uploadId);

  const session = {
    upload_id: uploadId,
    upload_token_hash: sha256Hex_(uploadToken),
    folder_id: tempFolder.getId(),
    gift_id: giftId,
    file_name: sanitizeFileName_(payload.file_name || 'memory-photo.jpg'),
    mime_type: mimeType,
    expected_bytes: expectedBytes,
    total_chunks: totalChunks,
    caption: sanitizeText_(payload.caption, 300),
    photo_type: sanitizeText_(payload.photo_type || 'photobooth', 40),
    created_at_ms: Date.now()
  };
  PropertiesService.getScriptProperties().setProperty('UPLOAD_' + uploadId, JSON.stringify(session));

  return {
    ok: true,
    upload_id: uploadId,
    upload_token: uploadToken,
    chunk_size: APP.IMAGE_CHUNK_SIZE,
    total_chunks: totalChunks,
    max_bytes: APP.MAX_IMAGE_BYTES,
    message: 'เริ่มอัปโหลดรูปภาพแล้ว'
  };
}

function uploadImageChunk_(params) {
  const uploadId = String(params.upload_id || '').trim();
  const uploadToken = String(params.upload_token || '').trim();
  const index = Number(params.index);
  const chunk = String(params.chunk || '');
  const session = getUploadSession_(uploadId, uploadToken);

  if (!Number.isInteger(index) || index < 0 || index >= session.total_chunks) {
    throw new Error('หมายเลขชิ้นข้อมูลไม่ถูกต้อง');
  }
  if (!chunk || chunk.length > APP.IMAGE_CHUNK_SIZE || !/^[A-Za-z0-9+/=]+$/.test(chunk)) {
    throw new Error('ข้อมูลรูปภาพส่วนนี้ไม่ถูกต้อง');
  }

  const folder = DriveApp.getFolderById(session.folder_id);
  const partName = 'part_' + String(index).padStart(4, '0') + '.txt';
  const existing = folder.getFilesByName(partName);
  if (existing.hasNext()) {
    existing.next().setContent(chunk);
  } else {
    folder.createFile(partName, chunk, MimeType.PLAIN_TEXT);
  }
  return { ok: true, upload_id: uploadId, received_index: index };
}

function finishImageUpload_(params) {
  const uploadId = String(params.upload_id || '').trim();
  const uploadToken = String(params.upload_token || '').trim();
  const session = getUploadSession_(uploadId, uploadToken);
  const tempFolder = DriveApp.getFolderById(session.folder_id);

  let base64 = '';
  for (let index = 0; index < session.total_chunks; index += 1) {
    const partName = 'part_' + String(index).padStart(4, '0') + '.txt';
    const files = tempFolder.getFilesByName(partName);
    if (!files.hasNext()) throw new Error('อัปโหลดรูปภาพยังไม่ครบ ขาดชิ้นที่ ' + index);
    base64 += files.next().getBlob().getDataAsString();
  }

  // Verify before decoding/writing to Drive.
  const approximateBytes = Math.floor((base64.length * 3) / 4) - (base64.endsWith('==') ? 2 : (base64.endsWith('=') ? 1 : 0));
  if (approximateBytes < 1 || approximateBytes > APP.MAX_IMAGE_BYTES ||
      Math.abs(approximateBytes - session.expected_bytes) > 3) {
    throw new Error('ขนาดรูปภาพไม่ตรงกับข้อมูลเริ่มต้น กรุณาเริ่มอัปโหลดใหม่');
  }

  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, session.mime_type, session.file_name);
  const assetsFolderId = PropertiesService.getScriptProperties().getProperty(PROP.ASSETS_FOLDER_ID);
  if (!assetsFolderId) throw new Error('ไม่พบโฟลเดอร์เก็บรูปภาพ กรุณารัน setupSystem()');

  const file = DriveApp.getFolderById(assetsFolderId).createFile(blob);
  let sharingEnabled = true;
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (sharingError) {
    sharingEnabled = false;
    console.warn('Drive sharing was not enabled: ' + sharingError.message);
  }

  const photoId = createId_('photo');
  const imageUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();
  const thumbnailUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200';
  appendObject_(APP.SHEETS.PHOTOS, {
    photo_id: photoId,
    gift_id: session.gift_id,
    drive_file_id: file.getId(),
    image_url: imageUrl,
    thumbnail_url: thumbnailUrl,
    caption: session.caption,
    photo_type: session.photo_type,
    created_at: nowIso_()
  });

  cleanupUploadSession_(uploadId, session);
  return {
    ok: true,
    photo: {
      photo_id: photoId,
      gift_id: session.gift_id,
      drive_file_id: file.getId(),
      image_url: imageUrl,
      thumbnail_url: thumbnailUrl,
      caption: session.caption,
      photo_type: session.photo_type
    },
    sharing_enabled: sharingEnabled,
    message: sharingEnabled ? 'บันทึกรูปภาพเรียบร้อยแล้ว' : 'บันทึกรูปภาพแล้ว แต่ Drive ไม่อนุญาตให้เปิดสิทธิ์ลิงก์สาธารณะ'
  };
}

function getUploadSession_(uploadId, uploadToken) {
  if (!/^[a-f0-9]{32}$/.test(uploadId)) throw new Error('รหัสอัปโหลดไม่ถูกต้อง');
  const raw = PropertiesService.getScriptProperties().getProperty('UPLOAD_' + uploadId);
  if (!raw) throw new Error('ไม่พบเซสชันอัปโหลด หรือเซสชันหมดอายุแล้ว');
  const session = JSON.parse(raw);
  if (!constantTimeEquals_(session.upload_token_hash, sha256Hex_(uploadToken))) {
    throw new Error('ไม่มีสิทธิ์อัปโหลดรูปภาพนี้');
  }
  if (Date.now() - Number(session.created_at_ms) > 86400000) {
    cleanupUploadSession_(uploadId, session);
    throw new Error('เซสชันอัปโหลดหมดอายุ กรุณาเริ่มใหม่');
  }
  return session;
}

function cleanupUploadSession_(uploadId, session) {
  try { DriveApp.getFolderById(session.folder_id).setTrashed(true); }
  catch (error) { console.warn('Temporary upload folder cleanup failed: ' + error.message); }
  PropertiesService.getScriptProperties().deleteProperty('UPLOAD_' + uploadId);
}

/** Run manually or from a time-driven trigger to clean abandoned uploads. */
function cleanupExpiredUploads() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  let cleaned = 0;
  Object.keys(all).forEach(function (key) {
    if (key.indexOf('UPLOAD_') !== 0) return;
    try {
      const session = JSON.parse(all[key]);
      if (Date.now() - Number(session.created_at_ms) > 86400000) {
        cleanupUploadSession_(key.substring('UPLOAD_'.length), session);
        cleaned += 1;
      }
    } catch (error) {
      props.deleteProperty(key);
    }
  });
  return { ok: true, cleaned_uploads: cleaned };
}

function deletePhoto_(params) {
  requireAdmin_(params.token);
  const photoId = sanitizeText_(params.photo_id || params.id, 100);
  if (!photoId) throw new Error('กรุณาระบุ photo_id');
  const found = getRecordById_(APP.SHEETS.PHOTOS, 'photo_id', photoId);
  if (!found) throw new Error('ไม่พบรูปภาพ');

  if (found.record.drive_file_id) {
    try { DriveApp.getFileById(String(found.record.drive_file_id)).setTrashed(true); }
    catch (error) { console.warn('Unable to trash photo file: ' + error.message); }
  }
  getSheet_(APP.SHEETS.PHOTOS).deleteRow(found.row_number);
  return { ok: true, message: 'ลบรูปภาพแล้ว' };
}

/* -------------------------------------------------------------------------- */
/* 09. SETTINGS AND PUBLIC CONFIG                                             */
/* -------------------------------------------------------------------------- */

function getPublicConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    ok: true,
    app_name: APP.NAME,
    version: APP.VERSION,
    public_creation_enabled: true,
    public_gift_types: ['flower', 'card', 'song', 'photobooth', 'memory'],
    site_url: props.getProperty(PROP.SITE_URL) || '',
    max_image_bytes: APP.MAX_IMAGE_BYTES,
    image_chunk_size: APP.IMAGE_CHUNK_SIZE
  };
}

function saveSettings_(params) {
  requireAdmin_(params.token);
  const payload = parsePayload_(params);
  if (Object.prototype.hasOwnProperty.call(payload, 'site_url')) {
    const url = normalizeSiteUrl_(payload.site_url);
    PropertiesService.getScriptProperties().setProperty(PROP.SITE_URL, url);
    upsertSetting_('site_url', url);
  }
  return { ok: true, settings: { site_url: PropertiesService.getScriptProperties().getProperty(PROP.SITE_URL) || '' } };
}

function upsertSetting_(key, value) {
  const record = {
    setting_key: String(key),
    setting_value: String(value),
    updated_at: nowIso_()
  };
  const existing = getRecordById_(APP.SHEETS.SETTINGS, 'setting_key', key);
  if (existing) updateRecordById_(APP.SHEETS.SETTINGS, 'setting_key', key, record);
  else appendObject_(APP.SHEETS.SETTINGS, record);
}

function normalizeSiteUrl_(siteUrl) {
  const value = String(siteUrl || '').trim().replace(/\/+$/, '');
  if (!/^https:\/\/[A-Za-z0-9.-]+(?::\d+)?(?:\/[A-Za-z0-9._~!$&'()*+,;=:@%-]*)*$/.test(value)) {
    throw new Error('กรุณาระบุ URL เว็บไซต์แบบ https:// ที่ถูกต้อง');
  }
  return value;
}

/* -------------------------------------------------------------------------- */
/* 10. GENERAL UTILITIES                                                      */
/* -------------------------------------------------------------------------- */

function nowIso_() {
  return new Date().toISOString();
}

function sanitizeText_(value, maxLength) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
    .trim()
    .substring(0, maxLength || APP.MAX_TEXT_LENGTH);
}

function sanitizeFileName_(fileName) {
  const name = String(fileName || 'image.jpg').replace(/[\\/:*?"<>|\u0000-\u001F]/g, '_').trim();
  return (name || 'image.jpg').substring(0, 120);
}

function safeHttpUrl_(url) {
  const value = String(url || '').trim();
  return /^https:\/\//i.test(value) ? value.substring(0, 2000) : '';
}

function createId_(prefix) {
  return String(prefix || 'id') + '-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toLowerCase();
}

function sha256Hex_(value) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  );
  return digest.map(function (byte) {
    const unsigned = (byte + 256) % 256;
    return ('0' + unsigned.toString(16)).slice(-2);
  }).join('');
}

function constantTimeEquals_(a, b) {
  const left = String(a || '');
  const right = String(b || '');
  if (left.length !== right.length) return false;
  let result = 0;
  for (let index = 0; index < left.length; index += 1) {
    result |= left.charCodeAt(index) ^ right.charCodeAt(index);
  }
  return result === 0;
}

/* -------------------------------------------------------------------------- */
/* 11. API CONTRACT FOR index.html / dashboard.html                           */
/* -------------------------------------------------------------------------- */
/**
 * Base URL:
 *   https://script.google.com/macros/s/DEPLOYMENT_ID/exec
 *
 * JSONP request format:
 *   <API_URL>?action=ping&callback=GiftApi.callbacks.cb1
 *
 * Every frontend request should use a unique callback name and remove its
 * script element after the callback runs. Use encodeURIComponent() for every
 * query-string value. Do not fetch() this endpoint from GitHub Pages.
 *
 * Public actions:
 *   ping
 *   getPublicConfig
 *   getPublicNonce
 *   createPublicGift        params: nonce, client_id, form_started_at, payload JSON
 *   searchYouTube           params: q, client_id (cached / public quota-limited)
 *   getGift                 params: id
 *   startImageUpload        payload: {gift_id, file_name, mime_type,
 *                             expected_bytes, total_chunks, caption, photo_type}
 *   uploadImageChunk        params: upload_id, upload_token, index, chunk
 *   finishImageUpload       params: upload_id, upload_token
 *
 * Authentication:
 *   1) getLoginChallenge -> {challenge_id, nonce, salt}
 *   2) Browser computes SHA-256(password + ':' + salt + ':' + nonce) as lowercase hex
 *      using crypto.subtle.digest('SHA-256', new TextEncoder().encode(text)).
 *   3) login params: challenge_id, proof -> {token, expires_in}
 *   4) Keep token in memory and send token=... with admin requests.
 *
 * Admin actions:
 *   logout                  params: token
 *   listGifts               params: token, optional type/status/q
 *   createGift              params: token, payload JSON string
 *   updateGift              params: token, id, payload JSON string
 *   deleteGift              params: token, id
 *   searchYouTube           params: token, q, optional max_results (admin bypasses public cap)
 *   saveSettings            params: token, payload JSON string
 *   startImageUpload        params: token + payload (admin may upload to drafts)
 *   deletePhoto             params: token, photo_id
 *
 * createGift payload example:
 * {
 *   "gift_type": "song",
 *   "title": "เพลงที่อยากส่งให้",
 *   "recipient_name": "คนพิเศษ",
 *   "sender_name": "คนส่ง",
 *   "status": "published",
 *   "theme": "pastel-pink",
 *   "message": "เพลงนี้ทำให้นึกถึงเธอ",
 *   "content": {
 *     "youtube_video_id": "VIDEO_ID_11_CHARS",
 *     "video_title": "ชื่อเพลง",
 *     "channel_title": "ชื่อช่อง",
 *     "thumbnail_url": "https://..."
 *   }
 * }
 *
 * Supported gift_type values:
 *   flower | card | song | photobooth | memory
 *
 * Note: Images are compressed to JPG/PNG/WebP and <= 1 MB in the browser.
 * The client sends the base64 string in 4,000-character chunks (IMAGE_CHUNK_SIZE).
 */

/* -------------------------------------------------------------------------- */
/* 12. MAINTENANCE / OPTIONAL MANUAL TESTS                                    */
/* -------------------------------------------------------------------------- */

/** Run manually to verify the basic services after setup. */
function testSystemHealth() {
  const spreadsheet = getSpreadsheet_();
  const report = {
    ok: true,
    app: APP.NAME,
    version: APP.VERSION,
    spreadsheet_url: spreadsheet.getUrl(),
    sheets: Object.keys(APP.HEADERS).map(function (name) {
      const sheet = spreadsheet.getSheetByName(name);
      return { name: name, exists: Boolean(sheet), rows: sheet ? Math.max(0, sheet.getLastRow() - 1) : 0 };
    }),
    youtube_api_key_configured: Boolean(PropertiesService.getScriptProperties().getProperty(PROP.YOUTUBE_API_KEY)),
    public_creation_enabled: true,
    assets_folder_configured: Boolean(PropertiesService.getScriptProperties().getProperty(PROP.ASSETS_FOLDER_ID)),
    temp_folder_configured: Boolean(PropertiesService.getScriptProperties().getProperty(PROP.TEMP_FOLDER_ID))
  };
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}
